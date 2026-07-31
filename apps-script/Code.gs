/**
 * Backend (Google Apps Script) para o app "Entrada até Baixas".
 * Deve ser vinculado a uma Google Sheet dedicada (ver apps-script/README.md
 * para o passo a passo de implantação como Web App).
 *
 * Guarda dois conjuntos de lançamentos na mesma estrutura de colunas:
 *   - aba "ERP"    -> última importação do export do TOTVS Protheus
 *   - aba "Manual" -> lançamentos incluídos manualmente pela tela
 *
 * Cada lançamento recebe um Status calculado ("OK" ou "Duplicado – revisar")
 * comparando Código Fornecedor + Valor + Vencimento entre as duas abas.
 */

var HEADERS = [
  'ID', 'Tipo', 'Data Emissão', 'Nº Documento', 'Vencimento Real',
  'Forma Pagamento', 'Código Fornecedor', 'Parcela', 'Data Baixa',
  'Vencimento', 'R$ Valor', 'Usuário Inclusor', 'Razão Social',
  'Aprovação Remessa', 'Remessa', 'Histórico', 'Origem', 'Status'
];

var PROP_DATA_BASE = 'dataBaseImportacao';

// Pasta do Drive onde o export diário do Protheus é salvo:
// Compartilhados comigo / 4 - CE 007_ADMINISTRATIVO / 4.11 - FINANCEIRO / 17 CONTAS A PAGAR
// Nome do arquivo esperado: AAAA.MM.DD.xlsx (ex.: 2026.07.30.xlsx)
var DRIVE_FOLDER_ID = '1sVlF29VGWDzHelgBGIeFvVK3OCpMjGmD';
var NOME_ARQUIVO_RE = /^(\d{4})\.(\d{2})\.(\d{2})\.xlsx$/i;

var SRC = {
  tipo: 'B', dt_emissao: 'C', no_titulo: 'D', form_pagto: 'E', fornecedor: 'F',
  nome_fornece: 'G', vencimento: 'H', vencto_real: 'I', vlr_titulo: 'J', parcela: 'K',
  dt_baixa: 'L', historico: 'M', aprovacao: 'S', usuario: 'AJ', remessa: 'AW',
};

function getDataBase_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_DATA_BASE);
}
function setDataBase_(valor) {
  if (valor) PropertiesService.getScriptProperties().setProperty(PROP_DATA_BASE, valor);
}

function colIdx_(letra) {
  var idx = 0;
  for (var i = 0; i < letra.length; i++) idx = idx * 26 + (letra.charCodeAt(i) - 64);
  return idx - 1;
}

function removerZerosEsquerda_(v) {
  if (typeof v === 'string' && /^[0-9]+$/.test(v)) {
    var cortado = v.replace(/^0+/, '');
    return cortado || '0';
  }
  return v;
}

function dataParaISO_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

// Encontra, na pasta do Drive, o arquivo AAAA.MM.DD.xlsx com a data mais recente no nome.
function encontrarArquivoMaisRecenteNoDrive_() {
  var pasta = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var arquivos = pasta.getFiles();
  var melhor = null;
  var melhorData = null;
  while (arquivos.hasNext()) {
    var f = arquivos.next();
    var m = f.getName().match(NOME_ARQUIVO_RE);
    if (!m) continue;
    var dataStr = m[1] + '-' + m[2] + '-' + m[3];
    if (!melhorData || dataStr > melhorData) {
      melhor = f;
      melhorData = dataStr;
    }
  }
  return melhor ? { file: melhor, dataBase: melhorData } : null;
}

// Converte o .xlsx (via Drive API avançada) para Google Sheets temporário,
// lê todas as linhas/colunas e apaga o arquivo temporário em seguida.
function lerLinhasDoXlsx_(driveFile) {
  var recurso = { name: 'tmp_import_' + Utilities.getUuid(), mimeType: MimeType.GOOGLE_SHEETS };
  var convertido = Drive.Files.create(recurso, driveFile.getBlob(), { fields: 'id' });
  try {
    var ss = SpreadsheetApp.openById(convertido.id);
    return ss.getSheets()[0].getDataRange().getValues();
  } finally {
    Drive.Files.remove(convertido.id);
  }
}

// Espelha o mapeamento de colunas usado no import manual (web/index.html):
// linha 2 do Excel = cabeçalho, dados a partir da linha 3.
function parseLinhasProtheus_(dados) {
  var idx = {};
  Object.keys(SRC).forEach(function (k) { idx[k] = colIdx_(SRC[k]); });
  var itens = [];
  for (var r = 2; r < dados.length; r++) {
    var linha = dados[r];
    var tipo = linha[idx.tipo];
    if (tipo === '' || tipo === null || tipo === undefined) continue;
    itens.push({
      'Tipo': tipo,
      'Data Emissão': dataParaISO_(linha[idx.dt_emissao]),
      'Nº Documento': removerZerosEsquerda_(linha[idx.no_titulo]),
      'Vencimento Real': dataParaISO_(linha[idx.vencto_real]),
      'Forma Pagamento': linha[idx.form_pagto],
      'Código Fornecedor': removerZerosEsquerda_(linha[idx.fornecedor]),
      'Parcela': linha[idx.parcela],
      'Data Baixa': dataParaISO_(linha[idx.dt_baixa]),
      'Vencimento': dataParaISO_(linha[idx.vencimento]),
      'R$ Valor': linha[idx.vlr_titulo],
      'Usuário Inclusor': linha[idx.usuario],
      'Razão Social': linha[idx.nome_fornece],
      'Aprovação Remessa': linha[idx.aprovacao],
      'Remessa': removerZerosEsquerda_(linha[idx.remessa]),
      'Histórico': linha[idx.historico],
    });
  }
  return itens;
}

function getSheet_(nome) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function lerAba_(nome) {
  var sh = getSheet_(nome);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return [];
  var valores = sh.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  return valores.map(function (linha) {
    var obj = {};
    HEADERS.forEach(function (h, i) { obj[h] = linha[i]; });
    return obj;
  });
}

function chaveDuplicidade_(item) {
  var fornecedor = String(item['Código Fornecedor'] || '').trim();
  var valor = Number(item['R$ Valor'] || 0).toFixed(2);
  var venc = item['Vencimento'];
  var vencStr = venc instanceof Date
    ? Utilities.formatDate(venc, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(venc || '');
  return fornecedor + '|' + valor + '|' + vencStr;
}

function recalcularStatus_() {
  var erpSheet = getSheet_('ERP');
  var manualSheet = getSheet_('Manual');
  var erp = lerAba_('ERP');
  var manual = lerAba_('Manual');

  var chavesErp = {};
  erp.forEach(function (item) { chavesErp[chaveDuplicidade_(item)] = true; });
  var chavesManual = {};
  manual.forEach(function (item) { chavesManual[chaveDuplicidade_(item)] = true; });

  function aplicarStatus(sheet, itens, chavesOutraAba) {
    if (!itens.length) return;
    var statusCol = HEADERS.indexOf('Status') + 1;
    var valores = itens.map(function (item) {
      var duplicado = !!chavesOutraAba[chaveDuplicidade_(item)];
      return [duplicado ? 'Duplicado – revisar' : 'OK'];
    });
    // Uma única chamada em lote em vez de milhares de setValue() individuais
    // (que é o que fazia a importação travar por vários minutos).
    sheet.getRange(2, statusCol, valores.length, 1).setValues(valores);
  }
  aplicarStatus(erpSheet, erp, chavesManual);
  aplicarStatus(manualSheet, manual, chavesErp);
}

// Visitar a URL do Web App direto no navegador serve a tela (Index.html).
// ?api=json mantém o retorno JSON antigo, usado pela versão local de
// web/index.html (que ainda faz fetch() em vez de google.script.run).
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'json') {
    return ContentService
      .createTextOutput(JSON.stringify(api_carregar()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // A página roda dentro de um iframe isolado (googleusercontent.com), então
  // location.href ali dentro NÃO é a URL pública do Web App — por isso
  // injetamos a URL certa (ScriptApp.getService().getUrl()) como variável de
  // template, para o Index.html usar nas chamadas fetch() ao backend.
  var template = HtmlService.createTemplateFromFile('Index');
  template.urlBackend = ScriptApp.getService().getUrl();
  // ?modo=leitura esconde as ações de escrita na tela (importar, lançar,
  // remover) — é uma restrição só de interface (o link compartilhado com
  // terceiros continua rodando com as mesmas permissões do dono do script;
  // ver aviso no apps-script/README.md).
  template.modoLeitura = !!(e && e.parameter && e.parameter.modo === 'leitura');
  return template.evaluate()
    .setTitle('Entrada até Baixas — Financeiro Protheus')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Substitui o conteúdo da aba ERP pelos itens informados e recalcula duplicidade.
function substituirErp_(itens, dataBase) {
  var sh = getSheet_('ERP');
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), HEADERS.length).clearContent();
  sh.appendRow(HEADERS);
  var linhas = itens.map(function (item) {
    var id = Utilities.getUuid();
    return HEADERS.map(function (h) {
      if (h === 'ID') return id;
      if (h === 'Origem') return 'ERP';
      if (h === 'Status') return 'OK';
      return item[h] != null ? item[h] : '';
    });
  });
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, HEADERS.length).setValues(linhas);
  }
  setDataBase_(dataBase);
  recalcularStatus_();
}

// ---------- Ações (chamadas via google.script.run pela tela hospedada, e
// via doPost/JSON pela versão local de web/index.html — mesma lógica). ----------

function api_carregar() {
  return { erp: lerAba_('ERP'), manual: lerAba_('Manual'), dataBase: getDataBase_() };
}

function api_importar(payload) {
  substituirErp_(payload.itens, payload.dataBase);
  return api_carregar();
}

function api_importarDoDrive() {
  var achado = encontrarArquivoMaisRecenteNoDrive_();
  if (!achado) throw new Error('Nenhum arquivo AAAA.MM.DD.xlsx encontrado na pasta do Drive.');
  var dadosPlanilha = lerLinhasDoXlsx_(achado.file);
  var itensDrive = parseLinhasProtheus_(dadosPlanilha);
  substituirErp_(itensDrive, achado.dataBase);
  var resultado = api_carregar();
  resultado.arquivoUsado = achado.file.getName();
  resultado.quantidadeImportada = itensDrive.length;
  return resultado;
}

function api_adicionarManual(payload) {
  var shM = getSheet_('Manual');
  var id = Utilities.getUuid();
  var linha = HEADERS.map(function (h) {
    if (h === 'ID') return id;
    if (h === 'Origem') return 'Manual';
    if (h === 'Status') return 'OK';
    return payload.item[h] != null ? payload.item[h] : '';
  });
  shM.appendRow(linha);
  recalcularStatus_();
  return api_carregar();
}

function api_removerManual(payload) {
  var shM = getSheet_('Manual');
  var idCol = HEADERS.indexOf('ID') + 1;
  var dados = shM.getRange(2, 1, Math.max(shM.getLastRow() - 1, 0), HEADERS.length).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) {
      shM.deleteRow(i + 2);
      break;
    }
  }
  recalcularStatus_();
  return api_carregar();
}

// doPost continua existindo só para a versão local de web/index.html
// (que fala HTTP/fetch); a tela hospedada usa google.script.run direto
// nas funções api_* acima, sem passar por doPost.
function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var acao = payload.action;
  try {
    var resultado;
    if (acao === 'importar') resultado = api_importar(payload);
    else if (acao === 'importarDoDrive') resultado = api_importarDoDrive();
    else if (acao === 'adicionarManual') resultado = api_adicionarManual(payload);
    else if (acao === 'removerManual') resultado = api_removerManual(payload);
    else throw new Error('Ação desconhecida: ' + acao);
    resultado.ok = true;
    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (erro) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, erro: erro.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
