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
  'Aprovação Remessa', 'Remessa', 'Histórico', 'Origem', 'Status',
  'Link Documento', 'Link Comprovante'
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
    return sh;
  }
  // Se HEADERS ganhou colunas novas desde a criação da planilha (ex.: Link
  // Documento/Comprovante), completa o cabeçalho — sem isso, as colunas
  // novas ficariam sem título na aba (os dados continuam sendo gravados
  // certo, já que lerAba_/escrita sempre usam HEADERS.length).
  if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

// Aba "Usuários" (Nome + Senha) — cadastro de quem pode editar o app. Uma
// planilha separada da HEADERS de ERP/Manual, editável direto pelo dono
// sem mexer em código: acrescentar, trocar ou apagar uma linha já
// cadastra/revoga o acesso daquela pessoa na próxima vez que ela tentar.
var USUARIOS_HEADERS = ['Nome', 'Senha'];
function getSheetUsuarios_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Usuários');
  if (!sh) {
    sh = ss.insertSheet('Usuários');
    sh.appendRow(USUARIOS_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}
function listarUsuarios_() {
  var sh = getSheetUsuarios_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return [];
  var valores = sh.getRange(2, 1, ultimaLinha - 1, 2).getValues();
  return valores
    .map(function (linha) { return { nome: String(linha[0] || '').trim(), senha: String(linha[1] || '').trim() }; })
    .filter(function (u) { return u.nome && u.senha; });
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
// Chave estável de um título entre uma importação e a próxima (o ID em si
// não é estável — é regenerado a cada import). Usada só para preservar os
// links de anexo (ver substituirErp_), não para duplicidade.
function chaveAnexo_(item) {
  return String(item['Nº Documento'] || '').trim() + '|' +
    String(item['Código Fornecedor'] || '').trim() + '|' +
    String(item['Parcela'] || '').trim();
}

function substituirErp_(itens, dataBase) {
  var sh = getSheet_('ERP');

  // Antes de apagar tudo, guarda os links de anexo já cadastrados
  // manualmente, pra não perdê-los na reimportação diária (o ERP inteiro é
  // substituído a cada import, então sem isso todo link se perderia).
  var linksAnteriores = {};
  var ultimaLinhaAntiga = sh.getLastRow();
  if (ultimaLinhaAntiga > 1) {
    var colDoc = HEADERS.indexOf('Link Documento');
    var colComp = HEADERS.indexOf('Link Comprovante');
    var dadosAntigos = sh.getRange(2, 1, ultimaLinhaAntiga - 1, HEADERS.length).getValues();
    dadosAntigos.forEach(function (linha) {
      var linkDoc = linha[colDoc], linkComp = linha[colComp];
      if (!linkDoc && !linkComp) return;
      var obj = {};
      HEADERS.forEach(function (h, i) { obj[h] = linha[i]; });
      linksAnteriores[chaveAnexo_(obj)] = { doc: linkDoc, comp: linkComp };
    });
  }

  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), HEADERS.length).clearContent();
  sh.appendRow(HEADERS);
  var linhas = itens.map(function (item) {
    var id = Utilities.getUuid();
    var preservado = linksAnteriores[chaveAnexo_(item)];
    return HEADERS.map(function (h) {
      if (h === 'ID') return id;
      if (h === 'Origem') return 'ERP';
      if (h === 'Status') return 'OK';
      if (h === 'Link Documento') return (preservado && preservado.doc) || '';
      if (h === 'Link Comprovante') return (preservado && preservado.comp) || '';
      return item[h] != null ? item[h] : '';
    });
  });
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, HEADERS.length).setValues(linhas);
  }
  setDataBase_(dataBase);
  recalcularStatus_();
}

// ---------- Ações (chamadas via doPost/fetch pela tela hospedada e pela
// versão local de web/index.html — mesma lógica). ----------

// Ações de escrita (tudo exceto carregar) exigem uma senha que bata com
// alguma linha da aba "Usuários". Enquanto essa aba estiver vazia, o app
// continua liberado (comportamento de antes) — assim a trava só entra em
// vigor quando o dono cadastrar o primeiro usuário.
function validarAcessoEdicao_(payload) {
  var usuarios = listarUsuarios_();
  if (!usuarios.length) return;
  var senhaFornecida = String((payload && payload.senha) || '').trim();
  var valido = usuarios.some(function (u) { return u.senha === senhaFornecida; });
  if (!valido) {
    throw new Error('Senha de edição incorreta ou não informada.');
  }
}

// Confere a senha e devolve de quem ela é — usado só na hora de "Habilitar
// edição" na tela, pra confirmar antes de liberar a interface (em vez de só
// descobrir que a senha está errada na primeira ação de escrita).
function api_validarEdicao(payload) {
  var usuarios = listarUsuarios_();
  if (!usuarios.length) return { nome: null, semRestricao: true };
  var senhaFornecida = String((payload && payload.senha) || '').trim();
  var encontrado = usuarios.filter(function (u) { return u.senha === senhaFornecida; })[0];
  if (!encontrado) throw new Error('Senha de edição incorreta ou não informada.');
  return { nome: encontrado.nome };
}

function api_carregar() {
  return {
    erp: lerAba_('ERP'),
    manual: lerAba_('Manual'),
    dataBase: getDataBase_(),
    usuariosConfigurados: listarUsuarios_().length > 0,
  };
}

function api_importar(payload) {
  validarAcessoEdicao_(payload);
  substituirErp_(payload.itens, payload.dataBase);
  return api_carregar();
}

function api_importarDoDrive(payload) {
  validarAcessoEdicao_(payload);
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
  validarAcessoEdicao_(payload);
  var chaveNovo = chaveDuplicidade_(payload.item);
  var jaExiste = lerAba_('ERP').concat(lerAba_('Manual')).some(function (i) {
    return chaveDuplicidade_(i) === chaveNovo;
  });
  if (jaExiste) {
    throw new Error('Já existe um lançamento com o mesmo Código Fornecedor, Valor e Vencimento (no ERP ou já lançado manualmente). Inclusão bloqueada.');
  }

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

// Remove de uma vez todos os lançamentos manuais marcados como duplicados
// (mesma chave de um título do ERP), evitando remover linha por linha.
function api_removerDuplicados(payload) {
  validarAcessoEdicao_(payload);
  var shM = getSheet_('Manual');
  var ultimaLinha = shM.getLastRow();
  if (ultimaLinha < 2) return api_carregar();

  var statusCol = HEADERS.indexOf('Status') + 1;
  var dados = shM.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  var linhasParaRemover = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][statusCol - 1] === 'Duplicado – revisar') linhasParaRemover.push(i + 2);
  }
  // Remove de baixo para cima para não invalidar os números de linha já calculados.
  linhasParaRemover.sort(function (a, b) { return b - a; });
  linhasParaRemover.forEach(function (linha) { shM.deleteRow(linha); });

  recalcularStatus_();
  var resultado = api_carregar();
  resultado.quantidadeRemovida = linhasParaRemover.length;
  return resultado;
}

// Grava o link de anexo (documento ou comprovante) digitado manualmente
// pelo usuário para um lançamento — a busca automática no Drive foi
// removida (não achava os arquivos de forma confiável). payload:
// { id, origem: 'ERP'|'Manual', tipo: 'documento'|'comprovante', url }.
function api_definirAnexo(payload) {
  validarAcessoEdicao_(payload);
  var origem = payload.origem === 'Manual' ? 'Manual' : 'ERP';
  var coluna = payload.tipo === 'comprovante' ? 'Link Comprovante' : 'Link Documento';
  var colIdx = HEADERS.indexOf(coluna) + 1;
  var idCol = HEADERS.indexOf('ID') + 1;

  var sh = getSheet_(origem);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) throw new Error('Lançamento não encontrado.');
  var ids = sh.getRange(2, idCol, ultimaLinha - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === payload.id) {
      sh.getRange(i + 2, colIdx).setValue(String(payload.url || ''));
      return api_carregar();
    }
  }
  throw new Error('Lançamento não encontrado (pode ter sido removido ou reimportado).');
}

function api_removerManual(payload) {
  validarAcessoEdicao_(payload);
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

// doPost é o único caminho usado pela tela (fetch), tanto para leitura
// (GET ?api=json) quanto para as ações de escrita abaixo.
function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var acao = payload.action;
  try {
    var resultado;
    if (acao === 'importar') resultado = api_importar(payload);
    else if (acao === 'importarDoDrive') resultado = api_importarDoDrive(payload);
    else if (acao === 'adicionarManual') resultado = api_adicionarManual(payload);
    else if (acao === 'removerManual') resultado = api_removerManual(payload);
    else if (acao === 'removerDuplicados') resultado = api_removerDuplicados(payload);
    else if (acao === 'definirAnexo') resultado = api_definirAnexo(payload);
    else if (acao === 'validarEdicao') resultado = api_validarEdicao(payload);
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
