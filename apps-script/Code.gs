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
var PROP_SENHA_EDICAO = 'senhaEdicao';

// Pasta do Drive onde o export diário do Protheus é salvo:
// Compartilhados comigo / 4 - CE 007_ADMINISTRATIVO / 4.11 - FINANCEIRO / 17 CONTAS A PAGAR
// Nome do arquivo esperado: AAAA.MM.DD.xlsx (ex.: 2026.07.30.xlsx)
var DRIVE_FOLDER_ID = '1sVlF29VGWDzHelgBGIeFvVK3OCpMjGmD';
var NOME_ARQUIVO_RE = /^(\d{4})\.(\d{2})\.(\d{2})\.xlsx$/i;

// Pastas do Drive com os anexos de cada lançamento (botões "Anexos" em
// Lançamentos). Nome dos arquivos: documentos = Nº Documento + nome do
// fornecedor; comprovantes = "comprovante" (fixo) + Nº Documento + (parte
// do) nome/razão social do fornecedor. A busca (buscarArquivosPorNumero_)
// usa só o Nº Documento como critério principal e o fornecedor só pra
// desempatar quando há mais de um resultado — então não depende de o nome
// ter exatamente esse formato/ordem.
var PASTAS_DOCUMENTOS = ['1P4oN7lQopWk2hgNXEDNIdxIgTp1tqeiK', '18nuIj4SsAEqPJ7jOxlUe8JzJ8TCNcueL'];
var PASTAS_COMPROVANTES = ['1leuDOfqLFDxdg3PECZ6aij2eIkPGKwq6'];

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

// ---------- Ações (chamadas via doPost/fetch pela tela hospedada e pela
// versão local de web/index.html — mesma lógica). ----------

// Ações de escrita (tudo exceto carregar/buscarAnexo) exigem a senha de
// edição configurada em Propriedades do script. Enquanto ninguém configurar
// essa propriedade, o app continua liberado (comportamento de antes) — assim
// a trava só entra em vigor quando o dono decide ativá-la.
function validarAcessoEdicao_(payload) {
  var senhaConfigurada = PropertiesService.getScriptProperties().getProperty(PROP_SENHA_EDICAO);
  if (!senhaConfigurada) return;
  var senhaFornecida = String((payload && payload.senha) || '');
  if (senhaFornecida !== senhaConfigurada) {
    throw new Error('Senha de edição incorreta ou não informada.');
  }
}

function api_carregar() {
  return {
    erp: lerAba_('ERP'),
    manual: lerAba_('Manual'),
    dataBase: getDataBase_(),
    senhaConfigurada: !!PropertiesService.getScriptProperties().getProperty(PROP_SENHA_EDICAO),
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

// Deixa só letras/números (maiúsculo) — usado pra comparar nomes de arquivo
// ignorando pontuação (ex.: "54.610", "Nº 54610" e "54610" batem entre si).
function normalizarNomeArquivo_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9À-Ÿ]/g, '');
}

// Lista os arquivos de uma pasta do Drive E de todas as subpastas dela
// (várias equipes organizam os documentos em subpastas por mês/fornecedor,
// e uma busca só na pasta-raiz nunca acharia esses arquivos). Via Drive API
// avançada, com paginação. Normaliza a resposta porque o serviço avançado
// "Drive API" pode ter sido habilitado como v2 (resposta.items, campos
// title/alternateLink) ou v3 (resposta.files, campos name/webViewLink) — e
// inclui os parâmetros que fazem a busca alcançar pastas dentro de Drives
// compartilhados (Shared Drives), que sem eles simplesmente não aparecem.
var MIME_PASTA_DRIVE = 'application/vnd.google-apps.folder';
function listarArquivosDaPasta_(folderId) {
  var arquivos = [];
  var filaPastas = [folderId];
  var visitadas = {};
  var pastasProcessadas = 0;

  while (filaPastas.length && pastasProcessadas < 200) { // trava de segurança
    var pastaAtual = filaPastas.shift();
    if (visitadas[pastaAtual]) continue;
    visitadas[pastaAtual] = true;
    pastasProcessadas++;

    var query = "'" + pastaAtual + "' in parents and trashed = false";
    var pageToken = null;
    var paginas = 0;
    do {
      var opcoesBase = {
        q: query,
        pageSize: 1000,
        maxResults: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        pageToken: pageToken || undefined,
      };
      var resposta;
      try {
        // Serviço avançado configurado como Drive API v3.
        resposta = Drive.Files.list(Object.assign({ fields: 'nextPageToken,files(id,name,webViewLink,mimeType)' }, opcoesBase));
      } catch (e1) {
        // Serviço avançado configurado como Drive API v2 (máscara de campos
        // de v3 é inválida em v2, então tenta de novo com os nomes de v2).
        resposta = Drive.Files.list(Object.assign({ fields: 'nextPageToken,items(id,title,alternateLink,mimeType)' }, opcoesBase));
      }
      var pagina = resposta.files || resposta.items || [];
      pagina.forEach(function (f) {
        if (f.mimeType === MIME_PASTA_DRIVE) {
          filaPastas.push(f.id);
        } else {
          arquivos.push({ nome: f.name || f.title, url: f.webViewLink || f.alternateLink });
        }
      });
      pageToken = resposta.nextPageToken || null;
      paginas++;
    } while (pageToken && paginas < 20); // trava de segurança: até 20.000 arquivos por pasta
  }
  return arquivos;
}

// Busca, numa pasta do Drive, arquivos cujo nome contenha o Nº Documento
// informado. Nome dos arquivos, por convenção da equipe: documentos = Nº
// Documento + nome do fornecedor; comprovantes = "comprovante" (fixo) + Nº
// Documento + (parte do) nome/razão social do fornecedor.
// Primeiro tenta uma consulta direta ao Drive (rápida); se não achar nada,
// cai para listar a pasta inteira e comparar os nomes ignorando pontuação
// — cobre casos como o nome do arquivo usar "54.610" para o Nº Documento
// "54610", que a consulta direta (substring exata) não encontraria.
function buscarArquivosPorNumero_(folderId, numDoc) {
  var escapado = numDoc.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var query = "'" + folderId + "' in parents and trashed = false and name contains '" + escapado + "'";
  var opcoesBase = {
    q: query,
    pageSize: 25,
    maxResults: 25,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  };
  var resposta;
  try {
    resposta = Drive.Files.list(Object.assign({ fields: 'files(id,name,webViewLink)' }, opcoesBase));
  } catch (e1) {
    resposta = Drive.Files.list(Object.assign({ fields: 'items(id,title,alternateLink)' }, opcoesBase));
  }
  var arquivos = resposta.files || resposta.items || [];
  var encontrados = arquivos.map(function (f) {
    return { nome: f.name || f.title, url: f.webViewLink || f.alternateLink };
  });
  if (encontrados.length) return { encontrados: encontrados, totalVarrido: encontrados.length };

  var todosDaPasta = listarArquivosDaPasta_(folderId);
  var alvo = normalizarNomeArquivo_(numDoc);
  var viaListagem = alvo
    ? todosDaPasta.filter(function (f) { return normalizarNomeArquivo_(f.nome).indexOf(alvo) !== -1; })
    : [];
  return { encontrados: viaListagem, totalVarrido: todosDaPasta.length };
}

function api_buscarAnexo(payload) {
  var numDoc = String(payload.numDoc || '').trim();
  var fornecedor = String(payload.fornecedor || '').trim();
  var tipo = payload.tipo === 'comprovante' ? 'comprovante' : 'documento';
  var pastas = tipo === 'comprovante' ? PASTAS_COMPROVANTES : PASTAS_DOCUMENTOS;
  if (!numDoc) return { encontrados: [], erros: [] };

  var encontrados = [];
  var erros = [];
  var totalVarrido = 0;
  pastas.forEach(function (folderId) {
    try {
      var resultado = buscarArquivosPorNumero_(folderId, numDoc);
      encontrados = encontrados.concat(resultado.encontrados);
      totalVarrido += resultado.totalVarrido;
    } catch (e) {
      erros.push('Pasta ' + folderId + ': ' + e.message);
    }
  });

  // Com mais de um resultado (Nº Documento repetido entre fornecedores),
  // refina pelo nome do fornecedor para achar o arquivo certo.
  if (encontrados.length > 1 && fornecedor) {
    var palavras = fornecedor.toUpperCase().split(/\s+/).filter(function (p) { return p.length > 2; });
    var refinado = encontrados.filter(function (a) {
      var nome = a.nome.toUpperCase();
      return palavras.some(function (p) { return nome.indexOf(p) !== -1; });
    });
    if (refinado.length) encontrados = refinado;
  }
  return { encontrados: encontrados, erros: erros, totalVarrido: totalVarrido };
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
    else if (acao === 'buscarAnexo') resultado = api_buscarAnexo(payload);
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
