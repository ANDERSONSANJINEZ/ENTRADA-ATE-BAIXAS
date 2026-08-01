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

// Novas colunas (ex.: Chave Pix) sempre vão no FIM da lista, nunca no meio —
// getSheet_ completa colunas novas só ao final da planilha existente; inserir
// no meio desalinharia todos os dados já gravados nas linhas antigas.
var HEADERS = [
  'ID', 'Tipo', 'Data Emissão', 'Nº Documento', 'Vencimento Real',
  'Forma Pagamento', 'Código Fornecedor', 'Parcela', 'Data Baixa',
  'Vencimento', 'R$ Valor', 'Usuário Inclusor', 'Razão Social',
  'Aprovação Remessa', 'Remessa', 'Histórico', 'Origem', 'Status',
  'Link Documento', 'Link Comprovante', 'Chave Pix'
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

// Aba "Log" — trilha de auditoria de toda ação de escrita (importar,
// lançar/remover manual, excluir duplicados, anexos). Só de leitura pela
// tela; cada linha é gravada por registrarLog_, chamado no fim de cada
// api_* que grava algo. Não passa por getSheet_ (que é específico do
// esquema ERP/Manual) — tem o próprio cabeçalho fixo.
var LOG_HEADERS = ['Data/Hora', 'Usuário', 'Ação', 'Detalhes'];
var LOG_MAX_LINHAS = 2000; // evita a aba crescer sem limite; mantém só as mais recentes

function getSheetLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Log');
  if (!sh) {
    sh = ss.insertSheet('Log');
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function registrarLog_(usuario, acao, detalhes) {
  var sh = getSheetLog_();
  sh.appendRow([new Date(), usuario || '(sem restrição de senha)', acao, detalhes || '']);
  // Poda o excesso só ocasionalmente (não a cada gravação) pra não pagar o
  // custo de deleteRows em toda ação de escrita.
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha > LOG_MAX_LINHAS + 200) {
    sh.deleteRows(2, ultimaLinha - LOG_MAX_LINHAS - 1);
  }
}

// Últimas N entradas do Log, mais recente primeiro — usado pela aba
// "Histórico" da tela. Leitura pura, sem exigir senha de edição (mesma
// lógica de api_carregar).
function api_carregarLog(payload) {
  var sh = getSheetLog_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return { log: [] };
  var limite = (payload && payload.limite) || 500;
  var primeiraLinha = Math.max(2, ultimaLinha - limite + 1);
  var valores = sh.getRange(primeiraLinha, 1, ultimaLinha - primeiraLinha + 1, LOG_HEADERS.length).getValues();
  var linhas = valores.map(function (linha) {
    return { dataHora: linha[0], usuario: linha[1], acao: linha[2], detalhes: linha[3] };
  });
  linhas.reverse(); // mais recente primeiro
  return { log: linhas };
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
  var fornecedor = removerZerosEsquerda_(String(item['Código Fornecedor'] || '').trim());
  var valor = Number(item['R$ Valor'] || 0).toFixed(2);
  var venc = item['Vencimento'];
  var vencStr = venc instanceof Date
    ? Utilities.formatDate(venc, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(venc || '');
  return fornecedor + '|' + valor + '|' + vencStr;
}

// Aceita os dados já lidos (evita reler ERP/Manual quando quem chamou —
// ex.: api_adicionarManual — já tem os dois em mãos).
function recalcularStatusComDados_(erp, manual) {
  var erpSheet = getSheet_('ERP');
  var manualSheet = getSheet_('Manual');

  var chavesErp = {};
  erp.forEach(function (item) { chavesErp[chaveDuplicidade_(item)] = true; });

  if (erp.length) {
    var chavesManual = {};
    manual.forEach(function (item) { chavesManual[chaveDuplicidade_(item)] = true; });
    var statusColErp = HEADERS.indexOf('Status') + 1;
    var valoresErp = erp.map(function (item) {
      return [chavesManual[chaveDuplicidade_(item)] ? 'Duplicado – revisar' : 'OK'];
    });
    // Uma única chamada em lote em vez de milhares de setValue() individuais
    // (que é o que fazia a importação travar por vários minutos).
    erpSheet.getRange(2, statusColErp, valoresErp.length, 1).setValues(valoresErp);
  }

  if (manual.length) {
    // Duplicado se bater com o ERP OU com outro lançamento manual (a partir
    // da 2ª ocorrência da mesma chave) — sem isso, dois lançamentos manuais
    // idênticos entre si (ex.: duplo clique no envio) ficavam os dois como
    // "OK" e o botão "Excluir duplicados" não tinha o que remover.
    var statusColManual = HEADERS.indexOf('Status') + 1;
    var vistosNoManual = {};
    var valoresManual = manual.map(function (item) {
      var chave = chaveDuplicidade_(item);
      var duplicado = !!chavesErp[chave] || !!vistosNoManual[chave];
      vistosNoManual[chave] = true;
      return [duplicado ? 'Duplicado – revisar' : 'OK'];
    });
    manualSheet.getRange(2, statusColManual, valoresManual.length, 1).setValues(valoresManual);
  }
}

function recalcularStatus_() {
  recalcularStatusComDados_(lerAba_('ERP'), lerAba_('Manual'));
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
// Devolve o nome de quem validou (para registrar no Log) — null quando a
// aba "Usuários" ainda está vazia (sem restrição configurada).
function validarAcessoEdicao_(payload) {
  var usuarios = listarUsuarios_();
  if (!usuarios.length) return null;
  var senhaFornecida = String((payload && payload.senha) || '').trim();
  var encontrado = usuarios.filter(function (u) { return u.senha === senhaFornecida; })[0];
  if (!encontrado) {
    throw new Error('Senha de edição incorreta ou não informada.');
  }
  return encontrado.nome;
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
  var nomeUsuario = validarAcessoEdicao_(payload);
  substituirErp_(payload.itens, payload.dataBase);
  registrarLog_(nomeUsuario, 'Importar (upload)', (payload.itens || []).length + ' título(s), base ' + (payload.dataBase || 'não identificada'));
  return api_carregar();
}

function api_importarDoDrive(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var achado = encontrarArquivoMaisRecenteNoDrive_();
  if (!achado) throw new Error('Nenhum arquivo AAAA.MM.DD.xlsx encontrado na pasta do Drive.');
  var dadosPlanilha = lerLinhasDoXlsx_(achado.file);
  var itensDrive = parseLinhasProtheus_(dadosPlanilha);
  substituirErp_(itensDrive, achado.dataBase);
  registrarLog_(nomeUsuario, 'Importar (Drive)', itensDrive.length + ' título(s) de "' + achado.file.getName() + '"');
  var resultado = api_carregar();
  resultado.arquivoUsado = achado.file.getName();
  resultado.quantidadeImportada = itensDrive.length;
  return resultado;
}

function api_adicionarManual(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  // Lê ERP/Manual uma única vez e reaproveita tanto para a conferência de
  // duplicidade quanto para recalcularStatusComDados_ logo abaixo — antes
  // isso lia as duas abas inteiras 3x numa mesma chamada (~3800 linhas do
  // ERP), o que dava a impressão de tela travada ao incluir um lançamento.
  var erp = lerAba_('ERP');
  var manual = lerAba_('Manual');
  var chaveNovo = chaveDuplicidade_(payload.item);
  var jaExiste = erp.concat(manual).some(function (i) {
    return chaveDuplicidade_(i) === chaveNovo;
  });
  if (jaExiste) {
    throw new Error('Já existe um lançamento com o mesmo Código Fornecedor, Valor e Vencimento (no ERP ou já lançado manualmente). Inclusão bloqueada.');
  }

  var shM = getSheet_('Manual');
  var id = Utilities.getUuid();
  var novoItem = {};
  var linha = HEADERS.map(function (h) {
    var valor;
    if (h === 'ID') valor = id;
    else if (h === 'Origem') valor = 'Manual';
    else if (h === 'Status') valor = 'OK';
    else valor = payload.item[h] != null ? payload.item[h] : '';
    novoItem[h] = valor;
    return valor;
  });
  shM.appendRow(linha);
  manual.push(novoItem);
  recalcularStatusComDados_(erp, manual);
  registrarLog_(nomeUsuario, 'Incluir lançamento manual',
    (novoItem['Razão Social'] || novoItem['Código Fornecedor'] || '') + ' | Nº ' + (novoItem['Nº Documento'] || '') +
    ' | Venc.: ' + (novoItem['Vencimento'] || '') + ' | R$ ' + (novoItem['R$ Valor'] || ''));
  return api_carregar();
}

// Remove de uma vez todos os lançamentos manuais marcados como duplicados
// (mesma chave de um título do ERP), evitando remover linha por linha.
function api_removerDuplicados(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var shM = getSheet_('Manual');
  var ultimaLinha = shM.getLastRow();
  if (ultimaLinha < 2) return api_carregar();

  var statusCol = HEADERS.indexOf('Status') + 1;
  var docCol = HEADERS.indexOf('Nº Documento') + 1;
  var razaoCol = HEADERS.indexOf('Razão Social') + 1;
  var dados = shM.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  var linhasParaRemover = [];
  var descricoes = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][statusCol - 1] === 'Duplicado – revisar') {
      linhasParaRemover.push(i + 2);
      descricoes.push((dados[i][razaoCol - 1] || '') + ' Nº ' + (dados[i][docCol - 1] || ''));
    }
  }
  // Remove de baixo para cima para não invalidar os números de linha já calculados.
  linhasParaRemover.sort(function (a, b) { return b - a; });
  linhasParaRemover.forEach(function (linha) { shM.deleteRow(linha); });

  recalcularStatus_();
  if (linhasParaRemover.length) {
    registrarLog_(nomeUsuario, 'Excluir duplicados', linhasParaRemover.length + ' lançamento(s): ' + descricoes.join('; '));
  }
  var resultado = api_carregar();
  resultado.quantidadeRemovida = linhasParaRemover.length;
  return resultado;
}

// Grava o link de anexo (documento ou comprovante) digitado manualmente
// pelo usuário para um lançamento — a busca automática no Drive foi
// removida (não achava os arquivos de forma confiável). payload:
// { id, origem: 'ERP'|'Manual', tipo: 'documento'|'comprovante', url }.
function api_definirAnexo(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var origem = payload.origem === 'Manual' ? 'Manual' : 'ERP';
  var coluna = payload.tipo === 'comprovante' ? 'Link Comprovante' : 'Link Documento';
  var colIdx = HEADERS.indexOf(coluna) + 1;
  var idCol = HEADERS.indexOf('ID') + 1;
  var docCol = HEADERS.indexOf('Nº Documento') + 1;
  var razaoCol = HEADERS.indexOf('Razão Social') + 1;

  var sh = getSheet_(origem);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) throw new Error('Lançamento não encontrado.');
  var dados = sh.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) {
      sh.getRange(i + 2, colIdx).setValue(String(payload.url || ''));
      registrarLog_(nomeUsuario, payload.url ? 'Definir anexo (' + payload.tipo + ')' : 'Remover anexo (' + payload.tipo + ')',
        (dados[i][razaoCol - 1] || '') + ' Nº ' + (dados[i][docCol - 1] || '') + (payload.url ? ': ' + payload.url : ''));
      return api_carregar();
    }
  }
  throw new Error('Lançamento não encontrado (pode ter sido removido ou reimportado).');
}

function api_removerManual(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var shM = getSheet_('Manual');
  var idCol = HEADERS.indexOf('ID') + 1;
  var docCol = HEADERS.indexOf('Nº Documento') + 1;
  var razaoCol = HEADERS.indexOf('Razão Social') + 1;
  var valorCol = HEADERS.indexOf('R$ Valor') + 1;
  var dados = shM.getRange(2, 1, Math.max(shM.getLastRow() - 1, 0), HEADERS.length).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) {
      registrarLog_(nomeUsuario, 'Remover lançamento manual',
        (dados[i][razaoCol - 1] || '') + ' | Nº ' + (dados[i][docCol - 1] || '') + ' | R$ ' + (dados[i][valorCol - 1] || ''));
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
    else if (acao === 'carregarLog') resultado = api_carregarLog(payload);
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

// ---------- Alerta diário por e-mail ----------
// Sem "_" no final do nome (diferente da maioria das funções internas deste
// arquivo) DE PROPÓSITO: essas duas funções (esta e configurarGatilhoDiario,
// mais abaixo) precisam aparecer nos menus "Executar" e "Adicionar gatilho"
// do editor do Apps Script para serem configuradas manualmente — e o Apps
// Script esconde desses menus, por convenção, qualquer função cujo nome
// termine em "_" (tratando como "privada"). Todas as outras funções deste
// arquivo mantêm o "_" de propósito (nunca precisam ser chamadas na mão).
//
// Roda 1x/dia via gatilho de tempo (ver configurarGatilhoDiario no fim
// deste arquivo) e avisa consorciovltce@gmail.com só quando há algo NOVO:
// título(s) vencido(s) que ainda não tinham entrado em nenhum alerta
// anterior, ou a base de dados desatualizada há alguns dias. Não manda
// e-mail todo dia à toa — só quando muda algo que precisa de atenção.
var EMAIL_ALERTA = 'consorciovltce@gmail.com';
var LIMITE_DIAS_BASE_DESATUALIZADA = 2; // mesmo limite usado no aviso da tela (badge-data-base)

function formatarDataServidor_(v) {
  var iso = dataParaISO_(v);
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? (m[3] + '/' + m[2] + '/' + m[1]) : '';
}

function formatarMoedaServidor_(v) {
  var n = Number(v || 0);
  var negativo = n < 0;
  n = Math.abs(n);
  var partes = n.toFixed(2).split('.');
  var inteiro = partes[0], comMilhar = '';
  for (var i = 0; i < inteiro.length; i++) {
    var posDireita = inteiro.length - i;
    comMilhar += inteiro[i];
    if (posDireita > 1 && posDireita % 3 === 1) comMilhar += '.';
  }
  return (negativo ? '-' : '') + 'R$ ' + comMilhar + ',' + partes[1];
}

// Compara datas "yyyy-MM-dd" como texto (funciona porque o formato tem
// tamanho fixo e ordem crescente de dígitos) — evita ficar convertendo
// pra timestamp só pra saber se uma data é anterior à outra.
function ehVencido_(vencimento, hojeIso) {
  var iso = dataParaISO_(vencimento);
  return !!iso && String(iso) < hojeIso;
}

function diasEntreIso_(isoAntiga, isoNova) {
  var pa = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(isoAntiga || ''));
  var pn = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(isoNova || ''));
  if (!pa || !pn) return 0;
  var a = Date.UTC(Number(pa[1]), Number(pa[2]) - 1, Number(pa[3]));
  var n = Date.UTC(Number(pn[1]), Number(pn[2]) - 1, Number(pn[3]));
  return Math.round((n - a) / 86400000);
}

// Aba oculta que guarda, a cada execução, a chave (Nº Documento + Código
// Fornecedor + Parcela — mesma chaveAnexo_ usada para preservar links entre
// reimportações) de cada título vencido já avisado — assim o próximo
// alerta só lista o que é NOVO, em vez de repetir a mesma lista todo dia.
// Usa uma aba (não PropertiesService) porque o limite de 9KB por
// propriedade estouraria fácil com centenas de títulos vencidos.
function getSheetEstadoAlerta_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_EstadoAlerta');
  if (!sh) {
    sh = ss.insertSheet('_EstadoAlerta');
    try { sh.hideSheet(); } catch (e) { /* segue sem esconder se não puder (ex.: única aba visível) */ }
  }
  return sh;
}
function lerChavesNotificadas_() {
  var sh = getSheetEstadoAlerta_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 1) return {};
  var valores = sh.getRange(1, 1, ultimaLinha, 1).getValues();
  var set = {};
  valores.forEach(function (l) { if (l[0]) set[l[0]] = true; });
  return set;
}
function gravarChavesNotificadas_(chaves) {
  var sh = getSheetEstadoAlerta_();
  sh.clearContent();
  if (chaves.length) sh.getRange(1, 1, chaves.length, 1).setValues(chaves.map(function (c) { return [c]; }));
}

function enviarAlertaDiario() {
  var erp = lerAba_('ERP');
  var manual = lerAba_('Manual').filter(function (m) { return m.Status !== 'Duplicado – revisar'; });
  var itens = erp.concat(manual);

  var hojeIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var vencidos = itens.filter(function (i) { return !i['Data Baixa'] && ehVencido_(i['Vencimento'], hojeIso); });

  var chavesAtuais = vencidos.map(chaveAnexo_);
  var notificadasAntes = lerChavesNotificadas_();
  var novos = vencidos.filter(function (i, idx) { return !notificadasAntes[chavesAtuais[idx]]; });

  var dataBaseIso = getDataBase_();
  var diasDefasagem = dataBaseIso ? diasEntreIso_(dataBaseIso, hojeIso) : 0;
  var baseDesatualizada = !!dataBaseIso && diasDefasagem >= LIMITE_DIAS_BASE_DESATUALIZADA;

  // Atualiza o estado sempre (mesmo sem mandar e-mail hoje), pra que um
  // título que já venceu antes não seja tratado como "novo" de novo amanhã.
  gravarChavesNotificadas_(chavesAtuais);

  if (!novos.length && !baseDesatualizada) return; // nada novo pra avisar hoje

  var linhas = [];
  linhas.push('Resumo automático — Entrada até Baixas');
  linhas.push(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  linhas.push('');

  if (baseDesatualizada) {
    linhas.push('⚠ A base de dados está desatualizada há ' + diasDefasagem + ' dia(s) — importe um arquivo mais recente da pasta do Drive.');
    linhas.push('');
  }

  if (novos.length) {
    var valorNovos = novos.reduce(function (s, i) { return s + Number(i['R$ Valor'] || 0); }, 0);
    linhas.push(novos.length + ' novo(s) título(s) venceram e continuam em aberto, somando ' + formatarMoedaServidor_(valorNovos) + ':');
    novos
      .slice()
      .sort(function (a, b) { return Number(b['R$ Valor'] || 0) - Number(a['R$ Valor'] || 0); })
      .slice(0, 25)
      .forEach(function (i) {
        linhas.push('- ' + (i['Razão Social'] || i['Código Fornecedor'] || '(sem fornecedor)') +
          ' | Nº ' + (i['Nº Documento'] || '') + ' | Venc.: ' + formatarDataServidor_(i['Vencimento']) +
          ' | ' + formatarMoedaServidor_(i['R$ Valor']));
      });
    if (novos.length > 25) linhas.push('... e mais ' + (novos.length - 25) + ' título(s) — veja a lista completa no app.');
    linhas.push('');
  }

  var totalVencidoValor = vencidos.reduce(function (s, i) { return s + Number(i['R$ Valor'] || 0); }, 0);
  linhas.push('Total geral em aberto e vencido no momento: ' + vencidos.length + ' título(s), ' + formatarMoedaServidor_(totalVencidoValor) + '.');
  linhas.push('');
  linhas.push('Acesse o app: ' + ScriptApp.getService().getUrl());

  var assuntoPartes = [];
  if (novos.length) assuntoPartes.push(novos.length + ' título(s) vencido(s) novo(s)');
  if (baseDesatualizada) assuntoPartes.push('base desatualizada');
  MailApp.sendEmail({
    to: EMAIL_ALERTA,
    subject: '[Entrada até Baixas] ' + assuntoPartes.join(' + '),
    body: linhas.join('\n'),
  });
  registrarLog_('(sistema)', 'Alerta automático enviado', assuntoPartes.join(' + ') + ' — para ' + EMAIL_ALERTA);
}

// Rode esta função UMA VEZ direto no editor do Apps Script (menu de funções
// no topo → selecione "configurarGatilhoDiario" → Executar) para instalar
// o gatilho diário. É seguro rodar de novo depois (ex.: pra trocar o
// horário) — ela sempre remove o gatilho anterior antes de criar um novo.
// (Alternativa sem rodar código nenhum: no ícone de relógio ⏰ do editor,
// "+ Adicionar gatilho" → função "enviarAlertaDiario" → Baseado em tempo →
// Timer diário → Salvar.)
function configurarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarAlertaDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarAlertaDiario').timeBased().everyDays(1).atHour(7).create();
}
