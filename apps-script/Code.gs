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
  'Link Documento', 'Link Comprovante', 'Chave Pix', 'Observação', 'Categoria'
];

// Perfis de acesso restrito por link próprio: ?usuario=<chave> na URL do
// Web App. Cada perfil força modo leitura, esconde da navegação todas as
// abas fora de "abas", e filtra ERP/Manual (dentro de api_carregar) pra
// nunca mandar ao navegador os títulos dos fornecedores listados em
// "fornecedoresExcluidos" — diferente do ?modo=leitura comum (só esconde
// botão na tela), aqui a exclusão de fornecedor é reforçada no servidor,
// antes do dado sair da planilha.
var PERFIS_RESTRITOS = {
  'sandro-costa': {
    nome: 'Sandro Costa',
    abas: ['lancamentos'],
    fornecedoresExcluidos: [
      '50458148', '63267110', '22407874', '5775268', '35021022',
      '46170722', '63067344', '53734399', '60634166', '67791821',
      '66769166', '63822521', '47811670', '67820752',
    ],
    // Simplifica a tela de "Todos os lançamentos" pra esse perfil: some
    // busca/filtros/exportar/totalizador e as colunas Origem, Status,
    // Aporte, Usuário e Remessa (não fazem sentido pra um link externo
    // travado numa aba só) — a tabela passa a ocupar a largura toda,
    // redistribuindo o espaço entre as colunas que sobraram.
    tabelaLancamentosSimplificada: true,
  },
};

function obterPerfilRestrito_(e) {
  var chave = e && e.parameter && e.parameter.usuario;
  if (!chave) return null;
  return PERFIS_RESTRITOS[String(chave).toLowerCase()] || null;
}

var PROP_DATA_BASE = 'dataBaseImportacao';

function getDataBase_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_DATA_BASE);
}
function setDataBase_(valor) {
  if (valor) PropertiesService.getScriptProperties().setProperty(PROP_DATA_BASE, valor);
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

// Aba "Conciliação Bancária" — histórico persistido de cada extrato já
// processado na tela (aba Conciliação Bancária), pra a aba Análise poder
// enxergar dados de banco de QUALQUER sessão passada, não só o último
// extrato ainda na memória de quem está com a tela aberta agora (o extrato
// em si nunca é salvo, só o resultado do cruzamento linha a linha). Cresce
// por acréscimo a cada processamento — igual ao Log — sem deduplicar
// reimportações do mesmo arquivo; é um histórico de auditoria, não um
// espelho do extrato mais recente.
var CONCILIACAO_HEADERS = [
  'Data Importação', 'Usuário', 'Situação', 'Data Movimento Extrato',
  'Valor Extrato', 'Histórico Extrato', 'Nº Documento Título',
  'Fornecedor Título', 'Data Baixa Título', 'Valor Título'
];
var CONCILIACAO_MAX_LINHAS = 20000;

function getSheetConciliacao_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Conciliação Bancária');
  if (!sh) {
    sh = ss.insertSheet('Conciliação Bancária');
    sh.appendRow(CONCILIACAO_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function api_salvarConciliacao(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var linhas = (payload && payload.linhas) || [];
  if (!linhas.length) throw new Error('Nenhuma linha de conciliação para salvar.');
  var agora = new Date();
  var sh = getSheetConciliacao_();
  var valores = linhas.map(function (l) {
    return [
      agora, nomeUsuario || '(sem restrição de senha)', l.situacao || '',
      l.dataExtrato || '', l.valorExtrato === '' ? '' : Number(l.valorExtrato || 0),
      l.historicoExtrato || '', l.nDoc || '', l.fornecedor || '',
      l.dataBaixaTitulo || '', l.valorTitulo === '' ? '' : Number(l.valorTitulo || 0),
    ];
  });
  sh.getRange(sh.getLastRow() + 1, 1, valores.length, CONCILIACAO_HEADERS.length).setValues(valores);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha > CONCILIACAO_MAX_LINHAS + 200) {
    sh.deleteRows(2, ultimaLinha - CONCILIACAO_MAX_LINHAS - 1);
  }
  registrarLog_(nomeUsuario, 'Salvar conciliação bancária', valores.length + ' lançamento(s) do extrato processados');
  return { resumoConciliacao: resumoConciliacaoBancaria_() };
}

// Agregado por situação (qtde + valor) pra a aba Análise mostrar sem
// precisar mandar pro cliente as dezenas de milhares de linhas cruas —
// lido sob demanda a cada api_carregar() porque a aba Conciliação Bancária
// pode crescer bastante e isso é só um resumo, calculado uma vez aqui.
function resumoConciliacaoBancaria_() {
  // getSheetByName (nunca getSheetConciliacao_/insertSheet) DE PROPÓSITO —
  // isto roda em todo api_carregar(), ou seja, em toda visita à tela; criar
  // a aba nova aqui faria até quem só está LENDO a tela disparar uma
  // mudança estrutural na planilha (mais lenta que uma leitura) sem nunca
  // ter processado extrato nenhum. A aba só é criada de fato quando
  // alguém salva uma conciliação de verdade (api_salvarConciliacao).
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conciliação Bancária');
  if (!sh) return { totalLinhas: 0, porSituacao: {}, qtdeImportacoes: 0, ultimaImportacao: null };
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return { totalLinhas: 0, porSituacao: {}, qtdeImportacoes: 0, ultimaImportacao: null };
  var valores = sh.getRange(2, 1, ultimaLinha - 1, CONCILIACAO_HEADERS.length).getValues();
  var porSituacao = {};
  var timestampsImportacao = {};
  var ultimaImportacao = null;
  valores.forEach(function (linha) {
    var dataImportacao = linha[0], situacao = linha[2], valorExtrato = linha[4], valorTitulo = linha[9];
    var valor = Number(valorExtrato || valorTitulo || 0);
    if (!porSituacao[situacao]) porSituacao[situacao] = { qtde: 0, valor: 0 };
    porSituacao[situacao].qtde += 1;
    porSituacao[situacao].valor += valor;
    var chaveImportacao = dataImportacao instanceof Date ? dataImportacao.getTime() : String(dataImportacao);
    timestampsImportacao[chaveImportacao] = true;
    if (dataImportacao instanceof Date && (!ultimaImportacao || dataImportacao > ultimaImportacao)) ultimaImportacao = dataImportacao;
  });
  return {
    totalLinhas: valores.length,
    porSituacao: porSituacao,
    qtdeImportacoes: Object.keys(timestampsImportacao).length,
    ultimaImportacao: ultimaImportacao,
  };
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
  var idCol = HEADERS.indexOf('ID');
  return valores
    // Ignora linha sem ID: toda linha de verdade (ERP ou Manual) sempre
    // ganha um ID (Utilities.getUuid()) na hora em que é criada — uma
    // linha em branco só existe por sobra na planilha (ex.: linha extra
    // que ficou "fantasma" depois de algum ajuste manual direto no Sheets).
    // Sem este filtro, essas linhas em branco eram lidas como um
    // lançamento de verdade com tudo vazio/zerado, e — como todas batem na
    // mesma chave de duplicidade ("" | "0.00" | "") — apareciam juntas como
    // "Duplicado a revisar" com R$ 0,00 no Dashboard.
    .filter(function (linha) { return linha[idCol] !== '' && linha[idCol] != null; })
    .map(function (linha) {
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
  var perfil = obterPerfilRestrito_(e);
  if (e && e.parameter && e.parameter.api === 'json') {
    return ContentService
      .createTextOutput(JSON.stringify(api_carregar(perfil)))
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
  // ver aviso no apps-script/README.md). Um perfil restrito (?usuario=...)
  // sempre força modo leitura também, além de limitar as abas.
  template.modoLeitura = !!(e && e.parameter && e.parameter.modo === 'leitura') || !!perfil;
  // 'null' (string, sem aspas) quando não há perfil — vira o literal
  // JS "null" ao ser colado no <script> do template; com perfil, JSON.stringify
  // já devolve a string entre aspas certa pra virar um array literal.
  template.chaveUsuarioRestrito = perfil ? JSON.stringify(String(e.parameter.usuario)) : 'null';
  template.abasPermitidasRestrito = perfil ? JSON.stringify(perfil.abas) : 'null';
  template.tabelaLancamentosSimplificada = !!(perfil && perfil.tabelaLancamentosSimplificada);
  return template.evaluate()
    .setTitle('Controle Financeiro')
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

  // Antes de apagar tudo, guarda os links de anexo, a Observação e a
  // Categoria (definida manualmente quando o "de-para" fixo não conhece o
  // fornecedor) já cadastrados, pra não perder na reimportação diária (o
  // ERP inteiro é substituído a cada import, então sem isso tudo isso se
  // perderia).
  var linksAnteriores = {};
  var ultimaLinhaAntiga = sh.getLastRow();
  if (ultimaLinhaAntiga > 1) {
    var colDoc = HEADERS.indexOf('Link Documento');
    var colComp = HEADERS.indexOf('Link Comprovante');
    var colObs = HEADERS.indexOf('Observação');
    var colCat = HEADERS.indexOf('Categoria');
    var dadosAntigos = sh.getRange(2, 1, ultimaLinhaAntiga - 1, HEADERS.length).getValues();
    dadosAntigos.forEach(function (linha) {
      var linkDoc = linha[colDoc], linkComp = linha[colComp], observacao = linha[colObs], categoria = linha[colCat];
      if (!linkDoc && !linkComp && !observacao && !categoria) return;
      var obj = {};
      HEADERS.forEach(function (h, i) { obj[h] = linha[i]; });
      linksAnteriores[chaveAnexo_(obj)] = { doc: linkDoc, comp: linkComp, obs: observacao, cat: categoria };
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
      if (h === 'Observação') return (preservado && preservado.obs) || '';
      if (h === 'Categoria') return (preservado && preservado.cat) || '';
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

// "perfil" (ver PERFIS_RESTRITOS) é opcional — chamadas existentes
// (api_importar, api_adicionarManual etc.) continuam passando nenhum
// argumento e recebem ERP/Manual completos, sem filtro nenhum.
function api_carregar(perfil) {
  var erp = lerAba_('ERP');
  var manual = lerAba_('Manual');
  if (perfil && perfil.fornecedoresExcluidos && perfil.fornecedoresExcluidos.length) {
    var excluidos = {};
    perfil.fornecedoresExcluidos.forEach(function (cod) {
      excluidos[removerZerosEsquerda_(String(cod).trim())] = true;
    });
    var mantemFornecedor_ = function (item) {
      return !excluidos[removerZerosEsquerda_(String(item['Código Fornecedor'] || '').trim())];
    };
    erp = erp.filter(mantemFornecedor_);
    manual = manual.filter(mantemFornecedor_);
  }
  return {
    erp: erp,
    manual: manual,
    dataBase: getDataBase_(),
    usuariosConfigurados: listarUsuarios_().length > 0,
    // Nunca deixa o resumo de conciliação (recurso novo, secundário) travar
    // ou quebrar o carregamento principal (ERP/Manual) se algo der errado
    // aqui — sem isso, um erro nessa parte tornaria a tela inteira incapaz
    // de carregar título nenhum.
    resumoConciliacao: tentarResumoConciliacaoBancaria_(),
  };
}

function tentarResumoConciliacaoBancaria_() {
  try {
    return resumoConciliacaoBancaria_();
  } catch (erro) {
    return { totalLinhas: 0, porSituacao: {}, qtdeImportacoes: 0, ultimaImportacao: null, erro: erro.message };
  }
}

// Planilha externa (mantida pela skill "consolidar-faturamento-vlt") com o
// faturamento do consórcio pro cliente — NFS-e emitidas por medição pelos 3
// prestadores (MPE, A Gaspar, VIPETRO). Usada só pra cruzar Faturado x Pago
// no Dashboard (ver insightsFaturadoPago_ no Index.html) — nunca escrita
// por aqui, só lida.
var FATURAMENTO_PLANILHA_ID_ = '1q6Q9Bc7TJsfUdShrAwJdLXGSUx-3RzfdhEE3LlVVS0A';
var FATURAMENTO_GID_ = 786051406;

// Soma o "Valor Total (R$)" de cada medição (várias linhas — uma por
// prestador — compartilham o mesmo número de Medição) e usa a emissão mais
// recente do grupo como data de referência daquela medição, pra dar pra
// cruzar com o período de baixas no Dashboard. Linhas sem Medição ou sem
// Valor Total numérico (a linha de TOTAL no fim, linhas em branco, notas
// "PENDENTE - PDF escaneado" sem valor extraído) ficam de fora.
//
// "Data de Emissao" pode vir como Date de verdade (célula formatada como
// data) OU como texto "dd/mm/aaaa"/"aaaa-mm-dd" (célula formatada como
// texto) — depende de como cada linha foi digitada na planilha de
// Faturamento. Sem aceitar as duas formas, uma medição inteira ficava sem
// nenhuma data de referência (todas as datas descartadas) e sumia do
// gráfico silenciosamente, mesmo sem erro nenhum na busca.
function paraDataFaturamento_(valor) {
  if (valor instanceof Date) return valor;
  if (typeof valor === 'string' && valor.trim()) {
    var texto = valor.trim();
    var bra = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (bra) return new Date(Number(bra[3]), Number(bra[2]) - 1, Number(bra[1]));
    var iso = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  return null;
}

function faturamentoPorMedicao_() {
  var planilha = SpreadsheetApp.openById(FATURAMENTO_PLANILHA_ID_);
  var aba = planilha.getSheets().filter(function (s) { return s.getSheetId() === FATURAMENTO_GID_; })[0];
  if (!aba) throw new Error('Aba de Faturamento (gid ' + FATURAMENTO_GID_ + ') não encontrada na planilha.');
  var dados = aba.getDataRange().getValues();
  if (!dados.length) return { medicoes: [], linhasBrutas: 0, linhasIgnoradas: 0 };
  var cabecalho = dados[0];
  var idx = {
    medicao: cabecalho.indexOf('Medicao'),
    dataEmissao: cabecalho.indexOf('Data de Emissao'),
    valorTotal: cabecalho.indexOf('Valor Total (R$)'),
    valorLiquido: cabecalho.indexOf('Valor Liquido (R$)'),
  };
  if (idx.medicao === -1 || idx.valorTotal === -1) {
    throw new Error('Cabeçalho da planilha de Faturamento mudou — colunas "Medicao"/"Valor Total (R$)" não encontradas.');
  }

  // Conta quantas linhas de dado existem e quantas foram ignoradas (sem
  // Medicao, ou "Valor Total (R$)" não numérico — célula formatada como
  // texto, por exemplo) — sem isso, uma planilha com linhas mas nenhuma
  // aproveitável vira silenciosamente medicoes:[], indistinguível de uma
  // planilha realmente vazia do lado do front-end.
  var linhasBrutas = 0;
  var linhasIgnoradas = 0;
  var porMedicao = {};
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var medicao = linha[idx.medicao];
    var valorTotal = linha[idx.valorTotal];
    var linhaEmBranco = linha.every(function (v) { return v === '' || v == null; });
    if (linhaEmBranco) continue;
    linhasBrutas++;
    if (medicao === '' || medicao == null || typeof valorTotal !== 'number') { linhasIgnoradas++; continue; }
    var chave = String(medicao);
    if (!porMedicao[chave]) porMedicao[chave] = { medicao: medicao, valorTotal: 0, valorLiquido: 0, datas: [] };
    porMedicao[chave].valorTotal += valorTotal;
    if (typeof linha[idx.valorLiquido] === 'number') porMedicao[chave].valorLiquido += linha[idx.valorLiquido];
    var dataEmissao = paraDataFaturamento_(linha[idx.dataEmissao]);
    if (dataEmissao) porMedicao[chave].datas.push(dataEmissao.getTime());
  }

  var medicoes = Object.keys(porMedicao).map(function (k) {
    var g = porMedicao[k];
    var dataRef = g.datas.length ? new Date(Math.max.apply(null, g.datas)) : null;
    return {
      medicao: g.medicao,
      valorTotal: g.valorTotal,
      valorLiquido: g.valorLiquido,
      dataReferencia: dataRef ? Utilities.formatDate(dataRef, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null,
    };
  }).sort(function (a, b) { return (a.dataReferencia || '').localeCompare(b.dataReferencia || ''); });

  return { medicoes: medicoes, linhasBrutas: linhasBrutas, linhasIgnoradas: linhasIgnoradas };
}

function tentarFaturamentoPorMedicao_() {
  try {
    var r = faturamentoPorMedicao_();
    return { medicoes: r.medicoes, linhasBrutas: r.linhasBrutas, linhasIgnoradas: r.linhasIgnoradas, erro: null };
  } catch (erro) {
    return { medicoes: [], linhasBrutas: 0, linhasIgnoradas: 0, erro: erro.message };
  }
}

// Ação própria (não dentro de api_carregar) de propósito: buscar a
// planilha de Faturamento é uma chamada externa (SpreadsheetApp.openById
// numa planilha diferente desta) — embuti-la em api_carregar faria TODA
// ação de escrita (importar, lançar manual, remover…) pagar esse custo
// extra de novo a cada vez, mesmo sem precisar do resultado. Como
// dado, o Faturamento muda raramente — o Dashboard busca 1x ao carregar.
function api_faturamentoPorMedicao(payload) {
  return tentarFaturamentoPorMedicao_();
}

function api_importar(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  substituirErp_(payload.itens, payload.dataBase);
  // payload.driveChave só vem preenchido quando esta importação veio do
  // Google Drive (ver api_buscarArquivoDriveNovo) — marca o arquivo como já
  // processado só agora, depois que a gravação deu certo, pra não perder um
  // arquivo se a importação falhar no meio do caminho.
  if (payload.driveChave) {
    PropertiesService.getScriptProperties().setProperty(PROP_ULTIMO_ARQUIVO_DRIVE_, payload.driveChave);
  }
  registrarLog_(nomeUsuario, 'Importar (' + (payload.driveChave ? 'Google Drive' : 'upload') + ')',
    (payload.itens || []).length + ' título(s), base ' + (payload.dataBase || 'não identificada'));
  return api_carregar();
}

// Pasta do Drive com o export do Protheus (Títulos/baixas) — o app checa
// aqui, na aba Extrator de Comprovante, se existe um .xlsx mais novo que o
// último já importado por este caminho (qualquer nome de arquivo, ver
// getFilesByType). Não baixa/processa a planilha aqui: só devolve o
// conteúdo em base64 pro navegador ler com o mesmo parser client-side já
// usado no upload manual (processarArquivo/parseArquivoProtheus, ver
// Index.html) — reaproveita 100% da lógica de leitura/normalização que já
// existia, em vez de duplicá-la aqui no servidor.
var PASTA_DRIVE_TITULOS_ID_ = '1sVlF29VGWDzHelgBGIeFvVK3OCpMjGmD';
var PROP_ULTIMO_ARQUIVO_DRIVE_ = 'drive_ultimo_arquivo_titulos';

// Percorre a pasta e devolve o arquivo .xlsx com getLastUpdated() mais
// recente, ou null se a pasta não tiver nenhum — usado tanto pela checagem
// de importação (api_buscarArquivoDriveNovo) quanto pelo download avulso
// (api_baixarArquivoDriveTitulos), que precisam do MESMO critério de "qual
// é o arquivo vigente" pra nunca divergir um do outro.
function arquivoDriveTitulosMaisRecente_() {
  var pasta = DriveApp.getFolderById(PASTA_DRIVE_TITULOS_ID_);
  var arquivos = pasta.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var maisRecente = null;
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    if (!maisRecente || arquivo.getLastUpdated().getTime() > maisRecente.getLastUpdated().getTime()) {
      maisRecente = arquivo;
    }
  }
  return maisRecente;
}

function api_buscarArquivoDriveNovo(payload) {
  validarAcessoEdicao_(payload);
  var maisRecente = arquivoDriveTitulosMaisRecente_();
  if (!maisRecente) return { encontrado: false };

  var chaveAtual = maisRecente.getId() + '|' + maisRecente.getLastUpdated().getTime();
  var ultimaChave = PropertiesService.getScriptProperties().getProperty(PROP_ULTIMO_ARQUIVO_DRIVE_);
  if (chaveAtual === ultimaChave) return { encontrado: false };

  var blob = maisRecente.getBlob();
  return {
    encontrado: true,
    nomeArquivo: maisRecente.getName(),
    conteudoBase64: Utilities.base64Encode(blob.getBytes()),
    driveChave: chaveAtual,
  };
}

// Baixa o mesmo arquivo que api_buscarArquivoDriveNovo usaria pra
// atualizar os dados — mas sem a checagem de "já importado" (sempre
// devolve o mais recente da pasta, já importado ou não) e sem marcar nada
// como processado, já que aqui é só um download, não uma importação.
function api_baixarArquivoDriveTitulos(payload) {
  validarAcessoEdicao_(payload);
  var maisRecente = arquivoDriveTitulosMaisRecente_();
  if (!maisRecente) throw new Error('Nenhum arquivo encontrado na pasta do Google Drive.');
  var blob = maisRecente.getBlob();
  return {
    nomeArquivo: maisRecente.getName(),
    conteudoBase64: Utilities.base64Encode(blob.getBytes()),
  };
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

// Grava o link de anexo (documento ou comprovante) de um lançamento —
// chamada tanto quando a pessoa cola o link à mão quanto quando a busca
// automática no Drive (ver api_buscarAnexoDrive) já encontra e salva
// sozinha. payload: { id, origem: 'ERP'|'Manual', tipo: 'documento'|
// 'comprovante', url }.
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

// Observação livre por lançamento — mesma lógica de api_definirAnexo (só
// grava 1 coluna, funciona pra ERP e Manual) mas pra texto livre em vez de
// link. Sobrevive a reimportação do ERP (ver substituirErp_).
function api_definirObservacao(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var origem = payload.origem === 'Manual' ? 'Manual' : 'ERP';
  var colIdx = HEADERS.indexOf('Observação') + 1;
  var idCol = HEADERS.indexOf('ID') + 1;
  var docCol = HEADERS.indexOf('Nº Documento') + 1;
  var razaoCol = HEADERS.indexOf('Razão Social') + 1;

  var sh = getSheet_(origem);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) throw new Error('Lançamento não encontrado.');
  var dados = sh.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) {
      sh.getRange(i + 2, colIdx).setValue(String(payload.valor || ''));
      registrarLog_(nomeUsuario, 'Definir observação',
        (dados[i][razaoCol - 1] || '') + ' Nº ' + (dados[i][docCol - 1] || '') + (payload.valor ? ': ' + payload.valor : ' (removida)'));
      return api_carregar();
    }
  }
  throw new Error('Lançamento não encontrado (pode ter sido removido ou reimportado).');
}

// Categoria do fornecedor definida manualmente — usada pela tela só quando
// o "de-para" fixo por Código Fornecedor (mantido no front-end) não conhece
// o fornecedor E o texto do Histórico também não dá pra usar como
// aproximação. Mesma lógica de api_definirObservacao, mas com um passo a
// mais: ao gravar uma categoria (não ao remover), propaga o mesmo valor
// pros demais títulos do mesmo Código Fornecedor — no ERP e no Manual,
// tanto em aberto quanto já baixados — pra não obrigar a digitar de novo
// em cada linha do mesmo fornecedor. Sobrevive a reimportação do ERP (ver
// substituirErp_).
function api_definirCategoria(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var origem = payload.origem === 'Manual' ? 'Manual' : 'ERP';
  var colIdx = HEADERS.indexOf('Categoria') + 1;
  var idCol = HEADERS.indexOf('ID') + 1;
  var docCol = HEADERS.indexOf('Nº Documento') + 1;
  var razaoCol = HEADERS.indexOf('Razão Social') + 1;
  var fornCol = HEADERS.indexOf('Código Fornecedor') + 1;

  var sh = getSheet_(origem);
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) throw new Error('Lançamento não encontrado.');
  var dados = sh.getRange(2, 1, ultimaLinha - 1, HEADERS.length).getValues();
  var linhaAlvo = -1;
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) { linhaAlvo = i; break; }
  }
  if (linhaAlvo === -1) throw new Error('Lançamento não encontrado (pode ter sido removido ou reimportado).');

  var valor = String(payload.valor || '');
  sh.getRange(linhaAlvo + 2, colIdx).setValue(valor);

  var qtdPropagada = 0;
  if (valor) {
    var fornecedorAlvo = removerZerosEsquerda_(String(dados[linhaAlvo][fornCol - 1] || '').trim());
    if (fornecedorAlvo) {
      ['ERP', 'Manual'].forEach(function (abaOrigem) {
        var mesmaAba = abaOrigem === origem;
        var shOrigem = mesmaAba ? sh : getSheet_(abaOrigem);
        var ultimaLinhaOrigem = shOrigem.getLastRow();
        if (ultimaLinhaOrigem < 2) return;
        var dadosOrigem = mesmaAba ? dados : shOrigem.getRange(2, 1, ultimaLinhaOrigem - 1, HEADERS.length).getValues();
        for (var j = 0; j < dadosOrigem.length; j++) {
          if (mesmaAba && j === linhaAlvo) continue; // já gravado acima
          var fornecedorLinha = removerZerosEsquerda_(String(dadosOrigem[j][fornCol - 1] || '').trim());
          if (fornecedorLinha !== fornecedorAlvo) continue;
          shOrigem.getRange(j + 2, colIdx).setValue(valor);
          qtdPropagada++;
        }
      });
    }
  }

  registrarLog_(nomeUsuario, 'Definir categoria',
    (dados[linhaAlvo][razaoCol - 1] || '') + ' Nº ' + (dados[linhaAlvo][docCol - 1] || '') + (valor ? ': ' + valor : ' (removida)') +
    (qtdPropagada ? ' — aplicada também a mais ' + qtdPropagada + ' título(s) do mesmo fornecedor' : ''));
  return api_carregar();
}

// Pastas do Drive onde os documentos/comprovantes recebidos (e-mail,
// WhatsApp etc.) são organizados manualmente — as mesmas duas pastas usadas
// pela rotina de importação de notas (ver skill controle-financeiro-pdf-import):
// "03 DOC FISCAL FATURA E RECIBO" e "CE-007 - ANALISE DE NOTAS -> GILBERTO".
// IDs fixos (não por nome) pra não depender de busca por texto, que falha
// se a pasta for renomeada — se algum dia mudarem de pasta, é só trocar o
// ID aqui.
var PASTAS_BUSCA_ANEXO = [
  '1srFnPfX6BCMYEShfDlF6H7nSiJD4aeef', // 03 DOC FISCAL FATURA E RECIBO (documentos — NF, FAT, BOL...)
  '18nuIj4SsAEqPJ7jOxlUe8JzJ8TCNcueL', // CE-007 - ANALISE DE NOTAS -> GILBERTO (documentos)
  '1leuDOfqLFDxdg3PECZ6aij2eIkPGKwq6', // 16 COMPROVANTE PAGTO, organizada por ano/mês (comprovantes de pagamento)
];

function normalizarTextoBusca_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // tira acento (Ç, ã, é...)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

// Confere se um arquivo está DENTRO de uma das pastas em PASTAS_BUSCA_ANEXO
// (em qualquer nível — ex.: 16 COMPROVANTE PAGTO/2026/07 2026/arquivo.pdf),
// subindo pela cadeia de pastas-pai até achar uma bater ou esgotar os
// níveis (a estrutura real tem no máximo 2-3 níveis; 6 é só uma folga de
// segurança). Só considera o 1º pai de cada nível — arquivo com mais de
// uma pasta-pai é raríssimo nessa estrutura organizada por ano/mês.
function arquivoDentroDasPastasAlvo_(arquivo) {
  var atual = arquivo;
  for (var nivel = 0; nivel < 6; nivel++) {
    var pais = atual.getParents();
    if (!pais.hasNext()) return false;
    var pai = pais.next();
    if (PASTAS_BUSCA_ANEXO.indexOf(pai.getId()) !== -1) return true;
    atual = pai;
  }
  return false;
}

// Busca no Drive um documento/comprovante pra um título — usada pelo botão
// "Buscar no Drive" do pop-up de anexo (só sugere; quem decide se aplica é
// sempre a pessoa, nunca grava nada sozinho). Os arquivos dessas pastas
// costumam seguir o padrão de nome "TIPO Nº_DOCUMENTO FORNECEDOR ...pdf"
// (ex.: "NF 351 RENOVE DISTRIBUIDORA DE MATERIAS.pdf"), mas nem todo
// arquivo real respeita essa ordem — por isso a condição mínima é, agora,
// independente de posição: exige o Nº Documento bater como palavra inteira
// EM QUALQUER LUGAR do nome E o Código Fornecedor aparecer em qualquer
// lugar do nome também. Achando os dois, já é suficiente pra sugerir o
// arquivo. Só quando o Código Fornecedor não está disponível pra comparar
// (raro) é que cai de volta pra exigir pelo menos uma palavra de 4+ letras
// da Razão Social no nome, pra não sugerir qualquer arquivo que só por
// acaso tenha o mesmo número.
//
// Usa DriveApp.searchFiles() (o índice de busca do Drive) em vez de varrer
// pasta por pasta — a 1ª versão fazia isso manualmente (getFolders/
// getFilesByType recursivo) e ficava lenta/travada demais pra ser usável:
// "16 COMPROVANTE PAGTO" sozinha tem dezenas de subpastas por ano/mês, cada
// uma com muitos arquivos. A busca aqui é ampla (Drive inteiro, pelo nome
// conter o Nº Documento) só pra achar candidatos rápido; o filtro de
// pontuação abaixo, MAIS a conferência de que o arquivo realmente está
// dentro de uma das pastas conhecidas (arquivoDentroDasPastasAlvo_), é que
// garante que só aparece o que é de verdade um documento/comprovante desse
// título — nunca mostra nada de fora dessas pastas.
var PREFIXOS_COMPROVANTE = ['COMPROVANTE', 'REC'];

// Código Fornecedor no Protheus é sempre um número de 8 ou 9 dígitos (com
// zero(s) à esquerda quando precisa) — mas o app guarda esse código já SEM
// esses zeros (removerZerosEsquerda_, aplicado na importação — ver
// parseArquivoProtheus no Index.html), enquanto o nome do arquivo no Drive
// tem o código como o Protheus mostra de verdade, COM os zeros. Pra achar
// de qualquer jeito, testa as duas larguras possíveis (8 e 9) em vez de só
// o código sem zeros — devolve 1 ou 2 strings, sempre com 8 ou 9 dígitos.
function paddingsFornecedor_(codigo) {
  var digitos = String(codigo || '').replace(/\D/g, '');
  if (!digitos) return [];
  var resultado = [];
  [8, 9].forEach(function (largura) {
    if (digitos.length > largura) return;
    var comZeros = ('000000000' + digitos).slice(-largura);
    if (resultado.indexOf(comZeros) === -1) resultado.push(comZeros);
  });
  if (!resultado.length) resultado.push(digitos); // código com mais de 9 dígitos — caso raro, usa como está
  return resultado;
}

function api_buscarAnexoDrive(payload) {
  validarAcessoEdicao_(payload);
  var nDocOriginal = removerZerosEsquerda_(String(payload.nDocumento || '')).trim();
  var nDocAlvo = normalizarTextoBusca_(nDocOriginal);
  if (!nDocAlvo) return { candidatos: [] };
  var codigosFornecedor = paddingsFornecedor_(payload.codigoFornecedor); // ex.: ["64781522"] ou ["00005678","000005678"]
  // Sem Código Fornecedor pra combinar na busca, um Nº Documento bem curto
  // (1-2 caracteres) sozinho bateria com uma fração enorme do Drive — nem
  // tenta, pra não travar sem achar nada de útil mesmo.
  if (!codigosFornecedor.length && nDocAlvo.length < 3) return { candidatos: [] };
  var palavrasFornecedor = normalizarTextoBusca_(payload.razaoSocial).split(' ')
    .filter(function (p) { return p.length >= 4; });
  var buscandoComprovante = payload.tipo === 'comprovante';

  // Combina Nº Documento + Código Fornecedor (nas duas larguras possíveis,
  // 8 e 9 dígitos) na busca sempre que o código estiver disponível (é o
  // caso normal) — os dois juntos aparecem em TODO nome de arquivo dessas
  // pastas (confirmado pelo Anderson: "NFS 1 LDA ENGENHARIA 64781522"),
  // então isso deixa a busca bem mais seletiva. Sem o código, um Nº
  // Documento curto e comum sozinho (ex.: "1", "2") bate com uma fração
  // enorme do Drive inteiro (qualquer nome com esse dígito em algum lugar
  // — data, quantidade etc.) e a busca nunca terminava de forma
  // perceptível.
  var escapar_ = function (v) { return String(v).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); };
  var query = "mimeType = 'application/pdf' and trashed = false and title contains '" + escapar_(nDocOriginal) + "'";
  if (codigosFornecedor.length) {
    query += ' and (' + codigosFornecedor.map(function (c) { return "title contains '" + escapar_(c) + "'"; }).join(' or ') + ')';
  }
  var resultadosBrutos;
  try {
    resultadosBrutos = DriveApp.searchFiles(query);
  } catch (e) {
    return { candidatos: [], erro: 'Não foi possível buscar no Drive: ' + e.message };
  }

  var candidatos = [];
  var verificados = 0;
  // Trava de segurança pra nunca processar um número absurdo de resultados
  // (ex.: um Nº Documento bem curto e genérico batendo com meio Drive) —
  // suficiente pra qualquer busca real: o filtro de pontuação abaixo já é
  // rigoroso, então os candidatos de verdade aparecem bem antes disso.
  while (resultadosBrutos.hasNext() && verificados < 300) {
    var arquivo = resultadosBrutos.next();
    verificados++;
    if (!arquivoDentroDasPastasAlvo_(arquivo)) continue;
    var nome = arquivo.getName();
    var nomeNorm = normalizarTextoBusca_(nome);
    var tokens = nomeNorm.split(' ');
    // Nº Documento como palavra inteira em QUALQUER posição do nome — não
    // mais só logo depois do Tipo/COMPROVANTE (tokens[0]), pois vários
    // arquivos reais não seguem essa ordem e a busca deixava de achar o
    // link. Ainda exige bater a palavra inteira (não substring solto), pra
    // não confundir com o mesmo número aparecendo por coincidência dentro
    // de outro campo do nome (ex.: dentro do próprio Código Fornecedor).
    if (tokens.indexOf(nDocAlvo) === -1) continue;
    var pontuacao = 1;
    var codigoBate = codigosFornecedor.some(function (c) { return nomeNorm.indexOf(c) !== -1; });
    if (codigoBate) pontuacao += 2;
    pontuacao += palavrasFornecedor.filter(function (p) { return nomeNorm.indexOf(p) !== -1; }).length;
    // Condição mínima: Nº Documento + Código Fornecedor no nome (em
    // qualquer posição) já basta. Sem código pra comparar, cai pra exigir
    // ao menos o nome do fornecedor (regra anterior).
    if (!codigoBate && pontuacao < 2) continue;
    // Cada botão só pode sugerir o arquivo do tipo certo pra ele — nunca os
    // dois juntos (documento e comprovante são coisas diferentes, mesmo
    // quando os dois existem pro mesmo título): comprovante só aceita
    // arquivo que comece com um dos PREFIXOS_COMPROVANTE; documento só
    // aceita o que NÃO comece com um deles (a 1ª palavra vira o Tipo do
    // título — NF, FAT, BOL... — então checar "não é comprovante" cobre
    // qualquer Tipo, sem precisar bater exatamente com o Tipo salvo nessa
    // linha, que pode variar).
    var primeiroToken = tokens[0];
    var ehComprovante = PREFIXOS_COMPROVANTE.indexOf(primeiroToken) !== -1;
    if (buscandoComprovante !== ehComprovante) continue;
    candidatos.push({ nome: nome, url: arquivo.getUrl(), pontuacao: pontuacao });
  }

  candidatos.sort(function (a, b) { return b.pontuacao - a.pontuacao; });
  return { candidatos: candidatos.slice(0, 5) };
}

// Campos que o modal de detalhamento pode alterar num lançamento MANUAL —
// tudo, exceto ID/Origem/Status (Status é sempre recalculado depois, nunca
// escrito direto pelo cliente). Só existe pra Manual: um título de ERP é
// sobrescrito por inteiro na próxima importação diária (substituirErp_),
// então uma edição feita aqui nele seria perdida sem aviso nenhum — por
// isso a tela nem oferece "Editar" pra linhas de Origem ERP.
var CAMPOS_EDITAVEIS_MANUAL = HEADERS.filter(function (h) {
  return h !== 'ID' && h !== 'Origem' && h !== 'Status';
});

function api_editarManual(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var shM = getSheet_('Manual');
  var idCol = HEADERS.indexOf('ID') + 1;
  var ultimaLinha = shM.getLastRow();
  var dados = shM.getRange(2, 1, Math.max(ultimaLinha - 1, 0), HEADERS.length).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][idCol - 1] === payload.id) {
      var linhaAtual = dados[i];
      var item = payload.item || {};
      var linhaNova = HEADERS.map(function (h, idx) {
        if (CAMPOS_EDITAVEIS_MANUAL.indexOf(h) === -1) return linhaAtual[idx];
        var novoValor = item[h];
        return novoValor != null ? novoValor : linhaAtual[idx];
      });
      shM.getRange(i + 2, 1, 1, HEADERS.length).setValues([linhaNova]);
      var razaoCol = HEADERS.indexOf('Razão Social');
      var docCol = HEADERS.indexOf('Nº Documento');
      registrarLog_(nomeUsuario, 'Editar lançamento manual',
        (linhaNova[razaoCol] || '') + ' | Nº ' + (linhaNova[docCol] || ''));
      recalcularStatus_();
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
    else if (acao === 'buscarArquivoDriveNovo') resultado = api_buscarArquivoDriveNovo(payload);
    else if (acao === 'baixarArquivoDriveTitulos') resultado = api_baixarArquivoDriveTitulos(payload);
    else if (acao === 'adicionarManual') resultado = api_adicionarManual(payload);
    else if (acao === 'editarManual') resultado = api_editarManual(payload);
    else if (acao === 'removerManual') resultado = api_removerManual(payload);
    else if (acao === 'removerDuplicados') resultado = api_removerDuplicados(payload);
    else if (acao === 'definirAnexo') resultado = api_definirAnexo(payload);
    else if (acao === 'buscarAnexoDrive') resultado = api_buscarAnexoDrive(payload);
    else if (acao === 'definirObservacao') resultado = api_definirObservacao(payload);
    else if (acao === 'definirCategoria') resultado = api_definirCategoria(payload);
    else if (acao === 'validarEdicao') resultado = api_validarEdicao(payload);
    else if (acao === 'carregarLog') resultado = api_carregarLog(payload);
    else if (acao === 'salvarConciliacao') resultado = api_salvarConciliacao(payload);
    else if (acao === 'faturamentoPorMedicao') resultado = api_faturamentoPorMedicao(payload);
    else if (acao === 'carregarRenomeacoes') resultado = api_carregarRenomeacoes(payload);
    else if (acao === 'aprovarRenomeacao') resultado = api_aprovarRenomeacao(payload);
    else if (acao === 'ignorarRenomeacao') resultado = api_ignorarRenomeacao(payload);
    else if (acao === 'identificarRenomeacoesAgora') resultado = api_identificarRenomeacoesAgora(payload);
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
    linhas.push('⚠ A base de dados está desatualizada há ' + diasDefasagem + ' dia(s) — importe um arquivo mais recente na aba "Atualizar Dados".');
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

// ---------- Padronização de nomes de arquivo (Drive) ----------
// Mantém os PDFs das pastas de PASTAS_BUSCA_ANEXO (documentos e comprovantes)
// no padrão "TIPO Nº_DOCUMENTO RAZÃO_SOCIAL CÓDIGO_FORNECEDOR.pdf" (ou
// "COMPROVANTE ..." na pasta de comprovantes) — o mesmo padrão que
// api_buscarAnexoDrive já depende para achar anexo automático. Roda 100%
// dentro do Google (Drive API + regex, sem chamar nenhuma IA/LLM): a
// extração de texto usa a conversão PDF -> Google Doc com OCR do serviço
// avançado do Drive (grátis, nativa do Apps Script), e a leitura de
// TIPO/Nº/Razão Social/CNPJ-CPF é feita com expressões regulares sobre esse
// texto. NENHUM arquivo é renomeado sozinho: cada sugestão cai na aba
// "Renomear Pendente" e só é aplicada de fato quando alguém marca a coluna
// "Aprovar" (aplicarRenomeacoesAprovadas cuida da aplicação) — evita
// renomear errado um documento de verdade só por causa de uma extração
// ambígua (razões sociais e leiautes de nota variam demais entre emissores).
//
// Uso (uma vez, no editor do Apps Script):
//  1. identificarArquivosForaDoPadrao — roda a varredura e enche a aba
//     "Renomear Pendente" com sugestões (repita algumas vezes se o
//     resultado disser "..." pra esgotar o acervo já existente — cada
//     execução processa um lote limitado pra não estourar o tempo máximo
//     de execução do Apps Script).
//  2. Revise a aba "Renomear Pendente" na planilha, corrija a coluna "Nome
//     Sugerido" quando a extração errou algo, e marque "Aprovar" (TRUE) nas
//     linhas que pode aplicar.
//  3. aplicarRenomeacoesAprovadas — renomeia de fato os arquivos aprovados.
//  4. configurarGatilhosRenomeacao — instala os gatilhos de tempo pra rodar
//     os passos 1 e 3 sozinhos daqui pra frente (identifica 1x/dia; aplica
//     de hora em hora), sem precisar abrir o editor de novo.

var RENOMEAR_HEADERS = [
  'Data Detecção', 'Pasta', 'ID Arquivo', 'Nome Atual', 'Nome Sugerido',
  'Confiança', 'CNPJ/CPF Encontrado', 'Aprovar', 'Status', 'Observação',
];
var RENOMEAR_LIMITE_OCR_POR_EXECUCAO = 25; // teto de conversões OCR por execução, pra nunca estourar o tempo máximo de um gatilho

// Nome já no padrão -> TIPO/COMPROVANTE conhecido, seguido de algo, seguido
// de um código de 8 ou 9 dígitos colado no ".pdf" — casa com todos os
// exemplos reais (ex.: "NFS 24610 COPHEL EXPRESS 26846738.pdf",
// "COMPROVANTE 1793 VOUDLOG TRANS 037295834.pdf"). Arquivo que já bate aqui
// nunca precisa de OCR — só os fora desse padrão entram na fila.
var REGEX_NOME_PADRAO_ = /^(COMPROVANTE|NF|NFS|NF3E|NFAG|DACTE|ND|BOL|FAT|FOL|FGTS|DARF|DAM|REC)\s+\S.*\s(\d{8,9})\.pdf$/i;

function getSheetRenomear_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Renomear Pendente');
  if (!sh) {
    sh = ss.insertSheet('Renomear Pendente');
    sh.appendRow(RENOMEAR_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// IDs já presentes na fila (pendente ou já renomeado) — evita reprocessar/
// reenfileirar o mesmo arquivo em toda execução do gatilho diário.
function idsJaNaFilaRenomear_() {
  var sh = getSheetRenomear_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return {};
  var idCol = RENOMEAR_HEADERS.indexOf('ID Arquivo') + 1;
  var valores = sh.getRange(2, idCol, ultimaLinha - 1, 1).getValues();
  var set = {};
  valores.forEach(function (l) { if (l[0]) set[l[0]] = true; });
  return set;
}

// As 3 pastas de PASTAS_BUSCA_ANEXO ficam dentro de uma estrutura
// compartilhada (Drive compartilhado/Team Drive, não "Meu Drive") — sem
// supportsTeamDrives + includeTeamDriveItems, Drive.Files.list() devolve
// SEMPRE 0 resultado pra qualquer pasta aí dentro, mesmo com a query certa
// (comportamento silencioso da API v2: não dá erro, só nunca acha nada).
// Confirmado pela URL real de "03 DOC FISCAL FATURA E RECIBO":
// .../folders/1srFnPfX6.../ com o breadcrumb "Compartilha... > 4 - CE
// 007_A... > 4.11 - FIN...", que é o padrão de pastas dentro de um Drive
// compartilhado. Repassado em toda chamada Drive.Files.list() abaixo.
var OPCOES_TEAM_DRIVE_ = { supportsTeamDrives: true, includeTeamDriveItems: true };

// Lista (id, subpastas incluídas) de todas as subpastas dentro de uma pasta
// raiz, via Drive API (mais rápido que DriveApp.getFolders() recursivo pra
// árvores grandes, como a "16 COMPROVANTE PAGTO" organizada por ano/mês).
function listarPastaEsubpastas_(pastaRaizId) {
  var pastas = [pastaRaizId];
  var fila = [pastaRaizId];
  while (fila.length) {
    var atual = fila.shift();
    var pageToken = null;
    do {
      var resp = Drive.Files.list(Object.assign({
        q: "'" + atual + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        maxResults: 200,
        pageToken: pageToken,
        fields: 'items(id),nextPageToken',
      }, OPCOES_TEAM_DRIVE_));
      (resp.items || []).forEach(function (f) { pastas.push(f.id); fila.push(f.id); });
      pageToken = resp.nextPageToken;
    } while (pageToken);
  }
  return pastas;
}

function listarPdfsDaPasta_(pastaId) {
  var arquivos = [];
  var pageToken = null;
  do {
    var resp = Drive.Files.list(Object.assign({
      q: "'" + pastaId + "' in parents and mimeType = 'application/pdf' and trashed = false",
      maxResults: 200,
      pageToken: pageToken,
      fields: 'items(id,title),nextPageToken',
    }, OPCOES_TEAM_DRIVE_));
    (resp.items || []).forEach(function (f) { arquivos.push({ id: f.id, nome: f.title }); });
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return arquivos;
}

// Converte o PDF num Google Doc temporário com OCR (serviço avançado do
// Drive) só pra extrair o texto, e sempre apaga o Doc temporário depois —
// funciona tanto pra PDF "nativo" (nota fiscal gerada em PDF) quanto pra PDF
// escaneado/print de tela (comprovante de banco, comum nas pastas
// monitoradas). Não usa nenhuma IA/LLM — é OCR nativo do Google Drive.
function extrairTextoPdfOcr_(idArquivo, nomeArquivo) {
  var recurso = {
    title: 'OCR temporário — ' + nomeArquivo,
    mimeType: MimeType.GOOGLE_DOCS,
  };
  var arquivoOriginal = DriveApp.getFileById(idArquivo);
  var docConvertido = Drive.Files.insert(recurso, arquivoOriginal.getBlob(),
    Object.assign({ ocr: true, ocrLanguage: 'pt' }, OPCOES_TEAM_DRIVE_));
  try {
    return DocumentApp.openById(docConvertido.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(docConvertido.id, OPCOES_TEAM_DRIVE_); } catch (e) { /* Doc temporário órfão, sem risco — só ocupa um pouco de Lixeira */ }
  }
}

// Extrai CNPJ (14 dígitos -> código = 8 primeiros) ou, na ausência de CNPJ,
// CPF (11 dígitos -> código = 9 primeiros) do texto do documento. Procura
// primeiro perto do rótulo "CNPJ"/"CPF" (mais confiável) e só cai pro
// primeiro número no formato certo em qualquer lugar do texto como reforço.
function extrairCnpjCpf_(texto) {
  var t = String(texto || '');
  var mCnpjRotulo = t.match(/CNPJ[^\d]{0,10}(\d[\d.\/-]{12,18}\d)/i);
  var mCnpjSolto = t.match(/\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}\b/);
  var candidatoCnpj = mCnpjRotulo ? mCnpjRotulo[1] : (mCnpjSolto ? mCnpjSolto[0] : null);
  if (candidatoCnpj) {
    var digitosCnpj = candidatoCnpj.replace(/\D/g, '');
    if (digitosCnpj.length === 14) {
      return { tipo: 'CNPJ', formatado: digitosCnpj, codigo: digitosCnpj.slice(0, 8) };
    }
  }
  var mCpfRotulo = t.match(/CPF[^\d]{0,10}(\d[\d.\-]{9,12}\d)/i);
  if (mCpfRotulo) {
    var digitosCpf = mCpfRotulo[1].replace(/\D/g, '');
    if (digitosCpf.length === 11) {
      return { tipo: 'CPF', formatado: digitosCpf, codigo: digitosCpf.slice(0, 9) };
    }
  }
  return null;
}

// Palavras-chave -> Tipo (mesmos prefixos usados na Conciliação Bancária e
// aceitos por api_buscarAnexoDrive). Testadas em ordem — a primeira que
// aparecer no texto do documento decide o Tipo.
var PALAVRAS_TIPO_DOCUMENTO_ = [
  [/NOTA\s+FISCAL\s+DE\s+SERVI[ÇC]OS|NFS-?E/i, 'NFS'],
  [/NF3E|NOTA\s+FISCAL\s+DE\s+ENERGIA/i, 'NF3E'],
  [/CT-?E|CONHECIMENTO\s+DE\s+TRANSPORTE|DACTE/i, 'DACTE'],
  [/BOLETO|FICHA\s+DE\s+COMPENSA[ÇC][ÃA]O/i, 'BOL'],
  [/FATURA/i, 'FAT'],
  [/NOTA\s+DE\s+D[ÉE]BITO/i, 'ND'],
  [/FOLHA\s+DE\s+PAGAMENTO/i, 'FOL'],
  [/GUIA\s+DO\s+FGTS|FGTS/i, 'FGTS'],
  [/DARF/i, 'DARF'],
  [/\bDAM\b/i, 'DAM'],
  [/RECIBO/i, 'REC'],
  [/NOTA\s+FISCAL/i, 'NF'],
];

function inferirTipoDocumento_(texto, ehPastaComprovante) {
  if (ehPastaComprovante) return 'COMPROVANTE'; // pasta "16 COMPROVANTE PAGTO" é sempre comprovante de pagamento
  for (var i = 0; i < PALAVRAS_TIPO_DOCUMENTO_.length; i++) {
    if (PALAVRAS_TIPO_DOCUMENTO_[i][0].test(texto)) return PALAVRAS_TIPO_DOCUMENTO_[i][1];
  }
  return null;
}

// Rótulos comuns em NFS-e/boleto/comprovante pra achar o Nº do documento —
// tenta cada um, na ordem, e usa o primeiro que achar um número junto.
var ROTULOS_NUMERO_DOCUMENTO_ = [
  /N[úu]mero\s+da\s+Nota[^\d]{0,10}(\d{1,15})/i,
  /N[°ºo]\s*\.?\s*da\s+Nota[^\d]{0,10}(\d{1,15})/i,
  /Nota\s+Fiscal\s+n[°ºo]?\.?[^\d]{0,10}(\d{1,15})/i,
  /Fatura\s+n[°ºo]?\.?[^\d]{0,10}(\d{1,15})/i,
  /Nosso\s+N[úu]mero[^\d]{0,10}(\d{1,15})/i,
  /N[°ºo]\s*\.?\s*Documento[^\d]{0,10}(\d{1,15})/i,
  /N[°ºo]\s*\.?\s*do\s+Documento[^\d]{0,10}(\d{1,15})/i,
];

function inferirNumeroDocumento_(texto) {
  for (var i = 0; i < ROTULOS_NUMERO_DOCUMENTO_.length; i++) {
    var m = texto.match(ROTULOS_NUMERO_DOCUMENTO_[i]);
    if (m) return removerZerosEsquerda_(m[1]);
  }
  return null;
}

// Rótulos comuns pra achar a razão social/nome de quem emitiu ou recebeu —
// pega o resto da linha depois do rótulo, corta em tamanho razoável e limpa
// caracteres que não fazem sentido num nome de arquivo.
var ROTULOS_RAZAO_SOCIAL_ = [
  /Prestador(?:\s+de\s+Servi[çc]os)?[:\s]+([^\n]{3,80})/i,
  /Raz[ãa]o\s+Social[:\s]+([^\n]{3,80})/i,
  /Nome\s*\/\s*Raz[ãa]o\s+Social[:\s]+([^\n]{3,80})/i,
  /Emitente[:\s]+([^\n]{3,80})/i,
  /Benefici[áa]rio[:\s]+([^\n]{3,80})/i,
  /Favorecido[:\s]+([^\n]{3,80})/i,
];

function inferirRazaoSocial_(texto) {
  for (var i = 0; i < ROTULOS_RAZAO_SOCIAL_.length; i++) {
    var m = texto.match(ROTULOS_RAZAO_SOCIAL_[i]);
    if (m) {
      var nome = m[1]
        .replace(/CNPJ.*$/i, '').replace(/CPF.*$/i, '')
        .replace(/[\\\/:*?"<>|]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (nome.length >= 3) return nome.slice(0, 60).toUpperCase();
    }
  }
  return null;
}

function sanitizarNomeArquivo_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
}

// Monta a sugestão completa pra um arquivo fora do padrão: OCR + regex, sem
// nenhuma chamada a IA/LLM. "Confiança: alta" só quando TIPO + Nº Documento
// + Razão Social + CNPJ/CPF foram todos identificados sem ambiguidade — o
// resto cai como "revisar" pra alguém completar/corrigir na planilha antes
// de aprovar.
function montarSugestaoRenomeacao_(idArquivo, nomeAtual, ehPastaComprovante) {
  var texto;
  try {
    texto = extrairTextoPdfOcr_(idArquivo, nomeAtual);
  } catch (e) {
    return { nomeSugerido: '', confianca: 'revisar', cnpjCpf: '', observacao: 'Falha ao ler o PDF (OCR): ' + e.message };
  }
  var docFiscal = extrairCnpjCpf_(texto);
  var tipo = inferirTipoDocumento_(texto, ehPastaComprovante);
  var numero = inferirNumeroDocumento_(texto);
  var razao = inferirRazaoSocial_(texto);

  var partes = [tipo, numero, razao ? sanitizarNomeArquivo_(razao) : null, docFiscal ? docFiscal.codigo : null]
    .filter(function (p) { return p; });
  var completo = !!(tipo && numero && razao && docFiscal);
  var nomeSugerido = partes.length ? sanitizarNomeArquivo_(partes.join(' ')) + '.pdf' : '';

  return {
    nomeSugerido: nomeSugerido,
    confianca: completo ? 'alta' : (partes.length ? 'revisar (incompleto)' : 'revisar (nada identificado)'),
    cnpjCpf: docFiscal ? (docFiscal.tipo + ' ' + docFiscal.formatado) : '',
    observacao: completo ? '' : 'Confira/complete antes de aprovar — campo(s) não identificado(s) automaticamente.',
  };
}

// Passo 1: varre PASTAS_BUSCA_ANEXO (+ subpastas) atrás de PDF fora do
// padrão e enche "Renomear Pendente" com sugestões. Não altera nenhum
// arquivo. Processa no máximo RENOMEAR_LIMITE_OCR_POR_EXECUCAO arquivos por
// execução (o resto fica pra próxima chamada/gatilho) pra nunca estourar o
// tempo máximo de execução do Apps Script.
function identificarArquivosForaDoPadrao() {
  var sh = getSheetRenomear_();
  var jaNaFila = idsJaNaFilaRenomear_();
  var agora = new Date();
  var linhasNovas = [];
  var processados = 0;

  for (var p = 0; p < PASTAS_BUSCA_ANEXO.length; p++) {
    var pastaRaizId = PASTAS_BUSCA_ANEXO[p];
    var ehPastaComprovante = p === PASTAS_BUSCA_ANEXO.length - 1; // "16 COMPROVANTE PAGTO" é sempre a última da lista (ver comentário acima de PASTAS_BUSCA_ANEXO)
    var todasAsPastas = listarPastaEsubpastas_(pastaRaizId);
    for (var i = 0; i < todasAsPastas.length && processados < RENOMEAR_LIMITE_OCR_POR_EXECUCAO; i++) {
      var pdfs = listarPdfsDaPasta_(todasAsPastas[i]);
      for (var j = 0; j < pdfs.length && processados < RENOMEAR_LIMITE_OCR_POR_EXECUCAO; j++) {
        var arq = pdfs[j];
        if (REGEX_NOME_PADRAO_.test(arq.nome)) continue; // já está no padrão — nem precisa de OCR
        if (jaNaFila[arq.id]) continue; // já sugerido (ou já renomeado) antes

        var sugestao = montarSugestaoRenomeacao_(arq.id, arq.nome, ehPastaComprovante);
        processados++;
        jaNaFila[arq.id] = true;
        linhasNovas.push([
          agora, pastaRaizId, arq.id, arq.nome, sugestao.nomeSugerido,
          sugestao.confianca, sugestao.cnpjCpf, false, 'Pendente', sugestao.observacao,
        ]);
      }
    }
  }

  if (linhasNovas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, linhasNovas.length, RENOMEAR_HEADERS.length).setValues(linhasNovas);
  }
  registrarLog_('(sistema)', 'Identificar arquivos fora do padrão',
    linhasNovas.length + ' novo(s) na fila "Renomear Pendente"' +
    (processados >= RENOMEAR_LIMITE_OCR_POR_EXECUCAO ? ' (limite do lote atingido — rode de novo pra continuar o restante)' : ''));
}

// Passo 3: aplica de fato as renomeações marcadas com "Aprovar" = TRUE na
// aba "Renomear Pendente" (Status ainda "Pendente"). Só mexe no arquivo cuja
// linha foi aprovada manualmente — nunca em massa sem revisão.
function aplicarRenomeacoesAprovadas() {
  var sh = getSheetRenomear_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return;
  var idCol = RENOMEAR_HEADERS.indexOf('ID Arquivo');
  var nomeSugeridoCol = RENOMEAR_HEADERS.indexOf('Nome Sugerido');
  var aprovarCol = RENOMEAR_HEADERS.indexOf('Aprovar');
  var statusCol = RENOMEAR_HEADERS.indexOf('Status');
  var observacaoCol = RENOMEAR_HEADERS.indexOf('Observação');
  var faixa = sh.getRange(2, 1, ultimaLinha - 1, RENOMEAR_HEADERS.length);
  var dados = faixa.getValues();
  var aplicados = 0;

  for (var i = 0; i < dados.length; i++) {
    var linha = dados[i];
    if (linha[statusCol] !== 'Pendente' || linha[aprovarCol] !== true) continue;
    var nomeSugerido = String(linha[nomeSugeridoCol] || '').trim();
    if (!nomeSugerido) continue;
    try {
      DriveApp.getFileById(linha[idCol]).setName(nomeSugerido);
      dados[i][statusCol] = 'Renomeado';
      dados[i][observacaoCol] = 'Renomeado em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      aplicados++;
    } catch (e) {
      dados[i][statusCol] = 'Erro';
      dados[i][observacaoCol] = 'Falha ao renomear: ' + e.message;
    }
  }

  if (aplicados) faixa.setValues(dados);
  if (aplicados) registrarLog_('(sistema)', 'Aplicar renomeações aprovadas', aplicados + ' arquivo(s) renomeado(s)');
}

// Rode esta função UMA VEZ no editor do Apps Script (mesmo jeito de
// configurarGatilhoDiario) pra deixar a padronização de nomes rodando
// sozinha dali pra frente: identifica 1x/dia (novo arquivo fora do padrão) e
// aplica as aprovações de hora em hora — sem precisar abrir o editor de novo
// nem gastar nenhum token de IA na execução do dia a dia.
// ---------- Ações da aba "Padronizar Nomes" na tela do app ----------
// Mesma fila "Renomear Pendente" usada pelos gatilhos automáticos acima,
// só que com ações imediatas pelo clique na tela — não depende de esperar
// o gatilho de hora em hora nem de abrir a planilha.
function localizarLinhaRenomear_(idArquivo) {
  var sh = getSheetRenomear_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return null;
  var idCol = RENOMEAR_HEADERS.indexOf('ID Arquivo');
  var statusCol = RENOMEAR_HEADERS.indexOf('Status');
  var faixa = sh.getRange(2, 1, ultimaLinha - 1, RENOMEAR_HEADERS.length);
  var dados = faixa.getValues();
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][idCol]) === String(idArquivo) && dados[i][statusCol] === 'Pendente') {
      return { faixa: faixa, dados: dados, indice: i };
    }
  }
  return null;
}

// Leitura pura (sem senha) — só as sugestões ainda não tratadas ("Pendente"),
// pra alimentar a tabela da aba "Padronizar Nomes".
function api_carregarRenomeacoes(payload) {
  var sh = getSheetRenomear_();
  var ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return { itens: [] };
  var valores = sh.getRange(2, 1, ultimaLinha - 1, RENOMEAR_HEADERS.length).getValues();
  var idCol = RENOMEAR_HEADERS.indexOf('ID Arquivo');
  var nomeAtualCol = RENOMEAR_HEADERS.indexOf('Nome Atual');
  var nomeSugeridoCol = RENOMEAR_HEADERS.indexOf('Nome Sugerido');
  var confiancaCol = RENOMEAR_HEADERS.indexOf('Confiança');
  var cnpjCpfCol = RENOMEAR_HEADERS.indexOf('CNPJ/CPF Encontrado');
  var observacaoCol = RENOMEAR_HEADERS.indexOf('Observação');
  var statusCol = RENOMEAR_HEADERS.indexOf('Status');
  var itens = valores
    .filter(function (l) { return l[statusCol] === 'Pendente'; })
    .map(function (l) {
      return {
        idArquivo: l[idCol],
        nomeAtual: l[nomeAtualCol],
        nomeSugerido: l[nomeSugeridoCol],
        confianca: l[confiancaCol],
        cnpjCpf: l[cnpjCpfCol],
        observacao: l[observacaoCol],
      };
    });
  return { itens: itens };
}

function api_aprovarRenomeacao(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var encontrado = localizarLinhaRenomear_(payload.idArquivo);
  if (!encontrado) throw new Error('Sugestão não encontrada (pode já ter sido tratada por outra pessoa — atualize a lista).');
  var nomeFinal = sanitizarNomeArquivo_(String(payload.nomeSugerido || '').trim());
  if (!nomeFinal) throw new Error('Informe um nome de arquivo válido antes de aprovar.');
  if (!/\.pdf$/i.test(nomeFinal)) nomeFinal += '.pdf';

  var idCol = RENOMEAR_HEADERS.indexOf('ID Arquivo');
  var nomeAtualCol = RENOMEAR_HEADERS.indexOf('Nome Atual');
  var nomeSugeridoCol = RENOMEAR_HEADERS.indexOf('Nome Sugerido');
  var statusCol = RENOMEAR_HEADERS.indexOf('Status');
  var observacaoCol = RENOMEAR_HEADERS.indexOf('Observação');
  var idArquivo = encontrado.dados[encontrado.indice][idCol];

  DriveApp.getFileById(idArquivo).setName(nomeFinal); // única linha que de fato mexe no arquivo — só roda depois do clique em "Aprovar"

  var nomeAntigo = encontrado.dados[encontrado.indice][nomeAtualCol];
  encontrado.dados[encontrado.indice][nomeSugeridoCol] = nomeFinal;
  encontrado.dados[encontrado.indice][statusCol] = 'Renomeado';
  encontrado.dados[encontrado.indice][observacaoCol] = 'Renomeado por ' + nomeUsuario + ' em ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' (via tela)';
  encontrado.faixa.setValues(encontrado.dados);

  registrarLog_(nomeUsuario, 'Renomear arquivo (Drive)', nomeAntigo + ' → ' + nomeFinal);
  return { ok: true };
}

function api_ignorarRenomeacao(payload) {
  var nomeUsuario = validarAcessoEdicao_(payload);
  var encontrado = localizarLinhaRenomear_(payload.idArquivo);
  if (!encontrado) throw new Error('Sugestão não encontrada (pode já ter sido tratada por outra pessoa — atualize a lista).');

  var nomeAtualCol = RENOMEAR_HEADERS.indexOf('Nome Atual');
  var statusCol = RENOMEAR_HEADERS.indexOf('Status');
  var observacaoCol = RENOMEAR_HEADERS.indexOf('Observação');
  encontrado.dados[encontrado.indice][statusCol] = 'Ignorado';
  encontrado.dados[encontrado.indice][observacaoCol] = 'Ignorado por ' + nomeUsuario + ' em ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  encontrado.faixa.setValues(encontrado.dados);

  registrarLog_(nomeUsuario, 'Ignorar sugestão de renomeação (Drive)', encontrado.dados[encontrado.indice][nomeAtualCol]);
  return { ok: true };
}

// Dispara a varredura na hora, pelo botão da tela (mesma
// identificarArquivosForaDoPadrao usada pelo gatilho diário) — exige edição
// habilitada porque grava sugestões novas na planilha. Pode levar até cerca
// de 1 minuto (lote de até RENOMEAR_LIMITE_OCR_POR_EXECUCAO arquivos com OCR).
function api_identificarRenomeacoesAgora(payload) {
  validarAcessoEdicao_(payload);
  identificarArquivosForaDoPadrao();
  return api_carregarRenomeacoes();
}

function configurarGatilhosRenomeacao() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'identificarArquivosForaDoPadrao' || fn === 'aplicarRenomeacoesAprovadas') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('identificarArquivosForaDoPadrao').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('aplicarRenomeacoesAprovadas').timeBased().everyHours(1).create();
}
