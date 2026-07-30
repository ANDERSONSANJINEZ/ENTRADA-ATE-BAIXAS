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

function getDataBase_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_DATA_BASE);
}
function setDataBase_(valor) {
  if (valor) PropertiesService.getScriptProperties().setProperty(PROP_DATA_BASE, valor);
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
    var statusCol = HEADERS.indexOf('Status') + 1;
    itens.forEach(function (item, i) {
      var duplicado = !!chavesOutraAba[chaveDuplicidade_(item)];
      var status = duplicado ? 'Duplicado – revisar' : 'OK';
      sheet.getRange(i + 2, statusCol).setValue(status);
    });
  }
  aplicarStatus(erpSheet, erp, chavesManual);
  aplicarStatus(manualSheet, manual, chavesErp);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ erp: lerAba_('ERP'), manual: lerAba_('Manual'), dataBase: getDataBase_() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var acao = payload.action;
  var resultado = { ok: true };

  if (acao === 'importar') {
    var sh = getSheet_('ERP');
    sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), HEADERS.length).clearContent();
    sh.appendRow(HEADERS);
    var linhas = payload.itens.map(function (item) {
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
    setDataBase_(payload.dataBase);
    recalcularStatus_();
  } else if (acao === 'adicionarManual') {
    var shM = getSheet_('Manual');
    var id2 = Utilities.getUuid();
    var linha = HEADERS.map(function (h) {
      if (h === 'ID') return id2;
      if (h === 'Origem') return 'Manual';
      if (h === 'Status') return 'OK';
      return payload.item[h] != null ? payload.item[h] : '';
    });
    shM.appendRow(linha);
    recalcularStatus_();
  } else if (acao === 'removerManual') {
    var shM2 = getSheet_('Manual');
    var idCol = HEADERS.indexOf('ID') + 1;
    var dados = shM2.getRange(2, 1, Math.max(shM2.getLastRow() - 1, 0), HEADERS.length).getValues();
    for (var i = 0; i < dados.length; i++) {
      if (dados[i][idCol - 1] === payload.id) {
        shM2.deleteRow(i + 2);
        break;
      }
    }
    recalcularStatus_();
  } else {
    resultado = { ok: false, erro: 'Ação desconhecida: ' + acao };
  }

  resultado.erp = lerAba_('ERP');
  resultado.manual = lerAba_('Manual');
  resultado.dataBase = getDataBase_();
  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}
