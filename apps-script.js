function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(['TESTE', new Date().toString()]);
  return ContentService.createTextOutput('OK - linha gravada na planilha');
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var dadosStr;
  if (e.parameter && e.parameter.dados) {
    dadosStr = e.parameter.dados;
  } else if (e.postData && e.postData.contents) {
    dadosStr = e.postData.contents;
  }

  if (!dadosStr) {
    sheet.appendRow(['ERRO', 'Nenhum dado recebido', new Date().toString()]);
    return ContentService.createTextOutput('erro: sem dados');
  }

  try {
    var dados = JSON.parse(dadosStr);

    if (sheet.getLastRow() == 0 && dados.cabecalho) {
      sheet.appendRow(dados.cabecalho);
    }

    sheet.appendRow(dados.valores);
    return ContentService.createTextOutput('ok');
  } catch(err) {
    sheet.appendRow(['ERRO', err.toString(), new Date().toString()]);
    return ContentService.createTextOutput('erro: ' + err.toString());
  }
}
