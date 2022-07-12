// Retorna o código HTML de um arquivo HTML passado nos parâmetros
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function render(file, argsObject) {
  // Cria um template a partir de um arquivo HTML passado nos parâmetros como file
  let tmp = HtmlService.createTemplateFromFile(file);

  // Caso tenha sido passado um objeto com propriedades a serem definidas...
  if (argsObject) {
    // Cria uma lista de chaves das propriedades do objeto
    let keys = Object.keys(argsObject);

    keys.forEach(function (key) {
      // Cria uma propriedade no objeto tmp igual à propriedade dentro do objeto argsObject, passado nos parâmetros
      tmp[key] = argsObject[key];
    });
  }

  let subtitle = "";
  switch (file) {
    case "addVod":
      subtitle = " - Adicionar VoD";
      break;
    case "editVod":
      subtitle = " - Editar VoD";
      break;
    default:
      break;
  }
  // Executa o template, incluindo suas propriedades previamente definidas
  return tmp.evaluate().setTitle("Vods do Felps" + subtitle);
}

function query(interval, qSearch) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsQuery = ss.getSheetByName("query");
  // Define uma variável como o texto da fórmula Query, passando os parâmetros dados e consulta iguais aos passados na definição da função
  const queryFormula = "=QUERY(" + interval + '; "' + qSearch + '")';
  // Define a célula A1 como uma fórmula Query com o texto definido previamente
  wsQuery.getRange(1, 1).setFormula(queryFormula);
  // Define um range a partir do intervalo passado
  const rangeData = wsQuery.getRange(interval);
  // Define uma variável com os dados extraídos da sheet de query
  const extractedData = wsQuery
    .getRange(
      1,
      1,
      wsQuery.getDataRange().getLastRow(),
      rangeData.getLastColumn()
    )
    .getValues();
  // Limpa os dados preenchidos na sheet query
  wsQuery.getDataRange().clearContent();
  // Retorna os dados extraídos
  return extractedData;
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
