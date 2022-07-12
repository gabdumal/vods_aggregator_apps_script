/** @OnlyCurrentDoc */

const Route = {};
// Adicionará uma propriedade ao objeto Route e define seu valor como uma função passada nos parâmetros
Route.path = function (route, callback) {
  Route[route] = callback;
};

function doGet(e) {
  if (e.parameters.v == "add") {
    return loadAddVod();
  } else if (e.parameters.v == "edit" && e.parameters.id > 0) {
    return loadEditVod(e.parameters.id);
  } else {
    return loadTable();
  }
}

function setPage(page) {
  // Cria o template, adiciona propriedades a ele e o executa
  // Tem como parâmtros o arquivo HTML que servirá de template, e o objeto com as propriedades a serem adicionadas
  page
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no"
    );
  return page;
}

// Inicia o HTML da página table
function loadTable() {
  const page = render("table");
  return setPage(page);
}

// Inicia o HTML da página addVod
function loadAddVod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("content_cat");

  // Obtém um array 2D com informações extraídas da sheet selecionada
  const list = ws
    .getRange(1, 7, ws.getRange("G1").getDataRegion().getLastRow(), 1)
    .getValues();
  // Monta uma lista de <option> com os valores extraídos, transformando o array 2D em um array 1D
  const htmlOptionList = list
    .map(function (r) {
      return "<option>" + r[0] + "</option>";
    })
    .join("");

  const page = render("addVod", { contentNamesList: htmlOptionList });
  return setPage(page);
}

// Inicia o HTML da página editVod
function loadEditVod(idVod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("content_cat");

  // Obtém um array 2D com informações extraídas da sheet selecionada
  const list = ws
    .getRange(1, 7, ws.getRange("G1").getDataRegion().getLastRow(), 1)
    .getValues();
  // Monta uma lista de <option> com os valores extraídos, transformando o array 2D em um array 1D
  const htmlOptionList = list
    .map(function (r) {
      return "<option>" + r[0] + "</option>";
    })
    .join("");

  const page = render("editVod", {
    contentNamesList: htmlOptionList,
    idVod: idVod,
  });
  return setPage(page);
}
