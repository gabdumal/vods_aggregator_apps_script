/** @OnlyCurrentDoc */

const Route = {};
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

// Constrói e retorna página
function setPage(page) {
  page
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no"
    );
  return page;
}

function loadTable() {
  const page = render("table");
  return setPage(page);
}

function loadAddVod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("operational");

  const contentCategoryOptionList = getContentCategoryOptionList(ws);

  const page = render("addVod", {
    contentCategoryOptionList: contentCategoryOptionList,
  });
  return setPage(page);
}

// Inicia o HTML da página editVod
function loadEditVod(idVod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("operational");

  const contentCategoryOptionList = getContentCategoryOptionList(ws);

  const page = render("editVod", {
    contentCategoryOptionList: contentCategoryOptionList,
    idVod: idVod,
  });
  return setPage(page);
}

function getContentCategoryOptionList(ws) {
  // Obtém array 2D com informações extraídas da sheet selecionada
  const list = ws.getRange(1, 7, ws.getLastRow(), 1).getValues();

  // Monta uma lista de <option>
  return list
    .map(function (data) {
      return "<option>" + data[0] + "</option>";
    })
    .join("");
}
