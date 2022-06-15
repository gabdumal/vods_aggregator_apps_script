/** @OnlyCurrentDoc */

function doGet() {
  return HtmlService.createTemplateFromFile("table").evaluate();
}
