/** @OnlyCurrentDoc */

function doGet() {
  /** Executa a função que cria o template, adiciona propriedades a ele e o executa,
  tendo como parâmtros o arquivo HTML que servirá de template, e o objeto com as propriedades a serem adicionadas */
  var page = render("table");
  page
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
  return page;
}
