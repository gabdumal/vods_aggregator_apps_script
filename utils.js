// Retorna o código HTML de um arquivo HTML passado nos parâmetros
function include(filename){
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}