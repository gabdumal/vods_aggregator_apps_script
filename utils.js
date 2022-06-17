// Retorna o código HTML de um arquivo HTML passado nos parâmetros
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function render(file, argsObject) {
  // Cria um template a partir de um arquivo HTML passado nos parâmetros como file
  var tmp = HtmlService.createTemplateFromFile(file);

  // Caso tenha sido passado um objeto com propriedades a serem definidas...
  if (argsObject) {
    // Cria uma lista de chaves das propriedades do objeto
    var keys = Object.keys(argsObject);

    // Para cada chave na lista...
    keys.forEach(function (key) {
      // Cria uma propriedade no objeto tmp igual à propriedade dentro do objeto argsObject, passado nos parâmetros
      tmp[key] = argsObject[key];
    });
  }

  // Executa o template, incluindo suas propriedades previamente definidas
  return tmp.evaluate().setTitle(file);
}
