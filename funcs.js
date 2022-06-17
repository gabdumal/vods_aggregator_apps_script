function getTableData() {
  // Obtém a spreadsheet ativa
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Obtém uma sheet a partir do nome vod
  let wsVod = ss.getSheetByName("vod");
  // Obtém um array 2D com informações extraídas da sheet selecionada (não ordenado)
  let dataVodNO = wsVod.getRange(1, 1, wsVod.getLastRow(), 7).getValues();

  // Declara um array 2D com as informações extraídas, mas ordenado pela segunda coluna de cada item na ordem decrescente
  let dataVod = dataVodNO.sort(function (a, b) {
    return b[1] - a[1];
  });

  // Obtém uma sheet a partir do nome content
  let wsContent = ss.getSheetByName("content");
  // Obtém um array 2D com informações extraídas da sheet selecionada (não ordenado)
  let dataContentNO = wsContent
    .getRange(1, 1, wsContent.getLastRow(), 5)
    .getValues();

  // Declara um array 2D com as informações extraídas, mas ordenado pela segunda coluna de cada item na ordem decrescente
  let dataContent = dataContentNO.sort(function (a, b) {
    return b[1] - a[1];
  });

  // Declara variável de array que guardará todos os dados da tabela
  tableData = [];

  // Para cada vod na lista de vods...
  dataVod.forEach(function (vod) {
    // Declara variável de objeto que guardará todos os dados de um único vod
    let vodObj = {
      id: vod[0],
      num: vod[1],
      sts: vod[2],
      tit: vod[3],
      cod: vod[4],
      obs: vod[5],
      part: vod[6],
    };
    // Função que verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
    function checkIdMatch(contPar) {
      return contPar[1] == vodObj.id;
    }
    // Define uma lista de objetos Content cujo atributo idVod é igual ao ID do Vod
    let contentListEqual = dataContent.filter(checkIdMatch);
    contentListEqual.forEach(function (content) {});
    let contentObj;
    // Adiciona a lista de objetos Content ao array de dados do vod
    vodObj.contents = contentListEqual;
    // Adiciona o array de dados do vod à lista de dados a serem exibidos na tabela
    tableData.push(vodObj);
  });
  return tableData;
}
