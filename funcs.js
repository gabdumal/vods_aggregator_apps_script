function getTableData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsVod = ss.getSheetByName("vod");
  let dataVodNO = wsVod.getRange(1, 1, wsVod.getLastRow(), 7).getValues();

  // Inicializa um array 2D com as informações extraídas, mas ordenado pela segunda coluna de cada item na ordem decrescente
  let dataVod = dataVodNO.sort(function (a, b) {
    return b[1] - a[1];
  });

  let wsContent = ss.getSheetByName("content");
  let dataContentNO = wsContent
    .getRange(1, 1, wsContent.getLastRow(), 5)
    .getValues();

  // Inicializa um array 2D com as informações extraídas, mas ordenado pela segunda coluna de cada item na ordem decrescente
  let dataContent = dataContentNO.sort(function (a, b) {
    return b[1] - a[1];
  });

  // Declara variável de array que guardará todos os dados da tabela
  tableData = [];

  dataVod.forEach(function (vod) {
    let vodObj = {
      id: vod[0],
      num: vod[1],
      sts: vod[2],
      tit: vod[3],
      cod: vod[4],
      obs: vod[5],
      part: vod[6],
    };

    // Verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
    function checkIdMatch(contPar) {
      return contPar[1] == vodObj.id;
    }

    // Define uma lista de objetos Content cujo atributo idVod é igual ao ID do Vod
    let contentListEqual = dataContent.filter(checkIdMatch);

    vodObj.contents = contentListEqual;
    tableData.push(vodObj);
  });
  return tableData;
}
