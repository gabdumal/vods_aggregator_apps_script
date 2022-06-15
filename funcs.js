function getTableData(){
    // Obtém a spreadsheet ativa
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Obtém uma sheet a partir do nome vod
    var wsVod = ss.getSheetByName("vod");
    // Obtém um array 2D com informações extraídas da sheet selecionada (não ordenado)
    var dataVodNO = wsVod.getRange(1, 1, wsVod.getLastRow(), 7).getValues();
    
    // Declara um array 2D com as informações extraídas, mas ordenado pela segunda coluna de cada item na ordem decrescente
    var dataVod = dataVodNO.sort(function(a,b) {
      return b[1] - a[1];
    });
  
    // Declara variável de array que guardará todos os dados da tabela
    tableData = [];
    
    // Para cada vod na lista de vods...
    dataVod.forEach(function(vod){
      // Declara variável de objeto que guardará todos os dados de um único vod
      var vodObj = {id: vod[0], num: vod[1], sts: vod[2], tit: vod[3], cod: vod[4], obs: vod[5], part: vod[6]};
      // Adiciona o array de dados do vod à lista de dados a serem exibidos na tabela
      tableData.push(vodObj);
    });
    return tableData;
  }