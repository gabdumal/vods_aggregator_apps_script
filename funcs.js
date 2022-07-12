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

  // Declara constante de array que guardará todos os dados da tabela
  const tableData = [];

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

    // Define uma lista de objetos Conteúdo cujo atributo idVod é igual ao ID do Vod
    let contentListEqual = dataContent.filter(checkIdMatch);

    vodObj.contents = contentListEqual;
    tableData.push(vodObj);
  });
  return tableData;
}

// Adiciona as informações do Vod e Conteúdos à Spreadsheet
function addVodSs(vodInfo, contentInfoList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsVod = ss.getSheetByName("vod");
  const wsContent = ss.getSheetByName("content");

  const lastRowV = wsVod.getRange("A1").getDataRegion().getLastRow();
  const lastRowC = wsContent.getRange("A1").getDataRegion().getLastRow();

  // Define a constante idVod como o maior valor de id encontrado para os Vods + 1
  const dataIdVods = wsVod.getRange(1, 1, lastRowV, 1).getValues();
  const idVodList = dataIdVods.map(function (r) {
    return r[0];
  });
  const maxIdVod = Math.max.apply(Math, idVodList);
  const idVod = maxIdVod + 1;

  // Cria uma nova linha na sheet wsVod
  wsVod.appendRow([""]);
  // Define o formato dos dados da nova linha
  wsVod.getRange(lastRowV + 1, 1, 1, 7).setNumberFormat("@");
  // Preenche a última linha criada com os dados do objeto Vod passado nos parâmetros
  wsVod
    .getRange(lastRowV + 1, 1, 1, 7)
    .setValues([
      [
        idVod,
        vodInfo.num,
        vodInfo.sts,
        vodInfo.tit,
        vodInfo.cod,
        vodInfo.obs,
        vodInfo.part,
      ],
    ]);
  // Define o formato dos dados da nova linha novamente
  wsVod.getRange(lastRowV + 1, 1, 1, 7).setNumberFormat("@");

  // Define a constante maxIdContent como o maior valor de id encontrado para os conteúdos
  const dataIdContent = wsContent.getRange(1, 1, lastRowC, 1).getValues();
  const idContentList = dataIdContent.map(function (r) {
    return r[0];
  });
  const maxIdContent = Math.max.apply(Math, idContentList);

  for (i = 0; i < contentInfoList.length; i++) {
    const contentInfo = contentInfoList[i];
    // Cria uma nova linha na sheet wsContent com os dados do objeto Conteúdo
    wsContent.appendRow([
      maxIdContent + (i + 1),
      idVod,
      contentInfo.sts,
      contentInfo.nome,
      contentInfo.mut,
    ]);
    // Define o formato dos dados da nova linha
    wsContent.getRange(lastRowC + 1 + i, 1, 1, 5).setNumberFormat("@");
  }
}

// Extrai os dados de um Vod a partir do seu ID
function getVodDataById(idPar) {
  // Query para buscar todas as colunas de um registro da sheet vod em que a coluna A seja igual ao parâmetro idPar
  const qSearchVod = "SELECT A, B, C, D, E, F, G WHERE (A='" + idPar + "')";
  // Busca todas as células que contenham o ID passado nos parâmetros
  const vodExtractedData = query("vod!$A$1:$G", qSearchVod);
  // Query para buscar todas as colunas de todos os registros da sheet content em que a coluna B seja igual ao parâmetro idPar
  const qSearchContent = "SELECT A, B, C, D, E WHERE (B='" + idPar + "')";
  // Busca todas as células que contenham o ID passado nos parâmetros
  const extractedDataContentList = query("content!$A$1:$E", qSearchContent);
  const vodData = vodExtractedData[0];
  vodData.push(extractedDataContentList);
  return vodData;
}

// Edita as informações do Vod e Conteúdo na Spreadsheet
function editVodSs(vodInfo, contentInfoList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsVod = ss.getSheetByName("vod");
  const wsContent = ss.getSheetByName("content");

  const idVod = String(vodInfo.id);

  const lastRowV = wsVod.getRange("A1").getDataRegion().getLastRow();

  // Obtém uma lista de todos os idVods na spreadsheet
  const dataIdVods = wsVod.getRange(1, 1, lastRowV, 1).getValues();
  const idVodList = dataIdVods.map(function (r) {
    return r[0];
  });

  // Procura a posição do idVod fornecido nos parâmetros dentro do array 1D de idVods
  const pV = idVodList.indexOf(idVod);
  const rowV = pV + 1;

  // Atualiza os dados da linha na sheet wsVod com os dados do objeto Vod passado nos parâmetros
  wsVod
    .getRange(rowV, 1, 1, 7)
    .setValues([
      [
        idVod,
        vodInfo.num,
        vodInfo.sts,
        vodInfo.tit,
        vodInfo.cod,
        vodInfo.obs,
        vodInfo.part,
      ],
    ]);
  // Define o formato dos dados da linha
  wsVod.getRange(rowV, 1, 1, 7).setNumberFormat("@");

  let lastRowC = wsContent.getRange("A1").getDataRegion().getLastRow();
  // Obtém uma lista de todos os idContents na spreadsheet
  const dataContent = wsContent.getRange(1, 1, lastRowC, 2).getValues();
  const idContentList = dataContent.map(function (r) {
    return r[0];
  });
  const maxIdContent = Math.max.apply(Math, idContentList);

  // Verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
  function checkIdMatch(contPar) {
    return contPar[1] == idVod;
  }
  const contentListEqual = dataContent.filter(checkIdMatch);

  let rowsCList = [];
  for (i = 0; i < contentListEqual.length; i++) {
    // Procura a posição do idContent fornecido nos parâmetros dentro do array 1D de idContents
    const pC = idContentList.indexOf(contentListEqual[i][0]);
    const rowC = pC + 1;
    rowsCList.push(String(rowC));
  }
  // Inverte o sentido de rowsCList
  rowsCList = rowsCList.reverse();
  for (i = 0; i < rowsCList.length; i++) {
    wsContent
      .getRange(rowsCList[i], 1, 1, 5)
      .deleteCells(SpreadsheetApp.Dimension.ROWS);
  }

  lastRowC = wsContent.getRange("A1").getDataRegion().getLastRow();
  for (i = 0; i < contentInfoList.length; i++) {
    const contentInfo = contentInfoList[i];
    // Cria uma nova linha na sheet wsContent com os dados do objeto Conteúdo
    wsContent.appendRow([
      maxIdContent + (i + 1),
      idVod,
      contentInfo.sts,
      contentInfo.nome,
      contentInfo.mut,
    ]);
    // Define o formato dos dados da nova linha
    wsContent.getRange(lastRowC + 1 + i, 1, 1, 5).setNumberFormat("@");
  }
}

function deleteVodSs(idVod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsVod = ss.getSheetByName("vod");
  const wsContent = ss.getSheetByName("content");

  const lastRowV = wsVod.getRange("A1").getDataRegion().getLastRow();

  // Obtém uma lista de todos os idVods na spreadsheet
  const dataIdVods = wsVod.getRange(1, 1, lastRowV, 1).getValues();
  const idVodList = dataIdVods.map(function (r) {
    return r[0];
  });

  // Procura a posição do idVod fornecido nos parâmetros dentro do array 1D de idVods
  const pV = idVodList.indexOf(idVod);
  const rowV = pV + 1;

  // Deleta os dados da linha na sheet wsVod
  wsVod.getRange(rowV, 1, 1, 7).deleteCells(SpreadsheetApp.Dimension.ROWS);

  const lastRowC = wsContent.getRange("A1").getDataRegion().getLastRow();
  // Obtém uma lista de todos os idContents na spreadsheet
  const dataContent = wsContent.getRange(1, 1, lastRowC, 2).getValues();
  const idContentList = dataContent.map(function (r) {
    return r[0];
  });

  // Verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
  function checkIdMatch(contPar) {
    return contPar[1] == idVod;
  }
  const contentListEqual = dataContent.filter(checkIdMatch);

  let rowsCList = [];
  for (i = 0; i < contentListEqual.length; i++) {
    // Procura a posição do idContent fornecido nos parâmetros dentro do array 1D de idContents
    const pC = idContentList.indexOf(contentListEqual[i][0]);
    const rowC = pC + 1;
    rowsCList.push(String(rowC));
  }
  // Inverte o sentido de rowsCList
  rowsCList = rowsCList.reverse();
  for (i = 0; i < rowsCList.length; i++) {
    wsContent
      .getRange(rowsCList[i], 1, 1, 5)
      .deleteCells(SpreadsheetApp.Dimension.ROWS);
  }
}
