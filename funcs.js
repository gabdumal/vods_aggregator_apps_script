function getTableData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsVod = ss.getSheetByName("server_vod");
  let wsContent = ss.getSheetByName("server_content");
  const tableData = []; // Guarda todos os dados da tabela

  let dataVod = wsVod.getRange(1, 1, wsVod.getLastRow(), 9).getValues();
  // Ordena matriz pela segunda coluna (número do VoD) em ordem decrescente
  dataVod = dataVod.sort(function (a, b) {
    return b[1] - a[1];
  });

  let dataContent = wsContent
    .getRange(1, 2, wsContent.getLastRow(), 6)
    .getValues();

  dataVod.forEach(function (vodData) {
    const vod = {
      id: vodData[0],
      number: vodData[1],
      title: vodData[2],
      link: vodData[3],
      observation: vodData[4],
      participants: vodData[5],
      watchStatus: vodData[6],
      comments: vodData[7],
      favorite: vodData[8] === "S",
      contentList: [],
    };

    // Verifica se o id do Vod é igual ao idVod registrado no Conteúdo
    function checkIdMatch(content) {
      return content[0] === vod.id;
    }

    // Lista de objetos Conteúdo cujo atributo idVod é igual ao ID do Vod
    const matchedContentList = dataContent.filter(checkIdMatch);

    for (const contentData of matchedContentList) {
      const content = {
        category: contentData[1],
        description: contentData[2],
        soundStatus: contentData[3],
        watchStatus: contentData[4],
      };
      vod.contentList.push(content);
    }
    tableData.push(vod);
  });
  return tableData;
}

// Adiciona as informações do Vod e Conteúdos à Spreadsheet
function addVodSs(vod, contentList) {
  const ssServer = SpreadsheetApp.openById(
    "1EByNGWjjCsvcSa1nlXCMck0JOM9V6MJ2EwVVD7UYmv8"
  );
  let ssUser = SpreadsheetApp.getActiveSpreadsheet();
  const wsConfiguration = ssServer.getSheetByName("configuration");
  const wsServerVod = ssServer.getSheetByName("server_vod");
  const wsUserVod = ssUser.getSheetByName("user_vod");
  const wsServerContent = ssServer.getSheetByName("server_content");
  const wsUserContent = ssUser.getSheetByName("user_content");

  const configurationData = wsConfiguration.getRange(1, 1, 1, 3).getValues();
  const newVodId = configurationData[0][0];
  const newContentId = configurationData[0][2];

  wsServerVod.appendRow([
    "'" + newVodId,
    "'" + vod.number,
    "'" + vod.title,
    "'" + vod.link,
    "'" + vod.observation,
    "'" + vod.participants,
  ]);
  wsUserVod.appendRow([
    "'" + newVodId,
    "'" + vod.watchStatus,
    "'" + vod.comments,
    "'" + vod.favorite,
  ]);

  for (i = 0; i < contentList.length; i++) {
    const content = contentList[i];
    wsServerContent.appendRow([
      "'" + (newContentId + i),
      "'" + newVodId,
      "'" + content.category,
      "'" + content.description,
      "'" + content.soundStatus,
    ]);
    wsUserContent.appendRow([
      "'" + (newContentId + i),
      "'" + newVodId,
      "'" + content.watchStatus,
    ]);
  }
}

// Extrai os dados de um Vod a partir do seu ID
function getVodDataById(idVod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsVod = ss.getSheetByName("server_vod");
  const wsContent = ss.getSheetByName("server_content");

  const dataIdVods = wsVod.getRange(1, 1, wsVod.getLastRow(), 1).getValues();
  const idVodList = dataIdVods.map(function (r) {
    return r[0];
  });

  // Procura a posição do idVod fornecido nos parâmetros dentro do array 1D de idVods
  const pV = idVodList.indexOf(idVod);
  const rowV = pV + 1;

  const vodData = wsVod.getRange(rowV, 1, 1, 9).getValues()[0];

  const vod = {
    id: vodData[0],
    number: vodData[1],
    title: vodData[2],
    link: vodData[3],
    observation: vodData[4],
    participants: vodData[5],
    watchStatus: vodData[6],
    comments: vodData[7],
    favorite: vodData[8] === "S",
    contentList: [],
  };

  const contentListData = wsContent
    .getRange(1, 2, wsContent.getLastRow(), 6)
    .getValues();

  // Verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
  function checkIdMatch(content) {
    return content[0] === idVod;
  }
  const matchedContentList = contentListData.filter(checkIdMatch);

  for (const contentData of matchedContentList) {
    const content = {
      category: contentData[1],
      description: contentData[2],
      soundStatus: contentData[3],
      watchStatus: contentData[4],
    };
    vod.contentList.push(content);
  }

  return vod;
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
    // Procura a posição do idContent fornecido nos parâmetros dentro do array 1D de idContent
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
  const ssServer = SpreadsheetApp.openById(
    "1EByNGWjjCsvcSa1nlXCMck0JOM9V6MJ2EwVVD7UYmv8"
  );
  let ssUser = SpreadsheetApp.getActiveSpreadsheet();
  const wsServerVod = ssServer.getSheetByName("server_vod");
  const wsUserVod = ssUser.getSheetByName("user_vod");
  const wsServerContent = ssServer.getSheetByName("server_content");
  const wsUserContent = ssUser.getSheetByName("user_content");

  const serverVodIdData = wsServerVod
    .getRange(1, 1, wsServerVod.getLastRow(), 1)
    .getValues();
  const serverVodIdList = serverVodIdData.map(function (r) {
    return r[0];
  });
  let pV = serverVodIdList.indexOf(idVod);
  let rowV = pV + 1;
  wsServerVod
    .getRange(rowV, 1, 1, 7)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);

  const userVodIdData = wsUserVod
    .getRange(1, 1, wsUserVod.getLastRow(), 1)
    .getValues();
  const userVodIdList = userVodIdData.map(function (r) {
    return r[0];
  });
  pV = userVodIdList.indexOf(idVod);
  rowV = pV + 1;
  wsUserVod.getRange(rowV, 1, 1, 7).deleteCells(SpreadsheetApp.Dimension.ROWS);

  const serverDataContent = wsServerContent
    .getRange(1, 1, wsServerContent.getLastRow(), 2)
    .getValues();
  const serverContentIdList = serverDataContent.map(function (r) {
    return r[0];
  });
  // Verifica se o ID do Vod é igual ao idVod registrado no Conteúdo
  function checkIdMatch(content) {
    return content[1] === idVod;
  }
  const matchedServerContentList = serverDataContent.filter(checkIdMatch);
  let serverContentRowsList = [];
  for (let i = 0; i < matchedServerContentList.length; i++) {
    // Procura a posição do idContent fornecido nos parâmetros dentro do array 1D de idContents
    const pC = serverContentIdList.indexOf(matchedServerContentList[i][0]);
    const rowC = pC + 1;
    serverContentRowsList.push(String(rowC));
  }
  // Inverte o sentido de rowsCList e deleta as linhas selecionadas
  serverContentRowsList = serverContentRowsList.reverse();
  for (let i = 0; i < serverContentRowsList.length; i++) {
    wsServerContent
      .getRange(serverContentRowsList[i], 1, 1, 5)
      .deleteCells(SpreadsheetApp.Dimension.ROWS);
  }

  const userDataContent = wsUserContent
    .getRange(1, 1, wsUserContent.getLastRow(), 2)
    .getValues();
  const userContentIdList = userDataContent.map(function (r) {
    return r[0];
  });
  const matchedUserContentList = userDataContent.filter(checkIdMatch);
  let userContentRowsList = [];
  for (let i = 0; i < matchedUserContentList.length; i++) {
    // Procura a posição do idContent fornecido nos parâmetros dentro do array 1D de idContents
    const pC = userContentIdList.indexOf(matchedUserContentList[i][0]);
    const rowC = pC + 1;
    userContentRowsList.push(String(rowC));
  }
  // Inverte o sentido de rowsCList e deleta as linhas selecionadas
  userContentRowsList = userContentRowsList.reverse();
  for (let i = 0; i < userContentRowsList.length; i++) {
    wsUserContent
      .getRange(userContentRowsList[i], 1, 1, 3)
      .deleteCells(SpreadsheetApp.Dimension.ROWS);
  }
}
