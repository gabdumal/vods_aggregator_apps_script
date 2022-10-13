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

  for (let i = 0; i < contentList.length; i++) {
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
    .getRange(1, 2, wsContent.getLastRow(), 5)
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
function editVodSs(idVod, vod, contentList) {
  const ssServer = SpreadsheetApp.openById(
    "1EByNGWjjCsvcSa1nlXCMck0JOM9V6MJ2EwVVD7UYmv8"
  );
  let ssUser = SpreadsheetApp.getActiveSpreadsheet();
  const wsConfiguration = ssServer.getSheetByName("configuration");
  const wsServerVod = ssServer.getSheetByName("server_vod");
  const wsUserVod = ssUser.getSheetByName("user_vod");
  const wsServerContent = ssServer.getSheetByName("server_content");
  const wsUserContent = ssUser.getSheetByName("user_content");

  const configurationData = wsConfiguration.getRange(1, 3, 1, 3).getValues();
  const newContentId = configurationData[0][0];

  // Edita wsServerVod
  const serverVodIdData = wsServerVod
    .getRange(1, 1, wsServerVod.getLastRow(), 1)
    .getValues();
  const serverVodIdList = serverVodIdData.map(function (r) {
    return r[0];
  });
  let pV = serverVodIdList.indexOf(idVod);
  let rowV = pV + 1;
  Logger.log(idVod);
  wsServerVod
    .getRange(rowV, 2, 1, 5)
    .setValues([
      [
        "'" + vod.number,
        "'" + vod.title,
        "'" + vod.link,
        "'" + vod.observation,
        "'" + vod.participants,
      ],
    ]);

  // Edita wsUserVod
  const userVodIdData = wsUserVod
    .getRange(1, 1, wsUserVod.getLastRow(), 1)
    .getValues();
  const userVodIdList = userVodIdData.map(function (r) {
    return r[0];
  });
  pV = userVodIdList.indexOf(idVod);
  rowV = pV + 1;
  wsUserVod
    .getRange(rowV, 2, 1, 3)
    .setValues([
      ["'" + vod.watchStatus, "'" + vod.comments, "'" + vod.favorite],
    ]);

  // Delete wsServerContent
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

  // Deleta wsUserContent
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

  // Preenche wsServerContent e wsUserContent
  for (let i = 0; i < contentList.length; i++) {
    const content = contentList[i];
    wsServerContent.appendRow([
      "'" + (newContentId + i),
      "'" + idVod,
      "'" + content.category,
      "'" + content.description,
      "'" + content.soundStatus,
    ]);
    wsUserContent.appendRow([
      "'" + (newContentId + i),
      "'" + idVod,
      "'" + content.watchStatus,
    ]);
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
    .getRange(rowV, 1, 1, 6)
    .deleteCells(SpreadsheetApp.Dimension.ROWS);

  const userVodIdData = wsUserVod
    .getRange(1, 1, wsUserVod.getLastRow(), 1)
    .getValues();
  const userVodIdList = userVodIdData.map(function (r) {
    return r[0];
  });
  pV = userVodIdList.indexOf(idVod);
  rowV = pV + 1;
  wsUserVod.getRange(rowV, 1, 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS);

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
