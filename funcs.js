function getTableData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsVod = ss.getSheetByName("server_vod");
  let wsContent = ss.getSheetByName("server_content");
  const tableData = []; // Guarda todos os dados da tabela

  const dataServerVod = wsVod.getRange(1, 1, wsVod.getLastRow(), 7).getValues();
  const dataServerContent = wsContent
    .getRange(1, 1, wsContent.getLastRow(), 3)
    .getValues();
  for (const serverVod of dataServerVod) {
    if (serverVod[6] !== "S") continue;
    // Verifica se o primeiro item de um array corresponde a um id dado, capturando a última correspondência
    function checkIdMatch(array) {
      return array[0] === serverVod[0];
    }
    const vod = {
      id: serverVod[0],
      number: serverVod[1],
      title: serverVod[2],
      link: serverVod[3],
      observation: serverVod[4],
      participants: serverVod[5],
      contentList: [],
    };

    const matchedServerContentList = dataServerContent.filter(checkIdMatch);
    for (const serverContent of matchedServerContentList) {
      const content = {
        name: serverContent[1],
        soundStatus: serverContent[2],
      };
      vod.contentList.push(content);
    }
    tableData.push(vod);
  }

  // Organiza pelo número do vod
  tableData.sort(function (a, b) {
    return b.number - a.number;
  });

  return tableData;
}
