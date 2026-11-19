
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("Architect");

    const arrowrange = sheet.getRangeByIndexes(4, 1, 50, 5);
    arrowrange.setValue("");

    const messagerane = sheet.getRangeByIndexes(5, 6, 50, 2);
    messagerane.setValue("");
    sheet.getCell(4,7).setValue(1);
    sheet.getCell(4, 6).setValue("");

      const sheets = workbook.getWorksheets();

      for (let i = sheets.length - 1; i >= 0; i--) {
        const sheet = sheets[i];
        const name = sheet.getName();

        if (name.includes("Blueprint")) {
          sheet.delete();
        }
      }


      let dashboardSheet = workbook.getWorksheet("Dashboard");

      if (dashboardSheet) {
        dashboardSheet.delete();
      }


}
