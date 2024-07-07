

function createSheetsBasedOnRange() {

var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetRef = ss.getSheetByName('參數'); // 指定來源工作表
var sheet = ss.getSheetByName('2024'); // 指定來源工作表

var rangeValues = sheetRef.getRange('G2:G32').getValues(); // 獲取 A1:A12 範圍的值

  for (var i = 0; i < rangeValues.length; i++) {
    var sheetName = rangeValues[i][0]; // 指定新工作表的名字
    if(sheetName && !ss.getSheetByName(sheetName)) { // 檢查名字是否存在，並且工作表是否已存在
      //ss.insertSheet(sheetName);
      
      var newsheet = ss.getSheetByName('2024').copyTo(ss);
      newsheet.setName(sheetName);
    }
  }
}
