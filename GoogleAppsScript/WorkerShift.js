/*--------------------------------------------------------------*/ 
//
//　ワーカーさんのシフト表を操作するためのGoogle Spreadsheet用マクロです。
//　シート名等はダミーになっています。
//　まずは2ヶ月分列を追加し、関数をオートフィルします。
//　その後古い列を削除しますが、シフト表に紐付いた関数に影響が出るため、
//　最後にそれらの関数の書き換えを行います。
//　ワーカーさんの数は増減があるため、増減に対応するための工夫をしました。
//
/*--------------------------------------------------------------*/


function WorkerShift() {
  const tableSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ワーカーシフト表');
  const lastColumn = tableSheet.getLastColumn(); //列追加前の最終列を取得
  const sourceColumn = lastColumn - 6;
  let maxRow = tableSheet.getMaxRows();
  let minus5Rows = maxRow - 5;
  console.log(maxRow)
  console.log(minus5Rows)

  tableSheet.insertColumnsAfter(lastColumn,62); //最後の列の後に62列追加

  const sourceRange = tableSheet.getRange(1,sourceColumn,maxRow,7); //元データ取得
  const fillRange = tableSheet.getRange(1,sourceColumn,maxRow,69); //オートフィルの対象範囲

  sourceRange.autoFill(fillRange,SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); //新しく追加された部分にオートフィル

  tableSheet.deleteColumns(7,62); //古い列を削除

  //関数書き換え　C列
  tableSheet.getRange("C6").setFormula("=xlookup($C$1,$G$1:$NG$1,$G6:$NG6,\"\",0,2)");
  tableSheet.getRange("C6").copyTo(tableSheet.getRange(6,3,minus5Rows,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);


  //関数書き換え　「紐付け関数あり」シート
  const sumSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('紐付け関数あり');
  sumSheet.getRange("D5").setFormula("=xlookup($C$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$4:$NG$4,\"一致なし\",0,2)");
  sumSheet.getRange("E5").setFormula("=xlookup($C$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$5:$NG$5,\"一致なし\",0,2)+xlookup($E$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$3:$NG$3,\"一致なし\",0,2)");
  sumSheet.getRange("G5").setFormula("=xlookup($E$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$4:$NG$4,\"一致なし\",0,2)");
  sumSheet.getRange("H5").setFormula("=xlookup($E$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$5:$NG$5,\"一致なし\",0,2)+xlookup($H$2,'ワーカーシフト表'!$G$1:$NG$1,'ワーカーシフト表'!$G$3:$NG$3,\"一致なし\",0,2)");
}
