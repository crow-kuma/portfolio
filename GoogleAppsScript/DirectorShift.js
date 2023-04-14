/*-----------------------------------------------------------------------*/
//
//　縦型のシフト表の更新用マクロです。
//　シート名等はダミーになっています。
//　
//　一年ごとの更新で、前年のシートをコピーして、手動入力箇所をクリアします。
//　人数によって手動入力箇所の列が異なりますが、
//　スプレッドシートの方に人数入力箇所を作ると見落とす可能性があると考えたため、
//　マクロ使用時に人数のみ確認してもらう仕様になっています。
//
/*-----------------------------------------------------------------------*/


function newShiftSchedule() {
  /*-----------------------------*/
  /* 人数追加の際は、27行目を変更！！ */
  /*-----------------------------*/
  //アクティブなブックとシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();
  //新しいシートに内容をコピー
  const newSheet = sheet.copyTo(spreadsheet).activate();
  //新しいシートに移動
  spreadsheet.moveActiveSheet(1);
  const getNewSheet = SpreadsheetApp.getActiveSheet();

  let date = new Date(); //今日の日付を取得
  let newYear = date.getFullYear() + 1; //今年に1を足して来年にする
  newSheet.setName('【シフト表' + newYear + '】'); //シート名変更
  Logger.log(newSheet.getSheetName()); //シート名確認

  //新シートの非表示になっている部分を表示させる
  const MaxColumn = getNewSheet.getMaxColumns();
  getNewSheet.showColumns(1,MaxColumn);

  //シート全体を新年にする(A3のみ変更でシート全体が変わります)
  newSheet.getRange('A3').setValue(newYear);

  //人数
  const numOfPeople = 7; ////ここに現在の人数を入れる。変更箇所はここだけ！////
  const NoPplus5 = numOfPeople + 5;
  const NoPplus6 = numOfPeople + 6;

  //昨年までのデータを消去する
  for (let i = 0; i <= 11; i++) {
    ////シフトの消去////
    newSheet.getRange(5,4+i*NoPplus6,124,numOfPeople).clearContent();
    ////休日設定の消去////
    newSheet.getRange(5,NoPplus5+i*NoPplus6,124,1).clearContent();
  }
  newSheet.getRange('A1').activate();
}