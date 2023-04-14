/*-------------------------------------------------------------------------------*/
//
//　ご依頼の受注情報を管理しているSpreadsheetがあり、
//　データがたまったときに更新に利用するマクロです。
//　シート名等はダミーになっています。
//　ユーザー様のご希望により、マクロとSpreadsheetの関数の一部は、
//　利用のたびに手動で書き換える仕様です。
//　
//　前身の管理マスターシートは、更新の際に膨大な手間と時間を要し、トラブルも多かったのですが、
//　受注管理マスターシートを自ら大規模改修し、更新時に問題が発生しにくいように工夫したことで、
//　非常にシンプルな構造にすることができました。
//　手動で入力しているデータのみ9000行分クリアし、
//　9002〜10000行目の入力データを2行目以下にコピーするだけでほぼ更新完了です。
//　更新の状況により、arrayやimportrangeなどの関数の入力位置や内容が多少異なるため、
//　その関数は手動入力のほうがシンプルだということで、手動操作になっています。
//
/*-------------------------------------------------------------------------------*/


function OrderMasterUpdate() {
  
  //！！テストする時は、31行目辺りの　const OMSheet = SpreadSheet.getSheetByName('依頼管理表');　のシート名書き換え！！

  /*使い方
    まずは以下を書き換えて、保存してください。
    通常の更新　→　両方false
    rawデータの更新と同時　→　const rawDataUpdate = true;
    関数が2行目にない(rawデータ更新の次の更新)　→　const formulasNotInSecondRow = true;
    両方trueのケースはありません。

    使い方は更に下に続きます。
  */
  const rawDataUpdate = false;
  const formulasNotInSecondRow = false;

  /*使い方　続き
  １．！！手動で管理表をコピーする(コピーがバックアップになります)！！
  ２．一応有料アカウントでマクロを動かす。(６分の壁問題回避のため)
  ３．importrangeとqueryの関数を手動で書き換え。B2, I2, AM2の関数の書き換えが必要になります。
    通常の更新　→　範囲のみ
    rawデータの更新と同時　→　URL,シート名、範囲を書き換え、空行の1行目に貼り付け。
    関数が2行目にない　→　空っぽになります。バックアップからコピーし、2行目に入力して、範囲のみ書き換え

  */

  var SpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const OMSheet = SpreadSheet.getSheetByName('管理マスター'); //テスト時はここのシート名をコピーシートに変更する。

  //関数に使う数字の取得　ここ9000行削除に対応のため、書き換える
  const NumFormula = OMSheet.getRange("A2").getFormula(); //現在のNoの関数
  const orderNumstr = NumFormula.substring(18,23); //現在のNoの関数から、数字取得。　※桁が変わったとき、要変更！ =sequence(10000,1,xxxxx)
  Logger.log(orderNumstr);
  const orderNum = Number(orderNumstr) + 9000; //新関数のNo
  Logger.log(orderNum);

  //No
  OMSheet.getRange("A2").setFormula("=sequence(10000,1," + orderNum + ")");

  if( rawDataUpdate == true ) {
    //importrange及びqueryの範囲のみ、9002行目〜10001行目の内容を2行目に貼り付け
    OMSheet.getRange("B9002:I10001").copyTo(OMSheet.getRange("B2:I1001"),{contentsOnly:true});
    OMSheet.getRange("AM9002:AT10001").copyTo(OMSheet.getRange("AM2:AT1001"),{contentsOnly:true});
  } else if (formulasNotInSecondRow == true) {
    //importrange及びqueryを消去
    OMSheet.getRange("B2:B1002").clearContent();
    OMSheet.getRange("I2:I1002").clearContent();
    OMSheet.getRange("AM2:AM1002").clearContent();
  }

  //2~9001行目の内容を消去
  OMSheet.getRange("P2:T9001").clearContent();
  OMSheet.getRange("Y2:AE9001").clearContent();

  //9002~10001行目の内容を２行目に貼り付け
  OMSheet.getRange("P9002:T10001").copyTo(OMSheet.getRange("P2:T1001"),{contentsOnly:true});
  OMSheet.getRange("Y9002:AE10001").copyTo(OMSheet.getRange("Y2:AE1001"),{contentsOnly:true});

  //9002~10001行目の内容を消去
  OMSheet.getRange("P9002:T10001").clearContent();
  OMSheet.getRange("Y9002:AE10001").clearContent();
}