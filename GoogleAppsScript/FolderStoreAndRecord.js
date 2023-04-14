/*-----------------------------------------------------------------------*/
//
//　Google Driveに入っている特定のフォルダを移動し、
//　移動したファイルの情報をSpreadsheetに記録するマクロです。
//　シート名等はダミーになっています。
//　
//　概ね移動して記録するだけですが、たくさんフォルダがある中で、
//　移動するものとしないものがあるため、振り分けが必要でした。
//　Spreadsheetにあるターゲット番号のリストとフォルダの一部にある番号を照会して
//　振り分けを行っています。
//
/*-----------------------------------------------------------------------*/

function FolderStoreAndRecord() {
  //シート取得
  const recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('記録用シート');
  const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ターゲット番号リストシート');

  //フォルダ取得
  const fromFolder = DriveApp.getFolderById("移動したいフォルダが入っているフォルダのID");
  const storeFolder = DriveApp.getFolderById("格納先のフォルダのID");

  //番号リスト取得
  const LastRow = infoSheet.getRange(2,10).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const NumArray = infoSheet.getRange(2,10,LastRow-1,1).getDisplayValues().flat();
  Logger.log(NumArray);

  //フォルダ取得
  let folders = fromFolder.getFolders();

  //whileループ 次のフォルダがあるなら
  while (folders.hasNext()) {
    let folder = folders.next();
    let folderId = folder.getId();
    let folderName = folder.getName();
    let folderNum = folderName.slice(3,8);

    //格納先フォルダ除外
    if ( folderId !== storeFolder ) {
     //フォルダ番号がリストにあるものなら
      if ( NumArray.indexOf(folderNum) > -1 ) {
        let recordRow = recordSheet.getRange(recordSheet.getMaxRows(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow()+1;
        let updateTime = folder.getLastUpdated();

        var contents = folder.getFiles();
        var i = 0;
        var j = 0;
        while(contents.hasNext()) {
          file = contents.next();
          var isImage  = file.getBlob().getContentType();
          var isFileName  = file.getName();
          console.log(isImage);
          console.log(isFileName);

            //jpg,psdファイルはカウントする
            if(isImage == 'image/jpeg' ){ i++}
            //if(isImage == 'image/x-photoshop' ){ j++ }
            if(isFileName !== '.DS_Store' && isImage !== 'image/jpeg' ){ j++}
        }

        //入稿用格納フォルダに移動
        folder.moveTo(storeFolder);

        //各情報を記録
        recordSheet.getRange(recordRow, 1).setValue(updateTime);
        recordSheet.getRange(recordRow, 2).setValue(folderName);
        recordSheet.getRange(recordRow, 3).setValue(i);
        recordSheet.getRange(recordRow, 4).setValue(j);
      }
    }
    else{}
  }
}