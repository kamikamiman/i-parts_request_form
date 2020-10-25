// ========================================================================= //
//                    不要なファイルと複製したフォームを削除                          //
// ========================================================================= //

function DeleteFile(sheetName) {
      
  // ルートフォルダ内に保存された添付ファイルを削除  
  const root = DriveApp.getRootFolder().getFiles();
  while(root.hasNext()) {
    const rootFile = root.next();
    rootFile.setTrashed(true);
  }      
  
  // 複製した依頼書フォームを削除
  const sheet = setForm.getSheetByName(sheetName);
  setForm.deleteSheet(sheet);  
  
}
