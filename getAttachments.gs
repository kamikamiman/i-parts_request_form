// ========================================================================= //
//                       回答フォームの添付ファイルを取得                            //
// ========================================================================= //

function GetAttachment() {
  
  let fileIds = [];　// 添付ファイルを格納する配列
  
  // 添付ファイルを取得
  if ( AO !== '' ) {
    
    const files = AO.split(',');
    
    fileIds = files.map( el => {
      const fileId = el.split('=')[1];
      return DriveApp.getFileById(fileId); // 配列[fileIds]に返す 
    });
  
  }

  return fileIds;

}
