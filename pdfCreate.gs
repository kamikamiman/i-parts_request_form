// ========================================================================= //
//                 複製したスプレットシートをPDF変換してファイルを出力                    //
// ========================================================================= //

function PdfCreate(sheetId) {
  
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ/export?exportFormat=pdf&gid=SID'.replace('SID', sheetId);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers:{
      'Authorization': 'Bearer '+token
    }
  });
  
  const now = DateFormat2();
  const blob = response.getBlob().setName( `${now}_${C}_iパーツ依頼書.pdf`);           // pdfの名前
  const folder = DriveApp.getFolderById('1mJKjawfQUio1ZEDsFq7re-t5cwm3xfyK');  // pdfの保存先フォルダを指定
  const form = folder.createFile(blob);                                        // フォルダ内にiパーツ依頼書を作成
  
  return form;
}
