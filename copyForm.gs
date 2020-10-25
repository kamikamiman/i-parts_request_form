// ========================================================================= //
//                           依頼書フォームを複製                                //
// ========================================================================= //

function CopyForm() {
  
  const now = DateFormat2();
  const sheetName = `${now}_iパーツ依頼書`;                          // 複製する依頼書フォーム名
  setSheetForm = setForm.duplicateActiveSheet().setName(sheetName); // 依頼書フォームを複製
  const sheetId = setSheetForm.getSheetId();                        // 複製した依頼書フォームのIDを取得
  setSheetForm.activate();                                          // 複製した依頼書フォームを開く
  
  let sheetItems = [ sheetName, sheetId ];
  
  return sheetItems;
}
