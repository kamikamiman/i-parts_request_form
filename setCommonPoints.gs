// ========================================================================= //
//                      依頼書フォームの共通の記入箇所                            //
// ========================================================================= //

function SetCommonPoints() {
  
  setSheetForm.getRange('B7').setValue(date);  // 依頼日
  setSheetForm.getRange('K7').setValue(B);     // 依頼部署
  setSheetForm.getRange('U7').setValue(C);     // 依頼者名
  setSheetForm.getRange('H9').setValue(D);     // 得意先名
  setSheetForm.getRange('V9').setValue(E);     // お客様担当者
  setSheetForm.getRange('AB9').setValue(G);    // 機種・号機・ユニット
  
}
