// ========================================================================= //
//                      （依頼種類）作業のみの場合の処理                            //
// ========================================================================= //

function WorkOnly(getPoints) {
  
  setSheetForm.getRange('P30').setBackground(fillColor);   //「作業のみ」塗り潰す。
  setSheetForm.getRange('W32').setValue(getPoints[0]);                // 作業日数
  setSheetForm.getRange('Z32').setValue(getPoints[1]);               // 作業人数
  setSheetForm.getRange('W35').setValue(getPoints[2]);               // 立会日数
  setSheetForm.getRange('Z35').setValue(getPoints[3]);               // 立会人数
  if ( getPoints[3] != '無し' && getPoints[3] != '' ) setSheetForm.getRange('S35').setBackground(fillColor);  // 立会「平日」塗り潰す。
  if ( getPoints[4] == '休日（前泊移動）' ) {
    setSheetForm.getRange('AA30').setBackground(fillColor); // 「前泊あり」塗り潰す。
    setSheetForm.getRange('S32').setBackground(fillColor);  // 「休日」塗り潰す。
  }
  if ( getPoints[4] == '平日（前泊移動）' ) {
    setSheetForm.getRange('AA30').setBackground(fillColor); // 「前泊あり」塗り潰す。
    setSheetForm.getRange('P32').setBackground(fillColor);  // 「平日」塗り潰す。
  }
  if ( getPoints[4] == '休日（当日移動）' ) {
    setSheetForm.getRange('AG30').setBackground(fillColor); // 「当日移動」塗り潰す。
    setSheetForm.getRange('S32').setBackground(fillColor);  // 「休日」塗り潰す。
  }
  if ( getPoints[4] == '平日（当日移動）' ) {
    setSheetForm.getRange('AG30').setBackground(fillColor); // 「当日移動」塗り潰す。
    setSheetForm.getRange('P32').setBackground(fillColor);  // 「平日」塗り潰す。
  }
  
}
