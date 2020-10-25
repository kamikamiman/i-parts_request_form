// ========================================================================= //
//                    （依頼種類）部品と作業の場合の処理                           //
// ========================================================================= //

function PartsAndWork(getPoints) {
  
  setSheetForm.getRange('P21').setBackground(fillColor);            //「部品と作業」塗り潰す
  setSheetForm.getRange('W23').setValue(getPoints[0]);              // 作業日数
  setSheetForm.getRange('Z23').setValue(getPoints[1]);              // 作業人数
  setSheetForm.getRange('W25').setValue(getPoints[2]);              // 立会日数
  setSheetForm.getRange('Z25').setValue(getPoints[3]);              // 立会人数
  if ( getPoints[2] != '無し' && getPoints[2] != '' ) setSheetForm.getRange('S25').setBackground(fillColor);  // 立会「平日」塗り潰す
  if ( getPoints[4] == '休日（前泊移動）' ) {
    setSheetForm.getRange('AA21').setBackground(fillColor); // 「前泊あり」塗り潰す。
    setSheetForm.getRange('S23').setBackground(fillColor);  // 「休日」塗り潰す。
  }
  if ( getPoints[4] == '平日（前泊移動）' ) {
    setSheetForm.getRange('AA21').setBackground(fillColor); // 「前泊あり」塗り潰す。
    setSheetForm.getRange('P23').setBackground(fillColor);  // 「平日」塗り潰す。
  }
  if ( getPoints[4] == '休日（当日移動）' ) {
    setSheetForm.getRange('AG21').setBackground(fillColor); // 「当日移動」塗り潰す。
    setSheetForm.getRange('S23').setBackground(fillColor);  // 「休日」塗り潰す。
  }
  if ( getPoints[4] == '平日（当日移動）' ) {
    setSheetForm.getRange('AG21').setBackground(fillColor); // 「当日移動」塗り潰す。
    setSheetForm.getRange('P23').setBackground(fillColor);  // 「平日」塗り潰す。
  }
  
}
