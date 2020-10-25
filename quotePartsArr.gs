// ========================================================================= //
//                    依頼要旨が見積もり + 部品手配の場合の処理                    //
// ========================================================================= //

function QuotePartsArr() {
  
    console.log("依頼要旨：見積もり依頼 + 部品手配");
    
    setSheetForm.getRange('H3').setBackground(fillColor);  // 「見積」塗り潰す
    setSheetForm.getRange('L3').setBackground(fillColor);  // 「発注」塗り潰す
    setSheetForm.getRange('I11').setBackground(fillColor); // 「見積作成と同時手配」塗り潰す
    setSheetForm.getRange('I11').setValue(checkMark);      // 「見積作成と同時手配」チェック
    setSheetForm.getRange('I19').setValue(W);              // 作成期限
    setSheetForm.getRange('I13').setValue(AP);             // 納期希望日
    setSheetForm.getRange('Q16').setValue(X);              // 受注予想確率
    setSheetForm.getRange('AG13').setValue(AD);            // 見積もり番号
    setSheetForm.getRange('B40').setValue(V);              // 依頼内容
    setSheetForm.getRange('I51').setValue(AU);             // お客様の状況
    
    // 依頼種類による処理
    if ( U == '部品のみ' ) setSheetForm.getRange('C21').setBackground(fillColor);  // 「部品のみ」塗り潰す
    getPoints = [ Z, AA, AJ, AK, Y ];　// [ 作業日数, 作業人数, 立会日数, 立会人数, 作業曜日 ]
    if ( U == '部品と作業' ) PartsAndWork(getPoints);
    if ( U == '作業のみ' ) WorkOnly(getPoints);
    
    // 部品発送方法の処理
    if ( AB == '直送' ) {
      setSheetForm.getRange('S13').setBackground(fillColor); // 「直送」塗り潰す
    } else {
      setSheetForm.getRange('V13').setBackground(fillColor); // 「工事持参」塗り潰す
    }
    
    // 事前連絡の有無の処理
    if ( AC == '必要' ) setSheetForm.getRange('Y13').setBackground(fillColor); // 「必要」塗り潰す
  
}
