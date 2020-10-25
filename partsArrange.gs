// ========================================================================= //
//                      依頼要旨が部品手配の場合の処理                            //
// ========================================================================= //

function PartsArrange() {
  
    console.log("依頼要旨：部品手配");
  
    const cell = setSheetForm.getRange('L3').setBackground(fillColor); // セル「発注」を黄色で塗り潰す。
    setSheetForm.getRange('I13').setValue(Q);                          // 納期希望日
    setSheetForm.getRange('AG13').setValue(T);                         // 見積もり番号
    setSheetForm.getRange('B40').setValue(P);                          // 依頼内容
    setSheetForm.getRange('I51').setValue(AT);                         // お客様の状況    
    
    // 部品発送方法の処理
    if ( R == '直送' ) {
      setSheetForm.getRange('S13').setBackground(fillColor); // 直送を黄色で塗り潰す。
    } else {
      setSheetForm.getRange('V13').setBackground(fillColor); // 工事持参を黄色で塗り潰す。
    }
    
    // 事前連絡の有無の処理
    if ( S == '必要' ) setSheetForm.getRange('Y13').setBackground(fillColor); // 事前連絡要を黄色で塗り潰す。
  
}
