// ========================================================================= //
//                    依頼要旨が見積もり依頼の場合の処理                           //
// ========================================================================= //

function ReqQuote() {
  
    console.log("依頼要旨：見積もり依頼");
  
    setSheetForm.getRange('H3').setBackground(fillColor); // 「見積」塗り潰す
    setSheetForm.getRange('I19').setValue(K);             // 作成期限
    setSheetForm.getRange('Q16').setValue(L);             // 受注予想確率
    setSheetForm.getRange('G27').setValue(AE);            // 作業費　　　（事後見積もり）
    setSheetForm.getRange('G30').setValue(AG);            // 部品費　　　（事後見積もり）
    setSheetForm.getRange('H32').setValue(AH);            // ブリッジ番号（事後見積もり）
    setSheetForm.getRange('G37').setValue(AF);            // 見積もり番号（作成済み見積もりの提出）
    setSheetForm.getRange('B40').setValue(J);             // 依頼内容
    setSheetForm.getRange('I51').setValue(AS);            // お客様の状況
    
    // 依頼種類の処理
    if ( I == '部品のみ' ) setSheetForm.getRange('C21').setBackground(fillColor); // 「部品のみ」塗り潰す
    if ( I == '事後見積もり' ) {
      setSheetForm.getRange('C25').setBackground(fillColor); // 「事後見積もり」塗り潰す
      setSheetForm.getRange('B40').setValue(AM);             // 備考
    }
    if ( I == '作成済み見積もりの確認と提出' ) {
      setSheetForm.getRange('C35').setBackground(fillColor); // 「作成済み見積もりの確認と提出」塗り潰す
      setSheetForm.getRange('B40').setValue(AN);             // 備考 
    }
    getPoints = [ N, O, AI, AL, M ];                         // [ 作業日数, 作業人数, 立会日数, 立会人数, 作業曜日 ]
    if ( I == '部品と作業' ) PartsAndWork(getPoints);
    if ( I == '作業のみ' ) WorkOnly(getPoints);

}
