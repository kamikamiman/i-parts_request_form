// ========================================================================= //
//                           スプレットシートを取得                              //
// ========================================================================= //

// 読込先スプレットシート
const getForm = SpreadsheetApp.openById('1RCwuHFfg4TcMsqY9dHcYpVa8BD52SmFMsPbJJn6Hu8Y');
const sheetForm = getForm.getSheetByName('iパーツ依頼書(回答)');  

// 読込先スプレットシートの最終行目
const lastRow = sheetForm.getLastRow();
  
// 書込先スプレットシート
const setForm = SpreadsheetApp.openById('1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ');
let setSheetForm = setForm.getSheetByName('iパーツ依頼フォーム');



// ========================================================================= //  
//                           Googleフォームの回答情報                           //
// ========================================================================= //

// 日付のフォーマット（パターン１：回答情報から取得）
function DateFormat1(cell) {
  if(cell !== '') {
    return Utilities.formatDate(cell,'JST' ,'yyyy/MM/dd');
  }
}

// 日付のフォーマット（パターン２：現在の時刻を取得）
function DateFormat2() {
    return Utilities.formatDate(new Date(),'JST' ,'yyyyMMdd_hhmmss');
}


// 共通記入箇所
let _date = new Date();                                   // 本日の日付
date = DateFormat1(_date);                                // 〃（表示を指定）
const  B = sheetForm.getRange(lastRow, 2).getValue();     // 依頼者部署
const  C = sheetForm.getRange(lastRow, 3).getValue();     // 依頼者名
const  D = sheetForm.getRange(lastRow, 4).getValue();     // お客様会社名
const  E = sheetForm.getRange(lastRow, 5).getValue();     // お客様担当者
const  F = sheetForm.getRange(lastRow, 6).getValue();     // 機械
const  G = sheetForm.getRange(lastRow, 7).getValue();     // 機種
const  H = sheetForm.getRange(lastRow, 8).getValue();     // 依頼要旨
const AO = sheetForm.getRange(lastRow,41).getValue();     // 添付ファイル
const AQ = sheetForm.getRange(lastRow,43).getValue();     // 送信アドレス(cc)

// 見積もり依頼
const  I = sheetForm.getRange(lastRow, 9).getValue();     // 依頼種類
const  J = sheetForm.getRange(lastRow,10).getValue();     // 依頼内容
const  k = sheetForm.getRange(lastRow,11).getValue();     // 作成期限
const  K = DateFormat1(k);                                // 〃（表示を指定）
const  L = sheetForm.getRange(lastRow,12).getValue();     // 受注予想確率
const  M = sheetForm.getRange(lastRow,13).getValue();     // 作業日
const  N = sheetForm.getRange(lastRow,14).getValue();     // 作業日数
const  O = sheetForm.getRange(lastRow,15).getValue();     // 作業人数
const AI = sheetForm.getRange(lastRow,35).getValue();     // 立会い日数
const AL = sheetForm.getRange(lastRow,38).getValue();     // 立会い人数
const AS = sheetForm.getRange(lastRow,45).getValue();     // お客様の状況

// 部品手配依頼
const checkMark = '✔';
const  P = sheetForm.getRange(lastRow,16).getValue();     // 依頼内容
const  q = sheetForm.getRange(lastRow,17).getValue();     // 発送希望納期
const  Q = DateFormat1(q);                                // 〃（表示を指定）
const  R = sheetForm.getRange(lastRow,18).getValue();     // 発送方法
const  S = sheetForm.getRange(lastRow,19).getValue();     // 事前連絡
const  T = sheetForm.getRange(lastRow,20).getValue();     // 見積もり番号
const AT = sheetForm.getRange(lastRow,46).getValue();     // お客様の状況

// 見積もり＋部品手配依頼
const  U = sheetForm.getRange(lastRow,21).getValue();     // 依頼種類
const  V = sheetForm.getRange(lastRow,22).getValue();     // 依頼内容
const  w = sheetForm.getRange(lastRow,23).getValue();     // 作成期限
const  W = DateFormat1(w);                                // 〃（表示を指定）   
const ap = sheetForm.getRange(lastRow,42).getValue();     // 発送希望納期
const AP = DateFormat1(ap);                               // 〃（表示を指定）     
const  X = sheetForm.getRange(lastRow,24).getValue();     // 受注予想確率
const  Y = sheetForm.getRange(lastRow,25).getValue();     // 作業日
const  Z = sheetForm.getRange(lastRow,26).getValue();     // 作業日数
const AA = sheetForm.getRange(lastRow,27).getValue();     // 作業人数
const AB = sheetForm.getRange(lastRow,28).getValue();     // 部品発送
const AC = sheetForm.getRange(lastRow,29).getValue();     // 事前連絡
const AD = sheetForm.getRange(lastRow,30).getValue();     // 見積もり番号
const AJ = sheetForm.getRange(lastRow,36).getValue();     // 立会い日数
const AK = sheetForm.getRange(lastRow,37).getValue();     // 立会い人数
const AU = sheetForm.getRange(lastRow,47).getValue();     // お客様の状況

// 事後見積もり、作成済み見積もり提出依頼
const AE = sheetForm.getRange(lastRow,31).getValue();     // 作業費
const AF = sheetForm.getRange(lastRow,32).getValue();     // 見積もり番号
const AG = sheetForm.getRange(lastRow,33).getValue();     // 部品費
const AH = sheetForm.getRange(lastRow,34).getValue();     // ブリッジ番号
const AM = sheetForm.getRange(lastRow,39).getValue();     // 備考欄 (事後見積もり）
const AN = sheetForm.getRange(lastRow,40).getValue();     // 備考欄 (作成済み見積もりの確認と提出依頼）

// その他
const AV = sheetForm.getRange(lastRow,48).getValue();     // 依頼内容

// 取得情報 → フォーム書込用
let getPoints = [];


// ========================================================================= //
//                           メイン関数を実行                                   //
// ========================================================================= //

// 変更項目
const fillColor = "yellow";                // 塗り潰しの色を指定
const destination = "support@isowa.co.jp"; // 送信先のアドレスを指定
const admin = "k.kamikura@isowa.co.jp"     // アドレス取得出来なかった場合の送信先 


function Main() {
  
  // 依頼書フォームを複製(スプレットシート)
  const formItems = CopyForm(); 
  
  // 複製したスプレットシートに情報を記入
  SetCommonPoints();                                 // 共通記入箇所
  if ( H == '見積もり依頼' ) ReqQuote();                // 見積もり依頼の場合に実行
  if ( H == '部品手配' ) PartsArrange();              // 部品手配の場合に実行
  if ( H == '見積もり依頼と部品手配' ) QuotePartsArr();   // 見積もり + 部品手配の場合に実行
  if ( H == 'その他' ) Other();                       // その他の場合に実行 
  
  // 回答フォームの添付ファイルを取得
  let attachments = GetAttachment();

  // 複製した依頼書フォームをPDF変換してファイル出力
  let requestForm = PdfCreate(formItems[1]);
  
  // 依頼書フォーム(PDF)を添付ファイルに追加
  attachments.push(requestForm);
  
  // 不要な添付ファイル ・ 複製した依頼書フォーム(スプレットシート)を削除
  DeleteFile(formItems[0]);
  
  // 依頼書を添付してメールを送信
  SendMail(attachments);

}

// ========================================================================= //
//                              ログの確認用                                   //
// ========================================================================= //

console.log(`B:${B}`);
console.log(`C:${C}`);
console.log(`D:${D}`);
console.log(`E:${E}`);
console.log(`F:${F}`);
console.log(`G:${G}`);
console.log(`H:${H}`);
console.log(`I:${I}`);
console.log(`G:${G}`);
console.log(`K:${K}`);
console.log(`L:${L}`);
console.log(`M:${M}`);
console.log(`N:${N}`);
console.log(`O:${O}`);
console.log(`P:${P}`);
console.log(`Q:${Q}`);
console.log(`R:${R}`);
console.log(`S:${S}`);
console.log(`T:${T}`);
console.log(`U:${U}`);
console.log(`V:${V}`);
console.log(`W:${W}`);
console.log(`X:${X}`);
console.log(`Y:${Y}`);
console.log(`Z:${Z}`);
console.log(`AA:${AA}`);
console.log(`AB:${AB}`);
console.log(`AC:${AC}`);
console.log(`AD:${AD}`);
console.log(`AE:${AE}`);
console.log(`AF:${AF}`);
console.log(`AG:${AG}`);
console.log(`AH:${AH}`);
console.log(`AI:${AI}`);
console.log(`AG:${AG}`);
console.log(`AK:${AK}`);
console.log(`AL:${AL}`);
console.log(`AM:${AM}`);
console.log(`AN:${AN}`);
console.log(`AO:${AO}`);
console.log(`AP:${AP}`);
console.log(`AQ:${AQ}`);
console.log(`AS:${AS}`);
console.log(`AT:${AT}`);
console.log(`AU:${AU}`);
console.log(`AV:${AV}`);
