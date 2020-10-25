// ========================================================================= //
//                   iパーツ依頼書を添付してメールを送信                           //
// ========================================================================= //

function SendMail(attachments) {
  
  // ISOWAのアドレスリストを取得
  const ssAdress = SpreadsheetApp.openById('1jx4T6lKn3tCwAFHq25JHuGctEFiehyAPFBnuVYeXimo');
  const sheetAdress = ssAdress.getSheetByName('メールアドレス一覧（ISOWA）');
  
  // メールアドレス一覧の情報を取得する。
  const names = sheetAdress.getRange(2, 2, 1100, 2).getValues();
  
  // シートの行と列を反転させる。
  const _ = Underscore.load();
  const namesTrans = _.zip.apply(_, names);
  
  // フォーム回答者
  const selectedName = C;
  
  // アドレス一覧からフォーム回答者と一致した番号を取得
  const namesNumber = namesTrans[0].indexOf(selectedName);
  
  // フォーム回答者のアドレスを取得
  let selectedAdress;
  if ( namesNumber !== -1 ) {
    selectedAdress = namesTrans[1][namesNumber];
  } else {
    selectedAdress = admin;
  }  
  
  // オプションでフォームから追加されたアドレス名を取得する。
  const opNames = AQ;  // フォームに記入された宛先
  let opAdress  = '';  // 指定アドレスの初期値
  
  // アドレス名が入っていたら実行する。
  if ( opNames !== '' ) {
    
    const op = opNames.split(",");       // アドレスを個別に分割（アドレスの数を取得するため）
    
    for ( i = 0; i < op.length; i++ ) {   // アドレスの数だけループして一致した番号を返す。
      const opName = opNames.split(", ")[i];
      const opNamesNumber = namesTrans[0].indexOf(opName);
      
      // 一致した番号があれば実行する。
      if ( opNamesNumber !== -1 ) {
        const _opAdress = namesTrans[1][opNamesNumber];
        
        // 初回のみ実行する。
        if ( opAdress == '' ) {
          opAdress = _opAdress;
          
          // 2回目以降実行する。
        } else {
          opAdress = `${opAdress}, ${_opAdress}`; 
        }
      }
    }
    
    // アドレス名が入ってなければ実行する。  
  } else {
    opAdress = admin;
  }
  
  
  // 送信先、タイトル、本文
  const to = destination;
  const subject = '【依頼】iパーツ依頼書 ${D} ${G}'
  .replace('${D}', D)
  .replace('${G}', G);
  const body = '\
ＩＳＯＷＡお客様サポート窓口　担当者様\n\n\n\n\
添付の依頼をしますので宜しくお願いします。\n\n\n\
以上、よろしくお願いします。'

  // オプション
  const options = {
    cc: opAdress,             // 送信者
    bcc: selectedAdress,      // 追加送信アドレス
    name: selectedName,       // 送信者名
    attachments: attachments  // 添付ファイル
  };
  
  // メール送信
  GmailApp.sendEmail(
    to,
    subject,
    body,
    options
  );
  
}