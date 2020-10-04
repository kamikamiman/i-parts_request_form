function test() {

// ======================  スプレットシート情報を取得  ============================= //
  
  // スプレットシートを取得する。
  const getForm = SpreadsheetApp.openById('1RCwuHFfg4TcMsqY9dHcYpVa8BD52SmFMsPbJJn6Hu8Y');
  // スプレットシート ==> シート名を取得する。
  const sheetForm = getForm.getSheetByName('iパーツ依頼書(回答)');
  // スプレットシートの最終行目の情報を取得する。
  const lastRow = sheetForm.getLastRow();
  
  // 読み込むシートと行を取得する。（上記をまとめた変数）
  let sss = 'getForm, sheetForm, lastRow'
  .replace('getForm', getForm)
  .replace('sheetForm', sheetForm)
  .replace('lastRow', lastRow)
 

  
// *********  【関数】回答のセル情報を取得    ********** // 
  function getSS(sss, colnum) {
    return sheetForm.getRange(lastRow, colnum).getValue();
  }
  
// *********  【関数】日付のフォーマット    ********** // 
  function dateFormat(cell) {
    if(cell !== '') {
      return Utilities.formatDate(cell,'JST' ,'yyyy/MM/dd');
    }
  }

  
// *********  【関数】複製したスプレットシートを削除  ********** // 

  function deleteSpreadSheet() {
    let sheet = setForm.getSheetByName(sheetName);
    setForm.deleteSheet(sheet); 
  }
  

// *********  回答フォームの情報を取得  ********** //   
  
  // 共通記入箇所
  let date = new Date();
  date = dateFormat(date); // 本日の日付
  // 変数を設定
  const B = getSS(sss, 2);   // 依頼者部署
  const C = getSS(sss, 3);   // 依頼者名
  const D = getSS(sss, 4);   // お客様会社名
  const E = getSS(sss, 5);   // お客様担当者
  const F = getSS(sss, 6);   // 機械
  const G = getSS(sss, 7);   // 機種
  const H = getSS(sss, 8);   // 依頼要旨
  const AO = getSS(sss,41);  // 添付ファイル
  const AQ = getSS(sss,43);  // 送信アドレス(cc)
  
  // 見積もり依頼  
  const I = getSS(sss, 9);   // 依頼種類
  const J = getSS(sss,10);   // 依頼内容
  const k = getSS(sss,11);   // 作成期限（元情報）
  const K = dateFormat(k);   // 〃（表示を指定）
  const L = getSS(sss,12);   // 受注予想確率
  const M = getSS(sss,13);   // 作業日
  const N = getSS(sss,14);   // 作業日数
  const O = getSS(sss,15);   // 作業人数
  const AI = getSS(sss,35);  // 立会い日数
  const AL = getSS(sss,38);  // 立会い人数
  
  // 部品手配依頼
  const checkMark = '✔';
  const P = getSS(sss,16);   // 依頼内容
  const q = getSS(sss,17);   // 発送希望納期
  const Q = dateFormat(q);   // 〃（表示を指定）
  const R = getSS(sss,18);   // 発送方法
  const S = getSS(sss,19);   // 事前連絡
  const T = getSS(sss,20);   // 見積もり番号
  
  // 見積もり＋部品手配依頼
  var U = getSS(sss,21);   // 依頼種類
  var V = getSS(sss,22);   // 依頼内容
  var w = getSS(sss,23);   // 作成期限
  var W = dateFormat(w);   // 〃（表示を指定）
  var ap = getSS(sss,42);  // 発送希望納期
  var AP = dateFormat(ap); // 〃（表示を指定）
  var X = getSS(sss,24);   // 受注予想確率
  var Y = getSS(sss,25);   // 作業日
  var Z = getSS(sss,26);   // 作業日数
  var AA = getSS(sss,27);  // 作業人数
  var AB = getSS(sss,28);  // 部品発送
  var AC = getSS(sss,29);  // 事前連絡
  var AD = getSS(sss,30);  // 見積もり番号
  var AJ = getSS(sss,36);  // 立会い日数
  var AK = getSS(sss,37);  // 立会い人数
  
  // 事後見積もり、作成済み見積もり提出依頼
  var AE = getSS(sss,31);  // 作業費
  var AF = getSS(sss,32);  // 見積もり番号
  var AG = getSS(sss,33);  // 部品費
  var AH = getSS(sss,34);  // ブリッジ番号
  var AM = getSS(sss,39);  // 備考欄 (事後見積もり）
  var AN = getSS(sss,40);  // 備考欄 (作成済み見積もりの確認と提出依頼）




  
  
  
// ======================  iパーツ依頼書フォームに記入する。  ============================= //  
  
  // iパーツ依頼フォームのシートを読込。
  const setForm = SpreadsheetApp.openById('1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ');
  let setSheetForm = setForm.getSheetByName('iパーツ依頼フォーム');
  
  // iパーツ依頼書フォームのシートを複製。
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_hhmmss');
  const sheetName = now+'_'+'iパーツ依頼書';
  setSheetForm = setForm.duplicateActiveSheet().setName(sheetName);
  const sheetId = setSheetForm.getSheetId();
  setSheetForm.activate(); // スプレットシートを開く（アクティブな状態にする）
  
// *********  【関数】回答の情報をフォームに記入  ********** // 
  function setSS(setCell, getCell) {
    return setSheetForm.getRange(setCell).setValue(getCell);
  }
  
// *********  【関数】回答の情報からフォームに色付け    ********** //
  function setColor(setCell, getCell) {
    return setSheetForm.getRange(setCell).setBackground(getCell);
  }
  
 
// *********  依頼フォームに記入  ********** // 
  
  // ---  共通記入箇所  -------------------------------- // 
  
  setSS('B7', date); // 依頼日
  setSS('K7', B);    // 依頼部署
  setSS('U7', C);    // 依頼者名
  setSS('H9', D);    // 得意先名
  setSS('V9', E);    // お客様担当者
  setSS('AE7', F);   // 機械
  setSS('AB9', G);   // 機種・号機・ユニット  
  
  // ---  依頼種類  ----------------------------------- //
  
  // *********  見積もり依頼のみの処理  ********** //
  
  if(H == '見積もり依頼') {
    setColor('H3', "yellow"); // セル「見積」を塗り潰す。
    setSS('I19', K);          // 作成期限
    setSS('Q16', L);          // 受注予想確率
    setSS('G27', AE);         // 作業費     （事後見積もり）
    setSS('G30', AG);         // 部品費     （事後見積もり）
    setSS('H32', AH);         // ブリッジ番号（事後見積もり）
    setSS('G37', AF);         // 見積もり番号（作成済み見積もりの提出）
    setSS('B40', J);          // 依頼内容
    
    // 依頼種類の処理
    if(I == '部品のみ') {
      setColor('C21', "yellow"); // セル「部品のみ」を塗り潰す。
      
    } else if(I == '事後見積もり') {
      setColor('C25', "yellow"); // セル「事後見積もり」を塗り潰す。
      setSS('B40', AM);          // 備考
      
    } else if(I == '作成済み見積もりの確認と提出') {
      setColor('C35', "yellow"); // セル「作成済み見積もりの確認と提出」を塗り潰す。
      setSS('B40', AN);          // 備考
      
    } else if(I == '部品と作業') {
      setColor('P21', "yellow"); // セル「部品と作業」を塗り潰す。
      setColor('S25', "yellow"); //  立会「平日」を塗り潰す。
      setSS('W23', N);           // 作業日数
      setSS('Z23', O);           // 作業人数
      setSS('W25', AI);          // 立会日数
      setSS('Z25', AL);          // 立会人数
      
      if(M == '休日（前泊移動）') {
        setColor('AA21', "yellow"); // 「前泊あり」を塗り潰す。
        setColor('S23', "yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（前泊移動）') {
        setColor('AA21', "yellow"); // 「前泊あり」塗り潰す。
        setColor('P23', "yellow");  // 「平日」塗り潰す。
      } else if(M == '休日（当日移動）') {
        setColor('AG21', "yellow"); // 「当日移動」塗り潰す。
        setColor('S23', "yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（当日移動）') {
        setColor('AG21', "yellow"); // 「当日移動」塗り潰す。
        setColor('P23', "yellow");  // 「平日」塗り潰す。
      }
      
    } else if(I == '作業のみ') {
      setColor('P30', "yellow");    // セル「作業のみ」を黄色で塗り潰す。
      setSS('W23', N);           // 作業日数
      setSS('Z23', O);           // 作業人数
      setSS('W25', AI);          // 立会日数
      setSS('Z25', AL);          // 立会人数
      
      if(M == '休日（前泊移動）') {
        setColor('AA30', "yellow"); // 「前泊あり」塗り潰す。
        setColor('S32', "yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（前泊移動）') {
        setColor('AA30', "yellow"); // 「前泊あり」塗り潰す。
        setColor('P32', "yellow");  // 「平日」塗り潰す。
      } else if(M == '休日（当日移動）') {
        setColor('AG30', "yellow"); // 「当日移動」塗り潰す。
        setColor('S32', "yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（当日移動）') {
        setColor('AG30', "yellow"); // 「当日移動」塗り潰す。
        setColor('P32', "yellow");  // 「平日」塗り潰す。
      }
    }
    
    
    // *********  部品手配のみの処理  ********** //
    
  } else if(H == '部品手配') {
    const cell = setColor('L3', "yellow"); // セル「発注」を黄色で塗り潰す。
    setSS('I13', Q);                  // 納期希望日
    setSS('AG13', T);                 // 見積もり番号
    setSS('B40', P);                  // 依頼内容    
    
    // 部品発送方法の処理
    if(R == '直送') {
      setSS('S13', "yellow"); // 直送
    } else {
      setSS('V13', "yellow"); // 工事持参
    }
    
    // 事前連絡の有無の処理
    if(S == '必要') {
      setSS('Y13', "yellow"); // 必要
    } 
    
        
    // *********  部品手配＋見積もり依頼の処理  ********** // 
    
  } else if(H == '見積もり依頼と部品手配') {
    setColor('H3', "yellow");      // セル「見積」を黄色で塗り潰す。
    setColor('L3', "yellow");      // セル「部品」を黄色で塗り潰す。
    setColor('I11', "yellow");     // セル「見積もり依頼と同時手配」を黄色で塗り潰す。
    setSS('I11', checkMark);    // 見積もり依頼と同時手配にチェックを入れる。
    setSS('I19', W);            // 作成期限
    setSS('I13', AP);           // 納期希望日
    setSS('Q16', X);            // 受注予想確率
    setSS('AG13', AD);          // 見積もり番号
    setSS('B40', V);            // 依頼内容
    
    // 依頼種類による処理
    if(U == '部品のみ') {
      setColor('C21', "yellow"); // セル「部品のみ」を黄色で塗り潰す。
    } else if(U == '事後見積もり') {
      setColor('C25', "yellow"); // セル「事後見積もり」を黄色で塗り潰す。
      
    } else if(U == '作成済み見積もりの確認と提出') {
      setColor('C35', "yellow"); // セル「作成済み見積もりの確認と提出」を黄色で塗り潰す。
      
    } else if(U == '部品と作業') {
      setColor('P21', "yellow"); // セル「部品と作業」を黄色で塗り潰す。
      setColor('S25', "yellow"); //  立会「平日」塗り潰す。
      setSS('W23', Z);        // 作業日数
      setSS('Z23', AA);       // 作業人数
      setSS('W25', AJ);       // 立会日数
      setSS('Z25', AK);       // 立会人数
      
      if(Y == '休日（前泊移動）') {
        setColor('AA21', "yellow"); // 「前泊あり」塗り潰す。
        setColor('S23', "yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（前泊移動）') {
        setColor('AA21', "yellow"); // 「前泊あり」塗り潰す。
        setColor('P23', "yellow");  // 「平日」塗り潰す。
      } else if(Y == '休日（当日移動）') {
        setColor('AG21', "yellow"); // 「当日移動」塗り潰す。
        setColor('S23', "yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（当日移動）') {
        setColor('AG21', "yellow"); // 「当日移動」塗り潰す。
        setColor('P23', "yellow");  // 「平日」塗り潰す。
      }
      
    } else {
      setColor('P30', "yellow");    // セル「作業のみ」を黄色で塗り潰す。
      setColor('S35', "yellow");    // 立会「平日」塗り潰す。
      setSS('W32', Z);           // 作業日数
      setSS('Z32', AA);          // 作業人数
      setSS('W35', AJ);          // 立会日数
      setSS('Z35', AK);          // 立会人数
      
      if(Y == '休日（前泊移動）') {
        setColor('AA30', "yellow"); // 「前泊あり」塗り潰す。
        setColor('S32', "yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（前泊移動）') {
        setColor('AA30', "yellow"); // 「前泊あり」塗り潰す。
        setColor('P32', "yellow");  // 「平日」塗り潰す。
      } else if(Y == '休日（当日移動）') {
        setColor('AG30', "yellow"); // 「当日移動」塗り潰す。
        setColor('S32', "yellow");   // 「休日」塗り潰す。
      } else if(Y == '平日（当日移動）') {
        setColor('AG30', "yellow"); // 「当日移動」塗り潰す。
        setColor('P32', "yellow");  // 「平日」塗り潰す。
      }
    }
    
    // 部品発送方法の処理
    if(AB == '直送') {
      setColor('S13', "yellow"); // 直送
    } else {
      setColor('V13', "yellow"); // 工事持参
    }
    
    // 事前連絡の有無の処理
    if(AC == '必要') {
      setColor('Y13', "yellow"); // 必要
    } 
  } 
  
  
// pdf作成
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ/export?exportFormat=pdf&gid=SID'.replace('SID', sheetId);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers:{
      'Authorization': 'Bearer '+token
    }
  });
  const blob = response.getBlob().setName(now + '_' + C + '_' + 'iパーツ依頼書.pdf');  // pdfの名前
  const folder = DriveApp.getFolderById('1mJKjawfQUio1ZEDsFq7re-t5cwm3xfyK');            // pdfの保存先フォルダを指定
  const requestForm = folder.createFile(blob);                                           // フォルダ内にiパーツ依頼書を作成
    
// 添付ファイル取得と保存したgoogleDrive直下の添付ファイルを削除
    if(AO !== '') {
        let files = AO.split(',');
        let fileIds = files.map(function(file){
        let fileId = file.split('=')[1]; // 個別のファイルIDを取得
        return DriveApp.getFileById(fileId);  
      });
          // 保存されたgoogleDrive直下の添付ファイルを削除  
      const root = DriveApp.getRootFolder().getFiles();
      while(root.hasNext()) {
        const rootFile = root.next();
        DriveApp.getRootFolder().removeFile(rootFile);
      }      
    } else {
      fileIds = [];
    }
  
  fileIds.push(requestForm);


  
// 複製したスプレットシートの削除
  deleteSpreadSheet();
    
 
  
  
  
// ======================  pdfファイルを添付してメールを送信する。  ============================= //
  
  // CCの送信先を指定する。
  const ssAdress = SpreadsheetApp.openById('1jx4T6lKn3tCwAFHq25JHuGctEFiehyAPFBnuVYeXimo');
  const sheetAdress = ssAdress.getSheetByName('メールアドレス一覧（ISOWA）');
  // メールアドレス一覧から送信対象者の情報を取得する。
  const names = sheetAdress.getRange(2, 2, 1100, 2).getValues();
    // スプレットシートの行と列を反転させる。
  const _ = Underscore.load();
  const namesTrans = _.zip.apply(_, names);
    // 選択された送信対象者から送信先メールアドレスの情報を取得する。
  const selectedName = C;
    // アドレスリストから名前が一致した番号を取得する。
  const namesNumber = namesTrans[0].indexOf(selectedName);
    // 一致した番号が取得できたら、その番号のアドレスを取得し、
    // 番号が取得できなかったら、特定のアドレスを取得する。
  if(namesNumber !== -1) {
    const selectedAdress = namesTrans[1][namesNumber];
  } else {
    const selectedAdress = 'k.kamikura@isowa.co.jp';
  }
  
  
  // 追加アドレスで選択されたアドレスをccに指定する。
  // オプションでフォームから追加されたアドレス名を取得する。
  
  const opNames = AQ; // フォームに記入されたアドレス
  let opAdress = '';// 指定アドレスの初期値

  // アドレス名が入っていたら実行する。
  if(opNames !== '') {
    var op = opNames.split(",");      // アドレスを個別に分割（アドレスの数を取得するため）

    for(i = 0; i < op.length; i++){   // アドレスの数だけループして一致した番号を返す。
      let opName = opNames.split(", ")[i];
      let opNamesNumber = namesTrans[0].indexOf(opName);
      // 一致した番号があれば実行する。
      if(opNamesNumber !== -1) {
        let _opAdress = namesTrans[1][opNamesNumber];
        // 初回のみ実行する。
        if(opAdress == '') {
        opAdress = _opAdress;
        // 2回目以降実行する。
        } else {
        opAdress = opAdress + ', ' + _opAdress; 
        }
      //一致した番号が無ければ実行する。
      } else {
        opAdress = 'k.kamikura@isowa.co.jp';
      }
    }
  // アドレス名が入ってなければ実行する。  
    } else {
      let opAdress = 'k.kamikura@isowa.co.jp';
    }

// 送信先、タイトル、本文  
  const to = 'k-m.natural-h-style@docomo.ne.jp';
  const subject = '【依頼】iパーツ依頼書 ${D} ${G}'
                  .replace('${D}', D)
                  .replace('${G}', G);
  const body = '\
ＩＳＯＷＡお客様サポート窓口　担当者様\n\n\n\n\
添付の依頼をしますので宜しくお願いします。\n\n\n\
以上、よろしくお願いします。'

// オプション
  const options = {
    cc:opAdress,          // 送信者
    name:selectedName,      // 送信者名
    attachments: fileIds    // 添付ファイル
  };
 
// メール送信
  GmailApp.sendEmail(
    to,
    subject,
    body,
    options
  );
  

  
  // ログ確認用
//  Logger.log('B= '+ B);
//  Logger.log('C= '+ C);
//  Logger.log('D= '+ D);
//  Logger.log('E= '+ E);
//  Logger.log('F= '+ F);
//  Logger.log('G= '+ G);
//  Logger.log('H= '+ H);
//  Logger.log('I= '+ I);
//  Logger.log('J= '+ J);
//  Logger.log('K= '+ K);
//  Logger.log('L= '+ L);
//  Logger.log('M= '+ M);
//  Logger.log('N= '+ N);
//  Logger.log('O= '+ O);
//  Logger.log('P= '+ P);
//  Logger.log('Q= '+ Q);
//  Logger.log('R= '+ R);
//  Logger.log('S= '+ S);
//  Logger.log('T= '+ T);  
//  Logger.log('U= '+ U);
//  Logger.log('V= '+ V);
//  Logger.log('W= '+ W);
//  Logger.log('X= '+ X);
//  Logger.log('Y= '+ Y);
//  Logger.log('Z= '+ Z);
//  Logger.log('AA='+ AA);
//  Logger.log('AB='+ AB);
//  Logger.log('AC='+ AC);
//  Logger.log('AD='+ AD);
//  Logger.log('AE='+ AE);
//  Logger.log('AF='+ AF);
//  Logger.log('AG='+ AG);
//  Logger.log('AH='+ AH);
//  Logger.log('AI='+ AI);
//  Logger.log('AJ='+ AJ);
//  Logger.log('AK='+ AK);
//  Logger.log('AL='+ AL);
//  Logger.log('AM='+ AM);
//  Logger.log('AN='+ AN);
//  Logger.log('AO='+ AO);
//  Logger.log('AP='+ AP);
//  Logger.log('AQ='+ AQ);    
}