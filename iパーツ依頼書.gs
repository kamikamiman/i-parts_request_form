function getInfo() {
  
// ======================  スプレットシートを取得  ============================= //
  
  // スプレットシートの情報を取得する。
  var getForm = SpreadsheetApp.openById('1RCwuHFfg4TcMsqY9dHcYpVa8BD52SmFMsPbJJn6Hu8Y');
  var sheetForm = getForm.getSheetByName('iパーツ依頼書(回答)');
  
  // スプレットシートの最終行目の情報を取得する。
  var lastRow = sheetForm.getLastRow();
  
  
// =====  ⅰパーツ依頼書フォームから書き込まれたスプレットシートの情報を取得する。  ===== //
  
  // フォームからスプレットシートに書き込まれた各セル情報を取得する。
  var A = sheetForm.getRange(lastRow, 1).getValue();     // タイムスタンプ
  // 共通記入箇所
  var date = new Date();
  date = Utilities.formatDate( date, 'Asia/Tokyo', 'yyyy/MM/dd'); // 本日の日付
  var B = sheetForm.getRange(lastRow, 2).getValue();     // 依頼者部署
  var C = sheetForm.getRange(lastRow, 3).getValue();     // 依頼者名
  var D = sheetForm.getRange(lastRow, 4).getValue();     // お客様会社名
  var E = sheetForm.getRange(lastRow, 5).getValue();     // お客様担当者
  var F = sheetForm.getRange(lastRow, 6).getValue();     // 機械
  var G = sheetForm.getRange(lastRow, 7).getValue();     // 機種
  var H = sheetForm.getRange(lastRow, 8).getValue();     // 依頼要旨
  var AO = sheetForm.getRange(lastRow, 41).getValue();   // 添付ファイル
  var AQ = sheetForm.getRange(lastRow, 43).getValue();   // 送信アドレス(cc)

  // 見積もり依頼
  var I = sheetForm.getRange(lastRow, 9).getValue();     // 依頼種類
  var J = sheetForm.getRange(lastRow,10).getValue();     // 依頼内容
  var k = sheetForm.getRange(lastRow,11).getValue();     // 作成期限
  if(k !== "") {
    var K = Utilities.formatDate(k,'JST' ,'yyyy/MM/dd'); // 時間表示を設定する。
  }
  var L = sheetForm.getRange(lastRow,12).getValue();     // 受注予想確率
  var M = sheetForm.getRange(lastRow,13).getValue();     // 作業日
  var N = sheetForm.getRange(lastRow,14).getValue();     // 作業日数
  var O = sheetForm.getRange(lastRow,15).getValue();     // 作業人数
  var AI = sheetForm.getRange(lastRow,35).getValue();    // 立会い日数
  var AL = sheetForm.getRange(lastRow,38).getValue();    // 立会い人数
  var AS = sheetForm.getRange(lastRow,45).getValue();   // お客様の状況
  
  // 部品手配依頼
  var P = sheetForm.getRange(lastRow,16).getValue();     // 依頼内容
  var q = sheetForm.getRange(lastRow,17).getValue();
  if(q !== "") {
  var Q = Utilities.formatDate(q,'JST' ,'yyyy/MM/dd');   // 発送希望納期
  }
  var R = sheetForm.getRange(lastRow,18).getValue();     // 発送方法
  var S = sheetForm.getRange(lastRow,19).getValue();     // 事前連絡
  var T = sheetForm.getRange(lastRow,20).getValue();     // 見積もり番号
  var AT = sheetForm.getRange(lastRow,46).getValue();    // お客様の状況
  
  // 見積もり＋部品手配依頼
  var U = sheetForm.getRange(lastRow,21).getValue();     // 依頼種類
  var V = sheetForm.getRange(lastRow,22).getValue();     // 依頼内容
  var w = sheetForm.getRange(lastRow,23).getValue();
  if(w !== "") {
  var W = Utilities.formatDate(w,'JST' ,'yyyy/MM/dd');   // 作成期限
  }
  var ap = sheetForm.getRange(lastRow,42).getValue();
  if(ap !== "") {
  var AP = Utilities.formatDate(ap,'JST' ,'yyyy/MM/dd'); // 発送希望納期
  } 
  var X = sheetForm.getRange(lastRow,24).getValue();     // 受注予想確率
  var Y = sheetForm.getRange(lastRow,25).getValue();     // 作業日
  var Z = sheetForm.getRange(lastRow,26).getValue();     // 作業日数
  var AA = sheetForm.getRange(lastRow,27).getValue();    // 作業人数
  var AB = sheetForm.getRange(lastRow,28).getValue();    // 部品発送
  var AC = sheetForm.getRange(lastRow,29).getValue();    // 事前連絡
  var AD = sheetForm.getRange(lastRow,30).getValue();    // 見積もり番号
  var AJ = sheetForm.getRange(lastRow,36).getValue();    // 立会い日数
  var AK = sheetForm.getRange(lastRow,37).getValue();    // 立会い人数
  var AU = sheetForm.getRange(lastRow,47).getValue();    // お客様の状況
  
  // 事後見積もり、作成済み見積もり提出依頼
  var AE = sheetForm.getRange(lastRow,31).getValue();    // 作業費
  var AF = sheetForm.getRange(lastRow,32).getValue();    // 見積もり番号
  var AG = sheetForm.getRange(lastRow,33).getValue();    // 部品費
  var AH = sheetForm.getRange(lastRow,34).getValue();    // ブリッジ番号
  var AM = sheetForm.getRange(lastRow,39).getValue();    // 備考欄 (事後見積もり）
  var AN = sheetForm.getRange(lastRow,40).getValue();    // 備考欄 (作成済み見積もりの確認と提出依頼）
  
  // その他
  var AV = sheetForm.getRange(lastRow,48).getValue();    // 依頼内容

  
  // 変数を設定
  var checkMark = '✔';
  
  // フォーム回答を書き出したスプレットシートのデータを別のシートに書き込む。
  var setForm = SpreadsheetApp.openById('1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ');
  var setSheetForm = setForm.getSheetByName('iパーツ依頼フォーム');
  
  // iパーツ依頼書フォームの複製
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_hhmmss');
  var sheetName = now+'_'+'iパーツ依頼書';
  var setSheetForm = setForm.duplicateActiveSheet().setName(sheetName);
  var sheetId = setSheetForm.getSheetId();
  setSheetForm.activate();
  
  
// ======================  iパーツ依頼書フォームに記入する。  ============================= //
  
  // ---  共通記入箇所  -------------------------------- // 
  
  setSheetForm.getRange('B7').setValue(date);  // 依頼日
  setSheetForm.getRange('K7').setValue(B);  // 依頼部署
  setSheetForm.getRange('U7').setValue(C);  // 依頼者名
  setSheetForm.getRange('H9').setValue(D);  // 得意先名
  setSheetForm.getRange('V9').setValue(E);  // お客様担当者
//  setSheetForm.getRange('AE7').setValue(F); // 機械
  setSheetForm.getRange('AB9').setValue(G); // 機種・号機・ユニット  
  
  // ---  依頼種類  ----------------------------------- //
  
  // *********  見積もり依頼のみの処理  ********** //
  
  if(H == '見積もり依頼') {
    setSheetForm.getRange('H3').setBackground("yellow");
    setSheetForm.getRange('I19').setValue(K);  // 作成期限
    setSheetForm.getRange('Q16').setValue(L);  // 受注予想確率
    setSheetForm.getRange('G27').setValue(AE); // 作業費　　　（事後見積もり）
    setSheetForm.getRange('G30').setValue(AG); // 部品費　　　（事後見積もり）
    setSheetForm.getRange('H32').setValue(AH); // ブリッジ番号（事後見積もり）
    setSheetForm.getRange('G37').setValue(AF); // 見積もり番号（作成済み見積もりの提出）
    setSheetForm.getRange('B40').setValue(J);  // 依頼内容
    setSheetForm.getRange('I51').setValue(AS); // お客様の状況
    
    // 依頼種類の処理
    if(I == '部品のみ') {
      setSheetForm.getRange('C21').setBackground("yellow"); // *** 「部品のみ」塗り潰す。*** //
      
    } else if(I == '事後見積もり') {
      setSheetForm.getRange('C25').setBackground("yellow"); // *** 「事後見積もり」塗り潰す。*** //
      setSheetForm.getRange('B40').setValue(AM);            // 備考
      
    } else if(I == '作成済み見積もりの確認と提出') {
      setSheetForm.getRange('C35').setBackground("yellow"); // *** 「作成済み見積もりの確認と提出」塗り潰す。*** //
      setSheetForm.getRange('B40').setValue(AN);             // 備考
      
    } else if(I == '部品と作業') {
      setSheetForm.getRange('P21').setBackground("yellow"); // *** 「部品と作業」塗り潰す。*** //
      setSheetForm.getRange('W23').setValue(N);             // 作業日数
      setSheetForm.getRange('Z23').setValue(O);             // 作業人数
      setSheetForm.getRange('W25').setValue(AI);            // 立会日数
      setSheetForm.getRange('Z25').setValue(AL);            // 立会人数
      
      if(AL != '無し') {
        if(AL != '') {
          setSheetForm.getRange('S25').setBackground("yellow");  // 立会「平日」塗り潰す。
        }
      }
      
      if(M == '休日（前泊移動）') {
        setSheetForm.getRange('AA21').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('S23').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（前泊移動）') {
        setSheetForm.getRange('AA21').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('P23').setBackground("yellow");  // 「平日」塗り潰す。
      } else if(M == '休日（当日移動）') {
        setSheetForm.getRange('AG21').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('S23').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（当日移動）') {
        setSheetForm.getRange('AG21').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('P23').setBackground("yellow");  // 「平日」塗り潰す。
      }
      
    } else if(I == '作業のみ') {
      setSheetForm.getRange('P30').setBackground("yellow");    // *** 「作業のみ」塗り潰す。*** //
      setSheetForm.getRange('W23').setValue(N);                // 作業日数
      setSheetForm.getRange('Z23').setValue(O);                // 作業人数
      setSheetForm.getRange('W25').setValue(AI);               // 立会日数
      setSheetForm.getRange('Z25').setValue(AL);               // 立会人数
      
      if(AL != '無し') {
        if(AL != '') {
          setSheetForm.getRange('S35').setBackground("yellow");  // 立会「平日」塗り潰す。
        }
      }
      if(M == '休日（前泊移動）') {
        setSheetForm.getRange('AA30').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('S32').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（前泊移動）') {
        setSheetForm.getRange('AA30').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('P32').setBackground("yellow");  // 「平日」塗り潰す。
      } else if(M == '休日（当日移動）') {
        setSheetForm.getRange('AG30').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('S32').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(M == '平日（当日移動）') {
        setSheetForm.getRange('AG30').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('P32').setBackground("yellow");  // 「平日」塗り潰す。
      }
      
    }
    
    
    // *********  部品手配のみの処理  ********** //
    
  } else if(H == '部品手配') {
    var cell = setSheetForm.getRange('L3').setBackground("yellow"); // セル「発注」を黄色で塗り潰す。
    setSheetForm.getRange('I13').setValue(Q);                       // 納期希望日
    setSheetForm.getRange('AG13').setValue(T);                      // 見積もり番号
    setSheetForm.getRange('B40').setValue(P);                       // 依頼内容
    setSheetForm.getRange('I51').setValue(AT);                      // お客様の状況    
    
    // 部品発送方法の処理
    if(R == '直送') {
      setSheetForm.getRange('S13').setBackground("yellow"); // 直送
    } else {
      setSheetForm.getRange('V13').setBackground("yellow"); // 工事持参
    }
    
    // 事前連絡の有無の処理
    if(S == '必要') {
      setSheetForm.getRange('Y13').setBackground("yellow"); // 必要
    } 
    
    
    
    // *********  部品手配＋見積もり依頼の処理  ********** // 
    
  } else if(H == '見積もり依頼と部品手配') {
    setSheetForm.getRange('H3').setBackground("yellow"); 
    setSheetForm.getRange('L3').setBackground("yellow"); 
    setSheetForm.getRange('I11').setBackground("yellow");
    setSheetForm.getRange('I11').setValue(checkMark);
    setSheetForm.getRange('I19').setValue(W);            // 作成期限
    setSheetForm.getRange('I13').setValue(AP);           // 納期希望日
    setSheetForm.getRange('Q16').setValue(X);            // 受注予想確率
    setSheetForm.getRange('AG13').setValue(AD);          // 見積もり番号
    setSheetForm.getRange('B40').setValue(V);            // 依頼内容
    setSheetForm.getRange('I51').setValue(AU);           // お客様の状況
    
    // 依頼種類による処理
    if(U == '部品のみ') {
      setSheetForm.getRange('C21').setBackground("yellow"); // *** 「部品のみ」塗り潰す。*** //
    } else if(U == '事後見積もり') {
      setSheetForm.getRange('C25').setBackground("yellow"); // *** 「事後見積もり」塗り潰す。*** //
      
    } else if(U == '作成済み見積もりの確認と提出') {
      setSheetForm.getRange('C35').setBackground("yellow"); // *** 「作成済み見積もりの確認と提出」塗り潰す。*** //
      
    } else if(U == '部品と作業') {
      setSheetForm.getRange('P21').setBackground("yellow"); // *** 「部品と作業」塗り潰す。*** //
      setSheetForm.getRange('W23').setValue(Z);             // 作業日数
      setSheetForm.getRange('Z23').setValue(AA);            // 作業人数
      setSheetForm.getRange('W25').setValue(AJ);            // 立会日数
      setSheetForm.getRange('Z25').setValue(AK);            // 立会人数
      
      if(AK != '無し') {
        if(AK != '') {
          setSheetForm.getRange('S25').setBackground("yellow");  // 立会「平日」塗り潰す。
        }
      }
      if(Y == '休日（前泊移動）') {
        setSheetForm.getRange('AA21').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('S23').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（前泊移動）') {
        setSheetForm.getRange('AA21').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('P23').setBackground("yellow");  // 「平日」塗り潰す。
      } else if(Y == '休日（当日移動）') {
        setSheetForm.getRange('AG21').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('S23').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（当日移動）') {
        setSheetForm.getRange('AG21').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('P23').setBackground("yellow");  // 「平日」塗り潰す。
      }
      
    } else {
      setSheetForm.getRange('P30').setBackground("yellow");    // *** 「作業のみ」塗り潰す。*** //
      setSheetForm.getRange('W32').setValue(Z);                // 作業日数
      setSheetForm.getRange('Z32').setValue(AA);               // 作業人数
      setSheetForm.getRange('W35').setValue(AJ);               // 立会日数
      setSheetForm.getRange('Z35').setValue(AK);               // 立会人数
      
      if(AK != '無し') {
        if(AK != '') {
          setSheetForm.getRange('S35').setBackground("yellow");  // 立会「平日」塗り潰す。
        }
      }
      if(Y == '休日（前泊移動）') {
        setSheetForm.getRange('AA30').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('S32').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（前泊移動）') {
        setSheetForm.getRange('AA30').setBackground("yellow"); // 「前泊あり」塗り潰す。
        setSheetForm.getRange('P32').setBackground("yellow");  // 「平日」塗り潰す。
      } else if(Y == '休日（当日移動）') {
        setSheetForm.getRange('AG30').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('S32').setBackground("yellow");  // 「休日」塗り潰す。
      } else if(Y == '平日（当日移動）') {
        setSheetForm.getRange('AG30').setBackground("yellow"); // 「当日移動」塗り潰す。
        setSheetForm.getRange('P32').setBackground("yellow");  // 「平日」塗り潰す。
      }
    }
    
    // 部品発送方法の処理
    if(AB == '直送') {
      setSheetForm.getRange('S13').setBackground("yellow"); // 直送
    } else {
      setSheetForm.getRange('V13').setBackground("yellow"); // 工事持参
    }
    
    // 事前連絡の有無の処理
    if(AC == '必要') {
      setSheetForm.getRange('Y13').setBackground("yellow"); // 必要
    }
    
  // その他の処理 
  } else if(H == 'その他') {
    setSheetForm.getRange('V3').setBackground("yellow"); // 「その他」塗り潰す。
    setSheetForm.getRange('B40').setValue(AV);           // 依頼内容  
  }
  
  
// *********  pdfファイルを出力  ********** // 
  
  SpreadsheetApp.flush();
  var url = 'https://docs.google.com/spreadsheets/d/1TaPz65cp5neRY7cOrCNCm_pQMeDGLMrEVdLwBxge7RQ/export?exportFormat=pdf&gid=SID'.replace('SID', sheetId);
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers:{
      'Authorization': 'Bearer '+token
    }
  });
  var blob = response.getBlob().setName(now+'_'+C+'_'+'iパーツ依頼書.pdf');    // pdfの名前
  var folder = DriveApp.getFolderById('1mJKjawfQUio1ZEDsFq7re-t5cwm3xfyK');  // pdfの保存先フォルダを指定
  var requestForm = folder.createFile(blob);                                 // フォルダ内にiパーツ依頼書を作成
  
  
// *********  添付ファイルを取得  ********** // 
  
  // 添付ファイルを取得（単体）
//  var fileId = AO.split('=')[1];
//  var file = DriveApp.getFileById(fileId));

  // 添付ファイルを取得（複数）
    if(AO !== '') {
      var files = AO.split(',');
      var fileIds = files.map(function(file){
        var fileId = file.split('=')[1]; // 個別のファイルIDを取得
        return DriveApp.getFileById(fileId);  
      });
   
//    Logger.log(fileId);
    
    
// *********  保存された余分なファイルの削除処理  ********** // 
    
    // 保存された添付ファイルを削除  
    var root = DriveApp.getRootFolder().getFiles();
    while(root.hasNext()) {
      var rootFile = root.next();
        rootFile.setTrashed(true);
//      DriveApp.getRootFolder().removeFile(rootFile);
    }      
  } else {
    var fileIds = [];
  }
//  Logger.log(fileIds);
//  Logger.log(AO);
  fileIds.push(requestForm);
  
  // 複製したスプレットシートを削除
  var sheet = setForm.getSheetByName(sheetName);
  setForm.deleteSheet(sheet);  
  
  
  
// ======================  pdfファイルを添付してメールを送信する。  ============================= //
  
  // CCの送信先を指定する。
  var ssAdress = SpreadsheetApp.openById('1jx4T6lKn3tCwAFHq25JHuGctEFiehyAPFBnuVYeXimo');
  var sheetAdress = ssAdress.getSheetByName('メールアドレス一覧（ISOWA）');
  // メールアドレス一覧から送信対象者の情報を取得する。
  var names = sheetAdress.getRange(2, 2, 1100, 2).getValues();
    // スプレットシートの行と列を反転させる。
  var _ = Underscore.load();
  var namesTrans = _.zip.apply(_, names);
    // 選択された送信対象者から送信先メールアドレスの情報を取得する。
  var selectedName = C;
    // アドレスリストから名前が一致した番号を取得する。
  var namesNumber = namesTrans[0].indexOf(selectedName);
    // 一致した番号が取得できたら、その番号のアドレスを取得し、
    // 番号が取得できなかったら、特定のアドレスを取得する。
  if(namesNumber !== -1) {
    var selectedAdress = namesTrans[1][namesNumber];
  } else {
    var selectedAdress = 'k.kamikura@isowa.co.jp';
  }  
//  Logger.log('------------');
//  Logger.log(selectedName);
//  Logger.log(namesNumber);
//  Logger.log(selectedAdress);
  
  
  
  // オプションでフォームから追加されたアドレス名を取得する。
  
  var opNames = AQ; // フォームに記入されたアドレス
  var opAdress = '';// 指定アドレスの初期値

  // アドレス名が入っていたら実行する。
  if(opNames !== '') {
    var op = opNames.split(",");      // アドレスを個別に分割（アドレスの数を取得するため）

    for(i = 0; i < op.length; i++){   // アドレスの数だけループして一致した番号を返す。
      var opName = opNames.split(", ")[i];
      var opNamesNumber = namesTrans[0].indexOf(opName);
      // 一致した番号があれば実行する。
      if(opNamesNumber !== -1) {
        var _opAdress = namesTrans[1][opNamesNumber];
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
      var opAdress = 'k.kamikura@isowa.co.jp';
    }


// 送信先、タイトル、本文  
  var to = 'support@isowa.co.jp';
//  var to = 'k.kamikura@isowa.co.jp';
  var subject = '【依頼】iパーツ依頼書 ${D} ${G}'
                .replace('${D}', D)
                .replace('${G}', G);
  var body = '\
ＩＳＯＷＡお客様サポート窓口　担当者様\n\n\n\n\
添付の依頼をしますので宜しくお願いします。\n\n\n\
以上、よろしくお願いします。'

// オプション
  var options = {
    cc:opAdress,            // 送信者
    bcc:selectedAdress,     // 追加送信アドレス
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
}