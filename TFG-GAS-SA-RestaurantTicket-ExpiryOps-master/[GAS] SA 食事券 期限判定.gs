//test 
//役割：デイリーで食事券の管理簿をチェック。期限切れがあれば管理簿を更新して、SA@にメール通知。
//トリガー：SA管理：毎日19時-20時。
//レポ：YSJP/TFG-GAS-SA-RestaurantTicket-ExpiryOps

function kigenHantei() {

  //管理簿シート
  const idKanribo = "1SRukLe3GHVyygXGZiIUxbJ2Eg8N71tuwgb9BcFSwSYw";
  const sheets = getBillSheets(idKanribo);

  //判定日=今日のミリ秒
  let dateCurrent = new Date();
  dateCurrent.setHours(0,0,0,0);
  const msRef = dateCurrent.getTime();

  //期限切れ食事券：判定と管理簿の更新
  let expired = "";　//期限切れ食事券の情報（メール本文用）
  for (let i = 0; i < sheets.length; i++) {
  //各シートを見ていくよ
    const lastRow = sheets[i].getRange("B:B").getValues().filter(String).length;
    const data = sheets[i].getRange("A1:O"+lastRow.toString()).getValues();
    for (let j = 1; j < data.length; j++) {
    //各行を見ていくよ
      var kigen = "";
      if(typeof data[j][5] == "string"){  //"破棄"などの文字列が.getTime()を呼ぶとエラーを起こすので型判定のif
        kigen = data[j][5];
      } else {
        kigen = data[j][5].getTime(); //有効期限
      }
      const shito = data[j][10];　//使途
      const shiyoubi = data[j][13];　//使用日
      if (kigen == msRef && shiyoubi == "") {
      //期限が今日 && 未使用の場合
        if (shito == "販売") {
        //販売した食事券の処理：雑収入
          sheets[i].getRange("N"+(j+1).toString()+":O"+(j+1).toString()).setValues([["無効","雑収入"]]).setFontColor('red');  //N列 使用日 & O列 経理処理方法
          expired += "\n シート："+sheets[i].getSheetName()+" 　No."+data[j][0]+" 　顧客: "+data[j][7]+" 　処理: 雑収入";
        } else {
        //販売以外の食事券の処理：無効化
          sheets[i].getRange("N" + (j + 1).toString() + ":O" + (j + 1).toString()).setValues([["無効","無効"]]).setFontColor('red');  //N列 使用日 & O列 経理処理方法
          expired += "\n シート："+sheets[i].getSheetName()+" 　No."+data[j][0]+" 　顧客: "+data[j][7]+" 　処理: 無効";
        }
      }
    }//次の行へ
  }//次のシートへ

  // バックアップ生成
  const ssBackup = getBackup(dateCurrent,idKanribo);
  const sheetsBackup = getBillSheets(ssBackup.getId());

  //集計用シート
  var shukei_sheets = []; //集計シートの配列
  for(i=0; i<sheetsBackup.length; i++){
    shukei_sheets += sheetsBackup[i].copyTo(ssBackup).setName(sheetsBackup[i].getName()+"集計用");
    ssBackup.moveActiveSheet(6+i);
  }

　//メール処理
  Logger.log("expired:"+expired);
  if(expired !== ""){

    //メール
    const from = "y-shinohara@fujiyagohonjin.co.jp";
    const to = "y-shinohara@fujiyagohonjin.co.jp,"; //テスト用
    //const to = "sa@fujiyagohonjin.co.jp,";
    const cc = "";
    const bcc = "";
    const nengappi = getNengappi(dateCurrent);
    const youbi = getYoubi(dateCurrent);
    const mailTitle = "[食事券] 有効期限処理を実行しました。 " + nengappi + "(" + youbi + ")";
    const mailBody = getMailBody(dateCurrent, expired);
    GmailApp.sendEmail(to, mailTitle, mailBody,
      {
        bcc: from,
        cc: cc,
        bcc: bcc,
        from: from,
        name: '篠原（システム配信）'
      }
    )
    Logger.log("本日で期限切れの食事券があります。データを sa@ にメールしました。");
/*
    Logger.log("テストなのでメールは送らないでおきます。");
    Logger.log("====================\n"+body+"====================\n");
*/
  } else {
    Logger.log("本日で期限切れの食事券はありませんでした。");
  }

} //main


function getBillSheets(idSS){
  const ss = SpreadsheetApp.openById(idSS);
  const billTypes = ["1","2","5","10"];
  const sheets = [];
  Logger.log("スプレッドシート『" + ss.getName() + "』を処理中。");
  for(let i=0;i<billTypes.length;i++){
    const sheet = ss.getSheetByName(billTypes[i] + "000円券");
    sheets.push(sheet);
    Logger.log(sheet.getName() + "を配列に追加しました。");
  }
  return sheets;
}

function getYoubi(d){
  const id = Utilities.formatDate(d, "JST", "u");
  const table = new Array("日", "月", "火", "水", "木", "金", "土", "日");
  const youbi = table[id];
  return youbi;
}

function getNengappi(d){
  const nengappi = Utilities.formatDate(d, "JST", "YYYY'年'M'月'd'日'");
  return nengappi;
}

function getTitle(d){
  const nengappi = getNengappi(d);
  const youbi = getYoubi(d);
  const title = "[バックアップ] 食事券 管理簿 " + nengappi + "(" + youbi + ")" + " 期限処理後";
  return title;
}

function getBackup(d,idKanribo){
  const idSaveToFolder = "0B86gPARieR2NX05KU0JIUUtWZnM";
  const saveToFolder = DriveApp.getFolderById(idSaveToFolder);
  const title = getTitle(d);
  const fileBackup = DriveApp.getFileById(idKanribo).makeCopy(title, saveToFolder);
  const idBackup = fileBackup.getId();
  const ssBackup = SpreadsheetApp.openById(idBackup);
  return ssBackup;
}

function getMailBody(d, expired){
  const nengappi = getNengappi(d);
  const youbi = getYoubi(d);
  let body = "SA 各位\n";
  body += "\n";
  body +=  nengappi + "(" + youbi + ")\n";
  body += "食事券の有効期限処理を、下記の通り実行しました。\n";
  body += "\n";
  body += "クラウド上の管理簿をご確認の上、内容がOKな場合は赤字を黒字に更新してください。\n"
  body += "\n";
  body += "▼管理簿URL(要ログイン)\n";
  body += "https://goo.gl/hG5jAd \n";
  body += "\n";
  body += "====================\n";
  body += "【実行一覧】\n";
  body += expired;
  body += "\n";
  body += "====================\n";
  body += "\n";
  body += "※ 本メールは自動配信されています。\n";
  body += "\n";
  body += "篠原（システム配信）\n";
  body += "\n";
  return body;
}