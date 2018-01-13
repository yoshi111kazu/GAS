function SendReport() {

  var mySS=SpreadsheetApp.getActiveSpreadsheet();
  var sheetDaily=mySS.getSheetByName("Daily"); //Dailyレポートシートを取得
  var rowDaily=sheetDaily.getDataRange().getLastRow(); //最終行を取得

  var yDate = sheetDaily.getRange(rowDaily,1).getValue();

  /* ①レポート :前日のPVとか */
  var strBody = "[info][title]www.icare.jpn.com " +
      Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + "のレポート[/title]" + 　//ga:date
      "[code]PV： " + sheetDaily.getRange(rowDaily,2).getValue() + "\n" + //ga:pageviews
      "User： " + sheetDaily.getRange(rowDaily,3).getValue() + "   /   " + //ga:users
      "Session： " + sheetDaily.getRange(rowDaily,4).getValue() + "\n" + //ga:sessions
      "直帰率： " + Number(sheetDaily.getRange(rowDaily,5).getValue()*100).toFixed(1) + "%   /   " + //ga:bounceRate
      "Session平均： " + Number(sheetDaily.getRange(rowDaily,6).getValue()).toFixed(1) + "秒[/code][hr]" ; //ga:avgSessionDuration

  /* ②レポート :前日の記事別Top10 */
  var sheetPost=mySS.getSheetByName("記事別"); //記事別レポートシートを取得
  var rowKiji=sheetPost.getDataRange().getLastRow();

  strBody = strBody + "[code]";
  for(var i=1;i<=10;i++){
    strBody = strBody + ("    " + sheetPost.getRange(i+15,3).getValue()).slice(-5) + " pv :  " + sheetPost.getRange(i+15,2).getValue() + "\n";
  }
  strBody = strBody + "[/code][hr][code]";

  /* ③レポート :LPのPV */
  var regexp = [
                /^\/$/,
                /^\/services\//,
                /^\/services\/carely\//,
                /^\/services\/stresscheck\//,
                /^\/services\/stresscheck_a\//,
                /^\/services\/stresscheck_b\//,
                /^\/services\/healthcheck\//,
                /^\/services\/sangyoui\//,
                /^\/services\/mysommnie\//,
                /^\/contact\//,
                /^\/contact_carely\//,
                /^\/contact_sangyoui\//,
                /^\/contact_healthcheck\//,
                /^\/contact_mysommnie\//,
                /^\/contact_scafterfollow\//,
                /^\/contact_womanact\//,
                /^\/contact-success\//,
                /^\/case\/$/,
                /^\/case\//
               ];

  for(var j=0;j<regexp.length;j++){
    var pv_total = 0;
    var ses_total = 0;
    var user_total = 0;

    for(var i=16;i<=rowKiji;i++){
      //Logger.log( regexp[j] );

      if( regexp[j].exec(sheetPost.getRange(i,1).getValue()) ) {
        pv_total += sheetPost.getRange(i,3).getValue();
        ses_total += sheetPost.getRange(i,5).getValue();
        user_total += sheetPost.getRange(i,4).getValue();
      }
    }

    //strBody = strBody + ("   " + pv_total).substr(-4) + " pv, " + ses_total + " ses, " + user_total + " user :  " + regexp[j] + "\n";
    strBody = strBody + ("                              " + regexp[j]).slice(-30) + " : " + ("   " + pv_total).slice(-3) + " pv, " + ("   " + ses_total).slice(-3) + " ses, " + ("   " + user_total).slice(-3) + " user\n";
  }
  strBody = strBody + "[/code][hr]";

  /* ④レポート :今月のPV合計 */
  var sheetMonthly=mySS.getSheetByName("Monthly"); //シートを取得
  var rowMonthly=sheetMonthly.getDataRange().getLastRow(); //最終行を取得

  var strBody = strBody +
    "[code]月間累計（" + sheetMonthly.getRange(rowMonthly,1).getValue() + "）\n" +
    "　・PV： " + sheetMonthly.getRange(rowMonthly,2).getValue() + "\n" +
    "　・問合せ(ID:19)： " + sheetMonthly.getRange(rowMonthly,7).getValue() + "件 (CVR: " + Math.round(sheetMonthly.getRange(rowMonthly,8).getValue()*100000)/1000 + "%)\n" +
    "　・資料DL(ID:14)： " + sheetMonthly.getRange(rowMonthly,9).getValue() + "件 (CVR: " + Math.round(sheetMonthly.getRange(rowMonthly,10).getValue()*100000)/1000 + "%)\n"

  strBody = strBody + "[/code][hr]" + mySS.getUrl() + "[/info]"; //スプレッドシートのURLを取得

  SendCW( strBody )
  //SendSlack( strBody )
}

function SendCW( strBody ) {
  var cwClient = ChatWorkClient.factory({token: 'xxx'}); //チャットワークAPI
  cwClient.sendMessage({
    room_id: xxx, //prod
    //room_id: xxx, //test
    body: strBody
  });
}

function SendSlack( strBody ) {
  var token = 'xoxb-xxx-xxx';
  var slackApp = SlackApp.create(token); //SlackApp インスタンスの取得
  var options = {
    channelId: "#bot_test", //チャンネル名
    userName: "mohi", //投稿するbotの名前
    message: strBody //投稿するメッセージ
  };
  slackApp.postMessage(options.channelId, options.message, {username: options.userName});
}
