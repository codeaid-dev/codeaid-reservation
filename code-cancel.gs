function sendCancel(e) {
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.openById('10hGRjVHrRo-4bJ-IicdiXMlYrerTVwPPw_5KQmR6Dgw').getSheetByName('予約状況');

    //予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById("ckvtvietabikeccicq31vpok10@group.calendar.google.com");

    var form = FormApp.getActiveForm(); // アクティブフォームを取得

    var formResponses = form.getResponses(); // 全回答内容を取得

    for (var i = 0; i < formResponses.length; i++) {
      var formResponse = formResponses[i]; // 回答ひとつ分を取得
      var itemResponses = formResponse.getItemResponses(); // 質問項目を取得
      for (var j = 0; j < itemResponses.length; j++) {　// 回答内容をひとつずつチェック
        var itemResponse = itemResponses[j];
        var question = itemResponse.getItem().getTitle();
        var answer = itemResponse.getResponse();
        if (question == '名前') {
          var name = answer;
        }
        if (question == 'メールアドレス') {
          var mail = answer;
        }
        if (question == 'キャンセル日時') {
          answer = answer.replace(/-/g, '/');
          var date = new Date(answer);
        }
      }
      if (deleteReserve(mail, date, sheet, cal)) {
        sendMailToUser(name, mail, date); // 成功のメール
      } else {
        sendFailureMail(name, mail, date); // 失敗のメール
      }
    }
    form.deleteAllResponses(); // キャンセルフォームの回答を削除

  } catch(exp){
    MailApp.sendEmail(mail, exp.message, exp.message);
  }

}

/***
 予約キャンセル失敗時のメール送信
***/
function sendFailureMail(username, mail, date) {
  var dateStr = date.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約キャンセルできませんでした';
  var cont = username + "様　\n\n";

  cont += '【予約キャンセルできなかった日時】' + dateStr + '\n\n';

  cont += '予約キャンセルできませんでした。\n予約キャンセル可能な項目が見つかりませんでした。\n'
   + 'お電話やお問い合わせフォームからもキャンセルすることができます。よろしくお願いします。\n\n';

  cont += '※本メールに心当たりのない方は、削除していただきますようよろしくお願いします。\n\n';

  cont += '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=\n'
  + '【住所】大阪府吹田市垂水町1-7-23-103\n'
  + '【電話番号】090-8193-2811\n'
  + '【メール】contact@codeaid.jp\n'
  + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=\n';

  MailApp.sendEmail(mail, title, cont);
}

/***
 予約キャンセル完了メール送信
***/
function sendMailToUser(username, mail, date){
  var dateStr = date.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約キャンセル完了';
  var message = '<html><body>' + username + '様<br><br>'
    + '予約がキャンセルされました。<br>'
    + '【予約キャンセル日時】' + dateStr + '<br><br>'
    + '※本メールに心当たりのない方は、削除していただきますようよろしくお願いします。<br><br>'
    + '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>'
    + '【住所】大阪府吹田市垂水町1-7-23-103<br>'
    + '【電話番号】090-8193-2811<br>'
    + '【メール】contact@codeaid.jp<br>'
    + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=</body></html>';
  MailApp.sendEmail({
    to: mail,
    subject: title,
    htmlBody: message,
  });
}

/***
 登録クラスを取得
***/
function getRegistedMailList() {
  var selectList = [];

  // マスタデータシートを取得
  var datasheet = SpreadsheetApp.openById('10hGRjVHrRo-4bJ-IicdiXMlYrerTVwPPw_5KQmR6Dgw').getSheetByName('登録');
  // B列2行目のデータからB列の最終行までを取得
  var lastRow = datasheet.getRange("B:B").getValues().filter(String).length - 1;
  Logger.log("lastRow = %s", lastRow);
  // B列2行目のデータからB列の最終行までを1列だけ取得
  selectList = datasheet.getRange(2, 2, lastRow, 1).getValues();
  Logger.log("selectList = %s", selectList);

  return selectList;
}

/***
 予約を削除
***/
function deleteReserve(mail, date, sheet, cal) {
  var uidList = [];
  var uid = mail+date.getFullYear()+date.getMonth()+date.getDate()+date.getHours();
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  date = new Date(date.getFullYear(), date.getMonth(), date.getDate());

  //  キャンセル日が昨日以前の場合False(日付だけを比較)
  if (date < today) {
    return false;
  }
  // L列2行目のデータからL列の最終行までを取得
  var lastRow = sheet.getRange("L:L").getValues().filter(String).length - 1;
  if (lastRow <= 0) {
    return false;
  }

  // L列2行目のデータからL列の最終行までを1列だけ取得(UIDリスト取得)
  uidList = sheet.getRange(2, 12, lastRow, 1).getValues();

  Logger.log("uid: %s", uid);
  for (var i=0; i < uidList.length; i++) {
    if (uidList[i] == uid) {
      var eid = sheet.getRange(i+2, 13).getValue();
      Logger.log("eid: %s", eid);
      cal.getEventById(eid).deleteEvent(); // カレンダーから予約を削除
      sheet.deleteRow(i+2); // スプレッドシートから予約を削除
      return true;
    }
  }
  return false;
}
