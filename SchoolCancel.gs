function doCancel() {
  // プレッドシートから予約状況を取得
  var sheet = SpreadsheetApp.openById('15ouF5hXgRblkEH0hzpGIuyiw5hDYSZU4--RG6CFlYoM').getSheetByName('予約状況');

  // プレッドシートから登録情報を取得
  var regsheet = SpreadsheetApp.openById('15ouF5hXgRblkEH0hzpGIuyiw5hDYSZU4--RG6CFlYoM').getSheetByName('登録');

  // 予約を記載するカレンダーを取得
  var cal = CalendarApp.getCalendarById("64e3647p1tlnd3qaa6v5fm4ag0@group.calendar.google.com");

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
    if (deleteReserve(mail, date, sheet, regsheet, cal)) {
      sendCompletionMail(name, mail, date); // 成功のメール
    } else {
      sendFailureMail(name, mail, date); // 失敗のメール
    }
  }
  form.deleteAllResponses(); // キャンセルフォームの回答を削除
}

/***
 予約を削除
***/
function deleteReserve(mail, rdate, sheet, regsheet, cal) {
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var yesterday = new Date(rdate.getFullYear(), rdate.getMonth(), rdate.getDate());

  // キャンセル日が昨日以前の場合False(日付だけを比較)
  if (yesterday.getTime() < today.getTime()) {
    return false;
  }

  // 登録メルアドとIDのリスト取得
  var idList = getIdList(regsheet);
  if (Object.keys(idList).length == 0) {
    return false;
  }

  // 予約状況シートのC列2行目のデータからC列の最終行までの行数を取得（見出しを除く）
  var num = sheet.getRange("C:C").getValues().filter(String).length - 1;
  // 予約状況シートのC列2行目のデータからC列の最終行までのデータを取得（予約メルアド取得）
  var mails = sheet.getRange(2, 3, num, 1).getValues();
  // 予約状況シートのD列2行目のデータからD列の最終行までの行数を取得（見出しを除く）
  var num = sheet.getRange("D:D").getValues().filter(String).length - 1;
  // 予約状況シートのD列2行目のデータからD列の最終行までのデータを取得（予約日取得）
  var recdates = sheet.getRange(2, 4, num, 1).getValues();
  // 予約状況シートのE列2行目のデータからE列の最終行までの行数を取得（見出しを除く）
  var num = sheet.getRange("E:E").getValues().filter(String).length - 1;
  // 予約状況シートのE列2行目のデータからE列の最終行までのデータを取得（予約時間取得）
  var rectimes = sheet.getRange(2, 5, num, 1).getValues();

  var end = new Date(rdate.getFullYear(), rdate.getMonth(), rdate.getDate(), rdate.getHours() + 1, rdate.getMinutes(), 0); // 終了時間

  if (mails.length > 0 && mails.length == recdates.length && mails.length == rectimes.length) {
    for (var i = 0; i < mails.length; i++) {
      recdates[i] = new Date(recdates[i]);
      if (rectimes[i] == '10:30 ~') {
        recdates[i].setHours(10, 30);
      } else if (rectimes[i] == '12:30 ~') {
        recdates[i].setHours(12, 30);
      } else if (rectimes[i] == '14:00 ~') {
        recdates[i].setHours(14, 00);
      } else if (rectimes[i] == '15:30 ~') {
        recdates[i].setHours(15, 30);
      } else if (rectimes[i] == '17:00 ~') {
        recdates[i].setHours(17, 00);
      } else if (rectimes[i] == '18:30 ~') {
        recdates[i].setHours(18, 30);
      } else if (rectimes[i] == '20:00 ~') {
        recdates[i].setHours(20, 00);
      } else {
        return false;
      }

      if (recdates[i].getTime() == rdate.getTime() && mail == mails[i]) {
        var events = cal.getEvents(rdate, end, {
          search: idList[mails[i]]
        });
        if (events.length != 0) {
          for (var j = 0; j < events.length; j++) {
            events[j].deleteEvent(); // カレンダーから予約を削除
            sheet.deleteRow(i+2); // スプレッドシートから予約を削除
          }
          return true;
        }
      }
    }
  } else {
    return false;
  }
  return false;
}

function getIdList(regsheet) {
  // 登録シートのB列2行目のデータからB列の最終行までの行数を取得（見出しを除く）
  var num = regsheet.getRange("B:B").getValues().filter(String).length - 1;
  // 登録シートのB列2行目のデータからB列の最終行までのデータを取得（登録メルアド取得）
  var regMails = regsheet.getRange(2, 2, num, 1).getValues();

  // 登録シートのD列2行目のデータからC列の最終行までの行数を取得（見出しを除く）
  var num = regsheet.getRange("D:D").getValues().filter(String).length - 1;
  // 登録シートのD列2行目のデータからC列の最終行までのデータを取得（ID取得）
  var regIds = regsheet.getRange(2, 4, num, 1).getValues();

  var idList = {};

  if (regMails.length == regIds.length) {
    for (var i = 0; i < regMails.length; i++) {
      idList[regMails[i].toString()] = regIds[i].toString();
    }
  }
  return idList;
}

/***
 予約キャンセル失敗時のメール送信
***/
function sendFailureMail(name, mail, date) {
  var strDate = date.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約キャンセルできませんでした';
  var cont = '<html><body>' + name + "様<br><br>";

  cont += '【予約キャンセルエラー】<br>'
    + 'キャンセルできなかった日時：　' + strDate + '<br><br>';

  cont += 'キャンセル可能な予約が見つかりませんでした。<br>'
    + 'お問い合わせフォーム、メール、電話でもキャンセルすることができます。<br><br>'
    + '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>'
    + '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>'
    + '【住所】大阪府吹田市垂水町1-7-23-103<br>'
    + '【電話番号】090-8193-2811<br>'
    + '【メール】school@codeaid.jp<br>'
    + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+</body></html>';

  GmailApp.sendEmail(mail, title, '予約キャンセルエラー', {
    htmlBody: cont,
    name: 'CodeAidプログラミング教室',
    bcc: 'codeaid.school@gmail.com'
  });
}

/***
 予約キャンセル完了メール送信
***/
function sendCompletionMail(name, mail, date){
  var strDate = date.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約キャンセル完了';
  var message = '<html><body>' + name + '様<br><br>'
    + '予約がキャンセルされました。<br>'
    + '【予約キャンセル日時】' + strDate + '<br><br>'
    + '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>'
    + '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>'
    + '【住所】大阪府吹田市垂水町1-7-23-103<br>'
    + '【電話番号】090-8193-2811<br>'
    + '【メール】school@codeaid.jp<br>'
    + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+</body></html>';

  GmailApp.sendEmail( mail, title, '予約キャンセル完了', {
      htmlBody: message,
      name: 'CodeAidプログラミング教室',
      bcc: 'codeaid.school@gmail.com'
  });
}
