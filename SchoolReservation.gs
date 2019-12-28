function makeReservation(e) {
  // プレッドシートから予約状況を取得
  var sheet = SpreadsheetApp.openById('15ouF5hXgRblkEH0hzpGIuyiw5hDYSZU4--RG6CFlYoM').getSheetByName('予約状況');

  // プレッドシートから登録情報を取得
  var regsheet = SpreadsheetApp.openById('15ouF5hXgRblkEH0hzpGIuyiw5hDYSZU4--RG6CFlYoM').getSheetByName('登録');

  // 予約を記載するカレンダーを取得
  var cal = CalendarApp.getCalendarById("64e3647p1tlnd3qaa6v5fm4ag0@group.calendar.google.com");

  const LIMIT_CLASS = 1; // 予約上限を設定(同一時間の上限)

  var lastRow = sheet.getLastRow(); // 新規予約された行番号を取得(フォーム送信時にシートへ自動保存されるデータ)

  var mailaddr = e.namedValues['メールアドレス'].toString(); // 予約者のメールアドレス
  var name = e.namedValues['名前'].toString(); // 予約者の名前
  var rdate = new Date(e.namedValues['予約日']); // 予約日
  var stime = e.namedValues['予約時間'].toString(); // 予約時間

  /***
   指定された時間を設定
  ***/
  if (stime == '10:30 ~') {
    rdate.setHours(10, 30);
  } else if (stime == '12:30 ~') {
    rdate.setHours(12, 30);
  } else if (stime == '14:00 ~') {
    rdate.setHours(14, 00);
  } else if (stime == '15:30 ~') {
    rdate.setHours(15, 30);
  } else if (stime == '17:00 ~') {
    rdate.setHours(17, 00);
  } else if (stime == '18:30 ~') {
    rdate.setHours(18, 30);
  } else if (stime == '20:00 ~') {
    rdate.setHours(20, 00);
  } else {
    sheet.deleteRow(lastRow);
    sendFailureMail(1, name, mailaddr, rdate, stime); // 失敗のメール（日時不可＊必須項目なのでありえないエラー）
    return;
  }

  /***
   登録メールのリストを取得
  ***/
  // 登録シートのB列2行目のデータからB列の最終行までの行数を取得（見出しを除く）
  var num = regsheet.getRange("B:B").getValues().filter(String).length - 1;
  // 登録シートのB列2行目のデータからB列の最終行までのデータを取得（登録メルアド取得）
  var regMails = regsheet.getRange(2, 2, num, 1).getValues();

  /***
   登録月上限のリストを取得
  ***/
  // 登録シートのC列2行目のデータからC列の最終行までの行数を取得（見出しを除く）
  var num = regsheet.getRange("C:C").getValues().filter(String).length - 1;
  // 登録シートのC列2行目のデータからC列の最終行までのデータを取得（月上限取得）
  var monLimits = regsheet.getRange(2, 3, num, 1).getValues();

  /***
   登録IDのリストを取得
  ***/
  // 登録シートのD列2行目のデータからC列の最終行までの行数を取得（見出しを除く）
  var num = regsheet.getRange("D:D").getValues().filter(String).length - 1;
  // 登録シートのD列2行目のデータからC列の最終行までのデータを取得（ID取得）
  var regIds = regsheet.getRange(2, 4, num, 1).getValues();

  /***
   登録されているメールが一致する時に予約を受け付けるようチェック
  ***/
  var limit = 4;
  var id = 'none';

  for (var i = 0; i < regMails.length; i++) {
    if (regMails[i] == mailaddr) {
      id = regIds[i].toString(); // 登録IDを取得
      limit = parseInt(monLimits[i]); // 月上限値を取得
      break;
    } else if (i == regMails.length - 1) { // 登録メールがない場合
      sheet.deleteRow(lastRow);
      sendFailureMail(3, name, mailaddr, rdate, stime); // 失敗のメール（登録メールなし）
      return;
    }
  }

  /***
   月上限を超えてるか確認
  ***/
  if (!checkLimit(sheet, mailaddr, rdate, limit)) {
    sheet.deleteRow(lastRow);
    sendFailureMail(4, name, mailaddr, rdate, stime); // 失敗のメール（月の上限）
    return;
  }

  /***
   指定された日が定休日か確認
  ***/
  if (isCloseday(cal, rdate)) {
    sheet.deleteRow(lastRow);
    sendFailureMail(6, name, mailaddr, rdate, stime); // 失敗のメール（定休日）
    return;
  }

  /***
   指定された日が昨日以前か確認
  ***/
  if (isBefore(rdate)) {
    sheet.deleteRow(lastRow);
    sendFailureMail(8, name, mailaddr, rdate, stime); // 失敗のメール（昨日以前）
    return;
  }

  /***
   指定された日が２ヶ月以内か確認
  ***/
  if (twoMonthsLater(rdate)) {
    sheet.deleteRow(lastRow);
    sendFailureMail(7, name, mailaddr, rdate, stime); // 失敗のメール（2ヶ月以上）
    return;
  }

  var end = new Date(rdate.getFullYear(), rdate.getMonth(), rdate.getDate(), rdate.getHours() + 1, rdate.getMinutes(), 0); // 終了時間

  /***
   同じ日時に予約が重複しているか確認
  ***/
  var exists = cal.getEvents(rdate, end, {
    search: id
  });
  if (exists.length != 0) {
    sheet.deleteRow(lastRow);
    sendFailureMail(5, name, mailaddr, rdate, stime); // 失敗のメール（予約の重複）
    return;
  }

  /***
   他に予約があるか確認（満席確認）
   * 同じ日時に２つ以上予約できない
  ***/
  var events = cal.getEvents(rdate, end); // 指定日時のイベントリスト取得
  if (events.length >= LIMIT_CLASS) {
    sheet.deleteRow(lastRow);
    sendFailureMail(2, name, mailaddr, rdate, stime); // 失敗のメール（満席）
    return;
  }

  /***
   予約完了通知送信
  ***/
  var item = "予約済";
  //予約情報をカレンダーに追加
  var res = cal.createEvent(item, rdate, end, {
    description: id
  });

  sendCompletionMail(rdate, name, mailaddr);
}

/***
 指定された日が定休日か確認
***/
function isCloseday(cal, date) {
  var events = cal.getEventsForDay(date);
  for (var i in events) {
    if (events[i].getTitle() == '定休日' || events[i].getTitle() == '休講' || events[i].getTitle() == '休日') {
      return true;
    }
  }
  return false;
}

/***
 指定された日が昨日以前か確認
***/
function isBefore(date) {
  var target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (target < today) {
    return true;
  } else {
    return false;
  }
}

/***
 指定された日が２ヶ月以内か確認
***/
function twoMonthsLater(date) {
  var target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var exdate = new Date();
  exdate = new Date(exdate.getFullYear(), exdate.getMonth()+2, exdate.getDate());

  if (target >= exdate) {
    return true;
  } else {
    return false;
  }
}

/***
 その月の上限を超えて予約しているか確認
***/
function checkLimit(sheet, mailaddr, rdate, limit) {
  // 予約状況シートのC列2行目のデータからC列の最終行までの行数を取得（見出しを除く）
  var num = sheet.getRange("C:C").getValues().filter(String).length - 1;
  // 予約状況シートのC列2行目のデータからC列の最終行までのデータを取得（予約メルアド取得）
  var mails = sheet.getRange(2, 3, num, 1).getValues();
  // 予約状況シートのD列2行目のデータからD列の最終行までの行数を取得（見出しを除く）
  var num = sheet.getRange("D:D").getValues().filter(String).length - 1;
  // 予約状況シートのD列2行目のデータからD列の最終行までのデータを取得（予約日取得）
  var recdates = sheet.getRange(2, 4, num, 1).getValues();

  if (mails.length == 0) {
    return true;
  }

  var count = 0;
  if (mails.length == recdates.length) {
    for (var i = 0; i < mails.length; i++) {
      var recdate = new Date(recdates[i]);
      if (mails[i] == mailaddr && recdate.getMonth() == rdate.getMonth()) {
        count++;
      }
    }
    if (count > limit) {
      return false;
    }
  }

  return true;
}

/***
 予約失敗時のメール送信
***/
function sendFailureMail(type, name, mailaddr, date, stime) {
  var title = '【CodeAid教室予約】予約できませんでした';
  var cont = '<html><body>' + name + "様<br><br>";

  cont += '【予約エラー】<br>';

  if (type == 1) {
    cont += '（予期せぬエラー）予約できない時間が選択されました。<br>予期せぬエラーのためメール・電話で予約を確認してください。<br><br>';
  } else if (type == 2) {
    cont += "指定した日時は満席のため予約できません。<br>申し訳ありませんが、他の日時で予約をお願いします。<br><br>";
  } else if (type == 3) {
    cont += "登録されていないメールアドレスで予約フォームが送信されました。<br>入力ミスも考えられます。<br><br>";
  } else if (type == 4) {
    cont += "今月の予約できる上限数を超えています。<br>来月以降に予約をお願いします。<br><br>";
  } else if (type == 5) {
    cont += "予約済の日時で予約しようとしています。<br>ご確認ください。<br><br>";
  } else if (type == 6) {
    cont += "指定した日はお休みとなります。<br>他の日時で予約をお願いします。<br><br>";
  } else if (type == 7) {
    cont += "本日から2ヶ月以上先の予約はできません。<br>本日から2ヶ月以内の日時で予約をお願いします。<br><br>";
  } else if (type == 8) {
    cont += "昨日以前の日付の予約はできません。<br>本日以降の日時で予約をお願いします。<br><br>";
  }

  if (type == 1 || type == 3) {
    GmailApp.sendEmail('codeaid.school@gmail.com', 'CodeAid教室予約エラー通知', cont, {
      name: 'CodeAidプログラミング教室'
    });
    return;
  }

  cont += '予約しようとした日時：　';
  cont += '' + date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + stime + '<br>';
  cont += 'お問い合わせフォーム、電話、メールでも予約することができます。<br><br>';
  cont += '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>';

  cont += '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>' +
    '【住所】大阪府吹田市垂水町1-7-23-103<br>' +
    '【電話番号】090-8193-2811<br>' +
    '【メール】contact@codeaid.jp<br>' +
    '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+</body></html>';

  GmailApp.sendEmail(mailaddr, title, '予約エラー', {
    htmlBody: cont,
    name: 'CodeAidプログラミング教室',
    bcc: 'codeaid.school@gmail.com'
  });
}

/***
 予約完了メール送信
***/
function sendCompletionMail(stime, name, mailaddr) {
  var strDate = stime.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約完了';
  var message = '<html><body>' + name + '様<br><br>' +
    '予約が完了しました。<br>' +
    '【予約日時】' + strDate + '<br><br>' +
    '予約をキャンセルする場合は、キャンセルフォームからキャンセルするか、お問い合わせフォーム/メールもしくは電話にてご連絡ください。<br><br>' +
    '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>' +
    '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>' +
    '【住所】大阪府吹田市垂水町1-7-23-103<br>' +
    '【電話番号】090-8193-2811<br>' +
    '【メール】contact@codeaid.jp<br>' +
    '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+</body></html>';

  GmailApp.sendEmail(mailaddr, title, '予約完了', {
    htmlBody: message,
    name: 'CodeAidプログラミング教室',
    bcc: 'codeaid.school@gmail.com'
  });
}
