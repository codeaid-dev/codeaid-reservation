function sendToCalendar(e) {
  try {
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.openById('10hGRjVHrRo-4bJ-IicdiXMlYrerTVwPPw_5KQmR6Dgw').getSheetByName('予約状況');

    //予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById("64e3647p1tlnd3qaa6v5fm4ag0@group.calendar.google.com");

    var form = FormApp.getActiveForm(); // アクティブフォームを取得

    var formResponses = form.getResponses(); // 全回答内容を取得

    for (var i = 0; i < formResponses.length; i++) {
      var formResponse = formResponses[i]; // 回答ひとつ分を取得
      var itemResponses = formResponse.getItemResponses(); // 質問項目を取得
      var mail = formResponse.getRespondentEmail(); // 予約者のメールアドレス取得
      Logger.log(mail);
      for (var j = 0; j < itemResponses.length; j++) {　// 回答内容をひとつずつチェック
        var itemResponse = itemResponses[j];
        var question = itemResponse.getItem().getTitle();
        var answer = itemResponse.getResponse();
        if (question == '名前') { // 予約者の名前を取得
          var name = answer;
          Logger.log(name);
        }
        if (question == '予約日') { // 予約日を取得
          var nDay = answer.replace(/-/g, '/');
          Logger.log(nDay);
          var nDate = new Date(nDay); // 予約日からDateクラスオブジェクト作成
        }
        if (question == '予約時間') { // 予約時間を取得
          var nTime = answer;
          Logger.log(nTime);
        }
      }
    }
    form.deleteAllResponses(); // 予約フォームの回答を削除

    var nFailure = false;
    var LIMIT_CLASS = 1; // 予約上限を設定(同一時間の上限)
    var num_row = sheet.getLastRow(); // 新規予約された行番号を取得

    /***
     登録されているメールとクラスが一致する時に予約を受け付けるようチェック
     * クラスチェックを削除 2019/05/06
    ***/
    var regMails = getRegistedMailList(); // 登録メールリストをシートから取得
    var monLimits = getMonthlyLimitList(); // 登録月上限リストをシートから取得
    var limit = 4;
    for (var i = 0; i < regMails.length; i++) {
      if (regMails[i] == mail) {
        limit = parseInt(monLimits[i]); // 月上限値を取得
        nFailure = false;
        break;
      }
      nFailure = true;
    }
    if (nFailure) { // 登録メールがない場合
      sheet.deleteRow(num_row);
      sendFailureMail('3', name, mail, nDate); // 失敗のメール（登録メールなし）
      return;
    }

    /***
     指定された日が定休日か確認
    ***/
    if (isCloseday(cal, nDate)) {
      sheet.deleteRow(num_row);
      sendFailureMail('6', name, mail, nDate); // 失敗のメール（定休日）
      return;
    }

    /***
     指定された日が昨日以前か確認
    ***/
    if (isBefore(nDate)) {
      sheet.deleteRow(num_row);
      sendFailureMail('8', name, mail, nDate); // 失敗のメール（昨日以前）
      return;
    }

    /***
     当日から2ヶ月以内であるか確認
    ***/
    var today = new Date();
    if (nDate.getMonth() >= (today.getMonth() + 2)) {
      if (nDate.getDate() > (today.getDate())) {
        sheet.deleteRow(num_row);
        sendFailureMail('7', name, mail, nDate); // 失敗のメール（2ヶ月以上）
        return;
      }
    }

    /***
     各クラスに応じて指定された時間を設定
     * クラスチェックを削除・全クラスに同じ時間割 2019/05/06
     * 土(6)・日(0)・月(1)・火(2)曜日のみ時間をチェック
    ***/
    if (nDate.getDay() == 6 || nDate.getDay() == 0 || nDate.getDay() == 1 || nDate.getDay() == 2) {
      if (nTime == '10:30 ~') {
        nDate.setHours(10, 30);
      } else if (nTime == '12:30 ~') {
        nDate.setHours(12, 30);
      } else if (nTime == '14:00 ~') {
        nDate.setHours(14, 00);
      } else if (nTime == '15:30 ~') {
        nDate.setHours(15, 30);
      } else if (nTime == '17:00 ~') {
        nDate.setHours(17, 00);
      } else if (nTime == '18:30 ~') {
        nDate.setHours(18, 30);
      } else if (nTime == '20:00 ~') {
        nDate.setHours(20, 00);
      } else {
        nFailure = true;
      }
    } else {
      nFailure = true;
    }

    if (nFailure) { // 各クラスに応じて予約できない日時を選択された時
      sheet.deleteRow(num_row);
      sendFailureMail('1', name, mail, nDate, nTime); // 失敗のメール（日時不可）
      return;
    }

    var rStart = new Date(nDate.getFullYear(), nDate.getMonth(), nDate.getDate(), nDate.getHours(), nDate.getMinutes(), 0);
    var rEnd = new Date(nDate.getFullYear(), nDate.getMonth(), nDate.getDate(), nDate.getHours() + 1, nDate.getMinutes(), 0);

    var events = cal.getEvents(rStart, rEnd); // 指定日時のイベント取得

    /***
     指定された日時にイベント（見学予約）があるか確認
    ***/
    var tour = cal.getEvents(rStart, rEnd, {
      search: 'イベント'
    });
    if (tour.length != 0) {
      sheet.deleteRow(num_row);
      sendFailureMail('9', name, mail, nDate, nTime); // 失敗のメール（イベント）
      return;
    }

    /***
     同じ日時に予約が重複しているか確認
    ***/
    var uid = mail + nDate.getFullYear() + nDate.getMonth() + nDate.getDate() + nDate.getHours();
    if (existTicket(uid, sheet)) {
      sheet.deleteRow(num_row);
      sendFailureMail('5', name, mail, nDate, nTime); // 失敗のメール（予約の重複）
      return;
    } else {
      sheet.getRange(num_row, 7).setValue(uid); // 重複チェック用UIDを追加
    }

    /***
     各クラスの枠内が既に上限の予約数に達しているか確認
    ***/
    if (events.length < LIMIT_CLASS) {
      // その月の上限を確認
      var mid = mail + nDate.getFullYear() + nDate.getMonth();
      if (validTicket(mid, limit, sheet)) {
        sheet.getRange(num_row, 6).setValue(mid);

        var item = "予約済";
        //予約情報をカレンダーに追加
        var res = cal.createEvent(item, rStart, rEnd);
        sheet.getRange(num_row, 8).setValue(res.getId()); // カレンダーのEvent IDを追加

        sendMailToUser(rStart, name, mail); // 成功のメール
      } else {
        sheet.deleteRow(num_row);
        sendFailureMail('4', name, mail, nDate, nTime); // 失敗のメール（月の上限）
      }
    } else { // 指定の時間が既に満席の時
      sheet.deleteRow(num_row);
      sendFailureMail('2', name, mail, nDate, nTime); // 失敗のメール（満席）
    }

  } catch (exp) {
    GmailApp.sendEmail(mail, exp.message, exp.message);
  }

}

/***
 予約失敗時のメール送信
***/
function sendFailureMail(type, username, mail, date, time) {
  var title = '【CodeAid教室予約】予約できませんでした';
  var cont = username + "様　\n\n";

  if (type == 1) {
    cont += '予約できない日時が選択されたため予約できませんでした。\n申し訳ございませんが、再度予約してください。\n';
  } else if (type == 2) {
    cont += "指定した日時は予約が満席です。\n申し訳ありませんが、他の日時で予約をお願いします。\n";
  } else if (type == 3) {
    cont += "登録されていないメールアドレスでは予約できません。\n登録しているメールアドレスで予約をお願いします。\n";
  } else if (type == 4) {
    cont += "今月の予約できる上限数を超えています。\n来月以降に予約をお願いします。\n";
  } else if (type == 5) {
    cont += "指定した日時で予約は完了しています。\n申し訳ありませんが、他の日時で予約をお願いします。\n";
  } else if (type == 6) {
    cont += "指定した日はお休みとなります。\n他の日時で予約をお願いします。\n";
  } else if (type == 7) {
    cont += "本日から2ヶ月以上先の予約はできません。\n本日から2ヶ月以内の日時で予約をお願いします。\n";
  } else if (type == 8) {
    cont += "昨日以前の日付の予約はできません。\n本日以降の日時で予約をお願いします。\n";
  } else if (type == 9) {
    cont += "指定した日時はイベントがあるため予約できません。\n申し訳ありませんが、他の日時で予約をお願いします。\n";
  }

  cont += '予約しようとした日時：　';
  cont += '' + date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate() + ' ' + time + '\n';
  cont += '問い合わせフォーム、電話、メールからでも予約することができます。\n\n';
  cont += '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。\n\n';

  cont += '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=\n' +
    '【住所】大阪府吹田市垂水町1-7-23-103\n' +
    '【電話番号】090-8193-2811\n' +
    '【メール】contact@codeaid.jp\n' +
    '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=\n';

  GmailApp.sendEmail(mail, title, cont, {
    name: 'CodeAidプログラミング教室'
  });
  GmailApp.sendEmail('codeaid.school@gmail.com', 'CodeAid教室予約エラー通知', '教室予約でエラーが発生しました。\n送信したエラー内容を確認してください。', {
    name: 'CodeAidプログラミング教室'
  });
}

/***
 予約完了メール送信
***/
function sendMailToUser(rStart, username, mail) {
  var dateStr = rStart.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約完了';
  var message = '<html><body>' + username + '様<br><br>' +
    '予約が完了しました。<br>' +
    '【予約日時】' + dateStr + '<br><br>' +
    '予約をキャンセルする場合は、キャンセルフォームからキャンセルするか、お問い合わせフォーム/メールもしくは電話にてご連絡ください。<br><br>' +
    '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>' +
    '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>' +
    '【住所】大阪府吹田市垂水町1-7-23-103<br>' +
    '【電話番号】090-8193-2811<br>' +
    '【メール】contact@codeaid.jp<br>' +
    '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=</body></html>';
  GmailApp.sendEmail(
    mail,
    title,
    '予約完了メール', {
      htmlBody: message,
      name: 'CodeAidプログラミング教室'
    });
  GmailApp.sendEmail('codeaid.school@gmail.com', 'CodeAid教室予約完了通知', '教室予約が完了しました。\n予約日時を確認してください。', {
    name: 'CodeAidプログラミング教室'
  });
}

/***
 案内する地図を作成（PNG形式）
***/
function getMapImage(point) {
  var map = Maps.newStaticMap().setSize(400, 300)
    .setCenter(point).setZoom(15).setLanguage('ja')
    .setMapType(Maps.StaticMap.Type.ROADMAP);
  map.addMarker(point);
  return map.getBlob().getAs(MimeType.PNG);
}

/***
 登録メールのリストを取得
***/
function getRegistedMailList() {
  var selectList = [];

  // マスタデータシートを取得
  var datasheet = SpreadsheetApp.openById('10hGRjVHrRo-4bJ-IicdiXMlYrerTVwPPw_5KQmR6Dgw').getSheetByName('登録');
  // B列2行目のデータからB列の最終行までを取得
  var lastRow = datasheet.getRange("B:B").getValues().filter(String).length - 1;
  // B列2行目のデータからB列の最終行までを1列だけ取得
  selectList = datasheet.getRange(2, 2, lastRow, 1).getValues();

  return selectList;
}
/***
 登録月上限のリストを取得
***/
function getMonthlyLimitList() {
  var selectList = [];

  // マスタデータシートを取得
  var datasheet = SpreadsheetApp.openById('10hGRjVHrRo-4bJ-IicdiXMlYrerTVwPPw_5KQmR6Dgw').getSheetByName('登録');
  // C列2行目のデータからC列の最終行を取得
  var lastRow = datasheet.getRange("C:C").getValues().filter(String).length - 1;
  // C列2行目のデータからC列の最終行までを1列だけ取得
  selectList = datasheet.getRange(2, 3, lastRow, 1).getValues();

  return selectList;
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
 指定された日が祝日か確認
***/
function isHoliday(date) {
  // 祝日カレンダーを取得
  var jcal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  var events = jcal.getEventsForDay(date);
  // 祝日カレンダーに何か予定が設定されていれば祝日とする
  if (events.length > 0) {
    return true;
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
 その月の上限を超えて予約しているか確認
 * クラスチェック削除・フォーム項目数減少でF列にMID保存 2019/05/06
***/
function validTicket(mid, limit, sheet) {
  var midList = [];
  var count = 0;
  //var LIMIT_COUNT = 4; // 予約上限を設定(受講回数の上限)

  // F列最終行までを取得(情報のある行数を算出)
  var lastRow = sheet.getRange("F:F").getValues().filter(String).length - 1;
  if (lastRow <= 0) {
    return true;
  }

  // F列2行目のデータからF列の最終行までを1列だけ取得
  midList = sheet.getRange(2, 6, lastRow, 1).getValues();

  for (var i = 0; i < midList.length; i++) {
    if (midList[i] == mid) {
      count++;
    }
  }
  if (count < limit) {
    return true;
  } else {
    return false;
  }
}

/***
 同じ日時で予約しているか確認
 * クラスチェック削除・フォーム項目数減少でG列にUID保存 2019/05/06
***/
function existTicket(uid, sheet) {
  var uidList = [];
  var count = 0;

  // G列最終行を取得
  var lastRow = sheet.getRange("G:G").getValues().filter(String).length - 1;
  if (lastRow <= 0) {
    return false;
  }

  // G列2行目のデータから最終行までを1列だけ取得
  uidList = sheet.getRange(2, 7, lastRow, 1).getValues();

  for (var i = 0; i < uidList.length; i++) {
    if (uidList[i] == uid) {
      return true;
    }
  }
  return false;
}
