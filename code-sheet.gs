function sendToCalendar(e) {
  try{
    //有効なGooglesプレッドシートを開く
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    //予約を記載するカレンダーを取得
    var cal = CalendarApp.getCalendarById("ckvtvietabikeccicq31vpok10@group.calendar.google.com");

    var nFailure = false;
    var LIMIT_CLASS = 2; // 予約上限を設定(同一時間の上限)
    var num_row = sheet.getLastRow(); // 新規予約された行番号を取得
    var mail = sheet.getRange(num_row, 2).getValue(); // メルアド
    var name = sheet.getRange(num_row, 3).getValue(); // 名前
    var className = sheet.getRange(num_row, 4).getValue(); // 予約クラス

    /***
     登録されているメールとクラスが一致する時に予約を受け付けるようチェック
    ***/
    var regMails = getRegistedMailList(); // 登録メールリストをシートから取得
    for (var i = 0; i < regMails.length; i++) {
        if (regMails[i] == mail) {
          if (getClassName(mail) == className) {
            nFailure = false;
            break;
          } else { // 登録しているクラスと違う予約送信した場合
            sheet.deleteRow(num_row);
            sendFailureMail('6', name, mail, className);// 失敗のメール（登録クラスと違う）
            return;
          }
        }
        nFailure = true;
    }
    if (nFailure) { // 登録メールがない場合
      sheet.deleteRow(num_row);
      sendFailureMail('3', name, mail, className);// 失敗のメール（登録メールなし）
      return;
    }

    /***
     各クラスに応じた場所から予約日と時間を取得
    ***/
    if (className == '初級クラス') {
      var nDay = sheet.getRange(num_row, 5).getValue(); // 予約日
      var nTime = sheet.getRange(num_row, 6).getValue(); // 予約時間
    } else if (className == '中級クラス') {
      var nDay = sheet.getRange(num_row, 7).getValue(); // 予約日
      var nTime = sheet.getRange(num_row, 8).getValue(); // 予約時間
    }
    var nDate = new Date(nDay);

    /***
     指定された日が定休日か確認
    ***/
    if (isCloseday(cal, nDate)) {
      sheet.deleteRow(num_row);
      sendFailureMail('7', name, mail, className);// 失敗のメール（定休日）
      return;
    }

    /***
     指定された日が昨日以前か確認
    ***/
    if (isBefore(nDate)) {
      sheet.deleteRow(num_row);
      sendFailureMail('9', name, mail, className);// 失敗のメール（昨日以前）
      return;
    }

    /***
     当日から2ヶ月以内であるか確認
    ***/
    var today = new Date();
    if (nDate.getMonth() >= (today.getMonth()+2)) {
      if (nDate.getDate() > (today.getDate())) {
        sheet.deleteRow(num_row);
        sendFailureMail('8', name, mail, className);// 失敗のメール（2ヶ月以上）
        return;
      }
    }

    /***
     各クラスに応じて指定された時間を設定
    ***/
    if (className == '初級クラス') {
      if (nDate.getDay() == 0 || nDate.getDay() == 6 /*|| isHoliday(nDate) == true*/) {
        if (nTime == '10:00 ~ (土日)') {
          nDate.setHours(10,00);
        } else if (nTime == '11:30 ~ (土日)') {
          nDate.setHours(11,30);
        } else if (nTime == '13:30 ~ (土日)') {
          nDate.setHours(13,30);
        } else if (nTime == '15:00 ~ (土日)') {
          nDate.setHours(15,00);
        } else {
          nFailure = true;
        }
      } else if (nDate.getDay() == 1 || nDate.getDay() == 2) {
        if (nTime == '10:30 ~ (月火)') {
          nDate.setHours(10,30);
        } else if (nTime == '12:30 ~ (月火)') {
          nDate.setHours(12,30);
        } else if (nTime == '14:00 ~ (月火)') {
          nDate.setHours(14,00);
        } else if (nTime == '15:30 ~ (月火)') {
          nDate.setHours(15,30);
        } else if (nTime == '17:00 ~ (月火)') {
          nDate.setHours(17,00);
        } else if (nTime == '18:30 ~ (月火)') {
          nDate.setHours(18,30);
        } else {
          nFailure = true;
        }
      } else {
        nFailure = true;
      }
    } else if (className == '中級クラス') {
      if (nDate.getDay() == 0 || nDate.getDay() == 6 /*|| isHoliday(nDate) == true*/) {
        if (nTime == '16:30 ~ (土日)') {
          nDate.setHours(16,30);
        } else if (nTime == '18:00 ~ (土日)') {
          nDate.setHours(18,00);
        } else if (nTime == '19:30 ~ (土日)') {
          nDate.setHours(19,30);
        } else {
          nFailure = true;
        }
      } else if (nDate.getDay() == 1 || nDate.getDay() == 2) {
        if (nTime == '20:00 ~ (月火)') {
          nDate.setHours(20,00);
        } else {
          nFailure = true;
        }
      } else {
        nFailure = true;
      }

    } else {
      nFailure = true;
    }

    if (nFailure) { // 各クラスに応じて予約できない日時を選択された時
      sheet.deleteRow(num_row);
      sendFailureMail('1', name, mail, className); // 失敗のメール（日時不可）
      return;
    }

    var rStart = new Date(nDate.getFullYear(),nDate.getMonth(),nDate.getDate(),nDate.getHours(),nDate.getMinutes(),0);
    var rEnd = new Date(nDate.getFullYear(),nDate.getMonth(),nDate.getDate(),nDate.getHours()+1,nDate.getMinutes(),0);

    var events = cal.getEvents(rStart, rEnd); // 指定日時のイベント取得

    /***
     指定された日時にイベント（見学予約）があるか確認
    ***/
    var tour = cal.getEvents(rStart, rEnd, {search:'イベント'});
    if (tour.length != 0) {
      sheet.deleteRow(num_row);
      sendFailureMail('10', name, mail, className);// 失敗のメール（イベント）
      return;
    }

    /***
     指定された日時にチケットクラスがあるか確認
    ***/
    var ticket = cal.getEvents(rStart, rEnd, {search:'チケットクラス'});
    if (ticket.length != 0) {
      sheet.deleteRow(num_row);
      sendFailureMail('11', name, mail, className);// 失敗のメール（チケットクラス）
      return;
    }

    /***
     同じ日時に予約が重複しているか確認
    ***/
    var uid = mail+nDate.getFullYear()+nDate.getMonth()+nDate.getDate()+nDate.getHours();
    if (existTicket(uid, sheet)) {
      sheet.deleteRow(num_row);
      sendFailureMail('5', name, mail, className); // 失敗のメール（予約の重複）
      return;
    } else {
      sheet.getRange(num_row, 12).setValue(uid); // 重複チェック用UIDを追加
    }

    /***
     各クラスの枠内が既に上限の予約数に達しているか確認
    ***/
    if(events.length < LIMIT_CLASS){
      // その月の上限を確認
      var mid = mail+nDate.getFullYear()+nDate.getMonth();
      if (validTicket(mid, sheet)) {
        sheet.getRange(num_row, 11).setValue(mid);

        var item = className + ": 予約済";
        //予約情報をカレンダーに追加
        var res = cal.createEvent(item, rStart, rEnd);
        sheet.getRange(num_row, 13).setValue(res.getId()); // カレンダーのEvent IDを追加

        sendMailToUser(rStart, name, mail, className); // 成功のメール
      } else {
        sheet.deleteRow(num_row);
        sendFailureMail('4', name, mail, className); // 失敗のメール（月の上限）
      }
    }
    else{ // 指定の時間が既に満席の時
      sheet.deleteRow(num_row);
      sendFailureMail('2', name, mail, className); // 失敗のメール（満席）
    }

  } catch(exp){
    MailApp.sendEmail(mail, exp.message, exp.message);
  }

}

/***
 予約失敗時のメール送信
***/
function sendFailureMail(type, username, mail, className) {
  var title = '【CodeAid教室予約】予約できませんでした';
  var cont = username + "様　\n\n";

  if (type == 1) {
    cont += '予約できない日時が選択されたため予約できませんでした。\n申し訳ございませんが、再度予約してください。\n'
     + '問い合わせフォームからでも予約することができます。よろしくお願いします。\n\n';
  } else if (type == 2) {
    cont += className + "のお申し込みの時間は満席となっています。\n申し訳ございませんが、予約できませんでした。\n再度日時を変更して予約をお願いします。\n\n";
  } else if (type == 3) {
    cont += "登録されていないメールアドレスでは予約できません。\n登録しているメールアドレスで予約をお願いします。\n\n";
  } else if (type == 4) {
    cont += "今月の予約できる上限数を超えています。\n来月以降に予約をお願いします。\n\n";
  } else if (type == 5) {
    cont += "指定した日時は既に予約が完了しています。\n他の日時で予約をお願いします。\n\n";
  } else if (type == 6) {
    cont += "登録クラスと異なるクラスで指定されました。\n登録しているクラスで予約をお願いします。\n\n";
  } else if (type == 7) {
    cont += "指定した日はお休みとなります。\n他の日時で予約をお願いします。\n\n";
  } else if (type == 8) {
    cont += "本日から2ヶ月以上先の予約はできません。\n本日から2ヶ月以内の日時で予約をお願いします。\n\n";
  } else if (type == 9) {
    cont += "昨日以前の日付の予約はできません。\n本日以降の日時で予約をお願いします。\n\n";
  } else if (type == 10) {
    cont += "指定した日時は各種イベントがあるため予約できません。\n申し訳ありませんが、他の日時で予約をお願いします。\n\n";
  } else if (type == 11) {
    cont += "指定した日時はすでに予約があります。\n申し訳ありませんが、他の日時で予約をお願いします。\n\n";
  }

  cont += '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。\n\n';

  cont += '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=\n'
  + '【住所】大阪府吹田市垂水町1-7-23-103\n'
  + '【電話番号】090-8193-2811\n'
  + '【メール】contact@codeaid.jp\n'
  + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=\n';

  MailApp.sendEmail(mail, title, cont);
}

/***
 予約完了メール送信
***/
function sendMailToUser(rStart, username, mail, className){
  var dateStr = rStart.toLocaleString("ja-JP");
  var title = '【CodeAid教室予約】予約完了';
//  var png = getMapImage("大阪府吹田市垂水町1-7-23"); // 住所または場所名などGoogleMapsでわかるもの
  var message = '<html><body>' + username + '様<br><br>'
    + className + 'の予約が完了しました。<br>'
    + '【予約日時】' + dateStr + '<br><br>'
    + '予約をキャンセルする場合は、キャンセルフォームからキャンセルするか、お問い合わせフォーム/メールもしくは電話にてご連絡ください。<br><br>'
    + '※本メールに心当たりのない方は、大変お手数ですが削除していただきますよう、よろしくお願いいたします。<br><br>'
    + '=+=+=+=+= CodeAidプログラミング教室 =+=+=+=+=<br>'
    + '【住所】大阪府吹田市垂水町1-7-23-103<br>'
    + '【電話番号】090-8193-2811<br>'
    + '【メール】contact@codeaid.jp<br>'
//    + '【Googleマップ】<br>'
//    + '<img src="cid:map" width="400px" height="300px"><br>'
    + '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=</body></html>';
  MailApp.sendEmail({
    to: mail,
    subject: title,
    htmlBody: message,
//    inlineImages:{ map: png},
//      attachments:[png]
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
  var datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('登録');
  // B列2行目のデータからB列の最終行までを取得
  var lastRow = datasheet.getRange("B:B").getValues().filter(String).length - 1;
  Logger.log("lastRow = %s", lastRow);
  // B列2行目のデータからB列の最終行までを1列だけ取得
  selectList = datasheet.getRange(2, 2, lastRow, 1).getValues();
  Logger.log("selectList = %s", selectList);

  return selectList;
}

/***
 登録クラスを取得
***/
function getClassName(mail) {
  var selectList = [];

  // マスタデータシートを取得
  var datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('登録');
  // B列2行目のデータからB列の最終行までを取得
  var lastRow = datasheet.getRange("B:B").getValues().filter(String).length - 1;
  Logger.log("lastRow = %s", lastRow);
  // B列2行目のデータからB列の最終行までを1列だけ取得
  selectList = datasheet.getRange(2, 2, lastRow, 1).getValues();
  Logger.log("selectList = %s", selectList);

  for (var i=0; i < selectList.length; i++) {
    if (selectList[i] == mail) {
      return datasheet.getRange(i+2, 3).getValue(); // 登録クラス
    }
  }
}

/***
 指定された日が定休日か確認
***/
function isCloseday (cal, date) {
  var events = cal.getEventsForDay(date);
  for (var i in events) {
    if (events[i].getTitle() == '定休日' || events[i].getTitle() == '臨時休講' || events[i].getTitle() == '休') {
      return true;
    }
  }
  return false;
}

/***
 指定された日が祝日か確認
***/
function isHoliday (date) {
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
***/
function validTicket(uid, sheet) {
  var uidList = [];
  var count = 0;
  var LIMIT_COUNT = 4; // 予約上限を設定(受講回数の上限)

  // K列2行目のデータからK列の最終行までを取得(情報のある行数を算出)
  var lastRow = sheet.getRange("K:K").getValues().filter(String).length - 1;
  if (lastRow <= 0) {
    return true;
  }

  // K列2行目のデータからK列の最終行までを1列だけ取得
  uidList = sheet.getRange(2, 11, lastRow, 1).getValues();

  for (var i=0; i < uidList.length; i++) {
    if (uidList[i] == uid) {
      count++;
    }
  }
  if (count < LIMIT_COUNT) {
    return true;
  } else {
    return false;
  }
}

/***
 同じ日時で予約しているか確認
***/
function existTicket(uid, sheet) {
  var uidList = [];
  var count = 0;

  // L列2行目のデータからK列の最終行までを取得
  var lastRow = sheet.getRange("L:L").getValues().filter(String).length - 1;
  if (lastRow <= 0) {
    return false;
  }

  // L列2行目のデータからL列の最終行までを1列だけ取得
  uidList = sheet.getRange(2, 12, lastRow, 1).getValues();

  for (var i=0; i < uidList.length; i++) {
    if (uidList[i] == uid) {
      return true;
    }
  }
  return false;
}

/***
 設定シートから指定した項目の値を返却
***/
function getValue(item) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  var values = sheet.getRange().getValues();
  for (var i in values) {
    if (values[i][0] == item) {
      var res = values[i][1].split(',');
      return res;
    }
  }
  return false;
}
