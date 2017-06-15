/**
 * ラベルが付いた未読のメール(スレッド)を探して返す
 * @return GmailThread[]
 */

function getMail() {
  var label = "_timer+gmail-✔agoda予約完了 ";
  var start = 0;
  var max = 500;
  return GmailApp.search('label:' + label + ' is:unread', start, max);
}

/**
 * メール本文を整形してスプレッドシートに保存するためのオブジェクトを返す
 * @return Object
 */
function getDatabyMailBody( body ) {
  var bookingID = body.match(/予約 ID.*/);
  var reservationInfo = body.match(/FMC.*/);
  var propertyID = body.match(/Property ID.*/);
  var firstName = body.match(/First Name.*/);
  var lastName = body.match(/Last Name.*/);
  var checkIn = body.match(/チェックイン.*/);
  var checkOut = body.match(/チェックアウト.*/);
  var roomInfo = body.match(/Apartment.*/);
  var payout = body.match(/JPY.*/);
  var customerNotes = body.match(/顧客情報.*/);
   return {
     bookingID: bookingID,
     reservationInfo: reservationInfo,
     propertyID: propertyID,
     firstName: firstName,
     lastName: lastName,
     checkIn: checkIn,
     checkOut: checkOut,
     roomInfo: roomInfo,
     payout: payout,
     customerNotes: customerNotes
  };
}

/**
 * gmailを取得してスプレッドシートに保存する
 */
function onSaveMailToSheet() {
  // データを保存するシートの名前
  var sheetName = 'シート1';
  var ss = SpreadsheetApp.getActive().getSheetByName( sheetName );
  var row = ss.getLastRow() + 1;
  var threads = getMail();

  for( var i in threads ) {
    var thread = threads[i];
    var msgs = thread.getMessages();
    // スレッド内のメールをそれぞれチェックする
    for( var j in msgs ) {
      var msg = msgs[j];
      // スレッド内の未読メッセージのみを処理
        var date = msg.getDate();
        var d = getDatabyMailBody( msg.getPlainBody() );
        var values = [
          [date, d.bookingID, d.reservationInfo, d.propertyID, d.firstName, d.lastName, d.checkIn, d.checkOut, d.roomInfo, d.payout, d.customerNotes]
        ];
        // シートに保存
        // ※ 3コラムなので A:C のRangeを取る。データ数に合わせて変更が必要
      Logger.log(values)
      ss.getRange("A" + row +":k" + row).setValues(values);
      row++;
    }
    // スレッドを既読にする,
    thread.markRead();
    Utilities.sleep(10000);
  }
}
