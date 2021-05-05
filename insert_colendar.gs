const CALENDAR_ID = '{calendar_ID}';
const CHANEL_ACCESS_TOKEN = "{access_token}";
const CALENDAR = CalendarApp.getCalendarById(CALENDAR_ID);
const SPREADSHEET = SpreadsheetApp.openById('{url}');
const LOG_SHEET = SPREADSHEET.getSheetByName('log');
const STATUS_SHEET = SPREADSHEET.getSheetByName('status');
const RG_USERID = STATUS_SHEET.getRange('B1');
const RG_STATUS = STATUS_SHEET.getRange('B2');
const RG_DATE = STATUS_SHEET.getRange('B3');
const RG_TITLE = STATUS_SHEET.getRange('B4');


// ボットにメッセージ送信/フォロー/アンフォローした時の処理
function doPost(e) {
  try{
    var events = JSON.parse(e.postData.contents).events;
    events.forEach(function(event) {
      if(event.type == "message") {
        main(event);
      } else if(event.type == "follow") {
        follow(event);
      } else if(event.type == "unfollow") {
        unFollow(event);
      } else if(event.type == 'postback') {
        main_postback(event);
      }
    });
  }catch(error){
    write_log("doPostでエラーが発生しました。：" + error.message);
    return false;
  }

}

//(postback)メイン処理
function main_postback(e){
  try{
    //ステータスが初期状態で
    //postback.dataがEnterDateの時だけ処理
    if (RG_STATUS.getValue() == '初期状態' &&
        e.postback.data == 'EnterDate'){
          rcv_select_date(e);
          return;
    }else{
      return;
    }
  }catch(error){
    write_log("main_postbackでエラーが発生しました。：" + error.message);
    return false;
  }
}

//メイン処理
function main(e){
  try{

    //ユーザが指定のユーザかどうか確認
    if (!chk_userid_auth(e)){
      send_message('申し訳ございませんがお使いのユーザはこの機能は使用できません。\n' + 
                   '管理者にお問い合わせください。', e);
      return;
    }

    //入力された値がテキストでなかったら処理終了
    if (e.message.type != 'text'){
      send_message("テキスト以外は受付できません…", e);
      return;
    }


    //入力された値がキャンセルだったら初期状態に戻して終了
    if (e.message.text == 'キャンセル'){
      RG_STATUS.setValue('初期状態');
      send_message('キャンセルされました。', e);
      return;
    }


    //ステータスに応じて処理を変更
    switch(RG_STATUS.getValue()){
      case '初期状態':
        if (e.message.text == '予定確認'){
          send_message("2週間分の予定は、\n" + 
          canvert_obj_to_str(get_2week_events()) + '\n' +
          "です。", e);
          break;
        }else{
          send_message("入力された内容が無効です。", e);
          break;
        }

      case 'タイトル受付中':
        rcv_enter_title(e);
        break;

      case '最終確認中':
        if(e.message.text == 'はい'){
          create_schedule();
          RG_STATUS.setValue('初期状態');
          send_message('日付：\n' + 
                       dayjs.dayjs(RG_DATE.getValue()).add(9, 'h').format('YYYY/MM/DD') + '\n' +
                       'タイトル：\n' +
                       RG_TITLE.getValue() + '\n' +
                       'で登録されました！', e);
          break;
        }else if(e.message.text == 'いいえ'){
          RG_STATUS.setValue("タイトル受付中");
          send_message("もう一度タイトルを入力してください。", e);
          break;
        }else{
          send_message('エラーが発生しました。管理者に問い合わせてください');
          write_log("mainの最終確認で予期せぬ動き");
          break;
        }

      default:
        RG_STATUS.setValue('初期状態');
        send_message('すみませんが、もう一度お試しください。', e);
        break;
    }
    return;
  }
  catch(error){
    write_log("mainでエラーが発生しました。：" + error.message);
    return false;
  }
}

//メッセージを受け取った時の動作
function reply(e){
  if (e.message.text == '予定追加'){
    if (!get_schedule()){
      send_message('スケジュール登録時にエラーが発生しました。', e);
    }else{
      send_message("テストスケジュールが登録されました。", e);
    }
  }
  if (e.message.text == '予定確認'){
    var events = get_2week_events();
    send_message("2週間分の予定は、\n" + 
                 canvert_obj_to_str(events) + '\n' +
                 "です。", e);
  }
}

//日付選択アクションされたときの関数
function rcv_select_date(e){
  try{
    RG_DATE.setValue(e.postback.params.date);
    RG_STATUS.setValue('タイトル受付中');
    mk_button_chk_date(e);
  }catch(error){
    write_log("rcv_select_dateでエラーが発生しました。：" + error.message);
    return false;
  }
}

//タイトルを入力された時の関数
function rcv_enter_title(e){
  try{
    RG_TITLE.setValue(e.message.text);
    RG_STATUS.setValue('最終確認中');
    mk_button_chk_final(e);
  }catch(error){
    write_log("rcv_enter_titleでエラーが発生しました。：" + error.message);
    return false;
  }
}

//フォローされたときの動作
function follow(e){
  var message;
  if (chk_userid_auth(e)){
    message = '友達登録ありがとうございます。\n' +
              'メニューから予定を登録できたり、\n' +
              '2週間分の予定を確認することができます。';
    //ステータスを初期状態に
    RG_STATUS.setValue('初期状態');
  }else{
    message = '友達登録ありがとうございます。\n' +
              '申し訳ございませんが、お使いのユーザはこの機能を利用することができません。\n' +
              '詳しくは管理者までお問い合わせください。';
  }
  send_message(message, e);
}
//アンフォローされた時の動作
function unfollow(e){
  let sentence = "アンフォローされました。"
  send_message(sentence, e);
}


function create_schedule() {
  try{
    CALENDAR.createAllDayEvent(RG_TITLE.getValue(), new Date(RG_DATE.getValue()));
  }
  catch(error){
    write_log("create_scheduleでエラーが発生しました。：" + error.message);
    return false;
  }
  return true;
}

//与えられた文字列をデバッグ列に追加する
function write_log(sentence){
  var last_row = LOG_SHEET.getRange(LOG_SHEET.getMaxRows() , 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  LOG_SHEET.getRange(last_row + 1 , 1).setValue(sentence);

  return true;
}

//与えられた文字列をデバッグ列に追加する
function write_log(sentence){
  var last_row = LOG_SHEET.getRange(LOG_SHEET.getMaxRows() , 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  LOG_SHEET.getRange(last_row + 1 , 1).setValue(sentence);
  LOG_SHEET.getRange(last_row + 1 , 2).setValue(dayjs.dayjs().add(13, 'h').format("YYYY/MM/DD HH:mm:ss"));
}

function test(){
  write_log("this is a test");
}
//メッセージを送信する関数
function send_message(sentence, e){
  var message = {
    "replyToken" : e.replyToken,
    "messages" : [
      {
        "type" : "text",
        "text" : sentence
      }
    ]
  }
  fetch_data(message);
}

//応答メッセージをlineへhhtp送信
function fetch_data(postData){
  var replyData = {
    "method" : "post",
    "headers" : {
      "Content-Type" : "application/json",
      "Authorization" : "Bearer " + CHANEL_ACCESS_TOKEN
    },
    "payload" : JSON.stringify(postData)
  };
  var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", replyData);
  if (response.getResponseCode() != 200){
    write_log(response.getResponseCode());
  }
}

//2週間分のスケジュール獲得し、返す関数
function get_2week_events(){
  try{
    var now = new Date();
    var span = new Date(now.getTime() + (14 * 24 * 60 *60 * 1000));
    var events = CALENDAR.getEvents(now, span);
  }
  catch(error){
    write_log("get_2week_eventsでエラーが発生しました。：" + error.message);
    return false;
  }
  return events;
}

//スケジュールデータをテキスト化する関数
function canvert_obj_to_str(events){
  try{
    var jpn_date;
    var rtn_str = "";
    events.forEach(function(event){
      jpn_date = dayjs.dayjs(event.getStartTime()).add(9, 'h').format('YYYY/MM/DD');
      rtn_str += jpn_date + "：\n"
      rtn_str += "　" + event.getTitle() + "\n";
    });
  }
  catch(error){
    write_log("canvert_obj_to_strでエラーが発生しました。：" + error.message);
    return false;
  }
  return rtn_str;
}

//日付の次にタイトルを入力させるメッセージを送信する関数
function mk_button_chk_date(e){
  var message = {
    "replyToken" : e.replyToken,
    "messages" : [
      {
        "type": "template",
        "altText": "This is a buttons template",
        "template": {
          "type": "buttons",
          "text": "選択された日付は、" + dayjs.dayjs(RG_DATE.getValue()).add(9, 'h').format('YYYY/MM/DD') + "です。\n" +
                  "次にタイトルを入力してください。",
        "actions": [
          {
            "type": "postback",
            "label": "Cancel",
            "data": "Cancel",
            "text": "キャンセル"
          },
        ] 
      }
    }
    ]
  }
  fetch_data(message);  
}

//日付の次にタイトルを入力させるメッセージを送信する関数
function mk_button_chk_final(e){
  var message = {
    "replyToken" : e.replyToken,
    "messages" : [
      {
        "type": "template",
        "altText": "This is a buttons template",
        "template": {
          "type": "buttons",
          "text": "選択された日付は、" + dayjs.dayjs(RG_DATE.getValue()).add(9, 'h').format('YYYY/MM/DD') + "で\n" +
                  "タイトルは、\n" + 
                  RG_TITLE.getValue() + "\n" +
                  "です。よろしいですか？",
        "actions": [
                    {
            "type": "postback",
            "label": "はい",
            "data": "Yes",
            "text": "はい"
          },
          {
            "type": "postback",
            "label": "いいえ",
            "data": "No",
            "text": "いいえ"
          },
          {
            "type": "postback",
            "label": "キャンセル",
            "data": "Cancel",
            "text": "キャンセル"
          }
        ] 
      }
    }
    ]
  }
  fetch_data(message);  
}

//ユーザIDが空だったら格納してTrue、同じだったらそのままTrue、違ってたらfalseを返す関数
function chk_userid_auth(e){
  try{
    if (RG_USERID.getValue() == ''){
      RG_USERID.setValue(e.source.userId);
      return true;
    }else if(RG_USERID.getValue() == e.source.userId){
      return true;
    }else{
      return false;
    }
  }
  catch(error){
    write_log("chk_userid_authでエラーが発生しました。：" + error.message);
    return false;
  }
}