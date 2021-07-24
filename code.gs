const CALENDER_ID = PropertiesService.getScriptProperties().getProperty('CALENDER_ID')
const URL = {
  MEN: 'https://www.gorin.jp/game/BKBMTEAM5-------------/',
  WOMEN: 'https://www.gorin.jp/game/BKBWTEAM5-------------/',
  MEN_3x3: 'https://www.gorin.jp/game/BK3MTEAM3-------------/',
  WOMEN_3x3: 'https://www.gorin.jp/game/BK3WTEAM3-------------/'
}
const CLASS = {
  MEN: '男子',
  WOMEN: '女子',
  MEN_3x3: '男子 3x3',
  WOMEN_3x3: '女子 3x3'
}
const SHEET_NAME = 'GorinJpから出力するシート'

/** URLから日程データを取得 */
function getCalender(url) {
  const html = UrlFetchApp.fetch(url).getContentText()
  const gameList = Parser.data(html).from('<div class="list-in">').to('</div>').iterate()
  let data = [];
  gameList.map((game, index) => {
    if(!game.match( url === URL.MEN ? /日本|米国/ : /日本/)) return // 男子だけはアメリカも含み、それ以外は日本以外無視
    game = game.replace('<span> - </span>', '') // 不要な部分を削除
    const time = Parser.data(game).from('<span class="ms">').to('</span>').build() // 時間を抜き出す ex: '7/29 13:40'
    const competition =  Parser.data(game).from('<span>').to('</span>').build() // 対戦を抜き出す ex: '日本 vs スロベニア'
    data = [...data, {time: time, competition: competition}]
  })
  return data
}

/** 日程データをSpreadsheetに入力 */
function postEventsToSheet() {
  const data = [{
    class: CLASS.MEN,
    calendar: getCalender(URL.MEN),
    url: URL.MEN
  },
  {
    class: CLASS.WOMEN,
    calendar: getCalender(URL.WOMEN),
    url: URL.WOMEN
  },
  {
    class: CLASS.MEN_3x3,
    calendar: getCalender(URL.MEN_3x3),
    url: URL.MEN_3x3
  },
  {
    class: CLASS.WOMEN_3x3,
    calendar: getCalender(URL.WOMEN_3x3),
    url: URL.WOMEN_3x3
  }]
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.clear() // シートの中身を消す
  sheet.appendRow(['種別', '開始時間', '対戦', 'URL'])
  data.map((data) => { // dataを1つずつ処理
    data.calendar.map((calender) => { // data.calenderを1つずつ処理
      sheet.appendRow([data.class, calender.time, calender.competition, data.url])
    })
  })
}

/** 日程データをカレンダーにセット */
function putCalenderEvents() {
  postEventsToSheet()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const calendar = CalendarApp.getCalendarById(CALENDER_ID) // 対象のカレンダーをCalenderオブジェクトとして取得
  const data = sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
  data.map((data, index) => {
    /** タイトル */
    const title = `東京オリンピック バスケットボール ${data[0]} ${data[2]}`
    /** 開始時間 */
    const startDate = new Date(data[1])
    /** 3x3ではないとき試合は2時間 */
    const gametimeHour = (data[0] === CLASS.MEN || data[0] === CLASS.WOMEN) ? 2 : 0
    /** 3x3の場合試合は20分 */
    const gametimeMinutes = (data[0] === CLASS.MEN || data[0] === CLASS.WOMEN) ? 0 : 20
    /** 終了時間 3x3かそうでないかで試合終了時間を計算 */
    const endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate(), startDate.getHours() + gametimeHour, startDate.getMinutes() + gametimeMinutes)
    const url = data[3]
    if(calendar.getEvents(startDate, endDate, {search: title}).length > 0) return // 既にイベント作ってあったら無視。開始時間から終了時間の間に同じタイトルのものがあれば無視
    calendar.createEvent(title, startDate, endDate, { description: `${url}` }) // イベント作成
  })
}
