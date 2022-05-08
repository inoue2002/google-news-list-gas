//検索したいワードを入力する
const searchWord = '滋賀県+実証実験'

/**
 * xmlを取得してまだ保存していないものがあれば追加する(トリガー時間型・間隔任意)
 */
const rss = () => {
  //最終更新日を取得してパラメータに追加する
  const mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  let targetUrl
  try {
    const lastUpdate = Utilities.formatDate(mySheet.getRange("E1").getValue(), "JST", "yyyy/MM/dd")
    console.log('最終更新日', lastUpdate)
    //最新のものだけアップデートしたい場合
    //targetUrl = `https://news.google.com/rss/search?q=${encodeURI(searchWord)}after:${lastUpdate}&gl=JP&ceid=JP:ja&hl=ja`
    //特に最新にはこだわらず、重複なくアップデートしたい場合
    targetUrl = `https://news.google.com/rss/search?q=${encodeURI(searchWord)}&gl=JP&ceid=JP:ja&hl=ja`
  } catch (e) {
    targetUrl = `https://news.google.com/rss/search?q=${encodeURI(searchWord)}&gl=JP&ceid=JP:ja&hl=ja`
  }
  const urls = getUrls()
  const xml = UrlFetchApp.fetch(targetUrl).getContentText()
  const document = XmlService.parse(xml)
  const root = document.getRootElement()
  const channel = root.getChild('channel')
  const entries = channel.getChildren('item')
  //古い順番に並べる
  entries.reverse()
  console.log(`${entries.length}件の記事を取得しました・・・・`)
  let count = 1
  for (const entry of entries) {
    const link = entry.getChild('link').getText()
    if (!urls.includes(link)) {
      const title = entry.getChild('title').getText()
      const pubDate = entry.getChild('pubDate').getText()
      const appendRow = mySheet.appendRow([, Utilities.formatDate(new Date(pubDate), "JST", "yyyy/MM/dd HH:mm:ss"), title, link])
      const ogpUrl = getOgpImageUrl(link)
      mySheet.getRange(appendRow.getLastRow(), 1).insertCheckboxes(true)
      if (ogpUrl) {
        try {
          const image = SpreadsheetApp.newCellImage().setSourceUrl(ogpUrl).setAltTextTitle("OGP").setAltTextDescription("OGP").build()
          mySheet.getRange(appendRow.getLastRow(), 5).setValue(image)
        } catch (e) {
        }
      }
      console.log(`${count}件目の追加を行いました・・・`)
      count++
      mySheet.getRange("E1").setValue(Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"))
    }
  }
  console.log(`処理終了：${count - 1}件/${entries.length}件中の記事を新しく追加しました。`)
}
/**
 * ogp画像URLを取得
 * @param {string} url - 取得したいページのURL
 * @return {string} ogpUrl - 対象のOGP画像URL
 */
const getOgpImageUrl = (url) => {
  try {
    const xml = UrlFetchApp.fetch(url).getContentText()
    const $ = Cheerio.load(xml)
    const ogpImageUrl = $('meta[property="og:image"]').attr('content')
    return ogpImageUrl
  } catch (e) {
    return
  }
}
/**
 * 変更を検知する関数（トリガー・変更時）
 */
function editMain() {
  //変更されたシートを取得
  const editSheet = SpreadsheetApp.getActiveSheet()
  const editSheetName = editSheet.getName()
  const editCell = editSheet.getActiveCell()
  const status = editCell.getValue()
  //変更がチェックボックスかつ、✔︎が入った時
  if (status === "TRUE" || status === true) {
    //チェックが入った行を取得
    const row = editCell.getRow()
    const column = editCell.getColumn()
    if (editSheetName === 'リーディングリスト') {
      const entry = {
        pubDate: editSheet.getRange(row, 2).getValue(),
        title: editSheet.getRange(row, 3).getValue(),
        link: editSheet.getRange(row, 4).getValue(),
        ogpUrl: editSheet.getRange(row, 5).getValue()
      }
      console.log('entry', entry)
      read(entry)
      editSheet.deleteRow(row)
    }
    else {
    }
  }
}
/**
 * 読んだ記事を読了シートへ移動
 * @param {number} row - 対象行
 * @param {object} entry - 記事
 * @param {string} entry.title - 記事タイトル
 * @param {string} entry.link - 記事リンク
 * @param {date} entry.pubDate - 公開日
 * @prama {string} entry.ogpUrl - OGP画像
 */
function read(entry) {
  console.log(entry)
  const readSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('読了')
  readSheet.appendRow([entry.pubDate, entry.title, entry.link, entry.ogpUrl])
}
/**
 * 現在あるURLを全て取得
 * @return string[] urls - 記事URL
 */
function getUrls() {
  const mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const mySheetLastRow = mySheet.getLastRow()
  const readSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('読了')
  const readSheetLastRow = readSheet.getLastRow()
  const entryUrls = mySheet.getRange(`D3:D${mySheetLastRow}`).getValues()
  const readUrls = readSheet.getRange(`C2:C${readSheetLastRow}`).getValues()
  const allUrls = entryUrls.concat(readUrls)
  const urls = allUrls.map(url => { return url[0] })
  return urls
}
function setUp() {
  //シート1を「リーディングシート」にリネーム
  //シートを1つ追加「読了」にリネーム

  //リーディングシート上2つを青色で塗りつぶす


  //変更トリガーを追加

  //rssトリガーを追加
/**
 * 新しいシートの作成&名前を変更
 * @param {string} name - 変えたい名前
 * @return {sheet} newSheet - 追加したシート
 */
function createNewSheet(name) {
  const mySheet = SpreadsheetApp.getActiveSpreadsheet()
  //スプレッドシートに新しいシートを追加挿入
  let newSheet = mySheet.insertSheet()
  //追加挿入したシートに名前を設定
  newSheet.setName(name)
  return newSheet
}
/**
 * 列をまとめて削除
 * @params {sheet} sheet - 対象のシート
 * @params {number} columnPosition - 開始列
 * @param {string} howMany - 終了列
 */
function deleteSheetColumns(sheet, columnPosition, howMany) {
  try {
    sheet.deleteColumns(columnPosition, howMany)
  } catch (e) {
    console.error('deleteSheetColumnsErr:', e)
  }
}