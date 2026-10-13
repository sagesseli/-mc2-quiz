/**
 * 時代華語2 完課測驗 — 成績接收程式（Google Apps Script）
 * 安裝方式請看 README.md。
 *
 * 這個程式會在你的 Google 試算表裡自動建立：
 *   「成績冊」：一位學生一列，每課一欄（L01、L02……），方便登記成績
 *   「作答紀錄」：每一次交卷一列（含各大題分數、錯題、作答時間、第幾次）
 */

// 成績要寫進的試算表（國專G春季班 成績）
const SHEET_ID = "1I2iHmCFYk8I2q--oeE5jNKpaohiScpP4YPJLxzGneoc";

// 成績冊要記哪一次的分數："first"＝第一次、"best"＝最高分、"latest"＝最後一次
const SCORE_POLICY = "first";

const LOG_SHEET = "作答紀錄";
const BOOK_SHEET = "成績冊";
const LOG_HEAD = ["交卷時間", "課別", "班級", "座號", "姓名", "第幾次", "得分", "滿分", "百分比", "作答分鐘", "各大題", "錯題"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = sheet_(ss, LOG_SHEET, LOG_HEAD);
    const secText = (d.sectionOrder || Object.keys(d.sections || {}))
      .map(k => k + " " + d.sections[k]).join("｜");
    log.appendRow([new Date(), d.lesson, d.klass, String(d.seat), d.name, d.attempt,
                   d.score, d.max, d.pct / 100, d.minutes, secText, d.wrong]);
    log.getRange(log.getLastRow(), 9).setNumberFormat("0.0%");
    updateBook_(ss, d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, msg: "成績接收程式正常運作中" });
}

function updateBook_(ss, d) {
  const lessons = [];
  for (let i = 1; i <= 16; i++) lessons.push("L" + ("0" + i).slice(-2));
  const book = sheet_(ss, BOOK_SHEET, ["班級", "座號", "姓名"].concat(lessons).concat(["平均"]));
  const col = 4 + lessons.indexOf(d.lesson);
  if (col < 4) return;
  const last = book.getLastRow();
  let row = -1;
  if (last > 1) {
    const vals = book.getRange(2, 1, last - 1, 3).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === String(d.klass) && vals[i][1] === String(d.seat) && vals[i][2] === String(d.name)) { row = i + 2; break; }
    }
  }
  if (row < 0) {
    row = last + 1;
    book.getRange(row, 1, 1, 3).setNumberFormat("@").setValues([[d.klass, String(d.seat), d.name]]);
    const a = book.getRange(row, 4).getA1Notation(), b = book.getRange(row, 3 + lessons.length).getA1Notation();
    book.getRange(row, 4 + lessons.length).setFormula('=IFERROR(ROUND(AVERAGE(' + a + ':' + b + '),1),"")');
  }
  const cell = book.getRange(row, col);
  const old = cell.getValue();
  const score = Math.round(d.pct);
  if (old === "" || SCORE_POLICY === "latest" || (SCORE_POLICY === "best" && score > Number(old))) {
    cell.setValue(score);
  }
  if (d.attempt > 1) cell.setNote("交卷 " + d.attempt + " 次，最新一次 " + score + " 分");
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#e6e9f6");
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
