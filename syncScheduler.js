// 設定一覧
const CONFIG = {
  // 同期先カレンダー名
  CALENDAR_NAME: "Brass Band ROAR! 練習日程",
  // 開始時刻セルが空のときの既定値
  DEFAULT_START: "19:30",
  // 終了時刻セルが空のときの既定値
  DEFAULT_END: "22:00",
};

// スプレッドシートのメニューに同期ボタンを追加する
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("操作用")
    .addItem("カレンダーを同期", "syncCurrentSheet")
    .addToUi();
}

// Googleカレンダーへの同期処理
function syncCurrentSheet() {
  // シートとカレンダーを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const calendar = getOrCreateCalendar_(CONFIG.CALENDAR_NAME);

  // ヘッダ行を取得して列インデックスを作成
  const [header, ...rows] = sheet.getDataRange().getValues();
  const col = header.reduce((m, v, i) => ((m[v] = i), m), {});

  rows.forEach((row, rIdx) => {
    // 日付と時刻を取得、空白なら何もしない
    const date = row[col["日付"]];
    if (!date) return;

    // 各データを取得
    const startRaw = row[col["開始時刻"]] || CONFIG.DEFAULT_START;
    const endRaw = row[col["終了時刻"]] || CONFIG.DEFAULT_END;
    const venue = row[col["会場"]] || "";
    const memo = row[col["備考(持ち物など)"]] || "";
    const drums = row[col["使用打楽器"]] || "";
    // 日付は Date オブジェクトに変換しておく
    const startAt = mergeDateTime_(date, startRaw);
    const endAt = mergeDateTime_(date, endRaw);

    // Googleカレンダー表示用の文字列を作成する
    const fmtTime = (d) =>
      Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
    const fmtDate = (d) =>
      Utilities.formatDate(d, Session.getScriptTimeZone(), "M月d日");
    const title = `練習 @ ${venue || "未定"}`;
    const desc = [
      `🎺 本日（${fmtDate(startAt)}）の練習案内 🎺`,
      "",
      `🕒 時間：${fmtTime(startAt)}～${fmtTime(endAt)}`,
      `📍 場所：${venue || "未定"}`,
      `🥁 使用打楽器：${drums || "—"}`,
      `🎼 備考(持ち物など)：${memo || "—"}`,
    ].join("\n");

    // イベントIDを取得
    const idCell = sheet.getRange(rIdx + 2, col["EventID"] + 1);
    const savedId = idCell.getValue();

    // 予定の作成・編集
    let ev;
    // イベントIDが存在する場合は、既存のイベントを取得を試みる
    // このとき、取得できなくてもエラーは無視する
    if (savedId) {
      try {
        ev = calendar.getEventById(savedId);
      } catch (e) {}
    }
    // イベントが存在する場合は更新
    if (ev) {
      ev.setTime(startAt, endAt)
        .setTitle(title)
        .setLocation(venue)
        .setDescription(desc);
      return;
    }
    // イベントが存在しない場合は新規作成
    ev = calendar.createEvent(title, startAt, endAt, {
      location: venue,
      description: desc,
    });
    // IDを反映しておく
    idCell.setValue(ev.getId());
  });

  // 全て完了したら完了アラートを表示する
  SpreadsheetApp.getUi().alert("カレンダー同期が完了しました ✅");
}

// 時刻と日付を結合するヘルパ関数
function mergeDateTime_(d, t) {
  // t は Date でも "HH:MM" 文字列でも OK
  let h, m;
  if (t instanceof Date) {
    h = t.getHours();
    m = t.getMinutes();
  } else {
    [h, m] = String(t).split(":").map(Number);
  }
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), h, m, 0);
}

// カレンダーを取得する、なければ作成するヘルパ関数
function getOrCreateCalendar_(name) {
  const [cal] = CalendarApp.getCalendarsByName(name);
  return cal || CalendarApp.createCalendar(name);
}
