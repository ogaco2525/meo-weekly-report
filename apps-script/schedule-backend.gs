// =============================================
// MEO相談 予約バックエンド
// ogaco2525@gmail.com のGoogleカレンダーと連携
// =============================================

const CALENDAR_ID = 'primary';          // メインカレンダー
const SLOT_HOURS  = 1;                  // 1コマ何時間
const START_HOUR  = 9;                  // 開始時刻（9:00）
const END_HOUR    = 14;                 // 終了時刻（14:00まで開始）
const DAYS_AHEAD  = 14;                 // 何日先まで表示するか
const NOTIFY_EMAIL = 'ogaco2525@gmail.com'; // 通知先メール
const EVENT_TITLE  = '【MEO相談】';     // カレンダーに登録されるタイトル

// =============================================
// GET: 空き枠一覧を返す
// =============================================
function doGet(e) {
  const slots = getAvailableSlots();
  return jsonResponse({ slots });
}

// =============================================
// POST: 予約を受け付ける
// =============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { start, name, phone, note } = data;

    if (!start || !name || !phone) {
      return jsonResponse({ ok: false, error: '必須項目が不足しています' });
    }

    const startDate = new Date(start);
    const endDate   = new Date(startDate.getTime() + SLOT_HOURS * 60 * 60 * 1000);

    // 再度空き確認（二重予約防止）
    const cal    = CalendarApp.getCalendarById(CALENDAR_ID);
    const events = cal.getEvents(startDate, endDate);
    if (events.length > 0) {
      return jsonResponse({ ok: false, error: 'この時間帯はすでに埋まっています。別の日程をお選びください。' });
    }

    // カレンダーに登録
    const desc = `氏名: ${name}\n電話: ${phone}\n備考: ${note || 'なし'}`;
    cal.createEvent(`${EVENT_TITLE}${name}様`, startDate, endDate, { description: desc });

    // メール通知
    const dateStr = Utilities.formatDate(startDate, 'Asia/Tokyo', 'M月d日(EEE) HH:mm');
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: `【MEO相談 予約】${dateStr} ${name}様`,
      body: `新しい予約が入りました。\n\n日時: ${dateStr}〜\n氏名: ${name}\n電話: ${phone}\n備考: ${note || 'なし'}\n\n※Googleカレンダーに自動登録済みです。`,
    });

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// =============================================
// 空き枠を計算して返す
// =============================================
function getAvailableSlots() {
  const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  const now   = new Date();
  const slots = [];

  for (let d = 0; d < DAYS_AHEAD; d++) {
    const day = new Date(now);
    day.setDate(now.getDate() + d + 1); // 翌日から
    day.setHours(0, 0, 0, 0);

    const dow = day.getDay();
    if (dow === 0 || dow === 6) continue; // 土日スキップ

    // その日の既存イベントを取得
    const dayEnd = new Date(day);
    dayEnd.setHours(23, 59, 59);
    const events = cal.getEvents(day, dayEnd);
    const busyRanges = events.map(ev => ({
      start: ev.getStartTime(),
      end:   ev.getEndTime(),
    }));

    // 9〜14時のスロットをチェック
    for (let h = START_HOUR; h < END_HOUR; h++) {
      const slotStart = new Date(day);
      slotStart.setHours(h, 0, 0, 0);
      const slotEnd = new Date(slotStart);
      slotEnd.setHours(h + SLOT_HOURS, 0, 0, 0);

      // 過去はスキップ
      if (slotStart <= now) continue;

      // 既存予定と重複チェック
      const isBusy = busyRanges.some(b => slotStart < b.end && slotEnd > b.start);
      if (!isBusy) {
        slots.push({
          start:   slotStart.toISOString(),
          label:   Utilities.formatDate(slotStart, 'Asia/Tokyo', 'M月d日(EEE) HH:mm'),
          endLabel: Utilities.formatDate(slotEnd, 'Asia/Tokyo', 'HH:mm'),
        });
      }
    }
  }

  return slots;
}

// =============================================
// ユーティリティ
// =============================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
