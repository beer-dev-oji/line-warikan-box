/**
 * ==============================================================================
 * プロジェクト名: みんなのわりかん箱 (Warikan Box)
 * ファイル名: コード.gs
 * 役割: バックエンドロジック (GAS Web App / LINE Messaging API Webhook / スプレッドシートDB操作)
 * 設計思想:
 *   - サーバーレス＆完全無料運用（GAS + スプレッドシート + Messaging API無料枠死の回避）
 *   - マルチテナント設計（LINEグループごとの完全データ分離）
 *   - データの整合性と監査ログ維持（論理削除アーキテクチャ）
 *   - 最小送金回数相殺アルゴリズムによる送金コストの極小化
 * ==============================================================================
 */

// ==========================================
// 環境設定・定数定義（GitHub公開用にマスキング済み）
// ==========================================

/** LINE Messaging API の長期チャネルアクセストークン */
const LINE_CHANNEL_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN';

/** Google Apps Script の公開WebアプリURL（末尾が /exec のもの） */
const WEB_APP_URL = 'YOUR_GAS_WEB_APP_URL';

// ==========================================
// 1. Web画面 表示用
// ==========================================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "";

  // ★ 規約・ポリシーページ（?page=terms のとき）
  if (page === "terms") {
    return HtmlService.createHtmlOutputFromFile("terms")
      .setTitle("利用規約・プライバシーポリシー | みんなのわりかん箱")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // ★ 通常の割り勘アプリ画面
  const template = HtmlService.createTemplateFromFile("index");
  template.serverGroupId = (e && e.parameter && e.parameter.groupId) || "G_TEST_GROUP";
  template.serverEventId = (e && e.parameter && e.parameter.eventId) || "";
  template.serverWebAppUrl = WEB_APP_URL; // 規約リンク用に渡す

  return template.evaluate()
    .setTitle("みんなのわりかん箱")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * 画面起動時に必要なデータ（指定グループのイベント・メンバー・履歴）を一括取得する
 */
function getInitialData(groupId, requestedEventId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. イベント一覧取得
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < eventRows.length; i++) {
    const [id, gId, name, status, createdAt] = eventRows[i];
    if (id && String(gId) === targetGroupId) {
      events.push({ id: String(id), name: String(name), status: String(status) });
    }
  }

  // 指定の eventId があれば優先、なければ active、それもなければ最後のイベント
  let currentEvent = null;
  if (requestedEventId) {
    currentEvent = events.find(e => e.id === String(requestedEventId));
  }
  if (!currentEvent) {
    currentEvent = events.filter(e => e.status === "active").pop();
  }
  if (!currentEvent && events.length > 0) {
    currentEvent = events[events.length - 1];
  }

  // 2. メンバー一覧取得（★ mergedTo を追加！）
  const memberSheet = ss.getSheetByName("Members");
  const memberRows = memberSheet.getDataRange().getValues();
  const members = [];
  const memberNameMap = {};

  for (let i = 1; i < memberRows.length; i++) {
    const [gId, id, name, lineId, paypay, status, mergedTo] = memberRows[i];
    if (id && String(gId) === targetGroupId) {
      const currentStatus = String(status || "active");
      members.push({
        id: String(id),
        name: String(name),
        paypay: String(paypay || ""),
        status: currentStatus,
        mergedTo: String(mergedTo || "") // ★ G列の合流先ID
      });
      memberNameMap[String(id)] = String(name);
    }
  }

  // 3. 支払い履歴一覧取得（deleted は status としてフロントに渡す）
  const paymentSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentSheet.getDataRange().getValues();
  const payments = [];
  const targetEventIds = events.map(e => e.id);

  for (let i = 1; i < paymentRows.length; i++) {
    const [pId, eId, payerId, amount, desc, targets, createdAt, status] = paymentRows[i];
    if (pId && targetEventIds.includes(String(eId))) {
      const targetIds = String(targets || "").split(",").map(s => s.trim()).filter(Boolean);
      const targetNames = targetIds.map(tId => memberNameMap[tId] || tId);
      const paymentStatus = String(status || "active");

      payments.unshift({
        id: String(pId),
        eventId: String(eId),
        payerName: memberNameMap[payerId] || String(payerId),
        amount: Number(amount),
        description: String(desc),
        targetNames: targetNames,
        isAllTargets: targetNames.length === members.filter(m => m.status === "active").length,
        date: createdAt ? String(createdAt) : "",
        status: paymentStatus
      });
    }
  }

  return {
    groupId: targetGroupId,
    currentEventId: currentEvent ? currentEvent.id : null,
    events: events,
    members: members,
    payments: payments
  };
}

/**
 * メンバーのステータス（active / inactive）を切り替える関数
 */
function toggleMemberStatus(groupId, memberId, newStatus) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === targetGroupId && String(rows[i][1]) === String(memberId)) {
      sheet.getRange(i + 1, 6).setValue(newStatus);
      return { status: "success", memberId: memberId, newStatus: newStatus };
    }
  }

  throw new Error("対象のメンバーが見つかりませんでした");
}

/**
 * 手入力メンバーを別のメンバー（自分）に統合する関数
 */
function mergeMember(groupId, fromMemberId, toMemberId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  if (!fromMemberId || !toMemberId || fromMemberId === toMemberId) {
    throw new Error("統合するメンバーの指定が不正です");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const rows = sheet.getDataRange().getValues();

  let targetRowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === targetGroupId && String(rows[i][1]) === String(fromMemberId)) {
      targetRowIndex = i + 1; // 1始まりの行番号
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("統合対象のメンバーが見つかりませんでした");
  }

  // 1. ステータス（6列目/F列）を "merged" に更新（入力欄やアクティブ一覧から非表示化）
  sheet.getRange(targetRowIndex, 6).setValue("merged");

  // 2. 合流先（7列目/G列）に toMemberId を書き込む！
  sheet.getRange(targetRowIndex, 7).setValue(String(toMemberId));

  console.log(`🔗 メンバー統合成功: ${fromMemberId} ➔ ${toMemberId}`);
  return { status: "success", fromMemberId: fromMemberId, toMemberId: toMemberId };
}

/**
 * 新しいメンバーを追加する（グループID紐付け）
 */
function addMemberFromWeb(groupId, name) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const cleanName = String(name).trim().slice(0, 30);
  if (!cleanName) throw new Error("名前が空です");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const lastRow = sheet.getLastRow();
  const newMemberId = "M" + String(lastRow).padStart(3, "0");

  sheet.appendRow([targetGroupId, newMemberId, cleanName, "", "", "active"]);
  return { id: newMemberId, name: cleanName };
}

/**
 * 新しいイベントを作成する（グループID紐付け）
 */
function createEventFromWeb(groupId, name, shouldClosePrevious) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const cleanName = String(name).trim().slice(0, 30);
  if (!cleanName) throw new Error("イベント名が空です");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Events");
  const rows = sheet.getDataRange().getValues();

  if (shouldClosePrevious) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === targetGroupId && rows[i][3] === "active") {
        sheet.getRange(i + 1, 4).setValue("closed");
      }
    }
  }

  const lastRow = sheet.getLastRow();
  const newEventId = "E" + String(lastRow).padStart(3, "0");
  const dateStr = new Date().toLocaleString("ja-JP");

  sheet.appendRow([newEventId, targetGroupId, cleanName, "active", dateStr]);
  return { id: newEventId, name: cleanName };
}

/**
 * 支払いを記録する
 */
function recordPaymentFromWeb(data) {
  const amount = Number(data.amount);
  if (isNaN(amount) || amount <= 0 || amount > 99999999) {
    throw new Error("金額が不正です（1円〜9,999万円まで）");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][0]) === String(data.eventId) && eventRows[i][3] === "closed") {
      throw new Error("このイベントは終了しているため記録できません");
    }
  }

  const sheet = ss.getSheetByName("Payments");
  const lastRow = sheet.getLastRow();
  const paymentId = "P" + String(lastRow).padStart(3, "0");
  const cleanDesc = String(data.description).trim().slice(0, 50);

  const newRow = [
    paymentId,
    String(data.eventId),
    String(data.payerId),
    amount,
    cleanDesc,
    String(data.targetMemberIds),
    new Date().toLocaleString("ja-JP"),
    "active" // status
  ];

  sheet.appendRow(newRow);
  return { status: "success" };
}

/**
 * 物理削除ではなく「論理削除（ステータスを deleted に更新）」する関数
 */
function cancelPaymentFromWeb(groupId, paymentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentSheet.getDataRange().getValues();

  for (let i = 1; i < paymentRows.length; i++) {
    if (String(paymentRows[i][0]) === String(paymentId)) {
      const targetEventId = String(paymentRows[i][1]);

      // イベントがすでに closed なら取り消し不可
      const eventSheet = ss.getSheetByName("Events");
      const eventRows = eventSheet.getDataRange().getValues();
      for (let j = 1; j < eventRows.length; j++) {
        if (String(eventRows[j][0]) === targetEventId && eventRows[j][3] === "closed") {
          throw new Error("終了済みのイベントの支払いは取り消せません");
        }
      }

      // 8列目（H列）のステータスを "deleted" に更新
      paymentSheet.getRange(i + 1, 8).setValue("deleted");
      return { status: "success" };
    }
  }
  throw new Error("対象の支払いデータが見つかりませんでした");
}

// ==========================================
// 2. LINE Messaging API 連携設定
// ==========================================

/**
 * LINEからのメッセージ（Webhook）を受け取る関数（Botメンション ＆ ボタン参加対応版）
 */
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const events = json.events;
    if (!events || events.length === 0) {
      return ContentService.createTextOutput("ok");
    }

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const currentGroupId = (event.source && (event.source.groupId || event.source.roomId || event.source.userId)) || "G_TEST_GROUP";
      const currentUserId = (event.source && event.source.userId) || "";

      // ★ 1. ボタンが押されたとき（Postbackイベント）
      if (event.type === "postback") {
        const postbackData = event.postback.data || "";

        if (postbackData.includes("action=join")) {
          // ボタンを押した人の名前をLINEから取得してスプシに登録！
          registerMemberFromLine(currentGroupId, currentUserId);
          // ※トーク荒れ防止のためBotは無言（200 OKだけ返して終了）
        }
        continue;
      }

      // 2. グループ招待時の挨拶
      if (event.type === "join") {
        // 招待時もボタン付きカードをドーンと返信！
        const flexMsg = createMainFlexMessage(currentGroupId);
        replyFlexToLine(event.replyToken, flexMsg);
        continue;
      }

      // 3. メッセージ受信
      if (event.type === "message" && event.message.type === "text") {
        const rawText = event.message.text.trim();
        const replyToken = event.replyToken;
        const text = rawText.replace(/＠/g, "@").toLowerCase();

        // Botメンション判定
        const isMentioned = 
          text.includes("@みんなのわりかん箱") || 
          text.includes("@わりかん箱") || 
          text.startsWith("@わりかん") || 
          text.startsWith("@warihako");

        if (!isMentioned) {
          continue; // 普段の雑談は完全スルー
        }

        // 「精算」が含まれているか判定
        const isSettlement = 
          text.includes("精算") || 
          text.includes("清算") || 
          text.includes("せいさん");

        if (isSettlement) {
          const settlementMessage = executeSettlement(currentGroupId);
          replyToLine(replyToken, settlementMessage);
        } else {
          // 基本（メンション時）：入力ボタン ＆ 参加ボタンが付いたカードを返信！
          const flexMsg = createMainFlexMessage(currentGroupId);
          replyFlexToLine(replyToken, flexMsg);
        }
      }
    }
  } catch (error) {
    console.error("❌ Webhookエラー:", error.stack || error.message);
  }

  return ContentService.createTextOutput("ok");
}

/**
 * 指定グループのアクティブイベントを精算して closed に倒す
 */
function executeSettlement(groupId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  let currentEvent = null;
  let eventRowIndex = -1;

  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][1]) === targetGroupId && eventRows[i][3] === "active") {
      currentEvent = { id: String(eventRows[i][0]), name: String(eventRows[i][2]) };
      eventRowIndex = i + 1; // 行番号（1始まり）
      break;
    }
  }
  if (!currentEvent) {
    return "このグループには現在進行中のイベントがありません💦\n「@みんなのわりかん箱」から新しいイベントを作成してください！";
  }

  // ★ 1. ステータスを即座に closed に更新！（これでDBもロックされる）
  eventSheet.getRange(eventRowIndex, 4).setValue("closed");

  // 2. メンバー情報取得（★ aliasMap を追加し、merged を除外！）
  const membersSheet = ss.getSheetByName("Members");
  const memberRows = membersSheet.getDataRange().getValues();
  const memberMap = {};
  const aliasMap = {}; // fromId -> toId（合流辞書）

  for (let i = 1; i < memberRows.length; i++) {
    const [gId, memberId, name, lineId, paypay, status, mergedTo] = memberRows[i];
    if (memberId && String(gId) === targetGroupId) {
      // 統合された人（merged）は精算一覧から除外
      if (status !== "merged") {
        memberMap[String(memberId)] = { name: String(name), paypay: String(paypay || "") };
      }
      // 合流先があれば辞書に登録
      if (mergedTo) {
        aliasMap[String(memberId)] = String(mergedTo);
      }
    }
  }

  // 3. 収支計算（★ aliasMap で合流先へ自動変換！）
  const balances = {};
  Object.keys(memberMap).forEach(id => balances[id] = 0);
  const paymentsSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentsSheet.getDataRange().getValues();
  let hasPayment = false;

  for (let i = 1; i < paymentRows.length; i++) {
    const [paymentId, eventId, payerId, amount, desc, targetsStr, createdAt, status] = paymentRows[i];
    const pStatus = String(status || "active");
    if (String(eventId) === currentEvent.id && amount && pStatus !== "deleted") {
      hasPayment = true;
      // 支払者を合流先へ変換
      const actualPayerId = aliasMap[String(payerId)] || String(payerId);
      
      // 対象者も合流先へ変換し、重複を除去
      const rawTargets = String(targetsStr).split(",").map(s => s.trim()).filter(Boolean);
      const actualTargets = [...new Set(rawTargets.map(tId => aliasMap[tId] || tId))];

      const splitAmount = Math.floor(Number(amount) / actualTargets.length);
      if (balances[actualPayerId] !== undefined) balances[actualPayerId] += Number(amount);
      actualTargets.forEach(tId => {
        if (balances[tId] !== undefined) balances[tId] -= splitAmount;
      });
    }
  }

  // 支払いが0件だった場合
  if (!hasPayment) {
    return `🎉【精算確定】${currentEvent.name}\n支払いの記録はありませんでした！イベントを終了しました📦`;
  }

  // 4. 最小送金計算
  let creditors = [], debtors = [];
  for (const id in balances) {
    if (balances[id] > 0) creditors.push({ id: id, amount: balances[id] });
    else if (balances[id] < 0) debtors.push({ id: id, amount: -balances[id] });
  }
  creditors.sort((a, b) => b.amount - a.amount);
  debtors.sort((a, b) => b.amount - a.amount);

  const settlements = [];
  let cIdx = 0, dIdx = 0;
  while (cIdx < creditors.length && dIdx < debtors.length) {
    const settleAmount = Math.min(creditors[cIdx].amount, debtors[dIdx].amount);
    settlements.push({
      fromName: memberMap[debtors[dIdx].id] ? memberMap[debtors[dIdx].id].name : "不明",
      toName: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].name : "不明",
      amount: settleAmount,
      toPayPay: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].paypay : ""
    });
    creditors[cIdx].amount -= settleAmount;
    debtors[dIdx].amount -= settleAmount;
    if (creditors[cIdx].amount === 0) cIdx++;
    if (debtors[dIdx].amount === 0) dIdx++;
  }

  // 5. メッセージ組み立て（★明細閲覧リンクを添える！）
  let msg = `🎉【精算確定】${currentEvent.name}\n─────────────────\n`;
  if (settlements.length === 0) {
    msg += `全員の支払いがピッタリ相殺されています！送金の必要はありません✨\n`;
  } else {
    settlements.forEach((s, idx) => {
      msg += `【送金 ${idx + 1}】\n👤 ${s.fromName} ➔ ${s.toName}\n💴 ¥${s.amount.toLocaleString()}\n`;
      if (s.toPayPay) msg += `🔗 ${s.toPayPay}\n`;
      msg += `\n`;
    });
    msg += `─────────────────\n※トーク画面左下の「＋」➔「送る・受け取る」からPayPayで送金できます！\n\n`;
  }

  // ★ 閲覧専用URLを添えて、全員が過去ログを安心して見返せるようにする！
  const detailUrl = `${WEB_APP_URL}?groupId=${targetGroupId}&eventId=${currentEvent.id}`;
  msg += `📜【支払い明細・履歴の確認はこちら】👇\n${detailUrl}`;

  return msg.trim();
}

function replyToLine(replyToken, text) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    replyToken: replyToken,
    messages: [{ type: "text", text: text }]
  };
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload)
  });
}

// ==========================================
// Web画面から精算を確定し、LINEに通知する関数
// ==========================================
function closeAndSettleEventFromWeb(groupId, eventId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();

  let eventName = "";
  let targetRow = -1;

  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][0]) === String(eventId) && String(eventRows[i][1]) === targetGroupId) {
      targetRow = i + 1;
      eventName = String(eventRows[i][2]);
      break;
    }
  }

  if (targetRow === -1) throw new Error("対象のイベントが見つかりませんでした");

  // 1. ステータスを closed に更新
  eventSheet.getRange(targetRow, 4).setValue("closed");

  // 2. 精算メッセージを作成（executeSettlement のロジックを活用）
  const settlementMsg = generateSettlementMessage(targetGroupId, eventId, eventName);

  // 3. 閲覧専用URLを添えて LINE グループにプッシュ送信！
  const detailUrl = `${WEB_APP_URL}?groupId=${targetGroupId}&eventId=${eventId}`;
  const finalMessage = 
    settlementMsg + `\n\n` +
    `📜【支払い明細・履歴の確認はこちら】👇\n${detailUrl}`;

  pushToLineGroup(targetGroupId, finalMessage);

  return { status: "success", eventId: eventId };
}

// 精算メッセージ本文を生成する共通ヘルパー（★ aliasMap対応版）
function generateSettlementMessage(groupId, eventId, eventName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName("Members");
  const memberRows = membersSheet.getDataRange().getValues();
  const memberMap = {};
  const aliasMap = {}; // fromId -> toId

  for (let i = 1; i < memberRows.length; i++) {
    const [gId, mId, name, lineId, paypay, status, mergedTo] = memberRows[i];
    if (mId && String(gId) === groupId) {
      if (status !== "merged") {
        memberMap[String(mId)] = { name: String(name), paypay: String(paypay || "") };
      }
      if (mergedTo) {
        aliasMap[String(mId)] = String(mergedTo);
      }
    }
  }

  const balances = {};
  Object.keys(memberMap).forEach(id => balances[id] = 0);

  const paymentsSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentsSheet.getDataRange().getValues();
  let hasPayment = false;

  for (let i = 1; i < paymentRows.length; i++) {
    const [pId, eId, payerId, amount, desc, targetsStr, createdAt, status] = paymentRows[i];
    const pStatus = String(status || "active");
    if (String(eId) === String(eventId) && amount && pStatus !== "deleted") {
      hasPayment = true;
      const actualPayerId = aliasMap[String(payerId)] || String(payerId);
      const rawTargets = String(targetsStr).split(",").map(s => s.trim()).filter(Boolean);
      const actualTargets = [...new Set(rawTargets.map(tId => aliasMap[tId] || tId))];

      const splitAmount = Math.floor(Number(amount) / actualTargets.length);
      if (balances[actualPayerId] !== undefined) balances[actualPayerId] += Number(amount);
      actualTargets.forEach(tId => {
        if (balances[tId] !== undefined) balances[tId] -= splitAmount;
      });
    }
  }

  if (!hasPayment) return `🎉【精算完了】${eventName}\n支払いの記録はありませんでした！`;

  let creditors = [], debtors = [];
  for (const id in balances) {
    if (balances[id] > 0) creditors.push({ id: id, amount: balances[id] });
    else if (balances[id] < 0) debtors.push({ id: id, amount: -balances[id] });
  }
  creditors.sort((a, b) => b.amount - a.amount);
  debtors.sort((a, b) => b.amount - a.amount);

  const settlements = [];
  let cIdx = 0, dIdx = 0;
  while (cIdx < creditors.length && dIdx < debtors.length) {
    const settleAmount = Math.min(creditors[cIdx].amount, debtors[dIdx].amount);
    settlements.push({
      fromName: memberMap[debtors[dIdx].id] ? memberMap[debtors[dIdx].id].name : "不明",
      toName: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].name : "不明",
      amount: settleAmount,
      toPayPay: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].paypay : ""
    });
    creditors[cIdx].amount -= settleAmount;
    debtors[dIdx].amount -= settleAmount;
    if (creditors[cIdx].amount === 0) cIdx++;
    if (debtors[dIdx].amount === 0) dIdx++;
  }

  let msg = `🎉【精算確定】${eventName}\n─────────────────\n`;
  if (settlements.length === 0) {
    msg += `全員の支払いがピッタリ相殺されています！送金の必要はありません✨\n`;
  } else {
    settlements.forEach((s, idx) => {
      msg += `【送金 ${idx + 1}】\n👤 ${s.fromName} ➔ ${s.toName}\n💴 ¥${s.amount.toLocaleString()}\n`;
      if (s.toPayPay) msg += `🔗 PayPay: ${s.toPayPay}\n`;
      msg += `\n`;
    });
    msg += `─────────────────\n※トーク画面左下の「＋」➔「送る・受け取る」からPayPayで送金できます！`;
  }
  return msg.trim();
}

// LINEグループ宛てプッシュ通知
function pushToLineGroup(groupId, text) {
  // テスト用ダミーグループの場合はLINE送信をスキップ
  if (!groupId || groupId.startsWith("G_TEST")) return;

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: groupId,
    messages: [{ type: "text", text: text }]
  };

  try {
    UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error("LINE Push送信エラー:", e.message);
  }
}

// ==========================================
// 3. メンバー自動参加（Postback & LINE API連携）
// ==========================================

/**
 * LINEグループ内のユーザープロフィールを取得する
 */
function getLineGroupMemberProfile(groupId, userId) {
  // グループIDがない場合（個別テスト時など）のフォールバック
  const url = (groupId && !groupId.startsWith("G_TEST"))
    ? `https://api.line.me/v2/bot/group/${groupId}/member/${userId}`
    : `https://api.line.me/v2/bot/profile/${userId}`;

  try {
    const res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() === 200) {
      return JSON.parse(res.getContentText());
    } else {
      console.warn("プロフィール取得失敗:", res.getContentText());
      return null;
    }
  } catch (e) {
    console.error("プロフィール取得エラー:", e.message);
    return null;
  }
}

/**
 * ボタンを押した人をスプレッドシートの Members に登録する（重複防止付き）
 */
function registerMemberFromLine(groupId, userId) {
  if (!userId) return null;

  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const rows = sheet.getDataRange().getValues();

  // 1. すでに登録済みかチェック（二重登録・連打防止）
  for (let i = 1; i < rows.length; i++) {
    const [gId, mId, name, lineId, paypay, status] = rows[i];
    if (String(gId) === targetGroupId && String(lineId) === String(userId)) {
      // お休み中になっていたら、せっかくボタンを押してくれたのでアクティブに戻してあげる神対応
      if (status !== "active") {
        sheet.getRange(i + 1, 6).setValue("active");
      }
      console.log("既存メンバーのため追加スキップ:", name);
      return { status: "already_registered", name: name };
    }
  }

  // 2. LINE APIから名前（displayName）を取得
  const profile = getLineGroupMemberProfile(groupId, userId);
  const displayName = (profile && profile.displayName) ? profile.displayName.trim().slice(0, 30) : "名無しさん";

  // 3. Membersシートの末尾に追加
  const lastRow = sheet.getLastRow();
  const newMemberId = "M" + String(lastRow).padStart(3, "0");

  sheet.appendRow([targetGroupId, newMemberId, displayName, userId, "", "active"]);
  console.log("🎉 新規メンバー登録完了:", displayName, userId);

  return { status: "created", name: displayName };
}

/**
 * 指定グループのアクティブなイベントを1件取得する
 */
function getActiveEventForGroup(groupId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();

  for (let i = 1; i < eventRows.length; i++) {
    const [id, gId, name, status] = eventRows[i];
    if (String(gId) === targetGroupId && String(status) === "active") {
      return { id: String(id), name: String(name) };
    }
  }
  return null;
}

/**
 * イベント状況（進行中 / 未作成）に応じてバナーとボタンが変化する動的Flex Message
 */
function createMainFlexMessage(groupId) {
  // 進行中のイベントがあるかチェック！
  const activeEvent = getActiveEventForGroup(groupId);

  // イベント有無で出し分ける要素
  let eventBadgeText = "";
  let eventBadgeColor = "";
  let eventBadgeBg = "";
  let mainBtnLabel = "";
  let targetUrl = "";

  if (activeEvent) {
    eventBadgeText = `🎉 進行中：${activeEvent.name}`;
    eventBadgeColor = "#4338ca"; // 濃いインディゴ
    eventBadgeBg = "#e0e7ff";    // 薄いインディゴ背景
    mainBtnLabel = "💸 支払いを記録する";
    targetUrl = `${WEB_APP_URL}?groupId=${groupId}&eventId=${activeEvent.id}`;
  } else {
    eventBadgeText = "💤 進行中のイベントはありません";
    eventBadgeColor = "#64748b"; // スレートグレー
    eventBadgeBg = "#f1f5f9";    // 薄いグレー背景
    mainBtnLabel = "🚀 新しいイベントを開始する";
    targetUrl = `${WEB_APP_URL}?groupId=${groupId}`;
  }

  return {
    type: "flex",
    altText: activeEvent ? `📦 【進行中】${activeEvent.name}` : "📦 みんなのわりかん箱",
    contents: {
      type: "bubble",
      size: "kilo",
      body: {
        type: "box",
        layout: "vertical",
        spacing: "md",
        contents: [
          // タイトル
          {
            type: "text",
            text: "📦 みんなのわりかん箱",
            weight: "bold",
            size: "md",
            color: "#1e293b"
          },
          // ★ 動的イベント状況バナー！
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: eventBadgeBg,
            cornerRadius: "md",
            paddingAll: "sm",
            contents: [
              {
                type: "text",
                text: eventBadgeText,
                weight: "bold",
                size: "xs",
                color: eventBadgeColor,
                wrap: true
              }
            ]
          },
          // 区切り線
          {
            type: "separator"
          },
          // つかいかたステップ
          {
            type: "box",
            layout: "vertical",
            spacing: "xs",
            contents: [
              {
                type: "text",
                text: "【つかいかた】",
                weight: "bold",
                size: "xs",
                color: "#334155"
              },
              {
                type: "text",
                text: "1️⃣ まず下の「割り勘に参加する」をタップ！",
                size: "xs",
                color: "#64748b",
                wrap: true
              },
              {
                type: "text",
                text: "2️⃣ 立替が出たら「支払いを記録する」で入力💸",
                size: "xs",
                color: "#64748b",
                wrap: true
              },
              {
                type: "text",
                text: "3️⃣ 精算は「@みんなのわりかん箱 精算」と送るだけ🎉",
                size: "xs",
                color: "#64748b",
                wrap: true
              }
            ]
          },
          // 同意注記
          {
            type: "text",
            text: "※「割り勘に参加する」を押すと、LINEの表示名がメンバー一覧に登録されます。",
            size: "xxs",
            color: "#94a3b8",
            wrap: true
          }
        ]
      },
      footer: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          // ① 参加ボタン
          {
            type: "button",
            style: "primary",
            color: "#4f46e5",
            height: "sm",
            action: {
              type: "postback",
              label: "🙋‍♂️ 割り勘に参加する",
              data: `action=join&groupId=${groupId}`,
              displayText: "🙋‍♂️ わりかん箱に参加したよ！"
            }
          },
          // ② 状況によって変わる動的ボタン！
          {
            type: "button",
            style: "secondary",
            height: "sm",
            action: {
              type: "uri",
              label: mainBtnLabel,
              uri: targetUrl
            }
          }
        ]
      }
    }
  };
}

/**
 * Flex Message対応のReply関数
 */
function replyFlexToLine(replyToken, flexMessage) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    replyToken: replyToken,
    messages: [flexMessage]
  };
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload)
  });
}



// ==========================================
// 1. Web画面 表示・ルーティング制御
// ==========================================

/**
 * HTTP GETリクエスト ハンドラー
 * URLパラメータ（page, groupId, eventId）に応じて適切なHTMLテンプレートを評価・返却する
 * 
 * @param {Object} e - GASのイベントオブジェクト (e.parameter)
 * @returns {HtmlOutput} レンダリングされたHTMLページ
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "";

  // 1. 規約・ポリシーページへのルーティング (?page=terms)
  if (page === "terms") {
    return HtmlService.createHtmlOutputFromFile("terms")
      .setTitle("利用規約・プライバシーポリシー | みんなのわりかん箱")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 2. メイン割り勘Webアプリのレンダリング
  // テンプレート変数にサーバー側パラメータを焼き込むことで、クライアント側でのURLパース迷子を構造的に防止
  const template = HtmlService.createTemplateFromFile("index");
  template.serverGroupId = (e && e.parameter && e.parameter.groupId) || "G_TEST_GROUP";
  template.serverEventId = (e && e.parameter && e.parameter.eventId) || "";
  template.serverWebAppUrl = WEB_APP_URL;

  return template.evaluate()
    .setTitle("みんなのわりかん箱")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * 画面初期表示用のデータを一括取得する
 * 通信回数を1回に抑えるため、該当グループのイベント一覧・メンバー一覧・支払い履歴をまとめて返却する
 * 
 * @param {string} groupId - 対象のLINEグループID
 * @param {string} [requestedEventId] - URL等で直接指定されたイベントID（任意）
 * @returns {Object} { groupId, currentEventId, events, members, payments }
 */
function getInitialData(groupId, requestedEventId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ----------------------------------------------------
  // 1. イベント一覧取得 (Eventsシート)
  // ----------------------------------------------------
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < eventRows.length; i++) {
    const [id, gId, name, status, createdAt] = eventRows[i];
    // 他グループのデータを混入させない（マルチテナント防壁）
    if (id && String(gId) === targetGroupId) {
      events.push({ id: String(id), name: String(name), status: String(status) });
    }
  }

  // 現在アクティブなイベントの特定（指定ID > 最新active > 最後のイベント）
  let currentEvent = null;
  if (requestedEventId) {
    currentEvent = events.find(e => e.id === String(requestedEventId));
  }
  if (!currentEvent) {
    currentEvent = events.filter(e => e.status === "active").pop();
  }
  if (!currentEvent && events.length > 0) {
    currentEvent = events[events.length - 1];
  }

  // ----------------------------------------------------
  // 2. メンバー一覧取得 (Membersシート)
  // ----------------------------------------------------
  const memberSheet = ss.getSheetByName("Members");
  const memberRows = memberSheet.getDataRange().getValues();
  const members = [];
  const memberNameMap = {}; // IDから名前に高速変換するための辞書

  for (let i = 1; i < memberRows.length; i++) {
    const [gId, id, name, lineId, paypay, status] = memberRows[i];
    if (id && String(gId) === targetGroupId) {
      const currentStatus = String(status || "active");
      members.push({
        id: String(id),
        name: String(name),
        paypay: String(paypay || ""),
        status: currentStatus
      });
      memberNameMap[String(id)] = String(name);
    }
  }

  // ----------------------------------------------------
  // 3. 支払い履歴取得 (Paymentsシート)
  // ----------------------------------------------------
  const paymentSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentSheet.getDataRange().getValues();
  const payments = [];
  const targetEventIds = events.map(e => e.id);

  for (let i = 1; i < paymentRows.length; i++) {
    const [pId, eId, payerId, amount, desc, targets, createdAt, status] = paymentRows[i];
    // このグループに属するイベントの支払いのみを抽出
    if (pId && targetEventIds.includes(String(eId))) {
      const targetIds = String(targets || "").split(",").map(s => s.trim()).filter(Boolean);
      const targetNames = targetIds.map(tId => memberNameMap[tId] || tId);
      const paymentStatus = String(status || "active");

      // 新しい支払いが上に来るように先頭へ追加 (unshift)
      payments.unshift({
        id: String(pId),
        eventId: String(eId),
        payerName: memberNameMap[payerId] || String(payerId),
        amount: Number(amount),
        description: String(desc),
        targetNames: targetNames,
        isAllTargets: targetNames.length === members.filter(m => m.status === "active").length,
        date: createdAt ? String(createdAt) : "",
        status: paymentStatus // 論理削除されたデータも "deleted" として渡し、フロントで打消線表示
      });
    }
  }

  return {
    groupId: targetGroupId,
    currentEventId: currentEvent ? currentEvent.id : null,
    events: events,
    members: members,
    payments: payments
  };
}


// ==========================================
// 2. メンバー・イベント管理 API
// ==========================================

/**
 * メンバーのステータス（active / inactive）を切り替える
 * ※ 物理削除せず論理状態を更新することで、過去の精算や履歴の名前解決クラッシュを完全に防ぐ
 * 
 * @param {string} groupId - グループID
 * @param {string} memberId - 対象メンバーID
 * @param {string} newStatus - "active" または "inactive"
 */
function toggleMemberStatus(groupId, memberId, newStatus) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === targetGroupId && String(rows[i][1]) === String(memberId)) {
      // F列 (6列目) の Status カラムを更新
      sheet.getRange(i + 1, 6).setValue(newStatus);
      return { status: "success", memberId: memberId, newStatus: newStatus };
    }
  }
  throw new Error("対象のメンバーが見つかりませんでした");
}

/**
 * Web画面から新規メンバーを追加する（初期Status: active）
 * オプトイン設計に基づき、LINE ID連携なしの「表示名（文字列）」だけで即参加可能
 * 
 * @param {string} groupId - グループID
 * @param {string} name - メンバー名（最大30文字）
 */
function addMemberFromWeb(groupId, name) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const cleanName = String(name).trim().slice(0, 30);
  if (!cleanName) throw new Error("名前が空です");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const lastRow = sheet.getLastRow();
  const newMemberId = "M" + String(lastRow).padStart(3, "0");

  sheet.appendRow([targetGroupId, newMemberId, cleanName, "", "", "active"]);
  return { id: newMemberId, name: cleanName };
}

/**
 * Web画面から新規イベントを作成する
 * 
 * @param {string} groupId - グループID
 * @param {string} name - イベント名（例: 北海道旅行、BBQ）
 * @param {boolean} shouldClosePrevious - 進行中だった既存イベントをクローズするかどうか
 */
function createEventFromWeb(groupId, name, shouldClosePrevious) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const cleanName = String(name).trim().slice(0, 30);
  if (!cleanName) throw new Error("イベント名が空です");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Events");
  const rows = sheet.getDataRange().getValues();

  // 進行中イベントの終了処理（単一有効イベントの保証）
  if (shouldClosePrevious) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === targetGroupId && rows[i][3] === "active") {
        sheet.getRange(i + 1, 4).setValue("closed");
      }
    }
  }

  const lastRow = sheet.getLastRow();
  const newEventId = "E" + String(lastRow).padStart(3, "0");
  const dateStr = new Date().toLocaleString("ja-JP");

  sheet.appendRow([newEventId, targetGroupId, cleanName, "active", dateStr]);
  return { id: newEventId, name: cleanName };
}


// ==========================================
// 3. 支払い記録・論理削除 API
// ==========================================

/**
 * Web画面から支払いを新規登録する
 * 
 * @param {Object} data - { eventId, payerId, amount, description, targetMemberIds }
 */
function recordPaymentFromWeb(data) {
  const amount = Number(data.amount);
  // 入力桁数の防弾バリデーション（1円〜9,999万円）
  if (isNaN(amount) || amount <= 0 || amount > 99999999) {
    throw new Error("金額が不正です（1円〜9,999万円まで）");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 終了済み（closed）イベントへの改ざん・追加入力を物理ブロック
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][0]) === String(data.eventId) && eventRows[i][3] === "closed") {
      throw new Error("このイベントは終了しているため記録できません");
    }
  }

  const sheet = ss.getSheetByName("Payments");
  const lastRow = sheet.getLastRow();
  const paymentId = "P" + String(lastRow).padStart(3, "0");
  const cleanDesc = String(data.description).trim().slice(0, 50);

  const newRow = [
    paymentId,
    String(data.eventId),
    String(data.payerId),
    amount,
    cleanDesc,
    String(data.targetMemberIds),
    new Date().toLocaleString("ja-JP"),
    "active" // H列 (8列目): ステータス（active / deleted）
  ];

  sheet.appendRow(newRow);
  return { status: "success" };
}

/**
 * 支払いの論理削除（Soft Delete）を行う
 * 行を物理削除せず status を "deleted" に更新することで、
 * 「誰かが勝手に履歴を消した」等の人間関係トラブルを防ぎ、監査ログとしての透明性を保つ
 * 
 * @param {string} groupId - グループID
 * @param {string} paymentId - 削除対象の支払いID
 */
function cancelPaymentFromWeb(groupId, paymentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentSheet.getDataRange().getValues();

  for (let i = 1; i < paymentRows.length; i++) {
    if (String(paymentRows[i][0]) === String(paymentId)) {
      const targetEventId = String(paymentRows[i][1]);

      // すでに精算確定・終了したイベントのデータは取り消し不可
      const eventSheet = ss.getSheetByName("Events");
      const eventRows = eventSheet.getDataRange().getValues();
      for (let j = 1; j < eventRows.length; j++) {
        if (String(eventRows[j][0]) === targetEventId && eventRows[j][3] === "closed") {
          throw new Error("終了済みのイベントの支払いは取り消せません");
        }
      }

      // H列（8列目）のステータスを "deleted" に更新
      paymentSheet.getRange(i + 1, 8).setValue("deleted");
      return { status: "success" };
    }
  }
  throw new Error("対象の支払いデータが見つかりませんでした");
}


// ==========================================
// 4. LINE Messaging API (Webhook ＆ 返信/通知)
// ==========================================

/**
 * LINE Messaging API Webhook 受信ハンドラー
 * グループ参加時の案内、Bot宛てメンション、精算コマンドに応答する
 * 
 * @param {Object} e - LINE PlatformからのHTTP POSTデータ (e.postData.contents)
 */
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const events = json.events;
    if (!events || events.length === 0) {
      return ContentService.createTextOutput("ok");
    }

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      // 送信元の識別子（グループID / ルームID / ユーザーID）を動的取得
      const currentGroupId = (event.source && (event.source.groupId || event.source.roomId || event.source.userId)) || "G_TEST_GROUP";

      // 1. グループ招待時の自動ウェルカム案内 (join イベント)
      if (event.type === "join") {
        const welcomeMsg = 
          `招待ありがとう！「みんなのわりかん箱」です📦✨\n` +
          `旅行や飲み会の割り勘を、最小の送金回数でスッキリ計算します！\n\n` +
          `【つかいかた】\n` +
          `トークで「@みんなのわりかん箱」と呼んでみてね👇\n\n` +
          `💸 支払いの記録・割り勘スタート\n` +
          `${WEB_APP_URL}?groupId=${currentGroupId}\n\n` +
          `※ 精算したい時は「@みんなのわりかん箱 精算」と送るだけ！`;

        replyToLine(event.replyToken, welcomeMsg);
        continue;
      }

      // 2. メッセージ受信イベントの処理
      if (event.type === "message" && event.message.type === "text") {
        const rawText = event.message.text.trim();
        const replyToken = event.replyToken;
        const text = rawText.replace(/＠/g, "@").toLowerCase();

        // メンション判定（普段のグループ会話を邪魔しない防音壁設計）
        const isMentioned = 
          text.includes("@みんなのわりかん箱") || 
          text.includes("@わりかん箱") || 
          text.startsWith("@わりかん") || 
          text.startsWith("@warihako");

        if (!isMentioned) {
          // 関係のない雑談は100%無視（既読スルー）
          continue;
        }

        // 精算コマンドの判定
        const isSettlement = 
          text.includes("精算") || 
          text.includes("清算") || 
          text.includes("せいさん");

        if (isSettlement) {
          // 精算を確定し、結果メッセージを返信
          const settlementMessage = executeSettlement(currentGroupId);
          replyToLine(replyToken, settlementMessage);
        } else {
          // メンション単体時：入力URLと案内カードを返信
          const targetUrl = `${WEB_APP_URL}?groupId=${currentGroupId}`;
          const msg = 
            `📦 みんなのわりかん箱\n\n` +
            `💸 支払いの記録・履歴の確認はこちら👇\n` +
            `${targetUrl}\n\n` +
            `💡 トークで結果を見るなら\n` +
            `「@みんなのわりかん箱 精算」と送ってね！`;

          replyToLine(replyToken, msg);
        }
      }
    }
  } catch (error) {
    console.error("❌ Webhookエラー:", error.stack || error.message);
  }

  return ContentService.createTextOutput("ok");
}

/**
 * LINEトーク上での精算を実行し、イベントを closed に更新する
 * 
 * @param {string} groupId - 対象のグループID
 * @returns {string} LINEトーク宛ての精算結果テキスト
 */
function executeSettlement(groupId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();
  let currentEvent = null;
  let eventRowIndex = -1;

  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][1]) === targetGroupId && eventRows[i][3] === "active") {
      currentEvent = { id: String(eventRows[i][0]), name: String(eventRows[i][2]) };
      eventRowIndex = i + 1;
      break;
    }
  }
  if (!currentEvent) {
    return "このグループには現在進行中のイベントがありません💦\n「@みんなのわりかん箱」から新しいイベントを作成してください！";
  }

  // ステータスを即座に closed に倒してDBロック
  eventSheet.getRange(eventRowIndex, 4).setValue("closed");

  // 精算メッセージ本文の生成
  const settlementMsg = generateSettlementMessage(targetGroupId, currentEvent.id, currentEvent.name);

  // 全員が納得して過去ログを見直せるよう、確定明細URL（?eventId=...）を添えて返却
  const detailUrl = `${WEB_APP_URL}?groupId=${targetGroupId}&eventId=${currentEvent.id}`;
  const finalMsg = `${settlementMsg}\n\n📜【支払い明細・履歴の確認はこちら】👇\n${detailUrl}`;

  return finalMsg.trim();
}

/**
 * Web画面上の精算ボタンから呼び出され、イベントを終了してLINEグループへプッシュ通知する
 * 
 * @param {string} groupId - グループID
 * @param {string} eventId - イベントID
 */
function closeAndSettleEventFromWeb(groupId, eventId) {
  const targetGroupId = groupId || "G_TEST_GROUP";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName("Events");
  const eventRows = eventSheet.getDataRange().getValues();

  let eventName = "";
  let targetRow = -1;

  for (let i = 1; i < eventRows.length; i++) {
    if (String(eventRows[i][0]) === String(eventId) && String(eventRows[i][1]) === targetGroupId) {
      targetRow = i + 1;
      eventName = String(eventRows[i][2]);
      break;
    }
  }

  if (targetRow === -1) throw new Error("対象のイベントが見つかりませんでした");

  // 1. ステータスを closed に更新
  eventSheet.getRange(targetRow, 4).setValue("closed");

  // 2. 精算メッセージを作成
  const settlementMsg = generateSettlementMessage(targetGroupId, eventId, eventName);

  // 3. 確定閲覧URLを添えてLINEグループへ通知
  const detailUrl = `${WEB_APP_URL}?groupId=${targetGroupId}&eventId=${eventId}`;
  const finalMessage = `${settlementMsg}\n\n📜【支払い明細・履歴の確認はこちら】👇\n${detailUrl}`;

  pushToLineGroup(targetGroupId, finalMessage);

  return { status: "success", eventId: eventId };
}


// ==========================================
// 5. 精算相殺アルゴリズム・送金メッセージ生成ヘルパー
// ==========================================

/**
 * 指定イベントの収支を相殺し、最小送金回数の精算メッセージ文字列を生成する
 * 【数学的アプローチ：債権債務の相殺アルゴリズム】
 * 各自の「立替総額 - 負担総額」を算出し、最大債権者と最大債務者を順次突き合わせることで送金回数を最小化
 * 
 * @param {string} groupId - グループID
 * @param {string} eventId - イベントID
 * @param {string} eventName - イベント名
 * @returns {string} フォーマットされた精算結果文面
 */
function generateSettlementMessage(groupId, eventId, eventName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName("Members");
  const memberRows = membersSheet.getDataRange().getValues();
  const memberMap = {};

  for (let i = 1; i < memberRows.length; i++) {
    const [gId, mId, name, lineId, paypay] = memberRows[i];
    if (mId && String(gId) === groupId) {
      memberMap[String(mId)] = { name: String(name), paypay: String(paypay || "") };
    }
  }

  const balances = {};
  Object.keys(memberMap).forEach(id => balances[id] = 0);

  const paymentsSheet = ss.getSheetByName("Payments");
  const paymentRows = paymentsSheet.getDataRange().getValues();
  let hasPayment = false;

  for (let i = 1; i < paymentRows.length; i++) {
    const [pId, eId, payerId, amount, desc, targetsStr, createdAt, status] = paymentRows[i];
    const pStatus = String(status || "active");
    // 論理削除されたデータ（deleted）は計算から完全に除外
    if (String(eId) === String(eventId) && amount && pStatus !== "deleted") {
      hasPayment = true;
      const targets = String(targetsStr).split(",").map(s => s.trim()).filter(Boolean);
      const splitAmount = Math.floor(Number(amount) / targets.length);

      balances[payerId] += Number(amount);
      targets.forEach(tId => {
        if (balances[tId] !== undefined) balances[tId] -= splitAmount;
      });
    }
  }

  if (!hasPayment) return `🎉【精算完了】${eventName}\n支払いの記録はありませんでした！`;

  // 債権者（もらいすぎ）と債務者（払い不足）のリストアップ
  let creditors = [], debtors = [];
  for (const id in balances) {
    if (balances[id] > 0) creditors.push({ id: id, amount: balances[id] });
    else if (balances[id] < 0) debtors.push({ id: id, amount: -balances[id] });
  }

  // 金額の大きい順に降順ソート（相殺の効率化）
  creditors.sort((a, b) => b.amount - a.amount);
  debtors.sort((a, b) => b.amount - a.amount);

  const settlements = [];
  let cIdx = 0, dIdx = 0;

  // グリーディ（貪欲）マッチングによる送金回数最小化ループ
  while (cIdx < creditors.length && dIdx < debtors.length) {
    const settleAmount = Math.min(creditors[cIdx].amount, debtors[dIdx].amount);
    settlements.push({
      fromName: memberMap[debtors[dIdx].id] ? memberMap[debtors[dIdx].id].name : "不明",
      toName: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].name : "不明",
      amount: settleAmount,
      toPayPay: memberMap[creditors[cIdx].id] ? memberMap[creditors[cIdx].id].paypay : ""
    });

    creditors[cIdx].amount -= settleAmount;
    debtors[dIdx].amount -= settleAmount;

    if (creditors[cIdx].amount === 0) cIdx++;
    if (debtors[dIdx].amount === 0) dIdx++;
  }

  // 文面フォーマット
  let msg = `🎉【精算確定】${eventName}\n─────────────────\n`;
  if (settlements.length === 0) {
    msg += `全員の支払いがピッタリ相殺されています！送金の必要はありません✨\n`;
  } else {
    settlements.forEach((s, idx) => {
      msg += `【送金 ${idx + 1}】\n👤 ${s.fromName} ➔ ${s.toName}\n💴 ¥${s.amount.toLocaleString()}\n`;
      if (s.toPayPay) msg += `🔗 PayPay: ${s.toPayPay}\n`;
      msg += `\n`;
    });
    msg += `─────────────────\n※トーク画面左下の「＋」➔「送る・受け取る」からPayPayで送金できます！`;
  }
  return msg.trim();
}


// ==========================================
// 6. LINE API 送信ラッパー関数
// ==========================================

/**
 * LINE Messaging API: Reply API による応答送信（月間無料枠を永久に消費しない）
 * 
 * @param {string} replyToken - Webhookから受領したリプライトークン
 * @param {string} text - 送信するテキスト
 */
function replyToLine(replyToken, text) {
  const url = "https://api.line.me/v2/bot/message/reply";
  const payload = {
    replyToken: replyToken,
    messages: [{ type: "text", text: text }]
  };
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload)
  });
}

/**
 * LINE Messaging API: Push API によるグループ宛て自発通知
 * 
 * @param {string} groupId - 送信先のLINEグループID
 * @param {string} text - 送信するテキスト
 */
function pushToLineGroup(groupId, text) {
  // テスト用ダミーグループの場合は外部通信を安全にスキップ
  if (!groupId || groupId.startsWith("G_TEST")) return;

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: groupId,
    messages: [{ type: "text", text: text }]
  };

  try {
    UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error("LINE Push送信エラー:", e.message);
  }
}
