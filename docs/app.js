const STORAGE_KEY = "bp_records";
const TIME_SLOTS = ["朝", "日中", "夜"];
const TIME_SLOT_ORDER = { "朝": 0, "日中": 1, "夜": 2 };

// JSH2019家庭血圧の目安を簡略化した6段階（参考情報であり診断ではありません）
const LEVELS = [
  ["至適血圧", "#27ae60"],
  ["正常血圧", "#2ecc71"],
  ["正常高値血圧", "#f4d03f"],
  ["I度高血圧", "#f39c12"],
  ["II度高血圧", "#e67e22"],
  ["III度高血圧", "#e74c3c"],
];

const LEVEL_MESSAGES = [
  "とても良い数値です。この調子で今の生活習慣を続けましょう。",
  "正常範囲です。今の生活習慣を維持しましょう。",
  "やや高めです。減塩や運動を意識すると安心です。",
  "高血圧の可能性があります。生活習慣の見直しをおすすめします。",
  "高血圧の可能性が高い数値です。継続するようであれば医療機関への相談をご検討ください。",
  "非常に高い数値です。早めに医療機関の受診をおすすめします。",
];

const GENERAL_TIPS = [
  "麺類の汁は残す（塩分の多くはスープに含まれています）",
  "醤油やソースは「かける」より「つける」で使用量を減らす",
  "ハム・練り物・インスタント食品などの加工食品は控えめに",
  "だしを効かせると、塩を減らしても美味しく食べられます",
  "生姜・ねぎ・大葉・柑橘類など香味野菜や酸味で味にアクセントを",
  "バナナ・ほうれん草・アボカドなどカリウムを含む野菜・果物を意識的に（腎機能に不安がある方は主治医に相談）",
  "漬物・佃煮・梅干しは量を控えめに",
  "味噌汁は具だくさんにして汁の量を減らす",
  "外食・コンビニ食は栄養成分表示の食塩相当量を確認する習慣を",
  "お酒はほどほどに（目安は1日あたり日本酒1合・ビール中瓶1本程度まで）",
  "減塩醤油・減塩味噌など減塩調味料を活用する",
  "スナック菓子や加工肉を減らし、素材そのものの味を楽しむ",
  "毎日同じ時間・同じ条件で血圧を測る習慣をつける（起床後1時間以内・朝食前や排尿後がおすすめ）",
  "ウォーキングなど軽い有酸素運動を1日30分程度心がける",
  "適正体重を維持する（肥満は血圧を上げやすくなります）",
  "禁煙・受動喫煙を避ける",
  "十分な睡眠とストレスケアを心がける",
  "湯船にゆっくり浸かるなど、リラックスできる時間を作る",
];

function classifyBp(sbp, dbp) {
  let sl;
  if (sbp < 115) sl = 0;
  else if (sbp < 125) sl = 1;
  else if (sbp < 135) sl = 2;
  else if (sbp < 145) sl = 3;
  else if (sbp < 160) sl = 4;
  else sl = 5;

  let dl;
  if (dbp < 75) dl = 0;
  else if (dbp < 85) dl = 2;
  else if (dbp < 90) dl = 3;
  else if (dbp < 100) dl = 4;
  else dl = 5;

  const level = Math.max(sl, dl);
  const [label, color] = LEVELS[level];
  return { level, label, color };
}

function sampleTips(n) {
  const pool = [...GENERAL_TIPS];
  const picked = [];
  for (let i = 0; i < n && pool.length; i++) {
    const idx = Math.floor(Math.random() * pool.length);
    picked.push(pool.splice(idx, 1)[0]);
  }
  return picked;
}

function buildAdvice(level) {
  const lines = [LEVEL_MESSAGES[level], "", "【今日の食事・生活アドバイス】"];
  for (const tip of sampleTips(3)) lines.push(`・${tip}`);
  return lines.join("\n");
}

function loadRecords() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}

function saveRecords(records) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

function sortKey(r) {
  return `${r.date}_${TIME_SLOT_ORDER[r.time_slot] ?? 1}`;
}

function sortedRecords(records) {
  return [...records].sort((a, b) => (sortKey(a) < sortKey(b) ? -1 : sortKey(a) > sortKey(b) ? 1 : 0));
}

function computeWeeklyStats(records) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const recent = records.filter((r) => {
    const d = new Date(r.date);
    const diffDays = (today - d) / 86400000;
    return diffDays >= 0 && diffDays < 7;
  });
  if (!recent.length) return null;
  const avgSbp = recent.reduce((s, r) => s + r.sbp, 0) / recent.length;
  const avgDbp = recent.reduce((s, r) => s + r.dbp, 0) / recent.length;
  return { avgSbp, avgDbp, n: recent.length };
}

let records = loadRecords();

const dateInput = document.getElementById("f-date");
dateInput.value = new Date().toISOString().slice(0, 10);

document.getElementById("record-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const date = dateInput.value;
  const slot = document.getElementById("f-slot").value;
  const sbp = parseInt(document.getElementById("f-sbp").value, 10);
  const dbp = parseInt(document.getElementById("f-dbp").value, 10);
  const pulseRaw = document.getElementById("f-pulse").value;
  const pulse = pulseRaw ? parseInt(pulseRaw, 10) : null;
  const memo = document.getElementById("f-memo").value.trim();

  if (!date || Number.isNaN(sbp) || Number.isNaN(dbp)) {
    alert("日付・収縮期・拡張期血圧を入力してください");
    return;
  }
  if (sbp < 50 || sbp > 300 || dbp < 30 || dbp > 200) {
    alert("血圧の値が現実的な範囲を超えています");
    return;
  }

  records.push({
    date, time_slot: slot, sbp, dbp, pulse, memo,
    created_at: new Date().toISOString(),
  });
  saveRecords(records);

  document.getElementById("f-sbp").value = "";
  document.getElementById("f-dbp").value = "";
  document.getElementById("f-pulse").value = "";
  document.getElementById("f-memo").value = "";

  refreshAll();
});

document.getElementById("btn-refresh-advice").addEventListener("click", () => {
  refreshAdvice();
});

document.getElementById("btn-export").addEventListener("click", () => {
  downloadFile(
    JSON.stringify(records, null, 2),
    "血圧記録.json",
    "application/json"
  );
});

document.getElementById("btn-export-csv").addEventListener("click", () => {
  const header = ["日付", "時間帯", "収縮期", "拡張期", "脈拍", "判定", "メモ"];
  const rows = sortedRecords(records).map((r) => {
    const { label } = classifyBp(r.sbp, r.dbp);
    return [r.date, r.time_slot, r.sbp, r.dbp, r.pulse ?? "", label, r.memo ?? ""];
  });
  const csv = [header, ...rows]
    .map((row) => row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
    .join("\n");
  downloadFile("﻿" + csv, "血圧記録.csv", "text/csv");
});

document.getElementById("btn-import").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result);
      if (!Array.isArray(imported)) throw new Error("invalid");
      records = imported;
      saveRecords(records);
      refreshAll();
      alert("インポートしました");
    } catch {
      alert("ファイルを読み込めませんでした");
    }
  };
  reader.readAsText(file);
  e.target.value = "";
});

function downloadFile(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function refreshAll() {
  refreshHistory();
  refreshStats();
  refreshAdvice();
  refreshChart();
}

function refreshHistory() {
  const container = document.getElementById("history-list");
  container.innerHTML = "";
  const sorted = sortedRecords(records);
  for (let i = sorted.length - 1; i >= 0; i--) {
    const r = sorted[i];
    const { label, color } = classifyBp(r.sbp, r.dbp);
    const item = document.createElement("div");
    item.className = "history-item";
    item.innerHTML = `
      <span class="label" style="background:${color}">${label}</span>
      <span class="meta">${r.date}（${r.time_slot}） ${r.sbp}/${r.dbp}${r.pulse ? ` 脈${r.pulse}` : ""}${r.memo ? ` - ${r.memo}` : ""}</span>
      <button class="delete-btn" data-index="${i}">✕</button>
    `;
    container.appendChild(item);
  }
  container.querySelectorAll(".delete-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.dataset.index, 10);
      const target = sorted[idx];
      records = records.filter((r) => r !== target);
      saveRecords(records);
      refreshAll();
    });
  });
}

function refreshStats() {
  const stats = computeWeeklyStats(records);
  const el = document.getElementById("weekly-stats");
  el.textContent = stats
    ? `直近7日間の平均: ${stats.avgSbp.toFixed(0)} / ${stats.avgDbp.toFixed(0)} mmHg（${stats.n}件）`
    : "直近7日間の記録はまだありません";
}

function refreshAdvice() {
  const judgementEl = document.getElementById("judgement");
  const adviceEl = document.getElementById("advice");
  if (!records.length) {
    judgementEl.textContent = "まだ記録がありません";
    judgementEl.style.background = "";
    adviceEl.textContent =
      "まずは今日の血圧を記録してみましょう。\n\n【今日の食事・生活アドバイス】\n・" +
      sampleTips(3).join("\n・");
    return;
  }
  const latest = sortedRecords(records).at(-1);
  const { level, label, color } = classifyBp(latest.sbp, latest.dbp);
  judgementEl.textContent = `最新の記録: ${latest.sbp}/${latest.dbp} mmHg → ${label}`;
  judgementEl.style.background = color;
  judgementEl.style.color = "#fff";
  adviceEl.textContent = buildAdvice(level);
}

function refreshChart() {
  const canvas = document.getElementById("chart");
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  const recent = sortedRecords(records).slice(-14);
  if (recent.length < 2) {
    ctx.fillStyle = "#888";
    ctx.font = "13px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("記録が2件以上になるとグラフを表示します", w / 2, h / 2);
    return;
  }

  const padL = 34, padR = 10, padT = 24, padB = 26;
  const plotW = w - padL - padR, plotH = h - padT - padB;
  const vMin = 50, vMax = 200;

  const yOf = (v) => padT + plotH * (1 - (Math.max(vMin, Math.min(vMax, v)) - vMin) / (vMax - vMin));
  const xOf = (i) => (recent.length === 1 ? padL : padL + (plotW * i) / (recent.length - 1));

  ctx.strokeStyle = "#999";
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, h - padB);
  ctx.lineTo(w - padR, h - padB);
  ctx.stroke();

  ctx.font = "9px sans-serif";
  ctx.fillStyle = "#999";
  ctx.textAlign = "right";
  [100, 135, 160].forEach((v) => {
    const y = yOf(v);
    ctx.strokeStyle = "#ddd";
    ctx.setLineDash([2, 2]);
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(w - padR, y);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillText(String(v), padL - 4, y + 3);
  });

  const drawLine = (values, color) => {
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = 2;
    ctx.beginPath();
    values.forEach((v, i) => {
      const x = xOf(i), y = yOf(v);
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
    values.forEach((v, i) => {
      const x = xOf(i), y = yOf(v);
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
    });
  };
  drawLine(recent.map((r) => r.sbp), "#e74c3c");
  drawLine(recent.map((r) => r.dbp), "#3498db");

  ctx.textAlign = "center";
  const step = Math.max(1, Math.floor(recent.length / 7));
  recent.forEach((r, i) => {
    if (i % step === 0) ctx.fillText(r.date.slice(5), xOf(i), h - padB + 14);
  });

  ctx.textAlign = "left";
  ctx.fillStyle = "#e74c3c";
  ctx.fillText("● 収縮期", padL + 4, 12);
  ctx.fillStyle = "#3498db";
  ctx.fillText("● 拡張期", padL + 70, 12);
}

refreshAll();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  });
}
