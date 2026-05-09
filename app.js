/* ===== Storage Helpers ===== */
const storage = {
  get: (key, fallback = []) => {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; }
    catch { return fallback; }
  },
  set: (key, val) => localStorage.setItem(key, JSON.stringify(val)),
};

/* ===== State ===== */
let goals = storage.get('wg_goals', []);
let records = storage.get('wg_records', []);
let reminder = storage.get('wg_reminder', null);

let progressChart = null;
let reminderTimerId = null;

/* ===== Utilities ===== */
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
const today = () => new Date().toISOString().slice(0, 10);

function formatDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${y}/${m}/${d}`;
}

function daysLeft(deadline) {
  if (!deadline) return null;
  const diff = Math.ceil((new Date(deadline) - new Date(today())) / 86400000);
  return diff;
}

/* ===== Save ===== */
function saveGoals() { storage.set('wg_goals', goals); }
function saveRecords() { storage.set('wg_records', records); }
function saveReminder() { storage.set('wg_reminder', reminder); }

/* ===== Tab Navigation ===== */
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
    if (btn.dataset.tab === 'progress') renderChart();
  });
});

/* ===== Modal Helpers ===== */
const overlay = document.getElementById('modal-overlay');

function openModal(id) {
  document.getElementById(id).classList.add('active');
  overlay.classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeModal(id) {
  document.getElementById(id).classList.remove('active');
  overlay.classList.remove('active');
  document.body.style.overflow = '';
}

document.querySelectorAll('.modal-close, [data-modal]').forEach(el => {
  el.addEventListener('click', () => closeModal(el.dataset.modal || el.closest('.modal').id));
});
overlay.addEventListener('click', () => {
  document.querySelectorAll('.modal.active').forEach(m => closeModal(m.id));
});

/* ===========================
   GOALS
   =========================== */
function renderGoals() {
  const list = document.getElementById('goals-list');
  const empty = document.getElementById('goals-empty');
  list.innerHTML = '';

  if (goals.length === 0) {
    empty.classList.add('visible');
    return;
  }
  empty.classList.remove('visible');

  goals.forEach(goal => {
    const pct = calcGoalProgress(goal);
    const dl = daysLeft(goal.deadline);
    const overdue = dl !== null && dl < 0;

    const card = document.createElement('div');
    card.className = 'card goal-card';
    card.innerHTML = `
      <span class="goal-category-badge">${goal.category || 'その他'}</span>
      <h3>${escHtml(goal.name)}</h3>
      <div class="goal-targets">
        ${goal.weight ? `<span>🎯 ${goal.weight}kg</span>` : ''}
        ${goal.reps   ? `<span>🔁 ${goal.reps}回</span>` : ''}
        ${goal.sets   ? `<span>📋 ${goal.sets}セット</span>` : ''}
      </div>
      <div class="progress-bar-wrap">
        <div class="progress-bar-fill" style="width:${pct}%"></div>
      </div>
      <div class="progress-pct">達成率 ${pct}%</div>
      ${goal.deadline ? `
        <div class="deadline ${overdue ? 'overdue' : ''}">
          ${overdue ? '⚠️ 期限切れ' : `📅 期限: ${formatDate(goal.deadline)} (残${dl}日)`}
        </div>` : ''}
      ${goal.memo ? `<div style="font-size:0.8rem;color:var(--text-muted);margin-bottom:0.75rem;">${escHtml(goal.memo)}</div>` : ''}
      <div class="card-actions">
        <button class="btn-secondary" data-edit="${goal.id}">編集</button>
        <button class="btn-danger" data-delete="${goal.id}">削除</button>
      </div>
    `;
    list.appendChild(card);
  });

  list.querySelectorAll('[data-edit]').forEach(btn => {
    btn.addEventListener('click', () => openGoalEdit(btn.dataset.edit));
  });
  list.querySelectorAll('[data-delete]').forEach(btn => {
    btn.addEventListener('click', () => deleteGoal(btn.dataset.delete));
  });
}

function calcGoalProgress(goal) {
  const relevant = records.filter(r => r.exercise === goal.name);
  if (relevant.length === 0) return 0;
  const latest = relevant.sort((a, b) => b.date.localeCompare(a.date))[0];

  if (goal.weight && latest.weight) {
    return Math.min(100, Math.round((latest.weight / goal.weight) * 100));
  }
  if (goal.reps && latest.reps) {
    return Math.min(100, Math.round((latest.reps / goal.reps) * 100));
  }
  return 0;
}

document.getElementById('open-goal-modal').addEventListener('click', () => {
  resetGoalForm();
  document.getElementById('goal-modal-title').textContent = '目標を追加';
  openModal('goal-modal');
});

function resetGoalForm() {
  document.getElementById('goal-form').reset();
  document.getElementById('goal-id').value = '';
  document.getElementById('goal-deadline').value = '';
}

function openGoalEdit(id) {
  const goal = goals.find(g => g.id === id);
  if (!goal) return;
  document.getElementById('goal-modal-title').textContent = '目標を編集';
  document.getElementById('goal-id').value = goal.id;
  document.getElementById('goal-name').value = goal.name;
  document.getElementById('goal-category').value = goal.category || 'その他';
  document.getElementById('goal-weight').value = goal.weight || '';
  document.getElementById('goal-reps').value = goal.reps || '';
  document.getElementById('goal-sets').value = goal.sets || '';
  document.getElementById('goal-deadline').value = goal.deadline || '';
  document.getElementById('goal-memo').value = goal.memo || '';
  openModal('goal-modal');
}

document.getElementById('goal-form').addEventListener('submit', e => {
  e.preventDefault();
  const id = document.getElementById('goal-id').value || uid();
  const goal = {
    id,
    name: document.getElementById('goal-name').value.trim(),
    category: document.getElementById('goal-category').value,
    weight: parseFloat(document.getElementById('goal-weight').value) || null,
    reps: parseInt(document.getElementById('goal-reps').value) || null,
    sets: parseInt(document.getElementById('goal-sets').value) || null,
    deadline: document.getElementById('goal-deadline').value || null,
    memo: document.getElementById('goal-memo').value.trim(),
  };
  const idx = goals.findIndex(g => g.id === id);
  if (idx >= 0) goals[idx] = goal;
  else goals.push(goal);

  saveGoals();
  closeModal('goal-modal');
  renderGoals();
  populateExerciseSelects();
});

function deleteGoal(id) {
  if (!confirm('この目標を削除しますか？')) return;
  goals = goals.filter(g => g.id !== id);
  saveGoals();
  renderGoals();
  populateExerciseSelects();
}

/* ===========================
   RECORDS
   =========================== */
function renderRecords() {
  const tbody = document.getElementById('records-tbody');
  const empty = document.getElementById('records-empty');
  const filterExercise = document.getElementById('record-filter-exercise').value;
  const filterDate = document.getElementById('record-filter-date').value;

  let filtered = [...records];
  if (filterExercise) filtered = filtered.filter(r => r.exercise === filterExercise);
  if (filterDate) filtered = filtered.filter(r => r.date === filterDate);
  filtered.sort((a, b) => b.date.localeCompare(a.date));

  tbody.innerHTML = '';
  if (filtered.length === 0) {
    empty.classList.add('visible');
    document.getElementById('records-list').style.display = 'none';
    return;
  }
  empty.classList.remove('visible');
  document.getElementById('records-list').style.display = '';

  filtered.forEach(rec => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${formatDate(rec.date)}</td>
      <td>${escHtml(rec.exercise)}</td>
      <td>${rec.weight != null ? rec.weight + ' kg' : '-'}</td>
      <td>${rec.reps}</td>
      <td>${rec.sets}</td>
      <td>${escHtml(rec.memo || '')}</td>
      <td>
        <button class="btn-secondary" data-edit-rec="${rec.id}">編集</button>
        <button class="btn-danger" data-delete-rec="${rec.id}">削除</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('[data-edit-rec]').forEach(btn => {
    btn.addEventListener('click', () => openRecordEdit(btn.dataset.editRec));
  });
  tbody.querySelectorAll('[data-delete-rec]').forEach(btn => {
    btn.addEventListener('click', () => deleteRecord(btn.dataset.deleteRec));
  });
}

document.getElementById('open-record-modal').addEventListener('click', () => {
  resetRecordForm();
  document.getElementById('record-modal-title').textContent = '記録を追加';
  document.getElementById('record-date').value = today();
  openModal('record-modal');
});

function resetRecordForm() {
  document.getElementById('record-form').reset();
  document.getElementById('record-id').value = '';
}

function openRecordEdit(id) {
  const rec = records.find(r => r.id === id);
  if (!rec) return;
  document.getElementById('record-modal-title').textContent = '記録を編集';
  document.getElementById('record-id').value = rec.id;
  document.getElementById('record-exercise').value = rec.exercise;
  document.getElementById('record-date').value = rec.date;
  document.getElementById('record-weight').value = rec.weight ?? '';
  document.getElementById('record-reps').value = rec.reps;
  document.getElementById('record-sets').value = rec.sets;
  document.getElementById('record-memo').value = rec.memo || '';
  openModal('record-modal');
}

document.getElementById('record-form').addEventListener('submit', e => {
  e.preventDefault();
  const id = document.getElementById('record-id').value || uid();
  const rec = {
    id,
    exercise: document.getElementById('record-exercise').value,
    date: document.getElementById('record-date').value,
    weight: parseFloat(document.getElementById('record-weight').value) || null,
    reps: parseInt(document.getElementById('record-reps').value),
    sets: parseInt(document.getElementById('record-sets').value),
    memo: document.getElementById('record-memo').value.trim(),
  };
  const idx = records.findIndex(r => r.id === id);
  if (idx >= 0) records[idx] = rec;
  else records.push(rec);

  saveRecords();
  closeModal('record-modal');
  renderRecords();
  renderGoals();
});

function deleteRecord(id) {
  if (!confirm('この記録を削除しますか？')) return;
  records = records.filter(r => r.id !== id);
  saveRecords();
  renderRecords();
  renderGoals();
}

document.getElementById('record-filter-exercise').addEventListener('change', renderRecords);
document.getElementById('record-filter-date').addEventListener('change', renderRecords);
document.getElementById('clear-filter').addEventListener('click', () => {
  document.getElementById('record-filter-exercise').value = '';
  document.getElementById('record-filter-date').value = '';
  renderRecords();
});

/* ===========================
   EXERCISE SELECTS
   =========================== */
function populateExerciseSelects() {
  const names = goals.map(g => g.name);
  const extraNames = [...new Set(records.map(r => r.exercise).filter(n => !names.includes(n)))];
  const allNames = [...names, ...extraNames];

  const selects = [
    document.getElementById('record-exercise'),
    document.getElementById('record-filter-exercise'),
    document.getElementById('chart-exercise-select'),
  ];

  selects.forEach((sel, i) => {
    const firstOpt = i === 0 ? '<option value="">選択してください</option>'
      : i === 1 ? '<option value="">すべて</option>'
      : '<option value="">選択してください</option>';
    sel.innerHTML = firstOpt + allNames.map(n => `<option value="${escAttr(n)}">${escHtml(n)}</option>`).join('');
  });
}

/* ===========================
   PROGRESS CHART
   =========================== */
document.getElementById('chart-exercise-select').addEventListener('change', renderChart);
document.getElementById('chart-metric-select').addEventListener('change', renderChart);

function renderChart() {
  const exercise = document.getElementById('chart-exercise-select').value;
  const metric = document.getElementById('chart-metric-select').value;
  const empty = document.getElementById('chart-empty');
  const canvas = document.getElementById('progress-chart');

  if (!exercise) {
    empty.style.display = 'block';
    canvas.style.display = 'none';
    document.getElementById('achievement-cards').innerHTML = '';
    if (progressChart) { progressChart.destroy(); progressChart = null; }
    return;
  }

  const data = records
    .filter(r => r.exercise === exercise)
    .sort((a, b) => a.date.localeCompare(b.date));

  if (data.length === 0) {
    empty.style.display = 'block';
    canvas.style.display = 'none';
    document.getElementById('achievement-cards').innerHTML = '';
    if (progressChart) { progressChart.destroy(); progressChart = null; }
    return;
  }

  empty.style.display = 'none';
  canvas.style.display = 'block';

  const labels = data.map(r => formatDate(r.date));
  const values = data.map(r => {
    if (metric === 'weight') return r.weight ?? 0;
    if (metric === 'reps') return r.reps;
    return (r.weight || 0) * r.reps * r.sets;
  });

  const metricLabel = metric === 'weight' ? '重量 (kg)' : metric === 'reps' ? '回数' : 'ボリューム';

  if (progressChart) progressChart.destroy();
  progressChart = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: `${exercise} - ${metricLabel}`,
        data: values,
        borderColor: '#6c63ff',
        backgroundColor: 'rgba(108,99,255,0.1)',
        pointBackgroundColor: '#6c63ff',
        pointRadius: 5,
        fill: true,
        tension: 0.3,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { display: true, position: 'top' },
        tooltip: { mode: 'index', intersect: false },
      },
      scales: {
        y: { beginAtZero: false, grid: { color: '#e5e7eb' } },
        x: { grid: { display: false } },
      },
    },
  });

  renderAchievementCards(exercise, data, metric);
}

function renderAchievementCards(exercise, data, metric) {
  const goal = goals.find(g => g.name === exercise);
  const container = document.getElementById('achievement-cards');
  container.innerHTML = '';

  const latest = data[data.length - 1];
  const maxWeight = Math.max(...data.map(r => r.weight || 0));
  const maxReps = Math.max(...data.map(r => r.reps));
  const totalSessions = data.length;

  const cards = [
    { label: '最新重量', value: latest.weight != null ? `${latest.weight}kg` : '-', sub: '直近の記録' },
    { label: '最高重量', value: maxWeight ? `${maxWeight}kg` : '-', sub: '全期間' },
    { label: '最高回数', value: `${maxReps}回`, sub: '全期間' },
    { label: '総セッション', value: `${totalSessions}回`, sub: 'の記録' },
  ];

  if (goal) {
    const pct = calcGoalProgress(goal);
    cards.push({ label: '目標達成率', value: `${pct}%`, sub: goal.weight ? `目標 ${goal.weight}kg` : goal.reps ? `目標 ${goal.reps}回` : '' });
  }

  cards.forEach(c => {
    const el = document.createElement('div');
    el.className = 'achievement-card';
    el.innerHTML = `
      <div class="ach-label">${c.label}</div>
      <div class="ach-value">${c.value}</div>
      <div class="ach-sub">${c.sub}</div>
    `;
    container.appendChild(el);
  });
}

/* ===========================
   REMINDER
   =========================== */
function checkNotificationPermission() {
  const banner = document.getElementById('notification-permission-banner');
  if (!('Notification' in window)) {
    banner.style.display = 'flex';
    banner.textContent = 'このブラウザは通知をサポートしていません。';
    return;
  }
  if (Notification.permission === 'default') {
    banner.style.display = 'flex';
  } else {
    banner.style.display = 'none';
  }
}

document.getElementById('request-permission-btn')?.addEventListener('click', async () => {
  const perm = await Notification.requestPermission();
  if (perm === 'granted') {
    document.getElementById('notification-permission-banner').style.display = 'none';
    showReminderStatus('通知が許可されました！');
  }
});

document.getElementById('save-reminder-btn').addEventListener('click', () => {
  const time = document.getElementById('reminder-time').value;
  const days = [...document.querySelectorAll('input[name="day"]:checked')].map(cb => parseInt(cb.value));
  const message = document.getElementById('reminder-message').value.trim();

  if (!time) { alert('時刻を設定してください。'); return; }
  if (days.length === 0) { alert('曜日を1つ以上選択してください。'); return; }

  reminder = { time, days, message: message || '今日も筋トレを頑張ろう！' };
  saveReminder();
  scheduleReminder();
  showReminderStatus(`リマインダーを保存しました。毎${days.map(d => '日月火水木金土'[d]).join('・')} ${time} に通知します。`);
  updateNextNotificationTime();
});

document.getElementById('clear-reminder-btn').addEventListener('click', () => {
  if (!confirm('リマインダーをクリアしますか？')) return;
  reminder = null;
  saveReminder();
  clearTimeout(reminderTimerId);
  reminderTimerId = null;
  document.getElementById('reminder-time').value = '07:00';
  document.querySelectorAll('input[name="day"]').forEach(cb => cb.checked = false);
  document.getElementById('reminder-message').value = '今日も筋トレを頑張ろう！';
  document.getElementById('next-notification-time').textContent = '設定されていません';
  showReminderStatus('リマインダーをクリアしました。');
});

function showReminderStatus(msg) {
  const el = document.getElementById('reminder-status');
  el.textContent = msg;
  el.style.display = 'flex';
  setTimeout(() => { el.style.display = 'none'; }, 5000);
}

function scheduleReminder() {
  clearTimeout(reminderTimerId);
  if (!reminder) return;

  const checkAndNotify = () => {
    const now = new Date();
    const [h, m] = reminder.time.split(':').map(Number);
    const target = new Date();
    target.setHours(h, m, 0, 0);

    const dayMatch = reminder.days.includes(now.getDay());
    const timeMatch = now.getHours() === h && now.getMinutes() === m;

    if (dayMatch && timeMatch && Notification.permission === 'granted') {
      new Notification('💪 筋トレの時間！', {
        body: reminder.message,
        icon: 'https://cdn.jsdelivr.net/npm/twemoji@14/assets/svg/1f4aa.svg',
      });
    }

    const msUntilNextMinute = (60 - now.getSeconds()) * 1000 - now.getMilliseconds();
    reminderTimerId = setTimeout(checkAndNotify, msUntilNextMinute);
  };

  const now = new Date();
  const msUntilNextMinute = (60 - now.getSeconds()) * 1000 - now.getMilliseconds();
  reminderTimerId = setTimeout(checkAndNotify, msUntilNextMinute);
}

function updateNextNotificationTime() {
  const el = document.getElementById('next-notification-time');
  if (!reminder) { el.textContent = '設定されていません'; return; }

  const dayNames = '日月火水木金土';
  const days = reminder.days.map(d => dayNames[d]).join('・');
  el.textContent = `毎週 ${days}曜日 ${reminder.time}`;
}

function loadReminderUI() {
  if (!reminder) return;
  document.getElementById('reminder-time').value = reminder.time || '07:00';
  document.getElementById('reminder-message').value = reminder.message || '';
  document.querySelectorAll('input[name="day"]').forEach(cb => {
    cb.checked = reminder.days?.includes(parseInt(cb.value));
  });
  updateNextNotificationTime();
  scheduleReminder();
}

/* ===========================
   XSS Prevention
   =========================== */
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function escAttr(str) { return escHtml(str); }

/* ===========================
   INIT
   =========================== */
function init() {
  renderGoals();
  populateExerciseSelects();
  renderRecords();
  checkNotificationPermission();
  loadReminderUI();
}

init();
