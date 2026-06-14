const state = {
  schedule: [],
  date: null,
  meal: null,
  options: [],
  selectedLabel: null,
  offline: false,
};

const MEAL_LABEL = { lunch: '午餐', dinner: '晚餐' };
const WEEKDAYS = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];

const el = (id) => document.getElementById(id);
const api = async (url, opts) => {
  const res = await fetch(url, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `请求失败 (${res.status})`);
  return data;
};

function nutriText(n) {
  if (!n || n.calories == null) return '营养数据待补充';
  return `${n.calories} kcal · 蛋白 ${n.protein}g · 脂肪 ${n.fat}g · 碳水 ${n.carbs}g`;
}

function formatDate(dateStr) {
  const [, m, d] = dateStr.split('-');
  return `${Number(m)}月${Number(d)}日`;
}

function toDateStr(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/* 后端不可用时（如 htmlpreview 静态预览）在前端计算排程，保证首屏可用。 */
function buildScheduleClient(days = 14) {
  const out = [];
  const base = new Date();
  base.setHours(0, 0, 0, 0);
  for (let i = 0; i < days; i += 1) {
    const d = new Date(base);
    d.setDate(base.getDate() + i);
    const day = d.getDay();
    const isWeekend = day === 0 || day === 6;
    out.push({
      date: toDateStr(d),
      weekday: WEEKDAYS[day],
      isWeekend,
      meals: isWeekend ? ['lunch', 'dinner'] : ['dinner'],
    });
  }
  return out;
}

/* ===================== 首屏 ===================== */
function setupHome() {
  const today = state.schedule[0];
  el('home-date').textContent = `${formatDate(today.date)} ${today.weekday} · ${
    today.isWeekend ? '週末' : '平日'
  }（${today.meals.map((m) => MEAL_LABEL[m]).join('／')}）`;

  el('home-start').addEventListener('click', enterApp);
  el('back-home').addEventListener('click', showHome);
}

function showHome() {
  el('app').classList.remove('active');
  el('home').classList.add('active');
  window.scrollTo(0, 0);
}

function enterApp() {
  el('home').classList.remove('active');
  el('app').classList.add('active');
  window.scrollTo(0, 0);
  clearStatus();
}

/* ===================== 主屏 ===================== */
async function init() {
  try {
    const { schedule } = await api('/api/schedule?days=14');
    state.schedule = schedule;
  } catch {
    state.offline = true;
    state.schedule = buildScheduleClient(14);
  }

  const sel = el('date-select');
  sel.innerHTML = state.schedule
    .map((d) => `<option value="${d.date}">${d.date}（${d.weekday}${d.isWeekend ? ' · 週末' : ''}）</option>`)
    .join('');
  sel.addEventListener('change', () => selectDate(sel.value));
  el('submit').addEventListener('click', submitOrder);

  setupHome();
  selectDate(state.schedule[0].date);
}

function selectDate(date) {
  state.date = date;
  const day = state.schedule.find((d) => d.date === date);
  el('appbar-member').textContent = `${formatDate(date)} · 全家共享`;
  const tabs = el('meal-tabs');
  tabs.innerHTML = day.meals.map((m) => `<button class="meal-tab" data-meal="${m}">${MEAL_LABEL[m]}</button>`).join('');
  tabs.querySelectorAll('.meal-tab').forEach((btn) => {
    btn.addEventListener('click', () => selectMeal(btn.dataset.meal));
  });
  selectMeal(day.meals[0]);
}

async function selectMeal(meal) {
  state.meal = meal;
  state.selectedLabel = null;
  document.querySelectorAll('.meal-tab').forEach((b) => b.classList.toggle('active', b.dataset.meal === meal));
  el('meal-title').textContent = `${formatDate(state.date)} ${MEAL_LABEL[meal]} · 选套餐`;
  clearStatus();
  updateOrderBar();

  if (state.offline) {
    state.options = [];
    el('options').innerHTML =
      '<p class="empty">📦 这是静态前端预览。套餐生成与下单需要运行后端：<br><code>npm install &amp;&amp; npm start</code>，再用本机地址打开。</p>';
    renderRecords([]);
    return;
  }

  try {
    const data = await api(`/api/meals?date=${state.date}&meal=${meal}`);
    state.options = data.options;
    renderOptions();
    renderRecords(data.existingOrders);
  } catch (err) {
    el('options').innerHTML = `<p class="empty">${err.message}</p>`;
  }
}

function renderOptions() {
  const grid = el('options');
  grid.innerHTML = state.options
    .map((opt) => {
      const ingredients = opt.ingredientGroups
        .map((g) => {
          const items = g.items.map((it) => `<li>${it}</li>`).join('');
          const label = g.label ? `<div class="ing-label">【${g.label}】</div>` : '';
          return `${label}<ul>${items}</ul>`;
        })
        .join('');
      const steps = (opt.steps || []).map((s) => `<li>${stripMd(s)}</li>`).join('');
      const sourceHtml = opt.source
        ? `<div class="source">来源：${
            opt.sourceUrl ? `<a href="${opt.sourceUrl}" target="_blank" rel="noopener">${opt.source}</a>` : opt.source
          }</div>`
        : '';
      const cal = opt.nutrition && opt.nutrition.calories != null ? `${opt.nutrition.calories} kcal` : '营养待补充';
      const photo = `<div class="option-photo">${
        opt.image ? `<img src="${opt.image}" alt="${escapeHtml(opt.title)}" loading="lazy" onerror="this.remove()" />` : ''
      }</div>`;
      return `
      <div class="option-card" data-label="${opt.label}">
        ${photo}
        <div class="option-head"><span class="option-label">${opt.label}</span><span class="option-cal">${cal}</span></div>
        <div class="option-title">${opt.title}</div>
        <div class="meal-meta">${opt.mealType || ''}</div>
        <div class="nutri"><span>${nutriText(opt.nutrition)}</span></div>
        ${sourceHtml}
        <details class="ingredients"><summary>🛒 食材与分量</summary>${ingredients}</details>
        ${steps ? `<details class="steps"><summary>👩‍🍳 做法步骤</summary><ol>${steps}</ol></details>` : ''}
      </div>`;
    })
    .join('');

  grid.querySelectorAll('.option-card').forEach((card) => {
    card.addEventListener('click', (e) => {
      if (e.target.closest('details') || e.target.closest('a')) return;
      state.selectedLabel = card.dataset.label;
      grid.querySelectorAll('.option-card').forEach((c) => c.classList.remove('selected'));
      card.classList.add('selected');
      updateOrderBar();
    });
  });
}

function updateOrderBar() {
  const info = el('order-bar-info');
  const btn = el('submit');
  const opt = state.options.find((o) => o.label === state.selectedLabel);
  if (opt) {
    info.textContent = `已选 ${opt.label} 餐 · ${opt.title}`;
    btn.disabled = false;
  } else {
    info.textContent = state.offline ? '静态预览模式（需后端才能点餐）' : '请选择一个套餐';
    btn.disabled = true;
  }
}

async function submitOrder() {
  const btn = el('submit');
  btn.disabled = true;
  btn.textContent = '提交中…';
  try {
    const data = await api('/api/orders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ date: state.date, meal: state.meal, optionLabel: state.selectedLabel }),
    });
    showSuccess(data);
    renderRecords(data.allOrders);
  } catch (err) {
    showError(err.message);
  } finally {
    btn.textContent = '保存点餐';
    state.selectedLabel = null;
    document.querySelectorAll('.option-card').forEach((c) => c.classList.remove('selected'));
    updateOrderBar();
  }
}

function showSuccess({ order, allOrders }) {
  const box = el('status');
  box.className = 'status ok';
  box.innerHTML = `✅ 已保存 <b>${order.option.label} 餐</b>：${order.option.title}（${nutriText(
    order.option.nutrition,
  )}）。当天共 <b>${(allOrders || []).length}</b> 笔记录（每餐只保留最新选择）。<br>📧 本周餐单将于 <b>周六下午 5:00</b> 统一发送备菜邮件。`;
  box.classList.remove('hidden');
  box.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function showError(msg) {
  const box = el('status');
  box.className = 'status err';
  box.textContent = `❌ ${msg}`;
  box.classList.remove('hidden');
}

function clearStatus() {
  el('status').classList.add('hidden');
}

function renderRecords(orders) {
  const panel = el('existing');
  const list = el('existing-list');
  const badge = el('ordered-badge');
  if (!orders || orders.length === 0) {
    panel.classList.add('hidden');
    badge.classList.add('hidden');
    return;
  }
  badge.textContent = `当天 ${orders.length} 笔`;
  badge.classList.remove('hidden');
  list.innerHTML = orders
    .map((o, i) => {
      const cal = o.option.nutrition && o.option.nutrition.calories != null ? `${o.option.nutrition.calories} kcal` : '营养待补充';
      const thumb = `<span class="rec-thumb">${
        o.option.image ? `<img src="${o.option.image}" alt="" loading="lazy" onerror="this.remove()" />` : '🍽️'
      }</span>`;
      return `<li>${thumb}<span class="rec-text"><b>${i + 1}.</b> ${MEAL_LABEL[o.meal] || o.meal} · ${o.option.label} 餐 — ${o.option.title}（${cal}）</span></li>`;
    })
    .join('');
  panel.classList.remove('hidden');
}

function stripMd(s) {
  return escapeHtml(String(s).replace(/\*\*/g, ''));
}
function escapeHtml(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

init().catch((err) => showError(err.message));
