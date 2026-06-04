const state = {
  schedule: [],
  date: null,
  meal: null,
  options: [],
  selectedLabel: null,
};

const el = (id) => document.getElementById(id);
const api = async (url, opts) => {
  const res = await fetch(url, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `请求失败 (${res.status})`);
  return data;
};

const MEAL_LABEL = { lunch: '午餐', dinner: '晚餐' };

function nutriText(n) {
  if (!n || n.calories == null) return '营养数据待补充';
  return `${n.calories} kcal · 蛋白 ${n.protein}g · 脂肪 ${n.fat}g · 碳水 ${n.carbs}g`;
}

async function init() {
  const { schedule } = await api('/api/schedule?days=14');
  state.schedule = schedule;
  const sel = el('date-select');
  sel.innerHTML = schedule
    .map((d) => `<option value="${d.date}">${d.date}（${d.weekday}${d.isWeekend ? ' · 周末' : ''}）</option>`)
    .join('');
  sel.addEventListener('change', () => selectDate(sel.value));
  el('member').addEventListener('input', updateOrderButtons);
  selectDate(schedule[0].date);
}

function selectDate(date) {
  state.date = date;
  const day = state.schedule.find((d) => d.date === date);
  const tabs = el('meal-tabs');
  tabs.innerHTML = day.meals
    .map((m) => `<button class="meal-tab" data-meal="${m}">${MEAL_LABEL[m]}</button>`)
    .join('');
  tabs.querySelectorAll('.meal-tab').forEach((btn) => {
    btn.addEventListener('click', () => selectMeal(btn.dataset.meal));
  });
  selectMeal(day.meals[0]);
}

async function selectMeal(meal) {
  state.meal = meal;
  state.selectedLabel = null;
  document.querySelectorAll('.meal-tab').forEach((b) => {
    b.classList.toggle('active', b.dataset.meal === meal);
  });
  el('meal-title').textContent = `${state.date} ${MEAL_LABEL[meal]} · 套餐选择（A/B/C/D）`;
  clearStatus();
  try {
    const data = await api(`/api/meals?date=${state.date}&meal=${meal}`);
    state.options = data.options;
    renderOptions();
    renderExisting(data.existingOrders);
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
      const steps = (opt.steps || []).map((s, i) => `<li>${stripMd(s)}</li>`).join('');
      const sourceHtml = opt.source
        ? `<div class="source">来源：${
            opt.sourceUrl ? `<a href="${opt.sourceUrl}" target="_blank" rel="noopener">${opt.source}</a>` : opt.source
          }</div>`
        : '';
      const cal = opt.nutrition && opt.nutrition.calories != null ? `${opt.nutrition.calories} kcal` : '营养待补充';
      return `
      <div class="option-card" data-label="${opt.label}">
        <div class="option-head">
          <span class="option-label">${opt.label}</span>
          <span class="option-cal">${cal}</span>
        </div>
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
      ensureOrderBar();
    });
  });
  ensureOrderBar();
}

function ensureOrderBar() {
  let bar = document.getElementById('order-bar');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'order-bar';
    bar.className = 'order-bar';
    el('options').after(bar);
  }
  bar.innerHTML = `<button id="submit" class="btn-primary">保存点餐并发送备菜邮件</button>`;
  document.getElementById('submit').addEventListener('click', submitOrder);
  updateOrderButtons();
}

function updateOrderButtons() {
  const btn = document.getElementById('submit');
  if (!btn) return;
  const member = el('member').value.trim();
  btn.disabled = !member || !state.selectedLabel;
}

async function submitOrder() {
  const member = el('member').value.trim();
  const btn = document.getElementById('submit');
  btn.disabled = true;
  btn.textContent = '提交中…';
  try {
    const data = await api('/api/orders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ member, date: state.date, meal: state.meal, optionLabel: state.selectedLabel }),
    });
    showSuccess(data);
    const meals = await api(`/api/orders?date=${state.date}&meal=${state.meal}`);
    renderExisting(meals.orders);
  } catch (err) {
    showError(err.message);
  } finally {
    btn.textContent = '保存点餐并发送备菜邮件';
    updateOrderButtons();
  }
}

function showSuccess({ order, mail }) {
  const box = el('status');
  box.className = 'status ok';
  let mailLine;
  if (mail.sent) mailLine = `备菜邮件已发送 ✅（${mail.mode}）`;
  else if (mail.mode === 'outbox') mailLine = `邮件已生成预览（${mail.reason}）。预览文件：${mail.file}`;
  else mailLine = `邮件发送失败：${mail.error || mail.reason || '未知错误'}`;

  box.innerHTML = `✅ <b>${order.member}</b> 已点 <b>${order.option.label} 餐</b>：${order.option.title}（${nutriText(
    order.option.nutrition,
  )}）。<br>${mailLine}`;
  if (mail.preview) {
    box.innerHTML += `<details><summary style="cursor:pointer;margin-top:8px;">查看邮件内容预览</summary><pre>${escapeHtml(
      mail.preview,
    )}</pre></details>`;
  }
  box.classList.remove('hidden');
  box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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

function renderExisting(orders) {
  const panel = el('existing');
  const list = el('existing-list');
  const badge = el('ordered-badge');
  if (!orders || orders.length === 0) {
    panel.classList.add('hidden');
    badge.classList.add('hidden');
    return;
  }
  badge.textContent = `已下单 ${orders.length} 人`;
  badge.classList.remove('hidden');
  list.innerHTML = orders
    .map(
      (o) =>
        `<li><b>${o.member}</b> · ${o.option.label} 餐 — ${o.option.title}（${
          o.option.nutrition && o.option.nutrition.calories != null ? `${o.option.nutrition.calories} kcal` : '营养待补充'
        }）</li>`,
    )
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
