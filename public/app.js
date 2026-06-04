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

async function init() {
  const { schedule } = await api('/api/schedule?days=14');
  state.schedule = schedule;
  const sel = el('date-select');
  sel.innerHTML = schedule
    .map(
      (d) =>
        `<option value="${d.date}">${d.date}（${d.weekday}${d.isWeekend ? ' · 周末' : ''}）</option>`,
    )
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
  el('meal-title').textContent = `${state.date} ${MEAL_LABEL[meal]} · 套餐选择`;
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
      const dishes = opt.dishes
        .map(
          (d) => `
        <div class="dish">
          <span class="tag ${d.type === 'meat' ? 'meat' : 'veg'}">${d.type === 'meat' ? '荤' : '素'}</span>
          <span class="name">${d.name}</span>
          <span class="cal">${d.calories} kcal</span>
        </div>`,
        )
        .join('');
      const ingredients = opt.ingredients
        .map((i) => `<li>${i.name}：${i.amount} ${i.unit}</li>`)
        .join('');
      const n = opt.nutrition;
      return `
      <div class="option-card" data-label="${opt.label}">
        <div class="option-head">
          <span class="option-label">${opt.label}</span>
          <span class="option-cal">${n.calories} kcal</span>
        </div>
        ${dishes}
        <div class="nutri">
          <span>蛋白 ${n.protein}g</span><span>脂肪 ${n.fat}g</span>
          <span>碳水 ${n.carbs}g</span><span>纤维 ${n.fiber}g</span><span>钠 ${n.sodium}mg</span>
        </div>
        <details class="ingredients"><summary>食材与分量</summary><ul>${ingredients}</ul></details>
      </div>`;
    })
    .join('');

  grid.querySelectorAll('.option-card').forEach((card) => {
    card.addEventListener('click', () => {
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
      body: JSON.stringify({
        member,
        date: state.date,
        meal: state.meal,
        optionLabel: state.selectedLabel,
      }),
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
  if (mail.sent) {
    mailLine = `备菜邮件已发送 ✅（${mail.mode}）`;
  } else if (mail.mode === 'outbox') {
    mailLine = `邮件已生成预览（${mail.reason}）。预览文件：${mail.file}`;
  } else {
    mailLine = `邮件发送失败：${mail.error || mail.reason || '未知错误'}`;
  }
  box.innerHTML = `✅ <b>${order.member}</b> 已点 <b>${order.option.label} 餐</b>（${order.option.dishes
    .map((d) => d.name)
    .join(' + ')}），合计 ${order.option.nutrition.calories} kcal。<br>${mailLine}`;
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
        `<li><b>${o.member}</b> · ${o.option.label} 餐 — ${o.option.dishes
          .map((d) => d.name)
          .join(' + ')} （${o.option.nutrition.calories} kcal）</li>`,
    )
    .join('');
  panel.classList.remove('hidden');
}

function escapeHtml(s) {
  return s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

init().catch((err) => showError(err.message));
