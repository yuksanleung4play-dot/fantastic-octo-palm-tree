import nodemailer from 'nodemailer';
import { writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { sumNutrition } from './recipes.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUTBOX_DIR = join(__dirname, '..', 'data', 'outbox');
const MEAL_LABEL = { lunch: '午餐', dinner: '晚餐' };
const WEEKDAYS = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];

/**
 * 根据环境变量创建邮件传输器。
 * 若未配置 SMTP，则返回 null，系统会改为把邮件落盘到 data/outbox 以便预览。
 */
function createTransport() {
  const { SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS } = process.env;
  if (!SMTP_HOST || !SMTP_USER || !SMTP_PASS) return null;
  return nodemailer.createTransport({
    host: SMTP_HOST,
    port: Number(SMTP_PORT || 587),
    secure: Number(SMTP_PORT) === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });
}

/** 按食谱分组记录，返回 [{ option, count, meals[] }]。 */
export function groupOrdersByRecipe(allOrders) {
  const map = new Map();
  for (const o of allOrders) {
    const id = o.option.recipeId;
    if (!map.has(id)) map.set(id, { option: o.option, count: 0, meals: new Set() });
    const g = map.get(id);
    g.count += 1;
    g.meals.add(MEAL_LABEL[o.meal] || o.meal);
  }
  return [...map.values()].map((g) => ({ option: g.option, count: g.count, meals: [...g.meals] }));
}

function nutriLine(n) {
  if (n.calories == null && n.hasMissing && !n.protein) return '营养数据待补充';
  const missing = n.hasMissing ? '（含待补充项）' : '';
  return `${n.calories} kcal｜蛋白质 ${n.protein}g｜脂肪 ${n.fat}g｜碳水 ${n.carbs}g${missing}`;
}

function weekdayOf(dateStr) {
  return WEEKDAYS[new Date(`${dateStr}T00:00:00`).getDay()];
}

/** 生成 [start, end]（含端点）的日期字符串数组。 */
export function dateRange(start, end) {
  const out = [];
  const d = new Date(`${start}T00:00:00`);
  const last = new Date(`${end}T00:00:00`);
  while (d <= last) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    out.push(`${y}-${m}-${day}`);
    d.setDate(d.getDate() + 1);
  }
  return out;
}

/**
 * 构建「本周餐单」邮件：覆盖 rangeStart~rangeEnd（通常为周一至周五）。
 * 含每日已点餐单、按食谱合并的备菜清单与营养合计。
 */
export function buildWeeklyEmail({ orders, rangeStart, rangeEnd }) {
  const subject = `🍽️ 家庭点餐｜${rangeStart} ~ ${rangeEnd} 本周餐单与备菜清单（共 ${orders.length} 餐）`;
  const days = dateRange(rangeStart, rangeEnd);
  const byDate = new Map();
  for (const o of orders) {
    if (!byDate.has(o.date)) byDate.set(o.date, []);
    byDate.get(o.date).push(o);
  }

  const groups = groupOrdersByRecipe(orders);
  const totalNutrition = sumNutrition(groups.map((g) => ({ nutrition: g.option.nutrition, _count: g.count })));

  const dayLines = days
    .map((date) => {
      const list = byDate.get(date) || [];
      if (list.length === 0) return `    ${date}（${weekdayOf(date)}）：未点餐`;
      const meals = list
        .map((o) => `${MEAL_LABEL[o.meal] || o.meal} ${o.option.label}餐 — ${o.option.title}`)
        .join('；');
      return `    ${date}（${weekdayOf(date)}）：${meals}`;
    })
    .join('\n');

  const shoppingBlocks = groups
    .map((g) => {
      const head = `  ▸ ${g.option.title} ×${g.count} 份`;
      const lines = g.option.ingredientGroups
        .map((grp) => {
          const items = grp.items.map((it) => `      - ${it}${g.count > 1 ? `  ×${g.count}` : ''}`).join('\n');
          return grp.label ? `    【${grp.label}】\n${items}` : items;
        })
        .join('\n');
      return `${head}\n${lines}`;
    })
    .join('\n\n');

  const text = [
    `家庭点餐 · 本周餐单`,
    ``,
    `周期：${rangeStart} ~ ${rangeEnd}`,
    ``,
    `———————————————`,
    `每日餐单：`,
    dayLines,
    ``,
    orders.length
      ? `🛒 一周备菜清单（按食谱合并，已标注份数）：\n${shoppingBlocks}\n\n全部营养合计：${nutriLine(totalNutrition)}`
      : `本周暂无已保存餐单。`,
  ]
    .filter((l) => l !== '')
    .join('\n');

  const html = renderHtml({ rangeStart, rangeEnd, days, byDate, groups, totalNutrition, count: orders.length });
  return { subject, text, html };
}

function renderHtml({ rangeStart, rangeEnd, days, byDate, groups, totalNutrition, count }) {
  const dayRows = days
    .map((date) => {
      const list = byDate.get(date) || [];
      const cell = list.length
        ? list.map((o) => `${MEAL_LABEL[o.meal] || o.meal} ${o.option.label}餐 — ${esc(o.option.title)}`).join('<br>')
        : '<span style="color:#9ca3af;">未点餐</span>';
      return `<tr><td style="white-space:nowrap;">${date}<br><span style="color:#6b7280;">${weekdayOf(date)}</span></td><td>${cell}</td></tr>`;
    })
    .join('');

  const shoppingBlocks = groups
    .map((g) => {
      const groupsHtml = g.option.ingredientGroups
        .map((grp) => {
          const items = grp.items
            .map((it) => `<li>${esc(it)}${g.count > 1 ? ` <b style="color:#ea580c;">×${g.count}</b>` : ''}</li>`)
            .join('');
          const label = grp.label ? `<div style="font-weight:600;color:#6b7280;margin-top:6px;">【${esc(grp.label)}】</div>` : '';
          return `${label}<ul style="margin:4px 0;padding-left:20px;">${items}</ul>`;
        })
        .join('');
      return `<div style="border:1px solid #fed7aa;border-radius:10px;padding:12px 14px;margin-bottom:10px;">
        <div style="font-weight:700;">${esc(g.option.title)} <span style="color:#ea580c;">×${g.count} 份</span></div>
        ${groupsHtml}
      </div>`;
    })
    .join('');

  return `<!doctype html><html lang="zh"><head><meta charset="utf-8"></head>
  <body style="font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;color:#1f2937;max-width:680px;margin:0 auto;padding:24px;">
    <h2 style="color:#ea580c;margin-bottom:4px;">🍽️ 家庭点餐 · 本周餐单</h2>
    <p style="color:#6b7280;margin-top:0;">${rangeStart} ~ ${rangeEnd} · 共 ${count} 餐</p>

    <h3 style="border-bottom:2px solid #fed7aa;padding-bottom:6px;">每日餐单</h3>
    <table style="width:100%;border-collapse:collapse;">${dayRows}</table>

    ${
      count
        ? `<h3 style="border-bottom:2px solid #bbf7d0;padding-bottom:6px;">🛒 一周备菜清单（按食谱合并）</h3>
    ${shoppingBlocks}
    <p style="background:#f0fdf4;padding:10px 14px;border-radius:8px;">全部营养合计：${nutriLine(totalNutrition)}</p>`
        : `<p style="background:#fff7ed;padding:10px 14px;border-radius:8px;">本周暂无已保存餐单。</p>`
    }
  </body></html>`;
}

function esc(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

/**
 * 发送「本周餐单」邮件。返回 { sent, mode, ... }。
 * - 已配置 SMTP：真实发送
 * - 未配置 SMTP：落盘到 data/outbox（开发/演示模式）
 */
export async function sendWeeklyEmail({ orders, rangeStart, rangeEnd }) {
  const to = process.env.NOTIFY_EMAIL;
  const message = buildWeeklyEmail({ orders, rangeStart, rangeEnd });
  const transport = createTransport();

  if (!transport || !to) {
    if (!existsSync(OUTBOX_DIR)) mkdirSync(OUTBOX_DIR, { recursive: true });
    const file = join(OUTBOX_DIR, `weekly-${rangeStart}_${rangeEnd}.html`);
    writeFileSync(file, `<!-- 主题：${message.subject} -->\n${message.html}`, 'utf-8');
    return {
      sent: false,
      mode: 'outbox',
      reason: !to ? '未配置 NOTIFY_EMAIL' : '未配置 SMTP',
      file,
      rangeStart,
      rangeEnd,
      count: orders.length,
    };
  }

  const info = await transport.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER,
    to,
    subject: message.subject,
    text: message.text,
    html: message.html,
  });
  return { sent: true, mode: 'smtp', messageId: info.messageId, rangeStart, rangeEnd, count: orders.length };
}
