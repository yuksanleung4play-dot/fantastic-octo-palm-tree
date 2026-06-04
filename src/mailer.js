import nodemailer from 'nodemailer';
import { writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { sumNutrition } from './recipes.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUTBOX_DIR = join(__dirname, '..', 'data', 'outbox');
const MEAL_LABEL = { lunch: '午餐', dinner: '晚餐' };

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

/** 按食谱分组所有订单，返回 [{ option, members[], count }]。 */
export function groupOrdersByRecipe(allOrders) {
  const map = new Map();
  for (const o of allOrders) {
    const id = o.option.recipeId;
    if (!map.has(id)) map.set(id, { option: o.option, members: [], count: 0 });
    const g = map.get(id);
    g.members.push(o.member);
    g.count += 1;
  }
  return [...map.values()];
}

function nutriLine(n) {
  if (n.calories == null && n.hasMissing && !n.protein) return '营养数据待补充';
  const missing = n.hasMissing ? '（含待补充项）' : '';
  return `${n.calories} kcal｜蛋白质 ${n.protein}g｜脂肪 ${n.fat}g｜碳水 ${n.carbs}g${missing}`;
}

/**
 * 构建备菜邮件内容。
 * 包含：本次下单详情 + 该餐所有家人按食谱汇总的备菜清单（×份数）。
 */
export function buildMealEmail({ order, allOrders }) {
  const mealName = MEAL_LABEL[order.meal] || order.meal;
  const subject = `🍽️ 家庭点餐｜${order.date} ${mealName}：${order.member} 点了「${order.option.title}」`;

  const groups = groupOrdersByRecipe(allOrders);
  const totalNutrition = sumNutrition(groups.map((g) => ({ nutrition: g.option.nutrition, _count: g.count })));

  // ---- 纯文本版 ----
  const shoppingBlocks = groups
    .map((g) => {
      const head = `  ▸ ${g.option.title} ×${g.count} 份（${g.members.join('、')}）`;
      const lines = g.option.ingredientGroups
        .map((grp) => {
          const items = grp.items.map((it) => `      - ${it}${g.count > 1 ? `  ×${g.count}` : ''}`).join('\n');
          return grp.label ? `    【${grp.label}】\n${items}` : items;
        })
        .join('\n');
      return `${head}\n${lines}`;
    })
    .join('\n\n');

  const orderSummary = allOrders
    .map((o) => `    • ${o.member}：${o.option.label} 餐 — ${o.option.title}`)
    .join('\n');

  const text = [
    `家庭点餐通知`,
    ``,
    `日期：${order.date}（${mealName}）`,
    `下单人：${order.member}`,
    `所选：${order.option.label} 餐 —— ${order.option.title}`,
    `营养：${nutriLine({ ...order.option.nutrition, hasMissing: order.option.nutrition.calories == null })}`,
    order.option.source ? `来源：${order.option.source}${order.option.sourceUrl ? ` ${order.option.sourceUrl}` : ''}` : '',
    ``,
    `———————————————`,
    `本餐已下单家人（共 ${allOrders.length} 人）：`,
    orderSummary,
    ``,
    `🛒 需要准备的食材与分量（按所点食谱分组，已标注份数）：`,
    shoppingBlocks,
    ``,
    `全部营养合计：${nutriLine(totalNutrition)}`,
  ]
    .filter((l) => l !== '')
    .join('\n');

  const html = renderHtml({ order, mealName, allOrders, groups, totalNutrition });
  return { subject, text, html };
}

function renderHtml({ order, mealName, allOrders, groups, totalNutrition }) {
  const orderRows = allOrders
    .map((o) => `<tr><td>${esc(o.member)}</td><td>${o.option.label} 餐</td><td>${esc(o.option.title)}</td></tr>`)
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
        <div style="font-size:12px;color:#6b7280;">${esc(g.members.join('、'))}</div>
        ${groupsHtml}
      </div>`;
    })
    .join('');

  return `<!doctype html><html lang="zh"><head><meta charset="utf-8"></head>
  <body style="font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;color:#1f2937;max-width:680px;margin:0 auto;padding:24px;">
    <h2 style="color:#ea580c;margin-bottom:4px;">🍽️ 家庭点餐通知</h2>
    <p style="color:#6b7280;margin-top:0;">${order.date} · ${mealName} · 下单人 <b>${esc(order.member)}</b></p>

    <h3 style="border-bottom:2px solid #fed7aa;padding-bottom:6px;">本次所选：${order.option.label} 餐</h3>
    <p style="font-size:16px;font-weight:700;margin:6px 0;">${esc(order.option.title)}</p>
    <p style="background:#fff7ed;padding:10px 14px;border-radius:8px;">营养：${nutriLine({ ...order.option.nutrition, hasMissing: order.option.nutrition.calories == null })}</p>
    ${order.option.source ? `<p style="font-size:12px;color:#6b7280;">来源：${order.option.sourceUrl ? `<a href="${esc(order.option.sourceUrl)}">${esc(order.option.source)}</a>` : esc(order.option.source)}</p>` : ''}

    <h3 style="border-bottom:2px solid #fed7aa;padding-bottom:6px;">本餐已下单家人（${allOrders.length} 人）</h3>
    <table style="width:100%;border-collapse:collapse;">${orderRows}</table>

    <h3 style="border-bottom:2px solid #bbf7d0;padding-bottom:6px;">🛒 备菜清单（按食谱分组 · 已标注份数）</h3>
    ${shoppingBlocks}
    <p style="background:#f0fdf4;padding:10px 14px;border-radius:8px;">全部营养合计：${nutriLine(totalNutrition)}</p>
  </body></html>`;
}

function esc(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

/**
 * 发送备菜邮件。返回 { sent, mode, ... }。
 * - 已配置 SMTP：真实发送
 * - 未配置 SMTP：落盘到 data/outbox（开发/演示模式）
 */
export async function sendMealEmail({ order, allOrders }) {
  const to = process.env.NOTIFY_EMAIL;
  const message = buildMealEmail({ order, allOrders });
  const transport = createTransport();

  if (!transport || !to) {
    if (!existsSync(OUTBOX_DIR)) mkdirSync(OUTBOX_DIR, { recursive: true });
    const safeId = (order.id || `${order.date}-${order.meal}-${order.member}`).replace(/[^\w\-]/g, '_');
    const file = join(OUTBOX_DIR, `${safeId}.html`);
    writeFileSync(file, `<!-- 主题：${message.subject} -->\n${message.html}`, 'utf-8');
    return {
      sent: false,
      mode: 'outbox',
      reason: !to ? '未配置 NOTIFY_EMAIL' : '未配置 SMTP',
      file,
      preview: message.text,
    };
  }

  const info = await transport.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER,
    to,
    subject: message.subject,
    text: message.text,
    html: message.html,
  });
  return { sent: true, mode: 'smtp', messageId: info.messageId };
}
