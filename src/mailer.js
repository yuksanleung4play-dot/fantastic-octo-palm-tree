import nodemailer from 'nodemailer';
import { writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { combineIngredients } from './mealGenerator.js';
import { sumNutrition } from './recipes.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUTBOX_DIR = join(__dirname, '..', 'data', 'outbox');

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

const MEAL_LABEL = { lunch: '午餐', dinner: '晚餐' };

/**
 * 构建备菜邮件内容。
 * 包含：本次下单详情 + 该餐所有家人的食材汇总与分量。
 */
export function buildMealEmail({ order, allOrders }) {
  const mealName = MEAL_LABEL[order.meal] || order.meal;
  const subject = `🍽️ 家庭点餐｜${order.date} ${mealName}：${order.member} 已下单（${order.option.label}餐）`;

  // 汇总该餐所有订单涉及的全部菜品，用于备菜清单。
  const allDishes = [];
  for (const o of allOrders) {
    for (const d of o.option.dishes) allDishes.push(d);
  }
  const shoppingList = combineIngredients(allDishes);
  const totalNutrition = sumNutrition(allDishes);

  const dishLines = order.option.dishes
    .map((d) => `    • ${d.name}（${d.type === 'meat' ? '荤' : '素'}，${d.calories} kcal）`)
    .join('\n');

  const orderSummary = allOrders
    .map((o) => `    • ${o.member}：${o.option.label}餐 — ${o.option.dishes.map((d) => d.name).join(' + ')}`)
    .join('\n');

  const shoppingLines = shoppingList
    .map((i) => `    • ${i.name}：${round(i.amount)} ${i.unit}`)
    .join('\n');

  const text = [
    `家庭点餐通知`,
    ``,
    `日期：${order.date}（${mealName}）`,
    `下单人：${order.member}`,
    `所选套餐：${order.option.label} 餐`,
    dishLines,
    ``,
    `本套餐营养合计：${order.option.nutrition.calories} kcal`,
    `  蛋白质 ${order.option.nutrition.protein}g｜脂肪 ${order.option.nutrition.fat}g｜碳水 ${order.option.nutrition.carbs}g｜膳食纤维 ${order.option.nutrition.fiber}g｜钠 ${order.option.nutrition.sodium}mg`,
    ``,
    `———————————————`,
    `本餐已下单家人（共 ${allOrders.length} 人）：`,
    orderSummary,
    ``,
    `🛒 需要准备的食材与分量（已按所有下单累加）：`,
    shoppingLines,
    ``,
    `全部菜品营养合计：${totalNutrition.calories} kcal`,
    `  蛋白质 ${totalNutrition.protein}g｜脂肪 ${totalNutrition.fat}g｜碳水 ${totalNutrition.carbs}g｜膳食纤维 ${totalNutrition.fiber}g｜钠 ${totalNutrition.sodium}mg`,
  ].join('\n');

  const html = renderHtml({
    order,
    mealName,
    allOrders,
    shoppingList,
    totalNutrition,
  });

  return { subject, text, html };
}

function round(n) {
  return Math.round(n * 10) / 10;
}

function renderHtml({ order, mealName, allOrders, shoppingList, totalNutrition }) {
  const dishRows = order.option.dishes
    .map(
      (d) => `<tr><td>${d.name}</td><td>${d.type === 'meat' ? '荤' : '素'}</td><td>${d.calories} kcal</td></tr>`,
    )
    .join('');
  const orderRows = allOrders
    .map(
      (o) =>
        `<tr><td>${o.member}</td><td>${o.option.label} 餐</td><td>${o.option.dishes
          .map((d) => d.name)
          .join(' + ')}</td></tr>`,
    )
    .join('');
  const shopRows = shoppingList
    .map((i) => `<tr><td>${i.name}</td><td>${round(i.amount)} ${i.unit}</td></tr>`)
    .join('');

  return `<!doctype html><html lang="zh"><head><meta charset="utf-8"></head>
  <body style="font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;color:#1f2937;max-width:640px;margin:0 auto;padding:24px;">
    <h2 style="color:#ea580c;margin-bottom:4px;">🍽️ 家庭点餐通知</h2>
    <p style="color:#6b7280;margin-top:0;">${order.date} · ${mealName} · 下单人 <b>${order.member}</b></p>

    <h3 style="border-bottom:2px solid #fed7aa;padding-bottom:6px;">本次所选：${order.option.label} 餐</h3>
    <table style="width:100%;border-collapse:collapse;">${dishRows}</table>
    <p style="background:#fff7ed;padding:10px 14px;border-radius:8px;">
      营养合计 <b>${order.option.nutrition.calories} kcal</b> ·
      蛋白质 ${order.option.nutrition.protein}g · 脂肪 ${order.option.nutrition.fat}g ·
      碳水 ${order.option.nutrition.carbs}g · 纤维 ${order.option.nutrition.fiber}g · 钠 ${order.option.nutrition.sodium}mg
    </p>

    <h3 style="border-bottom:2px solid #fed7aa;padding-bottom:6px;">本餐已下单家人（${allOrders.length} 人）</h3>
    <table style="width:100%;border-collapse:collapse;">${orderRows}</table>

    <h3 style="border-bottom:2px solid #bbf7d0;padding-bottom:6px;">🛒 备菜清单（按所有下单累加）</h3>
    <table style="width:100%;border-collapse:collapse;">${shopRows}</table>
    <p style="background:#f0fdf4;padding:10px 14px;border-radius:8px;">
      全部菜品营养合计 <b>${totalNutrition.calories} kcal</b> ·
      蛋白质 ${totalNutrition.protein}g · 脂肪 ${totalNutrition.fat}g ·
      碳水 ${totalNutrition.carbs}g · 纤维 ${totalNutrition.fiber}g · 钠 ${totalNutrition.sodium}mg
    </p>
  </body></html>`;
}

/**
 * 发送备菜邮件。返回 { sent, mode, info }。
 * - 已配置 SMTP：真实发送
 * - 未配置 SMTP：落盘到 data/outbox（开发/演示模式）
 */
export async function sendMealEmail({ order, allOrders }) {
  const to = process.env.NOTIFY_EMAIL;
  const message = buildMealEmail({ order, allOrders });
  const transport = createTransport();

  if (!transport || !to) {
    if (!existsSync(OUTBOX_DIR)) mkdirSync(OUTBOX_DIR, { recursive: true });
    const file = join(OUTBOX_DIR, `${order.id || `${order.date}-${order.meal}-${order.member}`}.html`);
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
