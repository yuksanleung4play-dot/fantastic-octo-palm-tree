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

/* ============== 备菜清单：解析 / 分类 / 汇总 ============== */

const FRACTIONS = { '½': 0.5, '¼': 0.25, '⅓': 1 / 3, '⅔': 2 / 3, '¾': 0.75, '⅛': 0.125 };

/** 解析单条食材文本为 { name, num, unit, vague }。兼容有/无空格："雞髀肉 2件（去骨切塊）"、"黑胡椒適量"、"鹽¼小匙" */
export function parseIngredient(text) {
  const raw = String(text).replace(/[（(].*?[）)]/g, '').trim();
  if (!raw) return { name: '', num: null, unit: '', vague: false };
  const pre = raw.match(/^(大量|適量|适量|少許|少许|少量|隨意|随意)(.+)/);
  if (pre) return { name: pre[2].trim(), num: null, unit: pre[1], vague: true };

  // 找到「分量」起始位置：空格 / 数字 / 分数 / 半 / 適量少許
  const m = raw.match(/[ \t0-9０-９½¼⅓⅔¾⅛]|半|適量|适量|少許|少许|少量|隨意|随意/);
  if (!m || m.index === 0) return { name: raw, num: null, unit: '', vague: false };
  const name = raw.slice(0, m.index).trim();
  let qty = raw.slice(m.index).trim();
  if (!name) return { name: raw, num: null, unit: '', vague: false };
  if (!qty) return { name, num: null, unit: '', vague: false };
  if (/^(適量|适量|少許|少许|少量|隨意|随意|酌量)/.test(qty)) {
    return { name, num: null, unit: qty, vague: true };
  }
  let num = null;
  let rest = qty;
  let mm;
  if ((mm = qty.match(/^(\d+(?:\.\d+)?)\s*\/\s*(\d+)/))) {
    num = parseFloat(mm[1]) / parseFloat(mm[2]);
    rest = qty.slice(mm[0].length);
  } else if ((mm = qty.match(/^(\d+(?:\.\d+)?)/))) {
    num = parseFloat(mm[1]);
    rest = qty.slice(mm[0].length);
  } else if (qty[0] in FRACTIONS) {
    num = FRACTIONS[qty[0]];
    rest = qty.slice(1);
  } else if (qty.startsWith('半')) {
    num = 0.5;
    rest = qty.slice(1);
  }
  return { name, num, unit: rest.trim(), vague: false };
}

const OILS = ['橄欖油', '菜籽油', '花生油', '麻油', '牛油', '香油', '粟米油', '牛油果油', '植物油', '芥花油', '辣椒油', '沙律油'];

// 按优先级：调味料 → 蔬菜 → 肉类（含蛋豆/海鲜）→ 其他
const SEASONING = [
  '鹽', '盐', '糖', '冰糖', '糖漿', '楓糖', '羅漢果', '醬油', '酱油', '生抽', '老抽', '蠔油', '蚝油', '魚露', '鱼露', '豉油',
  '醋', '酒', '味醂', '味噌', '大醬', '豆瓣', '辣椒醬', '辣醬', '蝦醬', 'XO醬', '沙茶', '沙嗲',
  '茄膏', '茄汁', '番茄醬', '咖喱', '咖哩', '五香粉', '十三香', '胡椒', '黑椒', '花椒', '八角', '桂皮', '香葉', '月桂',
  '香料', '孜然', '辣椒粉', '紅椒粉', '煙燻紅椒', '奧勒岡', '義大利香料', '香草',
  '生粉', '澱粉', '淀粉', '太白粉', '粟粉', '梳打粉', '蜂蜜', '芥末', '柚子胡椒',
  '燒汁', '滷水', '昆布', '高湯', '上湯', '雞湯', '蔬菜高湯', '豬骨高湯', '蒸魚豉油',
  '芝麻', '枸杞', '紅棗', '薑黃', '迷迭香', '百里香', '香茅', '九層塔', '羅勒', '芫荽', '香菜',
  '蔥', '葱', '蒜', '薑', '姜', '乾葱', '紅蔥', '乾蔥', '檸檬汁', '青檸汁', '橙汁', '柚子醋', '魚露檸檬',
  '沙律醬', '美乃滋', '蛋黃醬', '千島醬', '水',
];
const VEG = [
  '菜', '瓜', '椒', '茄', '蘿蔔', '萝卜', '甘筍', '紅蘿蔔', '洋蔥', '洋葱', '薯', '蕃薯', '番薯', '地瓜', '南瓜',
  '粟米', '玉米', '蘆筍', '芦笋', '西蘭花', '西兰花', '花椰菜', '椰菜', '西芹', '芹', '生菜', '菠菜', '韭', '芽',
  '蓮藕', '藕', '竹筍', '筍', '牛蒡', '秋葵', '百合', '菇', '菌', '耳', '松露', '海帶', '紫菜', '海苔',
  '酪梨', '牛油果', '蘋果', '芒果', '橙', '檸檬', '青檸', '番茄', '蕃茄', '彩椒', '甜椒', '青椒', '紅椒', '黃椒', '尖椒',
];
const MEAT = [
  '雞', '鸡', '豬', '猪', '牛', '羊', '魚', '鱼', '蝦', '虾', '蟹', '帶子', '元貝', '瑤柱', '蛤', '蜆', '蜊', '魷', '鱿',
  '中卷', '透抽', '吞拿', '鮭', '三文魚', '鯖', '鱸', '比目', '鯛', '龍利', '海鮮', '蛋', '豆腐', '豆乾', '豆干', '豆包',
  '腐皮', '腐竹', '天貝', '肉', '骰子', '扒', '排', '柳', '展', '腱', '免治', '絞',
];

export function categorize(name) {
  const n = name || '';
  const has = (list) => list.some((k) => n.includes(k));
  if (/洋蔥|洋葱/.test(n)) return '蔬菜'; // 洋蔥是蔬菜（避免被「葱」误判为调味料）
  if (n === '油' || OILS.some((o) => n.includes(o))) return '調味料';
  if (has(SEASONING)) return '調味料';
  if (has(VEG)) return '蔬菜';
  if (has(MEAT)) return '肉類';
  return '其他';
}

const CAT_ORDER = ['肉類', '蔬菜', '調味料', '其他'];

/** 汇总所有订单的食材，按分类合并。返回 { 肉類:[{name,qty}], 蔬菜:[...], ... }。 */
export function summarizeIngredients(orders) {
  // name -> { units: Map(unit->sum), vague:bool, cat }
  const map = new Map();
  for (const o of orders) {
    // 部分食材项可能用「、／，／＋」串接多种食材，先拆开
    const items = (o.option.ingredients || []).flatMap((it) => String(it).split(/[、，＋]/));
    for (const item of items) {
      const { name, num, unit, vague } = parseIngredient(item);
      if (!name) continue;
      if (!map.has(name)) map.set(name, { units: new Map(), vague: false, cat: categorize(name) });
      const rec = map.get(name);
      if (num != null) {
        rec.units.set(unit, (rec.units.get(unit) || 0) + num);
      } else if (vague || unit) {
        rec.vague = true;
      }
    }
  }
  const round = (x) => (Number.isInteger(x) ? x : Math.round(x * 100) / 100);
  const out = {};
  for (const cat of CAT_ORDER) out[cat] = [];
  for (const [name, rec] of map) {
    const parts = [...rec.units.entries()]
      .filter(([, v]) => v > 0)
      .map(([unit, v]) => `${round(v)}${unit}`);
    if (rec.vague) parts.push('適量');
    out[rec.cat].push({ name, qty: parts.join('、') });
  }
  for (const cat of CAT_ORDER) out[cat].sort((a, b) => a.name.localeCompare(b.name, 'zh-Hant'));
  return out;
}

/* ============== 邮件构建 ============== */

/**
 * 构建「本周餐单」邮件：覆盖 rangeStart~rangeEnd（通常为周一至周六）。
 * 含每日已点餐单 + 整周备菜清单（分类汇总：肉類 / 蔬菜 / 調味料 / 其他）+ 营养合计。
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
  const summary = summarizeIngredients(orders);

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

  const summaryText = CAT_ORDER.filter((c) => summary[c].length)
    .map((c) => {
      const items = summary[c].map((i) => `      - ${i.name}${i.qty ? `：${i.qty}` : ''}`).join('\n');
      return `  【${c}】\n${items}`;
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
      ? `🛒 一周备菜清单（分类汇总）：\n${summaryText}\n\n全部营养合计：${nutriLine(totalNutrition)}`
      : `本周暂无已保存餐单。`,
  ]
    .filter((l) => l !== '')
    .join('\n');

  const html = renderHtml({ rangeStart, rangeEnd, days, byDate, summary, totalNutrition, count: orders.length });
  return { subject, text, html };
}

function renderHtml({ rangeStart, rangeEnd, days, byDate, summary, totalNutrition, count }) {
  const dayRows = days
    .map((date) => {
      const list = byDate.get(date) || [];
      const cell = list.length
        ? list.map((o) => `${MEAL_LABEL[o.meal] || o.meal} ${o.option.label}餐 — ${esc(o.option.title)}`).join('<br>')
        : '<span style="color:#9ca3af;">未点餐</span>';
      return `<tr><td style="white-space:nowrap;">${date}<br><span style="color:#6b7280;">${weekdayOf(date)}</span></td><td>${cell}</td></tr>`;
    })
    .join('');

  const catColor = { 肉類: '#b91c1c', 蔬菜: '#16a34a', 調味料: '#b45309', 其他: '#6b7280' };
  const summaryHtml = CAT_ORDER.filter((c) => summary[c].length)
    .map((c) => {
      const items = summary[c]
        .map((i) => `<li>${esc(i.name)}${i.qty ? ` <b style="color:#ea580c;">${esc(i.qty)}</b>` : ''}</li>`)
        .join('');
      return `<div style="border:1px solid #e5e7eb;border-radius:10px;padding:10px 14px;margin-bottom:10px;">
        <div style="font-weight:800;color:${catColor[c]};margin-bottom:4px;">${c}</div>
        <ul style="margin:4px 0;padding-left:20px;columns:2;-webkit-columns:2;">${items}</ul>
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
        ? `<h3 style="border-bottom:2px solid #bbf7d0;padding-bottom:6px;">🛒 一周备菜清单（分类汇总）</h3>
    ${summaryHtml}
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
