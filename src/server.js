import './env.js';
import express from 'express';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { loadRecipes } from './recipes.js';
import { buildSchedule, mealsForDate, MEAL_LABELS, toDateStr } from './schedule.js';
import { generateMealOptions, findOption } from './mealGenerator.js';
import { saveOrder, ordersForDay } from './store.js';
import { sendMealEmail } from './mailer.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json());
app.use(express.static(join(__dirname, '..', 'public')));

let recipes = loadRecipes();

/** 食谱库元信息与全部食谱。 */
app.get('/api/recipes', (req, res) => {
  res.json(recipes);
});

/** 未来若干天的排程（平日一餐 / 周末午晚餐）。 */
app.get('/api/schedule', (req, res) => {
  const start = req.query.start || toDateStr(new Date());
  const days = Math.min(Number(req.query.days || 7), 30);
  res.json({ start, days, schedule: buildSchedule(start, days) });
});

/** 某日某餐的 A/B/C/D 选项；existingOrders 返回当天全部记录。 */
app.get('/api/meals', (req, res) => {
  const { date, meal } = req.query;
  if (!date || !meal) return res.status(400).json({ error: '缺少 date 或 meal 参数' });
  if (!['lunch', 'dinner'].includes(meal)) {
    return res.status(400).json({ error: 'meal 必须是 lunch 或 dinner' });
  }
  if (!mealsForDate(date).includes(meal)) {
    return res.status(400).json({ error: `${date} 不安排${MEAL_LABELS[meal]}（平日仅晚餐）` });
  }
  res.json({
    date,
    meal,
    mealLabel: MEAL_LABELS[meal],
    options: generateMealOptions(recipes, date, meal),
    existingOrders: ordersForDay(date),
  });
});

/** 保存一条记录并发送当天备菜邮件（不区分点餐人）。 */
app.post('/api/orders', async (req, res) => {
  const { date, meal, optionLabel } = req.body || {};
  if (!date || !meal || !optionLabel) {
    return res.status(400).json({ error: '缺少 date / meal / optionLabel' });
  }
  if (!mealsForDate(date).includes(meal)) {
    return res.status(400).json({ error: `${date} 不安排该餐次` });
  }
  const option = findOption(recipes, date, meal, optionLabel);
  if (!option) return res.status(404).json({ error: `未找到套餐 ${optionLabel}` });

  const record = saveOrder({ date, meal, option });
  const allOrders = ordersForDay(date);

  let mail;
  try {
    mail = await sendMealEmail({ order: record, allOrders });
  } catch (err) {
    mail = { sent: false, mode: 'error', error: err.message };
  }

  res.status(201).json({ order: record, allOrders, mail });
});

/** 查看某一天的全部记录。 */
app.get('/api/orders', (req, res) => {
  const { date } = req.query;
  if (!date) return res.status(400).json({ error: '缺少 date' });
  res.json({ orders: ordersForDay(date) });
});

const PORT = process.env.PORT || 3000;
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`🍽️  家庭点餐系统已启动： http://localhost:${PORT}`);
  });
}

export { app, recipes };
