import './env.js';
import express from 'express';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { loadRecipes } from './recipes.js';
import { buildSchedule, mealsForDate, MEAL_LABELS, toDateStr } from './schedule.js';
import { generateMealOptions, findOption } from './mealGenerator.js';
import { saveOrder, ordersForDay } from './store.js';
import { startScheduler, runWeeklySend, nextSendTime, upcomingWeekRange } from './scheduler.js';

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

/** 保存一条记录（每天每餐最多一条，重复保存只留最新）。不再即时发邮件——统一周六发送。 */
app.post('/api/orders', (req, res) => {
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
  res.status(201).json({ order: record, allOrders, nextEmail: nextSendTime(new Date()).toISOString() });
});

/** 手动触发本周餐单邮件（便于测试）。可选 body.start 指定某个周一。 */
app.post('/api/send-weekly', async (req, res) => {
  try {
    const result = await runWeeklySend(new Date(), (req.body && req.body.start) || null);
    res.json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/** 查看下一次发送时间与本周（即将到来的周一至周六）区间。 */
app.get('/api/weekly-info', (req, res) => {
  res.json({ nextSend: nextSendTime(new Date()).toISOString(), range: upcomingWeekRange(new Date()) });
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
    startScheduler();
  });
}

export { app, recipes };
