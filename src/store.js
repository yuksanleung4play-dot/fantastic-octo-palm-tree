import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(__dirname, '..', 'data');
const ORDERS_PATH = process.env.ORDERS_PATH || join(DATA_DIR, 'orders.json');

function ensureFile() {
  if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });
  if (!existsSync(ORDERS_PATH)) writeFileSync(ORDERS_PATH, '[]', 'utf-8');
}

export function readOrders() {
  ensureFile();
  try {
    return JSON.parse(readFileSync(ORDERS_PATH, 'utf-8'));
  } catch {
    return [];
  }
}

function writeOrders(orders) {
  ensureFile();
  writeFileSync(ORDERS_PATH, JSON.stringify(orders, null, 2), 'utf-8');
}

/**
 * 保存一条点餐记录。每天每餐最多一条：同一 (date, meal) 再次保存只覆盖为最新选择。
 */
export function saveOrder(order) {
  const orders = readOrders();
  const idx = orders.findIndex((o) => o.date === order.date && o.meal === order.meal);
  const now = new Date().toISOString();
  const record = {
    id: `${order.date}-${order.meal}`,
    createdAt: idx >= 0 ? orders[idx].createdAt : now,
    updatedAt: now,
    ...order,
  };
  record.id = `${order.date}-${order.meal}`;
  if (idx >= 0) orders[idx] = record;
  else orders.push(record);
  writeOrders(orders);
  return record;
}

/** 查询某一天的全部记录。 */
export function ordersForDay(date) {
  return readOrders().filter((o) => o.date === date);
}

/** 查询某 (date, meal) 记录。 */
export function ordersForMeal(date, meal) {
  return readOrders().filter((o) => o.date === date && o.meal === meal);
}

/** 查询某日期区间 [start, end]（含端点，YYYY-MM-DD 字符串比较）内的全部记录，按日期升序。 */
export function ordersForRange(start, end) {
  return readOrders()
    .filter((o) => o.date >= start && o.date <= end)
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}
