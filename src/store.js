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
 * 保存一条点餐记录。不再按人区分，每次保存都新增一条记录。
 */
export function saveOrder(order) {
  const orders = readOrders();
  const ts = Date.now();
  const record = {
    id: `${order.date}-${order.meal}-${ts}-${Math.random().toString(36).slice(2, 6)}`,
    createdAt: new Date().toISOString(),
    ...order,
  };
  orders.push(record);
  writeOrders(orders);
  return record;
}

/** 查询某一天的全部记录（同一天内的记录都可见并一起发送）。 */
export function ordersForDay(date) {
  return readOrders().filter((o) => o.date === date);
}

/** 查询某 (date, meal) 下所有记录。 */
export function ordersForMeal(date, meal) {
  return readOrders().filter((o) => o.date === date && o.meal === meal);
}
