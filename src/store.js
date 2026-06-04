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
 * 保存一条订单。若同一家人对同一 (date, meal) 已下单，则覆盖（改单）。
 */
export function saveOrder(order) {
  const orders = readOrders();
  const idx = orders.findIndex(
    (o) => o.member === order.member && o.date === order.date && o.meal === order.meal,
  );
  const record = {
    id: `${order.date}-${order.meal}-${order.member}`,
    createdAt: new Date().toISOString(),
    ...order,
  };
  if (idx >= 0) {
    orders[idx] = { ...orders[idx], ...record, updatedAt: new Date().toISOString() };
  } else {
    orders.push(record);
  }
  writeOrders(orders);
  return record;
}

/** 查询某 (date, meal) 下所有订单。 */
export function ordersForMeal(date, meal) {
  return readOrders().filter((o) => o.date === date && o.meal === meal);
}
