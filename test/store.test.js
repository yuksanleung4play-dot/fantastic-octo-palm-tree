import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TMP = join(__dirname, '..', 'data', `orders.test.${process.pid}.json`);
process.env.ORDERS_PATH = TMP;

const store = await import('../src/store.js');

before(() => {
  if (existsSync(TMP)) rmSync(TMP);
});
after(() => {
  if (existsSync(TMP)) rmSync(TMP);
});

const opt = (id) => ({ recipeId: id, label: 'A', title: `菜-${id}`, nutrition: { calories: 1 }, ingredientGroups: [] });

test('同一 (date, meal) 重复保存只留最新', () => {
  store.saveOrder({ date: '2026-06-08', meal: 'dinner', option: opt('r1') });
  store.saveOrder({ date: '2026-06-08', meal: 'dinner', option: opt('r2') });
  const day = store.ordersForDay('2026-06-08');
  assert.equal(day.length, 1);
  assert.equal(day[0].option.recipeId, 'r2');
  assert.equal(day[0].id, '2026-06-08-dinner');
});

test('周末午餐与晚餐分别各保留一条', () => {
  store.saveOrder({ date: '2026-06-13', meal: 'lunch', option: opt('L') });
  store.saveOrder({ date: '2026-06-13', meal: 'dinner', option: opt('D') });
  store.saveOrder({ date: '2026-06-13', meal: 'lunch', option: opt('L2') }); // 覆盖 lunch
  const day = store.ordersForDay('2026-06-13');
  assert.equal(day.length, 2);
  const lunch = day.find((o) => o.meal === 'lunch');
  assert.equal(lunch.option.recipeId, 'L2');
});

test('ordersForRange 返回区间内并按日期升序', () => {
  store.saveOrder({ date: '2026-06-10', meal: 'dinner', option: opt('w') });
  store.saveOrder({ date: '2026-06-09', meal: 'dinner', option: opt('v') });
  const r = store.ordersForRange('2026-06-08', '2026-06-12');
  const dates = r.map((o) => o.date);
  assert.deepEqual(dates, [...dates].sort());
  assert.ok(dates.includes('2026-06-08'));
  assert.ok(!dates.includes('2026-06-13'));
});
