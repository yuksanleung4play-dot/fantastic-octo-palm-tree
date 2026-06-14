import { test } from 'node:test';
import assert from 'node:assert/strict';
import { isWeekend, mealsForDate, buildSchedule } from '../src/schedule.js';

test('周末判断正确（2026-06-06 周六, 06-07 周日）', () => {
  assert.equal(isWeekend('2026-06-06'), true);
  assert.equal(isWeekend('2026-06-07'), true);
  assert.equal(isWeekend('2026-06-08'), false); // 周一
});

test('平日只有晚餐，周末有午餐和晚餐', () => {
  assert.deepEqual(mealsForDate('2026-06-08'), ['dinner']);
  assert.deepEqual(mealsForDate('2026-06-06'), ['lunch', 'dinner']);
});

test('buildSchedule 生成连续日期', () => {
  const s = buildSchedule('2026-06-08', 3);
  assert.equal(s.length, 3);
  assert.equal(s[0].date, '2026-06-08');
  assert.equal(s[2].date, '2026-06-10');
  assert.ok(s[0].weekday.startsWith('週'));
});
