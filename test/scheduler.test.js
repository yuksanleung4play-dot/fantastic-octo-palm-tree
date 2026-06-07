import { test } from 'node:test';
import assert from 'node:assert/strict';
import { nextSendTime, upcomingWeekRange } from '../src/scheduler.js';

const cfg = { dow: 6, hour: 17, minute: 0, offsetMin: 480 }; // 周六 17:00 UTC+8

test('nextSendTime：周日算出下一个周六 17:00 (UTC+8)', () => {
  const now = new Date('2026-06-07T00:00:00Z'); // UTC+8 为周日 08:00
  const next = nextSendTime(now, cfg);
  // 周六 2026-06-13 17:00 UTC+8 = 09:00 UTC
  assert.equal(next.toISOString(), '2026-06-13T09:00:00.000Z');
});

test('nextSendTime：当天已过发送点则顺延到下周', () => {
  const now = new Date('2026-06-13T10:00:00Z'); // 周六 18:00 UTC+8，已过 17:00
  const next = nextSendTime(now, cfg);
  assert.equal(next.toISOString(), '2026-06-20T09:00:00.000Z');
});

test('upcomingWeekRange：周日 → 即将到来的周一至周六', () => {
  const now = new Date('2026-06-07T00:00:00Z'); // UTC+8 周日
  const r = upcomingWeekRange(now, cfg);
  assert.deepEqual(r, { start: '2026-06-08', end: '2026-06-13' });
});

test('upcomingWeekRange：周六发送时 → 下一周周一至周六', () => {
  const now = new Date('2026-06-13T09:00:00Z'); // 周六 17:00 UTC+8
  const r = upcomingWeekRange(now, cfg);
  assert.deepEqual(r, { start: '2026-06-15', end: '2026-06-20' });
});
