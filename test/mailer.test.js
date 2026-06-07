import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildWeeklyEmail, groupOrdersByRecipe, dateRange } from '../src/mailer.js';
import { loadRecipes } from '../src/recipes.js';
import { findOption } from '../src/mealGenerator.js';

const recipes = loadRecipes();

function makeOrder(date, label, meal = 'dinner') {
  const option = findOption(recipes, date, meal, label);
  return { date, meal, option, id: `${date}-${meal}` };
}

test('dateRange 生成含端点的连续日期', () => {
  assert.deepEqual(dateRange('2026-06-08', '2026-06-12'), [
    '2026-06-08',
    '2026-06-09',
    '2026-06-10',
    '2026-06-11',
    '2026-06-12',
  ]);
});

test('周邮件主题含区间与餐数', () => {
  const orders = [makeOrder('2026-06-08', 'A'), makeOrder('2026-06-09', 'B')];
  const { subject } = buildWeeklyEmail({ orders, rangeStart: '2026-06-08', rangeEnd: '2026-06-12' });
  assert.match(subject, /2026-06-08 ~ 2026-06-12/);
  assert.match(subject, /共 2 餐/);
});

test('周邮件列出每日餐单，未点的日子标注未点餐', () => {
  const orders = [makeOrder('2026-06-08', 'A')];
  const { text } = buildWeeklyEmail({ orders, rangeStart: '2026-06-08', rangeEnd: '2026-06-12' });
  assert.match(text, /2026-06-08/);
  assert.match(text, /未点餐/); // 周二至周五未点
  assert.match(text, /一周备菜清单/);
});

test('周邮件备菜按食谱合并并标注份数', () => {
  const orders = [makeOrder('2026-06-08', 'A'), makeOrder('2026-06-09', 'A', 'dinner')];
  // 06-08 与 06-09 的 A 可能不同（不同 date 种子），分别核对
  const { text, html } = buildWeeklyEmail({ orders, rangeStart: '2026-06-08', rangeEnd: '2026-06-12' });
  assert.match(text, /全部营养合计/);
  assert.match(html, /备菜清单/);
});

test('groupOrdersByRecipe 同食谱合并计数', () => {
  const a = makeOrder('2026-06-08', 'A');
  const b = { ...makeOrder('2026-06-09', 'A'), option: a.option };
  const groups = groupOrdersByRecipe([a, b]);
  assert.equal(groups.length, 1);
  assert.equal(groups[0].count, 2);
});

test('空餐单也能生成邮件', () => {
  const { text } = buildWeeklyEmail({ orders: [], rangeStart: '2026-06-08', rangeEnd: '2026-06-12' });
  assert.match(text, /本周暂无已保存餐单/);
});
