import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildMealEmail } from '../src/mailer.js';
import { loadRecipes } from '../src/recipes.js';
import { findOption } from '../src/mealGenerator.js';

const recipes = loadRecipes();

function makeOrder(member, label) {
  const option = findOption(recipes, '2026-06-06', 'dinner', label);
  return { member, date: '2026-06-06', meal: 'dinner', option, id: `x-${member}` };
}

test('邮件主题包含日期、餐次、下单人', () => {
  const order = makeOrder('爸爸', 'A');
  const { subject } = buildMealEmail({ order, allOrders: [order] });
  assert.match(subject, /2026-06-06/);
  assert.match(subject, /晚餐/);
  assert.match(subject, /爸爸/);
});

test('备菜清单汇总所有家人订单的食材', () => {
  const a = makeOrder('爸爸', 'A');
  const b = makeOrder('妈妈', 'A'); // 与 A 相同 → 食材分量翻倍
  const { text } = buildMealEmail({ order: a, allOrders: [a, b] });
  assert.match(text, /需要准备的食材与分量/);
  assert.match(text, /共 2 人/);
});

test('邮件正文包含营养合计', () => {
  const order = makeOrder('宝宝', 'B');
  const { text, html } = buildMealEmail({ order, allOrders: [order] });
  assert.match(text, /营养合计/);
  assert.match(html, /kcal/);
});
