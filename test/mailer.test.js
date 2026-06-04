import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildMealEmail, groupOrdersByRecipe } from '../src/mailer.js';
import { loadRecipes } from '../src/recipes.js';
import { findOption } from '../src/mealGenerator.js';

const recipes = loadRecipes();

function makeOrder(label, meal = 'dinner') {
  const option = findOption(recipes, '2026-06-06', meal, label);
  return { date: '2026-06-06', meal, option, id: `x-${label}-${meal}-${Math.random()}` };
}

test('邮件主题含日期、本次餐次与食谱名，并标注当天笔数', () => {
  const order = makeOrder('A');
  const { subject } = buildMealEmail({ order, allOrders: [order] });
  assert.match(subject, /2026-06-06/);
  assert.match(subject, /晚餐/);
  assert.match(subject, /当天共 1 笔/);
  assert.match(subject, new RegExp(order.option.title.slice(0, 4)));
});

test('groupOrdersByRecipe 按食谱聚合并计数、记录餐次（不区分点餐人）', () => {
  const a = makeOrder('A');
  const b = makeOrder('A'); // 同一食谱 → 合并，count=2
  const c = makeOrder('B');
  const groups = groupOrdersByRecipe([a, b, c]);
  assert.equal(groups.length, 2);
  const ga = groups.find((g) => g.option.recipeId === a.option.recipeId);
  assert.equal(ga.count, 2);
  assert.deepEqual(ga.meals, ['晚餐']);
});

test('当天记录跨午晚餐一起聚合', () => {
  const lunch = makeOrder('A', 'lunch');
  const dinner = makeOrder('A', 'dinner');
  const { text } = buildMealEmail({ order: dinner, allOrders: [lunch, dinner] });
  assert.match(text, /当天已保存记录（共 2 笔）/);
  assert.match(text, /午餐/);
  assert.match(text, /晚餐/);
});

test('备菜清单含份数标注与营养合计', () => {
  const a = makeOrder('A');
  const b = makeOrder('A');
  const { text } = buildMealEmail({ order: a, allOrders: [a, b] });
  assert.match(text, /需要准备的食材与分量/);
  assert.match(text, /×2 份/);
  assert.match(text, /全部营养合计/);
});

test('邮件 HTML 含备菜清单与食谱标题', () => {
  const order = makeOrder('B');
  const { html } = buildMealEmail({ order, allOrders: [order] });
  assert.match(html, /备菜清单/);
  assert.ok(html.includes(order.option.title.slice(0, 3)));
});
