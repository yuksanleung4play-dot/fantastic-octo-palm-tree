import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildMealEmail, groupOrdersByRecipe } from '../src/mailer.js';
import { loadRecipes } from '../src/recipes.js';
import { findOption } from '../src/mealGenerator.js';

const recipes = loadRecipes();

function makeOrder(member, label) {
  const option = findOption(recipes, '2026-06-06', 'dinner', label);
  return { member, date: '2026-06-06', meal: 'dinner', option, id: `x-${member}` };
}

test('邮件主题包含日期、餐次、下单人与食谱名', () => {
  const order = makeOrder('爸爸', 'A');
  const { subject } = buildMealEmail({ order, allOrders: [order] });
  assert.match(subject, /2026-06-06/);
  assert.match(subject, /晚餐/);
  assert.match(subject, /爸爸/);
  assert.match(subject, new RegExp(order.option.title.slice(0, 4)));
});

test('groupOrdersByRecipe 按食谱聚合并计数', () => {
  const a = makeOrder('爸爸', 'A');
  const b = makeOrder('妈妈', 'A'); // 同一食谱 → 合并，count=2
  const c = makeOrder('宝宝', 'B'); // 不同食谱
  const groups = groupOrdersByRecipe([a, b, c]);
  assert.equal(groups.length, 2);
  const ga = groups.find((g) => g.option.recipeId === a.option.recipeId);
  assert.equal(ga.count, 2);
  assert.deepEqual(ga.members.sort(), ['妈妈', '爸爸']);
});

test('备菜清单含份数标注与营养合计', () => {
  const a = makeOrder('爸爸', 'A');
  const b = makeOrder('妈妈', 'A');
  const { text } = buildMealEmail({ order: a, allOrders: [a, b] });
  assert.match(text, /需要准备的食材与分量/);
  assert.match(text, /×2 份/);
  assert.match(text, /全部营养合计/);
});

test('邮件 HTML 含食谱标题', () => {
  const order = makeOrder('宝宝', 'B');
  const { html } = buildMealEmail({ order, allOrders: [order] });
  assert.match(html, /备菜清单/);
  assert.ok(html.includes(order.option.title.slice(0, 3)));
});
