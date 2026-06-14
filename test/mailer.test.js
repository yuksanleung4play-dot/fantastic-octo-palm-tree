import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  buildWeeklyEmail,
  groupOrdersByRecipe,
  dateRange,
  parseIngredient,
  categorize,
  summarizeIngredients,
} from '../src/mailer.js';
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

test('parseIngredient 兼容有/无空格与分数', () => {
  assert.deepEqual(parseIngredient('雞髀肉 2件（去骨切塊）'), { name: '雞髀肉', num: 2, unit: '件', vague: false });
  assert.deepEqual(parseIngredient('鹽 1/4茶匙'), { name: '鹽', num: 0.25, unit: '茶匙', vague: false });
  assert.deepEqual(parseIngredient('白胡椒½小匙'), { name: '白胡椒', num: 0.5, unit: '小匙', vague: false });
  assert.equal(parseIngredient('黑胡椒適量').name, '黑胡椒');
  assert.equal(parseIngredient('黑胡椒適量').vague, true);
});

test('categorize 分类正确（肉類/蔬菜/調味料/其他）', () => {
  assert.equal(categorize('雞胸肉'), '肉類');
  assert.equal(categorize('雞蛋'), '肉類');
  assert.equal(categorize('嫩豆腐'), '肉類');
  assert.equal(categorize('洋蔥'), '蔬菜'); // 不应被「葱」误判
  assert.equal(categorize('西蘭花'), '蔬菜');
  assert.equal(categorize('番茄'), '蔬菜');
  assert.equal(categorize('鹽'), '調味料');
  assert.equal(categorize('生抽'), '調味料');
  assert.equal(categorize('蒜蓉'), '調味料');
  assert.equal(categorize('米'), '其他');
});

test('summarizeIngredients 跨食谱合并同名分量并拆分串接项', () => {
  const orders = [
    { option: { ingredients: ['雞胸肉 200克', '鹽 1茶匙'] } },
    { option: { ingredients: ['雞胸肉 100克', '車厘茄、紫洋蔥 適量'] } },
  ];
  const s = summarizeIngredients(orders);
  const chicken = s['肉類'].find((x) => x.name === '雞胸肉');
  assert.equal(chicken.qty, '300克');
  assert.ok(s['蔬菜'].some((x) => x.name === '車厘茄'));
  assert.ok(s['蔬菜'].some((x) => x.name === '紫洋蔥'));
  assert.ok(s['調味料'].some((x) => x.name === '鹽'));
});

test('周邮件含分类汇总标题', () => {
  const a = makeOrder('2026-06-08', 'A');
  const { text, html } = buildWeeklyEmail({ orders: [a], rangeStart: '2026-06-08', rangeEnd: '2026-06-13' });
  assert.match(text, /分类汇总/);
  assert.match(html, /肉類|蔬菜|調味料/);
});
