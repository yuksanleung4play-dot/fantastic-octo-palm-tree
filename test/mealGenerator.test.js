import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadRecipes } from '../src/recipes.js';
import {
  generateMealOptions,
  findOption,
  combineIngredients,
  hashSeed,
  seededShuffle,
} from '../src/mealGenerator.js';

const recipes = loadRecipes();

test('生成 4 套 A/B/C/D 套餐', () => {
  const opts = generateMealOptions(recipes, '2026-06-06', 'dinner');
  assert.equal(opts.length, 4);
  assert.deepEqual(
    opts.map((o) => o.label),
    ['A', 'B', 'C', 'D'],
  );
});

test('每套套餐都是一荤一素', () => {
  const opts = generateMealOptions(recipes, '2026-06-06', 'lunch');
  for (const o of opts) {
    assert.equal(o.dishes.length, 2);
    const types = o.dishes.map((d) => d.type).sort();
    assert.deepEqual(types, ['meat', 'vegetable']);
  }
});

test('同一 date+meal 生成结果确定一致', () => {
  const a = generateMealOptions(recipes, '2026-06-07', 'dinner');
  const b = generateMealOptions(recipes, '2026-06-07', 'dinner');
  assert.deepEqual(a, b);
});

test('不同 meal 生成不同组合', () => {
  const lunch = generateMealOptions(recipes, '2026-06-06', 'lunch');
  const dinner = generateMealOptions(recipes, '2026-06-06', 'dinner');
  assert.notDeepEqual(
    lunch.map((o) => o.dishes.map((d) => d.id)),
    dinner.map((o) => o.dishes.map((d) => d.id)),
  );
});

test('四套套餐荤菜互不相同、素菜互不相同', () => {
  const opts = generateMealOptions(recipes, '2026-06-10', 'dinner');
  const meats = opts.map((o) => o.dishes.find((d) => d.type === 'meat').id);
  const veggies = opts.map((o) => o.dishes.find((d) => d.type === 'vegetable').id);
  assert.equal(new Set(meats).size, 4);
  assert.equal(new Set(veggies).size, 4);
});

test('套餐营养为两道菜之和', () => {
  const opt = generateMealOptions(recipes, '2026-06-06', 'dinner')[0];
  const sumCal = opt.dishes.reduce((s, d) => s + d.calories, 0);
  assert.equal(opt.nutrition.calories, sumCal);
});

test('findOption 能按 label 找回套餐', () => {
  const opt = findOption(recipes, '2026-06-06', 'dinner', 'C');
  assert.equal(opt.label, 'C');
});

test('combineIngredients 合并同名同单位食材', () => {
  const dishes = [
    { ingredients: [{ name: '食用油', amount: 10, unit: 'ml' }] },
    { ingredients: [{ name: '食用油', amount: 5, unit: 'ml' }] },
  ];
  const merged = combineIngredients(dishes);
  assert.equal(merged.length, 1);
  assert.equal(merged[0].amount, 15);
});

test('hashSeed 确定性、seededShuffle 不修改原数组', () => {
  assert.equal(hashSeed('abc'), hashSeed('abc'));
  const arr = [1, 2, 3, 4];
  const copy = [...arr];
  let i = 0;
  seededShuffle(arr, () => (i++ % 4) / 4);
  assert.deepEqual(arr, copy);
});
