import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadRecipes, mealsForSlot, sumNutrition } from '../src/recipes.js';
import { generateMealOptions, findOption, hashSeed, seededShuffle } from '../src/mealGenerator.js';

const recipes = loadRecipes();

test('食谱库为整餐模型且至少 4 道可作午/晚餐', () => {
  assert.ok(Array.isArray(recipes.meals));
  assert.ok(recipes.meals.length >= 4);
  assert.ok(mealsForSlot(recipes.meals, 'lunch').length >= 4);
  assert.ok(mealsForSlot(recipes.meals, 'dinner').length >= 4);
});

test('生成 4 套 A/B/C/D 餐', () => {
  const opts = generateMealOptions(recipes, '2026-06-06', 'dinner');
  assert.equal(opts.length, 4);
  assert.deepEqual(opts.map((o) => o.label), ['A', 'B', 'C', 'D']);
});

test('每套餐为一道完整食谱（含标题与食材）', () => {
  for (const o of generateMealOptions(recipes, '2026-06-06', 'lunch')) {
    assert.ok(o.title);
    assert.ok(o.recipeId);
    assert.ok(Array.isArray(o.ingredients) && o.ingredients.length > 0);
  }
});

test('四套餐互不相同', () => {
  const ids = generateMealOptions(recipes, '2026-06-10', 'dinner').map((o) => o.recipeId);
  assert.equal(new Set(ids).size, 4);
});

test('所选餐次的食谱都支持该餐次', () => {
  const lunch = generateMealOptions(recipes, '2026-06-06', 'lunch');
  for (const o of lunch) {
    const meal = recipes.meals.find((m) => m.id === o.recipeId);
    assert.ok(meal.mealSlots.includes('lunch'));
  }
});

test('同一 date+meal 生成结果确定一致', () => {
  const a = generateMealOptions(recipes, '2026-06-07', 'dinner');
  const b = generateMealOptions(recipes, '2026-06-07', 'dinner');
  assert.deepEqual(a, b);
});

test('不同 meal 生成不同组合', () => {
  const lunch = generateMealOptions(recipes, '2026-06-06', 'lunch').map((o) => o.recipeId);
  const dinner = generateMealOptions(recipes, '2026-06-06', 'dinner').map((o) => o.recipeId);
  assert.notDeepEqual(lunch, dinner);
});

test('findOption 能按 label 找回餐', () => {
  const opt = findOption(recipes, '2026-06-06', 'dinner', 'C');
  assert.equal(opt.label, 'C');
});

test('sumNutrition 按份数累加，缺失项标记 hasMissing', () => {
  const t = sumNutrition([
    { nutrition: { calories: 100, protein: 10, fat: 5, carbs: 8 }, _count: 2 },
    { nutrition: { calories: 200, protein: 20, fat: 10, carbs: 16 }, _count: 1 },
  ]);
  assert.equal(t.calories, 400);
  assert.equal(t.protein, 40);
  assert.equal(t.hasMissing, false);
  const t2 = sumNutrition([{ nutrition: { calories: null, protein: null, fat: null, carbs: null } }]);
  assert.equal(t2.hasMissing, true);
});

test('hashSeed 确定性、seededShuffle 不修改原数组', () => {
  assert.equal(hashSeed('abc'), hashSeed('abc'));
  const arr = [1, 2, 3, 4];
  const copy = [...arr];
  let i = 0;
  seededShuffle(arr, () => (i++ % 4) / 4);
  assert.deepEqual(arr, copy);
});
