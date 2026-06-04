import { mealsForSlot } from './recipes.js';

const OPTION_LABELS = ['A', 'B', 'C', 'D'];

/** 32 位确定性伪随机数生成器（mulberry32）。 */
function mulberry32(seed) {
  let a = seed >>> 0;
  return function next() {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** 将字符串散列为 32 位整数，作为随机种子。 */
export function hashSeed(str) {
  let h = 2166136261 >>> 0;
  for (let i = 0; i < str.length; i += 1) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

/** 基于种子的 Fisher–Yates 洗牌，返回新数组，不修改原数组。 */
export function seededShuffle(arr, rand) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i -= 1) {
    const j = Math.floor(rand() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/** 将一道食谱转换为下单/展示用的精简对象。 */
function toOption(meal, label) {
  return {
    label,
    recipeId: meal.id,
    no: meal.no,
    title: meal.title,
    source: meal.source,
    sourceUrl: meal.sourceUrl,
    mealType: meal.mealType,
    servings: meal.servings,
    nutrition: meal.nutrition,
    ingredientGroups: meal.ingredientGroups,
    ingredients: meal.ingredients,
    steps: meal.steps,
  };
}

/**
 * 为指定日期 + 餐次生成 A/B/C/D 四套餐（各为一道完整食谱）。
 * 同一 (date, meal) 始终生成相同结果（确定性），且四道互不相同。
 *
 * @param {object} recipes 食谱库（含 meals）
 * @param {string} dateStr 形如 "2026-06-06"
 * @param {string} meal "lunch" | "dinner"
 * @param {number} [count] 套餐数量，默认 4（A/B/C/D）
 */
export function generateMealOptions(recipes, dateStr, meal, count = 4) {
  const pool = mealsForSlot(recipes.meals, meal);
  const rand = mulberry32(hashSeed(`${dateStr}|${meal}`));
  const shuffled = seededShuffle(pool, rand);
  const picked = shuffled.slice(0, count);
  return picked.map((m, i) => toOption(m, OPTION_LABELS[i] || `选项${i + 1}`));
}

/** 根据所选 label 找回该餐（用于下单时还原详情）。 */
export function findOption(recipes, dateStr, meal, label) {
  return generateMealOptions(recipes, dateStr, meal).find((o) => o.label === label) || null;
}
