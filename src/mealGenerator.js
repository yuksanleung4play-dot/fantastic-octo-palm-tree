import { splitByType, sumNutrition } from './recipes.js';

const OPTION_LABELS = ['A', 'B', 'C', 'D'];

/**
 * 32 位确定性伪随机数生成器（mulberry32）。
 */
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

/**
 * 合并多道菜的食材（按食材名 + 单位累加分量）。
 */
export function combineIngredients(dishes) {
  const map = new Map();
  for (const dish of dishes) {
    for (const ing of dish.ingredients || []) {
      const key = `${ing.name}__${ing.unit}`;
      const existing = map.get(key);
      if (existing) {
        existing.amount += ing.amount;
      } else {
        map.set(key, { name: ing.name, unit: ing.unit, amount: ing.amount });
      }
    }
  }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name, 'zh'));
}

/**
 * 为指定日期 + 餐次生成 A/B/C/D 四套套餐，每套一荤一素。
 * 同一 (date, meal) 始终生成相同结果（确定性）。
 *
 * @param {object} recipes 食谱库（含 dishes）
 * @param {string} dateStr 形如 "2026-06-06"
 * @param {string} meal "lunch" | "dinner"
 * @param {number} [count] 套餐数量，默认 4（A/B/C/D）
 */
export function generateMealOptions(recipes, dateStr, meal, count = 4) {
  const { meats, veggies } = splitByType(recipes.dishes);
  const rand = mulberry32(hashSeed(`${dateStr}|${meal}`));
  const shuffledMeats = seededShuffle(meats, rand);
  const shuffledVeggies = seededShuffle(veggies, rand);

  const options = [];
  for (let i = 0; i < count; i += 1) {
    const meat = shuffledMeats[i % shuffledMeats.length];
    const veg = shuffledVeggies[i % shuffledVeggies.length];
    const dishes = [meat, veg];
    options.push({
      label: OPTION_LABELS[i] || `选项${i + 1}`,
      dishes: dishes.map((d) => ({
        id: d.id,
        name: d.name,
        type: d.type,
        calories: d.calories,
        nutrition: d.nutrition,
        tags: d.tags || [],
      })),
      nutrition: sumNutrition(dishes),
      ingredients: combineIngredients(dishes),
    });
  }
  return options;
}

/**
 * 根据所选套餐 label 找回该套餐（用于下单时还原详情）。
 */
export function findOption(recipes, dateStr, meal, label) {
  return generateMealOptions(recipes, dateStr, meal).find((o) => o.label === label) || null;
}
