import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DEFAULT_PATH = join(__dirname, '..', 'data', 'recipes.json');

/**
 * 加载并校验食谱库。
 * @param {string} [path] 食谱 JSON 路径，默认 data/recipes.json
 */
export function loadRecipes(path = process.env.RECIPES_PATH || DEFAULT_PATH) {
  const raw = readFileSync(path, 'utf-8');
  const data = JSON.parse(raw);
  validateRecipes(data);
  return data;
}

export function validateRecipes(data) {
  if (!data || !Array.isArray(data.dishes)) {
    throw new Error('食谱库格式错误：缺少 dishes 数组');
  }
  const ids = new Set();
  for (const dish of data.dishes) {
    if (!dish.id) throw new Error('存在缺少 id 的菜品');
    if (ids.has(dish.id)) throw new Error(`菜品 id 重复：${dish.id}`);
    ids.add(dish.id);
    if (!dish.name) throw new Error(`菜品缺少名称：${dish.id}`);
    if (!['meat', 'vegetable'].includes(dish.type)) {
      throw new Error(`菜品 ${dish.id} 的 type 必须是 meat 或 vegetable`);
    }
    if (typeof dish.calories !== 'number') {
      throw new Error(`菜品 ${dish.id} 缺少 calories`);
    }
    if (!Array.isArray(dish.ingredients)) {
      throw new Error(`菜品 ${dish.id} 缺少 ingredients`);
    }
  }
  const meats = data.dishes.filter((d) => d.type === 'meat');
  const veggies = data.dishes.filter((d) => d.type === 'vegetable');
  if (meats.length === 0 || veggies.length === 0) {
    throw new Error('食谱库需同时包含荤菜(meat)与素菜(vegetable)');
  }
  return true;
}

export function splitByType(dishes) {
  return {
    meats: dishes.filter((d) => d.type === 'meat'),
    veggies: dishes.filter((d) => d.type === 'vegetable'),
  };
}

const NUTRITION_KEYS = ['protein', 'fat', 'carbs', 'fiber', 'sodium'];

/**
 * 汇总一组菜品的总热量与营养。
 */
export function sumNutrition(dishes) {
  const total = {
    calories: 0,
    protein: 0,
    fat: 0,
    carbs: 0,
    fiber: 0,
    sodium: 0,
  };
  for (const dish of dishes) {
    total.calories += dish.calories || 0;
    const n = dish.nutrition || {};
    for (const key of NUTRITION_KEYS) {
      total[key] += n[key] || 0;
    }
  }
  return total;
}
