import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DEFAULT_PATH = join(__dirname, '..', 'data', 'recipes.json');

/**
 * 加载并校验食谱库（整餐食谱模型）。
 * 数据由 scripts/build-recipes.mjs 从 data/source-recipes.json 生成。
 * @param {string} [path] 食谱 JSON 路径，默认 data/recipes.json
 */
export function loadRecipes(path = process.env.RECIPES_PATH || DEFAULT_PATH) {
  const raw = readFileSync(path, 'utf-8');
  const data = JSON.parse(raw);
  validateRecipes(data);
  return data;
}

export function validateRecipes(data) {
  if (!data || !Array.isArray(data.meals)) {
    throw new Error('食谱库格式错误：缺少 meals 数组（请先运行 npm run build:recipes）');
  }
  const ids = new Set();
  for (const meal of data.meals) {
    if (!meal.id) throw new Error('存在缺少 id 的食谱');
    if (ids.has(meal.id)) throw new Error(`食谱 id 重复：${meal.id}`);
    ids.add(meal.id);
    if (!meal.title) throw new Error(`食谱缺少标题：${meal.id}`);
    if (!Array.isArray(meal.mealSlots) || meal.mealSlots.length === 0) {
      throw new Error(`食谱 ${meal.id} 缺少 mealSlots`);
    }
    if (!Array.isArray(meal.ingredients)) {
      throw new Error(`食谱 ${meal.id} 缺少 ingredients`);
    }
  }
  if (mealsForSlot(data.meals, 'lunch').length < 4) {
    throw new Error('可作午餐的完整食谱不足 4 道，无法生成 A/B/C/D');
  }
  if (mealsForSlot(data.meals, 'dinner').length < 4) {
    throw new Error('可作晚餐的完整食谱不足 4 道，无法生成 A/B/C/D');
  }
  return true;
}

/**
 * 返回适用于指定餐次（lunch/dinner）的「完整」食谱列表。
 * 完整 = 有食材数据，确保可生成备菜清单（自动跳过源数据缺失的条目）。
 */
export function mealsForSlot(meals, slot) {
  return meals.filter(
    (m) => m.mealSlots.includes(slot) && Array.isArray(m.ingredients) && m.ingredients.length > 0,
  );
}

const NUTRITION_KEYS = ['calories', 'protein', 'fat', 'carbs'];

/**
 * 汇总一组食谱（可带份数 _count 倍数）的营养。
 * null 营养按 0 计；返回 { calories, protein, fat, carbs, hasMissing }。
 * @param {Array<{nutrition:object, _count?:number}>} items
 */
export function sumNutrition(items) {
  const total = { calories: 0, protein: 0, fat: 0, carbs: 0 };
  let hasMissing = false;
  for (const item of items) {
    const n = item.nutrition || {};
    const count = item._count || 1;
    for (const key of NUTRITION_KEYS) {
      if (n[key] == null) {
        hasMissing = true;
      } else {
        total[key] += n[key] * count;
      }
    }
  }
  total.hasMissing = hasMissing;
  return total;
}
