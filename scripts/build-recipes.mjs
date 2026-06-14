#!/usr/bin/env node
/**
 * 重建系统使用的食谱库 data/recipes.json。
 *
 * 来源（均为「已结构化」的 {meta, meals}）：
 *   - data/recipes-base.json  整合基底库（家庭食谱 + 香港营养师协会主菜，精选合并）
 *   - data/lib-shipu.json     港式精选食谱库（由 scripts/gen-shipu.py 从 Shi-Pu 清单整理）
 *
 * 图片来自 data/recipe-images.json（由 scripts/fetch-images.mjs 抓取或 AI 生成），按 id 套用。
 *
 * 运行： npm run build:recipes
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA = join(__dirname, '..', 'data');
const OUT = join(DATA, 'recipes.json');
const IMAGES = join(DATA, 'recipe-images.json');

// 食谱来源（按顺序合并，按 id 去重）
const LIBRARIES = [
  { path: join(DATA, 'recipes-base.json'), name: '整合基底库' },
  { path: join(DATA, 'lib-shipu.json'), name: '港式精选食谱库' },
];

/** 套用图片缓存并补全字段（扁平 ingredients / no）。 */
function normalize(meal, images, index) {
  const e = images[meal.id];
  const groups = meal.ingredientGroups || [];
  const flat =
    meal.ingredients && meal.ingredients.length ? meal.ingredients : groups.flatMap((g) => g.items || []);
  return {
    ...meal,
    no: meal.no ?? index + 1,
    image: e && e.file ? `dish-images/${e.file}` : meal.image || null,
    imageSource: e && e.src ? e.src : meal.imageSource || null,
    ingredients: flat,
  };
}

function main() {
  const images = existsSync(IMAGES) ? JSON.parse(readFileSync(IMAGES, 'utf-8')) : {};
  const meals = [];
  const seen = new Set();
  const usedLibs = [];

  for (const lib of LIBRARIES) {
    if (!existsSync(lib.path)) continue;
    const data = JSON.parse(readFileSync(lib.path, 'utf-8'));
    let added = 0;
    for (const meal of data.meals || []) {
      if (!meal.id || seen.has(meal.id)) continue;
      meals.push(normalize(meal, images, meals.length));
      seen.add(meal.id);
      added += 1;
    }
    usedLibs.push(`${lib.name}(${added})`);
  }

  const data = {
    meta: {
      title: '家庭食谱库',
      version: '3.0.0',
      model: 'meal',
      servingNote: '每道食谱为一份完整主菜，营养与食材分量以食谱标注的人份计；部分营养为估算值。',
      libraries: LIBRARIES.map((l) => l.path.replace(`${DATA}/`, 'data/')),
      imagesFrom: 'data/recipe-images.json',
      generatedAt: new Date().toISOString(),
      count: meals.length,
      nutritionUnits: { calories: 'kcal', protein: 'g', fat: 'g', carbs: 'g' },
    },
    meals,
  };

  writeFileSync(OUT, `${JSON.stringify(data, null, 2)}\n`, 'utf-8');
  const lunch = meals.filter((m) => m.mealSlots.includes('lunch')).length;
  const dinner = meals.filter((m) => m.mealSlots.includes('dinner')).length;
  const withImg = meals.filter((m) => m.image).length;
  console.log(`✅ 生成 ${meals.length} 道食谱 → ${OUT}`);
  console.log(`   来源：${usedLibs.join('，')}`);
  console.log(`   可作午餐：${lunch}，可作晚餐：${dinner}，含图片：${withImg}/${meals.length}`);
}

main();
