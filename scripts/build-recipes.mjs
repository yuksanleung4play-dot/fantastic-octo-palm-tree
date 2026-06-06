#!/usr/bin/env node
/**
 * 将上传的原始食谱（data/source-recipes.json，整餐食谱列表）规范化为
 * 系统使用的 data/recipes.json（含解析后的餐次、人份、营养、食材分组、图片）。
 *
 * 图片来自 data/recipe-images.json（由 scripts/fetch-images.mjs 从 source 链接抓取）。
 *
 * 运行： npm run build:recipes
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SRC = join(__dirname, '..', 'data', 'source-recipes.json');
const OUT = join(__dirname, '..', 'data', 'recipes.json');
const IMAGES = join(__dirname, '..', 'data', 'recipe-images.json');

const FRACTIONS = { '½': 0.5, '¼': 0.25, '⅓': 1 / 3, '⅔': 2 / 3, '¾': 0.75, '⅛': 0.125 };

/** 去除 markdown 加粗与「約」等修饰，提取首个数字。 */
function parseNumber(value) {
  if (value == null) return null;
  const cleaned = String(value).replace(/[*約约\s]/g, '');
  const match = cleaned.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function parseNutrition(nut = {}) {
  return {
    calories: parseNumber(nut.calories),
    protein: parseNumber(nut.protein),
    fat: parseNumber(nut.fat),
    carbs: parseNumber(nut.carbs),
  };
}

/** 标题形如「食譜01｜煎嫩雞胸＋...」→ { no, title }。 */
function parseTitle(raw) {
  const parts = raw.split(/｜|\|/);
  if (parts.length >= 2) {
    const noMatch = parts[0].match(/\d+/);
    return { no: noMatch ? Number(noMatch[0]) : null, title: parts.slice(1).join('｜').trim() };
  }
  return { no: null, title: raw.trim() };
}

/** 从 source 文本中拆出展示名与链接（markdown [text](url)）。 */
function parseSource(raw = '') {
  const urlMatch = raw.match(/\[[^\]]*\]\((https?:\/\/[^)]+)\)/);
  const sourceUrl = urlMatch ? urlMatch[1] : null;
  const name = raw
    .replace(/\s*\[[^\]]*\]\([^)]*\)/g, '')
    .replace(/\s*\[[^\]]*\]/g, '')
    .trim();
  return { source: name, sourceUrl };
}

/** meal_type → 适用餐次（lunch/dinner）与基础人份。 */
function parseMealType(raw = '') {
  const slots = [];
  const hasLunch = /午餐|午／晚|午\/晚|brunch/i.test(raw);
  const hasDinner = /晚餐|午／晚|午\/晚/.test(raw);
  if (hasLunch) slots.push('lunch');
  if (hasDinner) slots.push('dinner');
  if (slots.length === 0) slots.push('lunch', 'dinner');
  const servingMatch = raw.match(/(\d+)\s*人份/);
  const servings = servingMatch ? Number(servingMatch[1]) : 1;
  return { mealSlots: slots, servings };
}

const GROUP_LABEL_RE = /^([^、：:0-9]{2,8})[：:]/;

/** 解析单行食材为 { label, items[] }。items 为去除前缀后的食材文本。 */
function parseIngredientLine(line) {
  const trimmed = line.trim();
  let label = null;
  let rest = trimmed;
  const m = trimmed.match(GROUP_LABEL_RE);
  if (m) {
    label = m[1].trim();
    rest = trimmed.slice(m[0].length).trim();
  }
  const items = rest
    .split(/、|，|,/)
    .map((s) => s.trim())
    .filter(Boolean);
  return { label, items };
}

function buildRecipe(entry, index, images) {
  const { no, title } = parseTitle(entry.title || '');
  const { source, sourceUrl } = parseSource(entry.source || '');
  const { mealSlots, servings } = parseMealType(entry.meal_type || '');
  const nutrition = parseNutrition(entry.nutrition || {});

  const ingredientGroups = (entry.ingredients || []).map(parseIngredientLine);
  const ingredients = ingredientGroups.flatMap((g) => g.items);

  const id = `recipe-${String(no ?? index + 1).padStart(2, '0')}`;
  const imgEntry = images[id];
  const image = imgEntry && imgEntry.file ? `dish-images/${imgEntry.file}` : null;

  return {
    id,
    no: no ?? index + 1,
    title,
    source,
    sourceUrl,
    image,
    imageSource: imgEntry && imgEntry.src ? imgEntry.src : null,
    mealType: entry.meal_type || '',
    mealSlots,
    servings,
    nutrition,
    ingredientGroups,
    ingredients,
    steps: entry.steps || [],
  };
}

function main() {
  const src = JSON.parse(readFileSync(SRC, 'utf-8'));
  const images = existsSync(IMAGES) ? JSON.parse(readFileSync(IMAGES, 'utf-8')) : {};
  const meals = src.map((entry, i) => buildRecipe(entry, i, images));

  const data = {
    meta: {
      title: '家庭食谱库',
      version: '2.1.0',
      model: 'meal',
      servingNote: '每道食谱为一份完整的餐（含荤与素），营养与食材分量以食谱标注的人份计。',
      generatedFrom: 'data/source-recipes.json',
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
  console.log(`   可作午餐：${lunch}，可作晚餐：${dinner}，含图片：${withImg}/${meals.length}`);
}

main();
