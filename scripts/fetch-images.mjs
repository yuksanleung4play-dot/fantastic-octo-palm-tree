#!/usr/bin/env node
/**
 * 从每道食谱的 source 链接中提取一张代表图片，并下载到 public/dish-images/。
 * 结果缓存到 data/recipe-images.json（id → { file, src, from }），
 * 之后由 build-recipes.mjs 合并进 data/recipes.json 的 image 字段。
 *
 * 用法：
 *   node scripts/fetch-images.mjs           # 只抓取尚未成功缓存的
 *   node scripts/fetch-images.mjs --force   # 全部重新抓取
 *   node scripts/fetch-images.mjs --retry   # 仅重试此前失败（无图）的
 *
 * 多策略：YouTube 缩略图 → og:image/twitter:image → <link image_src>
 *        → JSON-LD image → 正文首图；下载时按文件头嗅探图片类型并重试。
 *        抓不到则留空，前端用占位图兜底。
 */
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const RECIPES = join(__dirname, '..', 'data', 'recipes.json');
const CACHE = join(__dirname, '..', 'data', 'recipe-images.json');
const IMG_DIR = join(__dirname, '..', 'public', 'dish-images');

const UA =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const FORCE = process.argv.includes('--force');
const RETRY = process.argv.includes('--retry');
const LOGO_RE = /(logo|cropped|avatar|favicon|icon|placeholder|sprite|blank)/i;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function decodeEntities(s = '') {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&#x2F;/gi, '/')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function absolutize(url, base) {
  try {
    return new URL(decodeEntities(url), base).href;
  } catch {
    return null;
  }
}

function youtubeId(url) {
  const m =
    url.match(/[?&]v=([\w-]{6,})/) ||
    url.match(/youtu\.be\/([\w-]{6,})/) ||
    url.match(/youtube\.com\/(?:embed|shorts)\/([\w-]{6,})/);
  return m ? m[1] : null;
}

function pickMeta(html, key, attr = 'property') {
  const re = new RegExp(
    `<meta[^>]+${attr}=["']${key}["'][^>]+content=["']([^"']+)["']|<meta[^>]+content=["']([^"']+)["'][^>]+${attr}=["']${key}["']`,
    'i',
  );
  const m = html.match(re);
  return m ? m[1] || m[2] : null;
}

function fromJsonLd(html) {
  const blocks = [...html.matchAll(/<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi)];
  for (const b of blocks) {
    try {
      const data = JSON.parse(b[1].trim());
      const arr = Array.isArray(data) ? data : [data, ...(data['@graph'] || [])];
      for (const node of arr) {
        const img = node && node.image;
        if (!img) continue;
        if (typeof img === 'string') return img;
        if (Array.isArray(img)) return typeof img[0] === 'string' ? img[0] : img[0]?.url;
        if (img.url) return img.url;
      }
    } catch {
      /* ignore malformed json-ld */
    }
  }
  return null;
}

function firstContentImage(html, base) {
  const imgs = [...html.matchAll(/<img[^>]+>/gi)].map((m) => m[0]);
  for (const tag of imgs) {
    const src =
      (tag.match(/\sdata-src=["']([^"']+)["']/i) || [])[1] ||
      (tag.match(/\ssrc=["']([^"']+)["']/i) || [])[1];
    if (!src || src.startsWith('data:')) continue;
    if (LOGO_RE.test(src)) continue;
    if (/uploads|images?|cdn|media|photo/i.test(src)) return absolutize(src, base);
  }
  return null;
}

async function fetchRetry(url, opts, tries = 3) {
  let lastErr;
  for (let i = 0; i < tries; i += 1) {
    try {
      const res = await fetch(url, opts);
      if (res.status === 403 || res.status === 429 || res.status >= 500) {
        lastErr = new Error(`http-${res.status}`);
      } else {
        return res;
      }
    } catch (e) {
      lastErr = e;
    }
    await sleep(800 * (i + 1));
  }
  throw lastErr;
}

async function extractImageUrl(pageUrl) {
  const yt = youtubeId(pageUrl);
  if (yt) return { url: `https://i.ytimg.com/vi/${yt}/hqdefault.jpg`, from: 'youtube' };

  let res;
  try {
    res = await fetchRetry(pageUrl, {
      headers: { 'User-Agent': UA, Accept: 'text/html', Referer: 'https://www.google.com/' },
      redirect: 'follow',
    });
  } catch (e) {
    return { url: null, from: `fetch-error:${e.message}` };
  }
  if (!res.ok) return { url: null, from: `http-${res.status}` };
  const html = await res.text();
  const base = res.url || pageUrl;

  const candidates = [
    ['og:image:secure_url', () => pickMeta(html, 'og:image:secure_url')],
    ['og:image:url', () => pickMeta(html, 'og:image:url')],
    ['og:image', () => pickMeta(html, 'og:image')],
    ['og:image', () => pickMeta(html, 'og:image', 'name')],
    ['twitter:image', () => pickMeta(html, 'twitter:image', 'name')],
    ['twitter:image', () => pickMeta(html, 'twitter:image')],
    ['twitter:image:src', () => pickMeta(html, 'twitter:image:src', 'name')],
    ['image_src', () => (html.match(/<link[^>]+rel=["']image_src["'][^>]+href=["']([^"']+)["']/i) || [])[1]],
    ['json-ld', () => fromJsonLd(html)],
    ['content-img', () => firstContentImage(html, base)],
  ];

  for (const [from, fn] of candidates) {
    let raw = fn();
    if (!raw) continue;
    raw = decodeEntities(raw);
    if (LOGO_RE.test(raw)) continue;
    const abs = absolutize(raw, base);
    if (abs) return { url: abs, from };
  }
  return { url: null, from: 'not-found' };
}

const EXT_BY_TYPE = { 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif' };

/** 按文件头魔数判断图片类型（应对服务器返回 octet-stream / 错误类型）。 */
function sniffExt(buf) {
  if (buf.length < 12) return null;
  if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'jpg';
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'png';
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) return 'gif';
  if (
    buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
    buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50
  ) {
    return 'webp';
  }
  return null;
}

function extFromUrl(url) {
  const m = url.split('?')[0].match(/\.(jpe?g|png|webp|gif)$/i);
  return m ? m[1].toLowerCase().replace('jpeg', 'jpg') : null;
}

async function download(imgUrl, id, referer) {
  const res = await fetchRetry(imgUrl, {
    headers: { 'User-Agent': UA, Referer: referer || '', Accept: 'image/*,*/*' },
    redirect: 'follow',
  });
  if (!res.ok) throw new Error(`http-${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.length < 2048) throw new Error(`图片过小:${buf.length}B`);
  if (buf.length > 3_000_000) throw new Error(`图片过大:${buf.length}B`);
  const type = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
  const ext = sniffExt(buf) || EXT_BY_TYPE[type] || extFromUrl(imgUrl);
  if (!ext) throw new Error(`无法识别图片类型:${type}`);
  const file = `${id}.${ext}`;
  writeFileSync(join(IMG_DIR, file), buf);
  return { file, bytes: buf.length };
}

async function main() {
  if (!existsSync(IMG_DIR)) mkdirSync(IMG_DIR, { recursive: true });
  const recipes = JSON.parse(readFileSync(RECIPES, 'utf-8'));
  const cache = existsSync(CACHE) ? JSON.parse(readFileSync(CACHE, 'utf-8')) : {};

  let ok = 0;
  let fail = 0;
  for (const meal of recipes.meals) {
    const cached = cache[meal.id];
    const haveFile = cached && cached.file && existsSync(join(IMG_DIR, cached.file));
    if (!FORCE && haveFile) {
      ok += 1;
      continue;
    }
    if (RETRY && haveFile) {
      ok += 1;
      continue;
    }

    const { url, from } = await extractImageUrl(meal.sourceUrl || '');
    if (!url) {
      cache[meal.id] = { file: null, src: null, from };
      fail += 1;
      console.log(`✗ ${meal.id}  ${from}  ${meal.title.slice(0, 14)}`);
      continue;
    }
    try {
      const { file, bytes } = await download(url, meal.id, meal.sourceUrl);
      cache[meal.id] = { file, src: url, from };
      ok += 1;
      console.log(`✓ ${meal.id}  ${from}  ${(bytes / 1024) | 0}KB  ${meal.title.slice(0, 14)}`);
    } catch (e) {
      cache[meal.id] = { file: null, src: url, from: `dl-fail:${e.message}` };
      fail += 1;
      console.log(`✗ ${meal.id}  下载失败 ${e.message}  ${meal.title.slice(0, 14)}`);
    }
  }

  writeFileSync(CACHE, `${JSON.stringify(cache, null, 2)}\n`, 'utf-8');
  console.log(`\n完成：有图 ${ok}，无图 ${fail}。缓存 → ${CACHE}`);
}

main();
