import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DEFAULT_ENV_PATH = process.env.ENV_PATH || join(__dirname, '..', '.env');

/**
 * 轻量 .env 加载器（零依赖）。
 * Node 默认不会读取 .env，必须显式加载，否则 process.env 里取不到 SMTP_* 等配置。
 * 解析规则：
 *   - 跳过空行与以 # 开头的注释行
 *   - 按第一个 = 分割 key/value（value 中可含空格、< > 等）
 *   - 去除 value 两侧成对引号
 *   - 不覆盖已存在的真实环境变量（系统环境优先级更高）
 */
export function loadEnv(path = DEFAULT_ENV_PATH) {
  if (!existsSync(path)) return { loaded: false, path, count: 0 };
  const content = readFileSync(path, 'utf-8');
  let count = 0;
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const eq = line.indexOf('=');
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    if (!key) continue;
    let val = line.slice(eq + 1).trim();
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    if (!(key in process.env)) {
      process.env[key] = val;
      count += 1;
    }
  }
  return { loaded: true, path, count };
}

// 作为副作用在被 import 时立即加载，确保后续模块能读到配置。
loadEnv();
