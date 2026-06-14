# 家庭点餐系统 🍽️

一个家庭内部使用的点餐系统：基于**真实整餐食谱库**自动生成 A/B/C/D 餐，家人在网页上挑选，保存后自动给你发送「备菜邮件」。配有为 Safari（含 iPhone）优化的前置首屏。

## 功能特性

- **真实食谱库驱动**：共 **86 道**完整主菜食谱，由两个已结构化的来源库合并而成（`data/recipes-base.json` 整合基底库 + `data/lib-shipu.json` 港式精选库），经 `npm run build:recipes` 生成 `data/recipes.json`，每道均配有按食谱内容生成的效果图。
- **智能排程**：
  - 平日（周一至周五）：**一餐（晚餐）**
  - 周末（周六、周日）：**午餐 + 晚餐**
- **自动生成 A/B/C/D 四套餐**：每套是一道完整食谱，按餐次（午/晚）从可用食谱中**确定性**挑选——同一餐次始终一致，四套互不相同。
- **标明热量与营养成分**：每道食谱显示热量、蛋白质、脂肪、碳水。
- **不区分点餐人 · 每餐只留最新**：无需登录或选择「我是谁」，任何人选好套餐保存即可。**每天每餐最多保留一条记录，重复保存只覆盖为最新选择**。
- **每周统一发邮件（周六 17:00）**：不再每次保存就发邮件，而是**每周六下午 5 点统一发送一封「下周一至周六」的餐单与备菜邮件**，内容包括：
  - 每日已点餐单（未点的日子标注「未点餐」）
  - **一周备菜清单（分类汇总）**：把整周全部食材合并后**分类总结为「肉類 / 蔬菜 / 調味料 / 其他」**，同名食材自动累加分量，方便一次过采购
  - 全部营养合计
  - 发送时间/星期/时区可在 `.env` 调整（见下）。手动测试可调用 `POST /api/send-weekly`。

## 快速开始

```bash
npm install
npm run build:recipes   # 由源数据生成 data/recipes.json（已附带生成结果）
cp .env.example .env     # 按需填写邮箱与 SMTP
npm start                # 访问 http://localhost:3000
```

用手机浏览器打开 → 进入首屏点「开始点餐」→ 选日期与餐次 → 点选 A/B/C/D → 点击「保存点餐」。当天每次保存都会刷新记录并发送备菜邮件。iPhone Safari 可「添加到主屏幕」当作 App 使用。

## 邮件配置

> 启动时会自动加载项目根目录的 `.env`（由 `src/env.js` 零依赖实现，无需 dotenv 或 `--env-file`）。系统环境变量优先级高于 `.env`。

在 `.env` 中配置（参考 `.env.example`）：

| 变量 | 说明 |
| --- | --- |
| `NOTIFY_EMAIL` | 收件邮箱（备菜通知发到这里） |
| `SMTP_HOST` / `SMTP_PORT` | SMTP 服务器，如 `smtp.gmail.com` / `587` |
| `SMTP_USER` / `SMTP_PASS` | 发信账号与密码（Gmail 需用「应用专用密码」） |
| `SMTP_FROM` | 可选，发件人显示名 |
| `WEEKLY_SEND_DOW` / `WEEKLY_SEND_HOUR` / `WEEKLY_SEND_MINUTE` | 每周发送的星期(0–6)/小时/分钟，默认 6/17/0（周六 17:00） |
| `WEEKLY_SEND_TZ_OFFSET` | 时区分钟偏移，默认 480（UTC+8 香港） |
| `WEEKLY_SEND_ENABLED` | 设为 `false` 可关闭定时发送 |

> 未配置 SMTP 时，系统不会报错，而是把邮件以 HTML 形式落盘到 `data/outbox/` 供预览，方便先在本地体验。

## 食谱库与自定义

食谱库由若干**已结构化的来源库**（`{meta, meals}`）合并而成，在 `scripts/build-recipes.mjs` 的 `LIBRARIES` 中按顺序列出，构建时按 `id` 去重合并并套用图片缓存。当前来源：

- `data/recipes-base.json`：整合基底库（家庭食谱 + 香港营养师协会主菜，精选合并，43 道）
- `data/lib-shipu.json`：港式精选食谱库（43 道；由 `scripts/gen-shipu.py` 依据 `data/source-shipu.md` 的菜名/链接整理而成，营养为估算值）

`data/recipes.json` 即合并后的最终数据（共 86 道），系统实际读取它。要新增来源库，把同结构 JSON 放进 `data/` 并加入 `LIBRARIES` 即可，运行 `npm run build:recipes` 重建。

### 菜品图片（效果图）

由于从来源网页抓取的图片常与菜名不符，改为**根据每道食谱内容用 AI 生成效果图**，确保图文相符。86 道菜全部配图，存放于 `public/dish-images/<id>.jpg`（统一压缩为约 900px 宽的 JPEG，整体约 11MB），图片路径记录在 `data/recipe-images.json`（`from: "generated"`），由 `npm run build:recipes` 合并进 `data/recipes.json` 的 `image` 字段。

> 备用方案：`scripts/fetch-images.mjs`（`npm run fetch:images`）仍可从食谱 `source` 链接抓取真实图片（YouTube 缩略图 / `og:image` / JSON-LD / 正文首图）。前端在缺图时会显示带 🍽️ 的占位图。

## API

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/recipes` | 食谱库全部食谱 |
| `GET` | `/api/schedule?start=YYYY-MM-DD&days=14` | 排程（平日/周末餐次） |
| `GET` | `/api/meals?date=YYYY-MM-DD&meal=lunch\|dinner` | 该餐 A/B/C/D 选项；`existingOrders` 为当天全部记录 |
| `POST` | `/api/orders` | 保存记录（每餐只留最新，不即时发邮件），body：`{date,meal,optionLabel}` |
| `GET` | `/api/orders?date=YYYY-MM-DD` | 查看某一天的全部记录 |
| `GET` | `/api/weekly-info` | 下次发送时间与即将到来的周一至周六区间 |
| `POST` | `/api/send-weekly` | 手动触发本周餐单邮件，可选 body：`{start:"YYYY-MM-DD"}`（指定周一） |

## 项目结构

```
data/recipes-base.json    整合基底库（家庭+HKDA 精选）
data/lib-shipu.json       港式精选库（gen-shipu.py 整理）
data/source-shipu.md      港式食谱清单（菜名 + 来源链接）
data/recipes.json         构建生成的结构化食谱库
data/orders.json          记录存储（自动生成，已 gitignore）
data/outbox/              未配置 SMTP 时的邮件预览（已 gitignore）
scripts/build-recipes.mjs 合并各来源库 → data/recipes.json
scripts/gen-shipu.py      由 source-shipu.md 整理港式食谱库
src/recipes.js            食谱加载/校验/营养汇总
src/schedule.js           平日一餐 / 周末午晚餐排程
src/mealGenerator.js      A/B/C/D 确定性挑选
src/store.js              记录持久化（每餐只留最新）
src/mailer.js             每周餐单邮件构建与发送
src/scheduler.js          每周六 17:00 定时发送（时区/星期/时间可配置）
src/env.js                零依赖 .env 加载器
src/server.js             Express 服务与 API
public/                   前端页面（首屏 + 点餐主屏，Safari 优化 / PWA）
test/                     单元测试（node --test）
```

## 测试

```bash
npm test
```
