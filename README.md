# 家庭点餐系统 🍽️

一个家庭内部使用的点餐系统：基于**真实整餐食谱库**自动生成 A/B/C/D 餐，家人在网页上挑选，保存后自动给你发送「备菜邮件」。配有为 Safari（含 iPhone）优化的前置首屏。

## 功能特性

- **真实食谱库驱动**：50 道完整食谱（每道为一份含荤与素的均衡正餐，部分带汤），来自 `data/source-recipes.json`，经 `npm run build:recipes` 规范化为 `data/recipes.json`。
- **智能排程**：
  - 平日（周一至周五）：**一餐（晚餐）**
  - 周末（周六、周日）：**午餐 + 晚餐**
- **自动生成 A/B/C/D 四套餐**：每套是一道完整食谱，按餐次（午/晚）从可用食谱中**确定性**挑选——同一餐次始终一致，四套互不相同。
- **标明热量与营养成分**：每道食谱显示热量、蛋白质、脂肪、碳水。
- **不区分点餐人 · 按天汇总**：无需登录或选择「我是谁」，任何人选好套餐保存即可。**同一天内保存的所有记录都会一起显示，并在每次保存时发送备菜邮件**。
- **下单即发邮件**：保存后自动发送邮件到你的邮箱，列出：
  - 本次新增的食谱与营养
  - **当天已保存的全部记录**（跨午餐/晚餐）
  - **需要准备的全部食材与分量**：按食谱合并、按份数标注（如 `×2 份`），方便买菜备菜。

## 快速开始

```bash
npm install
npm run build:recipes   # 由源数据生成 data/recipes.json（已附带生成结果）
cp .env.example .env     # 按需填写邮箱与 SMTP
npm start                # 访问 http://localhost:3000
```

用手机浏览器打开 → 进入首屏点「开始点餐」→ 选日期与餐次 → 点选 A/B/C/D → 点击「保存点餐」。当天每次保存都会刷新记录并发送备菜邮件。iPhone Safari 可「添加到主屏幕」当作 App 使用。

## 邮件配置

在 `.env` 中配置（参考 `.env.example`）：

| 变量 | 说明 |
| --- | --- |
| `NOTIFY_EMAIL` | 收件邮箱（备菜通知发到这里） |
| `SMTP_HOST` / `SMTP_PORT` | SMTP 服务器，如 `smtp.gmail.com` / `587` |
| `SMTP_USER` / `SMTP_PASS` | 发信账号与密码（Gmail 需用「应用专用密码」） |
| `SMTP_FROM` | 可选，发件人显示名 |

> 未配置 SMTP 时，系统不会报错，而是把邮件以 HTML 形式落盘到 `data/outbox/` 供预览，方便先在本地体验。

## 食谱库与自定义

食谱库分两层：

- `data/source-recipes.json`：**原始食谱**（人类可读，整餐列表）。新增/修改菜谱在这里编辑。
- `data/recipes.json`：由构建脚本生成的**结构化数据**（解析后的餐次、人份、营养、食材分组），系统实际读取它。

编辑源数据后运行 `npm run build:recipes` 重新生成。缺少食材的食谱会自动从可点餐池中跳过（保证备菜清单可用）。

### 菜品图片（效果图）

由于从来源网页抓取的图片常与菜名不符，现改为**根据每道食谱内容用 AI 生成效果图**，确保图文相符。50 道菜全部配图，存放于 `public/dish-images/recipe-XX.jpg`（统一压缩为约 900px 宽的 JPEG，整体约 6MB），图片路径记录在 `data/recipe-images.json`（`from: "generated"`），由 `npm run build:recipes` 合并进 `data/recipes.json` 的 `image` 字段。

> 备用方案：`scripts/fetch-images.mjs`（`npm run fetch:images`）仍可从食谱 `source` 链接抓取真实图片（YouTube 缩略图 / `og:image` / JSON-LD / 正文首图）。前端在缺图时会显示带 🍽️ 的占位图。

## API

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/recipes` | 食谱库全部食谱 |
| `GET` | `/api/schedule?start=YYYY-MM-DD&days=14` | 排程（平日/周末餐次） |
| `GET` | `/api/meals?date=YYYY-MM-DD&meal=lunch\|dinner` | 该餐 A/B/C/D 选项；`existingOrders` 为当天全部记录 |
| `POST` | `/api/orders` | 保存记录并发当天备菜邮件，body：`{date,meal,optionLabel}` |
| `GET` | `/api/orders?date=YYYY-MM-DD` | 查看某一天的全部记录 |

## 项目结构

```
data/source-recipes.json  原始食谱（可编辑）
data/recipes.json         构建生成的结构化食谱库
data/orders.json          记录存储（自动生成，已 gitignore）
data/outbox/              未配置 SMTP 时的邮件预览（已 gitignore）
scripts/build-recipes.mjs 源数据 → 结构化食谱库的构建脚本
src/recipes.js            食谱加载/校验/营养汇总
src/schedule.js           平日一餐 / 周末午晚餐排程
src/mealGenerator.js      A/B/C/D 确定性挑选
src/store.js              记录持久化（按天，不区分点餐人）
src/mailer.js             备菜邮件构建与发送
src/server.js             Express 服务与 API
public/                   前端页面（首屏 + 点餐主屏，Safari 优化 / PWA）
test/                     单元测试（node --test）
```

## 测试

```bash
npm test
```
