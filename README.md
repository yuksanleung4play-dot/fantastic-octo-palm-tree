# 家庭点餐系统 🍽️

一个家庭内部使用的点餐系统：根据食谱库自动生成套餐，家人在网页上挑选，保存后自动给你发送「备菜邮件」。

## 功能特性

- **食谱库驱动**：所有菜品来自 `data/recipes.json`，可自由增删替换成自家菜谱。
- **智能排程**：
  - 平日（周一至周五）：**一餐（晚餐）**
  - 周末（周六、周日）：**午餐 + 晚餐**
- **自动生成 A/B/C/D 四套套餐**：每套**一荤一素**，同一餐次对所有家人展示一致（确定性生成）。
- **标明热量与营养**：每道菜与每套套餐均显示总热量、蛋白质、脂肪、碳水、膳食纤维、钠。
- **下单即发邮件**：任一家人保存点餐后，系统自动发送邮件到你的邮箱，列出：
  - 本次所点套餐与营养
  - 本餐所有已下单家人
  - **需要准备的全部食材与分量**（按所有订单自动累加，方便买菜备菜）

## 快速开始

```bash
npm install
cp .env.example .env   # 按需填写邮箱与 SMTP
npm start              # 访问 http://localhost:3000
```

打开浏览器 → 输入名字 → 选日期与餐次 → 点选 A/B/C/D → 点击「保存点餐并发送备菜邮件」。

## 邮件配置

在 `.env` 中配置（参考 `.env.example`）：

| 变量 | 说明 |
| --- | --- |
| `NOTIFY_EMAIL` | 收件邮箱（备菜通知发到这里） |
| `SMTP_HOST` / `SMTP_PORT` | SMTP 服务器，如 `smtp.gmail.com` / `587` |
| `SMTP_USER` / `SMTP_PASS` | 发信账号与密码（Gmail 需用「应用专用密码」） |
| `SMTP_FROM` | 可选，发件人显示名 |

> 未配置 SMTP 时，系统不会报错，而是把邮件以 HTML 形式落盘到 `data/outbox/` 供预览，方便先在本地体验。

## 自定义食谱

编辑 `data/recipes.json`，每道菜结构如下（营养与食材均以**单人份**计，系统按点餐人数自动累加）：

```json
{
  "id": "m-gongbao-chicken",
  "name": "宫保鸡丁",
  "type": "meat",                       // meat=荤, vegetable=素
  "tags": ["川菜", "下饭"],
  "calories": 380,
  "nutrition": { "protein": 28, "fat": 22, "carbs": 16, "fiber": 2, "sodium": 760 },
  "ingredients": [
    { "name": "鸡胸肉", "amount": 180, "unit": "g" }
  ]
}
```

食谱库需同时包含至少 4 道荤菜与 4 道素菜，才能保证 A/B/C/D 四套套餐荤素都不重复。

## API

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/recipes` | 食谱库全部菜品 |
| `GET` | `/api/schedule?start=YYYY-MM-DD&days=14` | 排程（平日/周末餐次） |
| `GET` | `/api/meals?date=YYYY-MM-DD&meal=lunch\|dinner` | 该餐 A/B/C/D 套餐 |
| `POST` | `/api/orders` | 下单并发邮件，body：`{member,date,meal,optionLabel}` |
| `GET` | `/api/orders?date=&meal=` | 查看某餐已下单情况 |

## 项目结构

```
data/recipes.json     食谱库（可自定义）
data/orders.json      订单存储（自动生成，已 gitignore）
data/outbox/          未配置 SMTP 时的邮件预览（已 gitignore）
src/recipes.js        食谱加载/校验/营养汇总
src/schedule.js       平日一餐 / 周末午晚餐排程
src/mealGenerator.js  A/B/C/D 套餐确定性生成
src/store.js          订单持久化
src/mailer.js         备菜邮件构建与发送
src/server.js         Express 服务与 API
public/               前端页面（HTML/CSS/JS）
test/                 单元测试（node --test）
```

## 测试

```bash
npm test
```
