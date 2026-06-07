import { ordersForRange } from './store.js';
import { sendWeeklyEmail } from './mailer.js';

/**
 * 每周定时发送「本周餐单」邮件（默认周六 17:00，香港时区 UTC+8）。
 * 时区通过分钟偏移量配置，独立于服务器系统时区，确保发送时间稳定。
 */
export function getConfig() {
  return {
    enabled: process.env.WEEKLY_SEND_ENABLED !== 'false',
    dow: Number(process.env.WEEKLY_SEND_DOW ?? 6), // 0=周日 … 6=周六
    hour: Number(process.env.WEEKLY_SEND_HOUR ?? 17),
    minute: Number(process.env.WEEKLY_SEND_MINUTE ?? 0),
    offsetMin: Number(process.env.WEEKLY_SEND_TZ_OFFSET ?? 480), // UTC+8 = 480 分钟
  };
}

function fmt(y, m, d) {
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

/** 计算给定时刻之后的下一次发送时间（UTC 瞬时 Date）。 */
export function nextSendTime(now, cfg = getConfig()) {
  const offMs = cfg.offsetMin * 60000;
  const tz = new Date(now.getTime() + offMs); // 其 UTC 字段即目标时区的“墙上时间”
  const delta = (cfg.dow - tz.getUTCDay() + 7) % 7;
  const candWall = Date.UTC(
    tz.getUTCFullYear(),
    tz.getUTCMonth(),
    tz.getUTCDate() + delta,
    cfg.hour,
    cfg.minute,
    0,
  );
  let instant = candWall - offMs;
  if (instant <= now.getTime()) instant += 7 * 86400000;
  return new Date(instant);
}

/** 计算给定时刻之后“即将到来的”周一至周六日期区间（按目标时区）。 */
export function upcomingWeekRange(now, cfg = getConfig()) {
  const offMs = cfg.offsetMin * 60000;
  const tz = new Date(now.getTime() + offMs);
  const daysUntilMon = ((1 - tz.getUTCDay() + 7) % 7) || 7; // 严格的下一个周一
  const mon = new Date(Date.UTC(tz.getUTCFullYear(), tz.getUTCMonth(), tz.getUTCDate() + daysUntilMon));
  const sat = new Date(mon.getTime() + 5 * 86400000);
  return {
    start: fmt(mon.getUTCFullYear(), mon.getUTCMonth(), mon.getUTCDate()),
    end: fmt(sat.getUTCFullYear(), sat.getUTCMonth(), sat.getUTCDate()),
  };
}

/** 立即执行一次本周发送。可传 startMonday 指定某个周一（否则取即将到来的周一）。 */
export async function runWeeklySend(now = new Date(), startMonday = null) {
  const cfg = getConfig();
  let start;
  let end;
  if (startMonday) {
    start = startMonday;
    const m = new Date(`${startMonday}T00:00:00Z`);
    const sat = new Date(m.getTime() + 5 * 86400000);
    end = fmt(sat.getUTCFullYear(), sat.getUTCMonth(), sat.getUTCDate());
  } else {
    ({ start, end } = upcomingWeekRange(now, cfg));
  }
  const orders = ordersForRange(start, end);
  const result = await sendWeeklyEmail({ orders, rangeStart: start, rangeEnd: end });
  console.log(
    `📧 周餐单邮件 ${start}~${end}：${result.sent ? '已发送' : `未发送(${result.reason || result.mode})`}，共 ${orders.length} 餐`,
  );
  return result;
}

let timer = null;

/** 启动定时器：到点发送后自动续订下一周。 */
export function startScheduler() {
  const cfg = getConfig();
  if (!cfg.enabled) {
    console.log('📅 周餐单定时邮件已禁用（WEEKLY_SEND_ENABLED=false）');
    return;
  }
  const schedule = () => {
    const now = new Date();
    const next = nextSendTime(now, cfg);
    const ms = next.getTime() - now.getTime();
    timer = setTimeout(async () => {
      try {
        await runWeeklySend(new Date());
      } catch (e) {
        console.error('周餐单邮件发送出错：', e.message);
      }
      schedule();
    }, ms);
    if (timer.unref) timer.unref();
    const tzLabel = `UTC${cfg.offsetMin >= 0 ? '+' : ''}${cfg.offsetMin / 60}`;
    console.log(
      `📅 下次周餐单邮件：${next.toISOString()}（${tzLabel} 周${'日一二三四五六'[cfg.dow]} ${String(
        cfg.hour,
      ).padStart(2, '0')}:${String(cfg.minute).padStart(2, '0')}，约 ${(ms / 3600000).toFixed(1)} 小时后）`,
    );
  };
  schedule();
}

export function stopScheduler() {
  if (timer) clearTimeout(timer);
  timer = null;
}
