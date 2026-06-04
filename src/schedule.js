/**
 * 排程逻辑：
 *  - 平日（周一至周五）：一餐（晚餐）
 *  - 周末（周六、周日）：午餐 + 晚餐
 */

export const MEAL_LABELS = {
  dinner: '晚餐',
  lunch: '午餐',
};

/**
 * 判断给定日期是否周末。dateStr 形如 "2026-06-06"。
 */
export function isWeekend(dateStr) {
  const day = new Date(`${dateStr}T00:00:00`).getDay();
  return day === 0 || day === 6;
}

/**
 * 返回某一天需要安排的餐次列表。
 * 平日只有 dinner；周末有 lunch 与 dinner。
 */
export function mealsForDate(dateStr) {
  return isWeekend(dateStr) ? ['lunch', 'dinner'] : ['dinner'];
}

/**
 * 生成从 startDate 起 days 天的排程概要。
 */
export function buildSchedule(startDate, days = 7) {
  const result = [];
  const base = new Date(`${startDate}T00:00:00`);
  for (let i = 0; i < days; i += 1) {
    const d = new Date(base);
    d.setDate(base.getDate() + i);
    const dateStr = toDateStr(d);
    result.push({
      date: dateStr,
      weekday: weekdayName(d),
      isWeekend: isWeekend(dateStr),
      meals: mealsForDate(dateStr),
    });
  }
  return result;
}

export function toDateStr(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

const WEEKDAY_NAMES = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
export function weekdayName(date) {
  return WEEKDAY_NAMES[date.getDay()];
}
