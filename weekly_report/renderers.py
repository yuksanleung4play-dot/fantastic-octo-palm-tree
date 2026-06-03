"""把 :class:`~weekly_report.models.WeeklyReport` 輸出成不同格式。

提供三種格式：
    - ``text``：純文字，貼到郵件或聊天視窗都好讀。
    - ``markdown``：適合貼到支援 Markdown 的工具（Notion、GitHub…）。
    - ``html``：可直接於瀏覽器開啟或貼進郵件。

所有渲染器都遵循同一份週報結構：
    1) 姓名、日期、部門
    2) 日常工作內容
    3) 其他事項
"""

from __future__ import annotations

import html
from typing import Callable, Dict

from .models import WeeklyReport, WorkDay


def _item_line(description: str, recipients_text: str) -> str:
    if recipients_text:
        return f"{description}（對象：{recipients_text}）"
    return description


# --------------------------------------------------------------------------- #
# 純文字
# --------------------------------------------------------------------------- #
def render_text(report: WeeklyReport) -> str:
    cfg = report.config
    lines = []
    lines.append("=" * 48)
    lines.append(cfg.title.center(40))
    lines.append("=" * 48)
    lines.append("")
    lines.append("一、基本資訊")
    lines.append(f"  姓名：{cfg.name}")
    lines.append(f"  部門：{cfg.department}")
    lines.append(f"  日期：{cfg.period_text}")
    lines.append("")
    lines.append("二、日常工作內容")
    if report.work_days:
        for wd in report.work_days:
            lines.append(f"  ▍{wd.day:%Y-%m-%d}（{wd.weekday_zh}）")
            if wd.items:
                for idx, item in enumerate(wd.items, 1):
                    lines.append(
                        f"    {idx}. {_item_line(item.description, item.recipients_text)}"
                    )
            else:
                lines.append("    （無）")
    else:
        lines.append("  （本期間沒有工作內容）")
    lines.append("")
    lines.append("三、其他事項")
    if cfg.other_matters:
        for idx, matter in enumerate(cfg.other_matters, 1):
            lines.append(f"  {idx}. {matter}")
    else:
        lines.append("  （無）")
    lines.append("")
    lines.append("-" * 48)
    lines.append(f"產生時間：{report.generated_at:%Y-%m-%d %H:%M}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Markdown
# --------------------------------------------------------------------------- #
def render_markdown(report: WeeklyReport) -> str:
    cfg = report.config
    lines = []
    lines.append(f"# {cfg.title}")
    lines.append("")
    lines.append("## 一、基本資訊")
    lines.append("")
    lines.append("| 姓名 | 部門 | 日期 |")
    lines.append("| --- | --- | --- |")
    lines.append(f"| {cfg.name} | {cfg.department} | {cfg.period_text} |")
    lines.append("")
    lines.append("## 二、日常工作內容")
    lines.append("")
    if report.work_days:
        for wd in report.work_days:
            lines.append(f"### {wd.day:%Y-%m-%d}（{wd.weekday_zh}）")
            lines.append("")
            if wd.items:
                for item in wd.items:
                    lines.append(f"- {_item_line(item.description, item.recipients_text)}")
            else:
                lines.append("- （無）")
            lines.append("")
    else:
        lines.append("（本期間沒有工作內容）")
        lines.append("")
    lines.append("## 三、其他事項")
    lines.append("")
    if cfg.other_matters:
        for matter in cfg.other_matters:
            lines.append(f"- {matter}")
    else:
        lines.append("- （無）")
    lines.append("")
    lines.append(f"> 產生時間：{report.generated_at:%Y-%m-%d %H:%M}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
def render_html(report: WeeklyReport) -> str:
    cfg = report.config
    e = html.escape

    day_blocks = []
    if report.work_days:
        for wd in report.work_days:
            items_html = ""
            if wd.items:
                lis = "".join(
                    f"<li>{e(_item_line(item.description, item.recipients_text))}</li>"
                    for item in wd.items
                )
                items_html = f"<ul>{lis}</ul>"
            else:
                items_html = "<p class='empty'>（無）</p>"
            day_blocks.append(
                f"<div class='day'><h3>{e(f'{wd.day:%Y-%m-%d}')}"
                f"<span class='wd'>{e(wd.weekday_zh)}</span></h3>{items_html}</div>"
            )
    else:
        day_blocks.append("<p class='empty'>（本期間沒有工作內容）</p>")

    if cfg.other_matters:
        others_html = "<ul>" + "".join(
            f"<li>{e(m)}</li>" for m in cfg.other_matters
        ) + "</ul>"
    else:
        others_html = "<p class='empty'>（無）</p>"

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(cfg.title)} - {e(cfg.name)}</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{
    font-family: -apple-system, "PingFang TC", "Microsoft JhengHei",
      "Helvetica Neue", Arial, sans-serif;
    max-width: 820px; margin: 32px auto; padding: 0 20px; line-height: 1.7;
    color: #1f2933; background: #f7f9fc;
  }}
  .card {{
    background: #fff; border-radius: 14px; padding: 28px 32px;
    box-shadow: 0 8px 28px rgba(15, 23, 42, .08);
  }}
  h1 {{ margin: 0 0 4px; font-size: 26px; }}
  h2 {{
    font-size: 18px; margin: 28px 0 12px; padding-bottom: 8px;
    border-bottom: 2px solid #2563eb; color: #2563eb;
  }}
  h3 {{ font-size: 15px; margin: 18px 0 8px; color: #334155; }}
  .wd {{
    display: inline-block; margin-left: 8px; font-size: 12px;
    background: #e0ecff; color: #2563eb; padding: 2px 8px; border-radius: 999px;
    vertical-align: middle;
  }}
  table.info {{ border-collapse: collapse; width: 100%; }}
  table.info th, table.info td {{
    text-align: left; padding: 8px 12px; border: 1px solid #e2e8f0;
  }}
  table.info th {{ background: #f1f5f9; width: 90px; white-space: nowrap; }}
  ul {{ margin: 6px 0; padding-left: 22px; }}
  li {{ margin: 4px 0; }}
  .day {{ padding: 4px 0; }}
  .empty {{ color: #94a3b8; }}
  footer {{ margin-top: 24px; font-size: 12px; color: #94a3b8; }}
</style>
</head>
<body>
  <div class="card">
    <h1>{e(cfg.title)}</h1>
    <h2>一、基本資訊</h2>
    <table class="info">
      <tr><th>姓名</th><td>{e(cfg.name)}</td></tr>
      <tr><th>部門</th><td>{e(cfg.department)}</td></tr>
      <tr><th>日期</th><td>{e(cfg.period_text)}</td></tr>
    </table>
    <h2>二、日常工作內容</h2>
    {''.join(day_blocks)}
    <h2>三、其他事項</h2>
    {others_html}
    <footer>產生時間：{e(f'{report.generated_at:%Y-%m-%d %H:%M}')}</footer>
  </div>
</body>
</html>"""


RENDERERS: Dict[str, Callable[[WeeklyReport], str]] = {
    "text": render_text,
    "markdown": render_markdown,
    "md": render_markdown,
    "html": render_html,
}


def render(report: WeeklyReport, fmt: str = "text") -> str:
    fmt = fmt.lower()
    if fmt not in RENDERERS:
        raise ValueError(
            f"不支援的輸出格式：{fmt}（可用：{', '.join(sorted(RENDERERS))}）"
        )
    return RENDERERS[fmt](report)
