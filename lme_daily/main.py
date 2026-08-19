"""LME 每日報價自動化進入點。

流程：讀 config → 算日期 → 跑 VBA 產生 yyyymmdd.xlsx → 抓 BBG 數據 → 生成最終報告。
任一步失敗會印出明確錯誤並以非零狀態碼結束，不會靜默跳過。
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

from lme_daily.config import AppConfig, load_config, validate_required_files
from lme_daily.dates import calc_lme_dates
from lme_daily.exceptions import LMEAutomationError

logger = logging.getLogger("lme_daily")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="LME 每日報價自動化（Excel COM + Bloomberg + 遠期曲線報告）",
    )
    parser.add_argument("--config", type=Path, default=None, help="config.yaml 路徑")
    parser.add_argument(
        "--as-of",
        dest="as_of",
        default=None,
        help="覆寫執行日（YYYY-MM-DD），預設為系統當天",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只檢查設定、計算日期，不開啟 Excel / Bloomberg",
    )
    parser.add_argument(
        "--skip-vba",
        action="store_true",
        help="跳過巨集，直接使用 working_dir 裡既有的 yyyymmdd.xlsx",
    )
    parser.add_argument(
        "--skip-bbg",
        action="store_true",
        help="跳過 Bloomberg 刷新（BBG快照會是空表）",
    )
    return parser.parse_args(argv)


def parse_as_of(text: str | None) -> date:
    if not text:
        return date.today()
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError as exc:
        raise LMEAutomationError(f"--as-of 必須是 YYYY-MM-DD，收到 {text!r}") from exc


def setup_logging(config: AppConfig) -> None:
    level_name = (config.logging.level or "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(level)

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    root.addHandler(console)

    if config.logging.file:
        log_path = Path(config.logging.file)
        if not log_path.is_absolute():
            log_path = config.paths.working_dir / log_path
        try:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            file_handler = logging.FileHandler(log_path, encoding="utf-8")
            file_handler.setFormatter(fmt)
            root.addHandler(file_handler)
            logger.info("日誌檔：%s", log_path)
        except OSError as exc:
            logger.warning("無法寫入日誌檔 %s：%s（改只輸出到 console）", log_path, exc)


@contextmanager
def log_step(name: str):
    started = datetime.now()
    t0 = time.perf_counter()
    logger.info("=== START %s  at %s ===", name, started.isoformat(timespec="seconds"))
    try:
        yield
    except Exception:
        elapsed = time.perf_counter() - t0
        logger.exception(
            "=== FAILED %s  at %s（耗時 %.1fs）===",
            name,
            datetime.now().isoformat(timespec="seconds"),
            elapsed,
        )
        raise
    elapsed = time.perf_counter() - t0
    logger.info(
        "=== END %s  at %s（耗時 %.1fs）===",
        name,
        datetime.now().isoformat(timespec="seconds"),
        elapsed,
    )


def run(config: AppConfig, *, as_of: date, dry_run: bool, skip_vba: bool, skip_bbg: bool) -> Path | None:
    with log_step("讀取設定並檢查檔案"):
        validate_required_files(
            config,
            require_ref_workbook=not dry_run and not skip_vba,
            require_bbg_workbook=not dry_run and not skip_bbg,
        )
        logger.info("working_dir=%s", config.paths.working_dir)
        logger.info("chart.engine=%s  forward_months=%s", config.chart.engine, config.chart.forward_months)
        logger.info("vba.macro_name=%s  use_param_injection=%s", config.vba.macro_name, config.vba.use_param_injection)

    with log_step("計算上日日期與 3M date"):
        prev_date, three_m_date = calc_lme_dates(
            as_of,
            holiday_list=config.holidays,
            date_format=config.vba.date_format,
        )
        logger.info("as_of=%s  上日日期=%s  3M date=%s", as_of.isoformat(), prev_date, three_m_date)
        logger.info("公休日筆數=%d（未列入者僅排除週末）", len(config.holidays))

    if dry_run:
        logger.info("dry-run：到此結束，不呼叫 Excel / Bloomberg")
        return None

    step2_path = config.step2_workbook_path(as_of)
    if skip_vba:
        logger.warning("略過 VBA（--skip-vba），改用既有檔案：%s", step2_path)
        if not step2_path.is_file():
            raise LMEAutomationError(f"--skip-vba 但找不到 {step2_path}")
    else:
        from lme_daily.vba_runner import run_reference_macro

        with log_step("執行 VBA 產生 yyyymmdd.xlsx"):
            step2_path = run_reference_macro(
                config,
                as_of=as_of,
                prev_date=prev_date,
                three_m_date=three_m_date,
            )
            logger.info("步驟二產出：%s", step2_path)

    if skip_bbg:
        logger.warning("略過 Bloomberg 刷新（--skip-bbg），BBG快照將為空")
        bbg_values: tuple[tuple[object, ...], ...] = ()
        bbg_formats: tuple[tuple[str, ...], ...] = ()
    else:
        from lme_daily.bbg_fetch import fetch_bloomberg_snapshot

        with log_step("刷新 Bloomberg 並讀取快照"):
            bbg_values, bbg_formats = fetch_bloomberg_snapshot(config)

    from lme_daily.report_builder import build_report

    output_path: Path
    with log_step("產生最終報告 Excel"):
        output_path = build_report(
            config,
            as_of=as_of,
            step2_path=step2_path,
            bbg_values=bbg_values,
            bbg_formats=bbg_formats,
        )

    logger.info("最終輸出檔案：%s", output_path.resolve())
    print(str(output_path.resolve()))
    return output_path


def run_cli(argv: list[str] | None = None) -> tuple[int, Path | None]:
    """執行 CLI 流程，回傳 ``(exit_code, 報告路徑)``。dry-run 時路徑為 None。"""
    from lme_daily.bootstrap import configure_windows_console

    configure_windows_console()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    args = parse_args(argv)
    try:
        as_of = parse_as_of(args.as_of)
        config = load_config(args.config)
        setup_logging(config)
        logger.info("LME 每日報價自動化啟動（as_of=%s）", as_of.isoformat())
        output = run(
            config,
            as_of=as_of,
            dry_run=args.dry_run,
            skip_vba=args.skip_vba,
            skip_bbg=args.skip_bbg,
        )
    except LMEAutomationError as exc:
        logging.getLogger("lme_daily").error("%s", exc)
        print(f"錯誤：{exc}", file=sys.stderr)
        return 1, None
    except KeyboardInterrupt:
        print("已中斷", file=sys.stderr)
        return 130, None
    except Exception as exc:  # 未預期例外也要明確印出，不可沉默
        logging.getLogger("lme_daily").exception("未預期錯誤：%s", exc)
        print(f"未預期錯誤：{exc}", file=sys.stderr)
        return 1, None
    return 0, output


def main(argv: list[str] | None = None) -> int:
    code, _output = run_cli(argv)
    return code


if __name__ == "__main__":
    sys.exit(main())
