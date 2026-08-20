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
from lme_daily.excel_com import get_workbook_open_count, reset_workbook_open_count
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
        help="跳過巨集，直接使用 vba_dir（working_dir\\yyyymmdd）裡既有的 yyyymmdd.xlsx",
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


def attach_run_dir_log(run_dir: Path) -> None:
    """每天一份 lme_daily.log 寫進 run_dir（最終報告資料夾）。"""
    log_path = run_dir / "lme_daily.log"
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    try:
        run_dir.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setFormatter(fmt)
        logging.getLogger().addHandler(handler)
        logger.info("當日日誌：%s", log_path)
    except OSError as exc:
        logger.warning("無法寫入當日日誌 %s：%s", log_path, exc)


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
        vba_dir, run_dir = config.ensure_run_dirs(as_of)
        logger.info("working_dir=%s", config.paths.working_dir)
        logger.info("vba_dir=%s（VBA 中繼檔，不受 output_dir 影響）", vba_dir)
        logger.info(
            "run_dir=%s（最終報告；output_dir=%s）",
            run_dir,
            config.paths.output_dir or "(空，與 vba_dir 相同)",
        )
        logger.info("chart.engine=%s  forward_months=%s", config.chart.engine, config.chart.forward_months)
        logger.info("vba.macro_name=%s  use_param_injection=%s", config.vba.macro_name, config.vba.use_param_injection)
        logger.info(
            "excel.reuse_running=%s  quit_on_exit=%s  bloomberg.source=%s",
            config.excel.reuse_running,
            config.excel.quit_on_exit,
            config.bloomberg.source,
        )

    attach_run_dir_log(run_dir)

    with log_step("計算上日日期與 3M date"):
        prev_date, three_m_date = calc_lme_dates(
            as_of,
            holiday_list=config.holidays,
            date_format=config.vba.date_format,
        )
        logger.info("as_of=%s  上日日期=%s  3M date=%s", as_of.isoformat(), prev_date, three_m_date)
        ibox_prev, ibox_three = calc_lme_dates(
            as_of,
            holiday_list=config.holidays,
            date_format=config.vba.inputbox_date_format,
        )
        logger.info("InputBox 將填入 上日=%s  3M=%s（格式 %s）", ibox_prev, ibox_three, config.vba.inputbox_date_format)
        logger.info("公休日筆數=%d（未列入者僅排除週末）", len(config.holidays))

    if dry_run:
        logger.info("dry-run：到此結束，不呼叫 Excel / Bloomberg")
        logger.info("若實際執行：VBA 中繼檔 → %s", config.step2_workbook_path(as_of))
        logger.info("若實際執行：最終報告 → %s", config.output_workbook_path(as_of))
        return None

    reset_workbook_open_count()

    step2_path = config.step2_workbook_path(as_of)
    if skip_vba:
        from lme_daily.vba_runner import relocate_step2_workbook, resolve_existing_step2

        found = resolve_existing_step2(config, as_of)
        logger.warning("略過 VBA（--skip-vba），改用既有檔案：%s", found or step2_path)
        if found is None:
            raise LMEAutomationError(
                f"--skip-vba 但找不到 VBA 中繼檔 {step2_path}"
                f"（也沒有 {config.step2_legacy_path(as_of)}）。"
                "遠期走勢圖與原始數據只讀 vba_dir 下的 yyyymmdd.xlsx，不會生成缺資料的報告。"
            )
        step2_path = relocate_step2_workbook(found, step2_path)
    else:
        from lme_daily.vba_runner import run_reference_macro

        with log_step("執行 VBA 產生 yyyymmdd.xlsx"):
            step2_path = run_reference_macro(
                config,
                as_of=as_of,
                prev_date=prev_date,
                three_m_date=three_m_date,
            )
            logger.info("步驟二產出（vba_dir）：%s", step2_path)

    try:
        step2_ready = step2_path.is_file()
    except OSError:
        step2_ready = False
    if not step2_ready:
        raise LMEAutomationError(
            f"找不到 VBA 中繼檔 {step2_path}。"
            "「遠期走勢圖」與「原始數據」只讀此檔，中斷而不產生缺資料報告。"
        )

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
            output_path=config.output_workbook_path(as_of),
        )

    pdf_path: Path | None = None
    with log_step("匯出合併版面 PDF"):
        pdf_path = _export_print_pdf(config, output_path, as_of=as_of)

    logger.info("本次 Workbook.Open 次數：%d（沿用已開啟檔不計；用來對照 Bloomberg 上鎖頻率）", get_workbook_open_count())
    logger.info("最終輸出檔案：%s", output_path.resolve())
    print(str(output_path.resolve()))
    if pdf_path is not None:
        logger.info("PDF：%s", pdf_path.resolve())
        print(str(pdf_path.resolve()))
    return output_path


def _export_print_pdf(config: AppConfig, workbook_path: Path, *, as_of: date) -> Path | None:
    """Windows + Excel 才把「合併版面」匯出 PDF；其他平台略過，xlsx 仍算成功。"""
    from lme_daily.excel_com import export_sheet_as_pdf
    from lme_daily.exceptions import ExcelComError
    from lme_daily.report_builder import SHEET_PRINT

    pdf_path = config.output_pdf_path(as_of)
    if sys.platform != "win32":
        logger.warning("非 Windows，略過 PDF 匯出（xlsx 已完成）：%s", pdf_path.name)
        return None
    try:
        return export_sheet_as_pdf(
            workbook_path=workbook_path,
            sheet_name=SHEET_PRINT,
            pdf_path=pdf_path,
            visible=config.excel.visible,
            display_alerts=config.excel.display_alerts,
            reuse_running=config.excel.reuse_running,
            quit_on_exit=config.excel.quit_on_exit,
            new_instance=config.excel.new_instance,
        )
    except ExcelComError as exc:
        raise LMEAutomationError(f"無法匯出 PDF：{exc}") from exc


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
