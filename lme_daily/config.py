"""讀取並驗證 config.yaml，組合路徑、載入公休日。"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

from lme_daily.exceptions import ConfigError

logger = logging.getLogger(__name__)

CONFIG_FILENAME = "config.yaml"


@dataclass(frozen=True)
class PathsConfig:
    working_dir: Path
    ref_workbook: Path
    bbg_workbook: Path
    output_prefix: str
    holidays_file: Path | None


@dataclass(frozen=True)
class VbaConfig:
    macro_name: str
    use_param_injection: bool
    date_format: str
    output_timeout_seconds: float
    poll_interval_seconds: float
    inputbox_timeout_seconds: float


@dataclass(frozen=True)
class ExcelUiConfig:
    visible: bool
    display_alerts: bool


@dataclass(frozen=True)
class BloombergConfig:
    copy_range: str
    bbg_sheet_name: str
    refresh_wait_seconds: float
    calculation_timeout_seconds: float


@dataclass(frozen=True)
class ChartConfig:
    forward_months: int
    engine: str
    image_width: int
    image_height: int


@dataclass(frozen=True)
class LoggingConfig:
    level: str
    file: str | None


@dataclass(frozen=True)
class AppConfig:
    source_path: Path
    paths: PathsConfig
    vba: VbaConfig
    excel: ExcelUiConfig
    bloomberg: BloombergConfig
    chart: ChartConfig
    holidays: frozenset[date]
    logging: LoggingConfig

    def output_workbook_path(self, as_of: date) -> Path:
        name = f"{self.paths.output_prefix}{as_of.strftime('%Y%m%d')}.xlsx"
        return self.paths.working_dir / name

    def step2_workbook_path(self, as_of: date) -> Path:
        return self.paths.working_dir / f"{as_of.strftime('%Y%m%d')}.xlsx"


def discover_config_path(explicit: str | Path | None = None) -> Path:
    """尋找 config.yaml。

    優先順序：``--config`` → 目前工作目錄 → 專案根目錄（本套件上一層）→ 套件目錄。
    """
    if explicit is not None:
        path = Path(explicit).expanduser().resolve()
        if not path.is_file():
            raise ConfigError(f"指定的設定檔不存在：{path}")
        return path

    candidates = [
        Path.cwd() / CONFIG_FILENAME,
        Path(__file__).resolve().parent.parent / CONFIG_FILENAME,
        Path(__file__).resolve().parent / CONFIG_FILENAME,
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    looked = "\n  ".join(str(p) for p in candidates)
    raise ConfigError(
        "找不到 config.yaml。請將檔案放在執行目錄或專案根目錄，或用 --config 指定。\n"
        f"已嘗試：\n  {looked}"
    )


def _require_mapping(raw: Any, context: str) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise ConfigError(f"{context} 必須是 mapping，實際為 {type(raw).__name__}")
    return raw


def _get(section: dict[str, Any], key: str, *, context: str, default: Any = ...):
    if key in section and section[key] is not None:
        return section[key]
    if default is not ...:
        return default
    raise ConfigError(f"設定缺少必填欄位：{context}.{key}")


def _as_bool(value: Any, context: str) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str) and value.strip().lower() in {"true", "false", "1", "0", "yes", "no"}:
        return value.strip().lower() in {"true", "1", "yes"}
    raise ConfigError(f"{context} 必須是布林值，收到 {value!r}")


def _as_path(value: Any, context: str) -> Path:
    if not isinstance(value, (str, Path)) or str(value).strip() == "":
        raise ConfigError(f"{context} 必須是非空路徑字串")
    return Path(str(value)).expanduser()


def _parse_holiday_token(token: Any, *, source: str) -> date:
    text = str(token).strip()
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ConfigError(f"{source} 的公休日必須是 YYYY-MM-DD，收到 {token!r}") from exc


def _load_holidays_file(path: Path) -> set[date]:
    if not path.is_file():
        raise ConfigError(f"公休日檔不存在：{path}")
    try:
        payload = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        raise ConfigError(f"無法解析公休日檔 {path}：{exc}") from exc

    if isinstance(payload, list):
        raw_dates = payload
    elif isinstance(payload, dict):
        raw_dates = payload.get("holidays") or payload.get("dates") or []
    else:
        raise ConfigError(f"{path} 格式無法識別（需為 list 或含 holidays/dates 的 mapping）")

    if not isinstance(raw_dates, list):
        raise ConfigError(f"{path} 的 holidays 必須是清單")
    return {_parse_holiday_token(item, source=str(path)) for item in raw_dates}


def load_holidays(
    inline_dates: Any,
    holidays_file: Path | None,
) -> frozenset[date]:
    collected: set[date] = set()
    if inline_dates:
        if not isinstance(inline_dates, list):
            raise ConfigError("holidays.dates 必須是清單")
        collected.update(_parse_holiday_token(item, source="config.holidays.dates") for item in inline_dates)
    if holidays_file is not None:
        collected.update(_load_holidays_file(holidays_file))
    logger.info("已載入 %d 個公休日", len(collected))
    return frozenset(collected)


def load_config(path: str | Path | None = None) -> AppConfig:
    """讀取 YAML、組合 working_dir 與各檔名，並在執行前檢查檔案是否存在。"""
    config_path = discover_config_path(path)
    logger.info("讀取設定檔：%s", config_path)
    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        raise ConfigError(f"無法解析 {config_path}：{exc}") from exc
    raw = _require_mapping(raw, "config")

    paths_raw = _require_mapping(_get(raw, "paths", context="config"), "paths")
    vba_raw = _require_mapping(_get(raw, "vba", context="config"), "vba")
    bbg_raw = _require_mapping(_get(raw, "bloomberg", context="config"), "bloomberg")
    chart_raw = _require_mapping(_get(raw, "chart", context="config"), "chart")
    excel_raw = _require_mapping(raw.get("excel") or {}, "excel")
    holidays_raw = _require_mapping(raw.get("holidays") or {}, "holidays")
    logging_raw = _require_mapping(raw.get("logging") or {}, "logging")

    working_dir = _as_path(_get(paths_raw, "working_dir", context="paths"), "paths.working_dir")
    if not working_dir.is_absolute():
        working_dir = (config_path.parent / working_dir).resolve()
    else:
        working_dir = working_dir.resolve()

    ref_name = str(_get(paths_raw, "ref_workbook_name", context="paths"))
    bbg_name = str(_get(paths_raw, "bbg_workbook_name", context="paths"))
    output_prefix = str(_get(paths_raw, "output_prefix", context="paths"))

    holidays_file_value = paths_raw.get("holidays_file") or holidays_raw.get("file")
    holidays_file: Path | None = None
    if holidays_file_value:
        holidays_file = _as_path(holidays_file_value, "paths.holidays_file")
        if not holidays_file.is_absolute():
            holidays_file = (config_path.parent / holidays_file).resolve()

    paths = PathsConfig(
        working_dir=working_dir,
        ref_workbook=working_dir / ref_name,
        bbg_workbook=working_dir / bbg_name,
        output_prefix=output_prefix,
        holidays_file=holidays_file,
    )

    engine = str(_get(chart_raw, "engine", context="chart", default="matplotlib")).strip().lower()
    if engine not in {"matplotlib", "xlsxwriter"}:
        raise ConfigError("chart.engine 只能是 matplotlib 或 xlsxwriter")

    config = AppConfig(
        source_path=config_path,
        paths=paths,
        vba=VbaConfig(
            macro_name=str(_get(vba_raw, "macro_name", context="vba")),
            use_param_injection=_as_bool(
                _get(vba_raw, "use_param_injection", context="vba"),
                "vba.use_param_injection",
            ),
            date_format=str(_get(vba_raw, "date_format", context="vba", default="%Y/%m/%d")),
            output_timeout_seconds=float(
                _get(vba_raw, "output_timeout_seconds", context="vba", default=180)
            ),
            poll_interval_seconds=float(
                _get(vba_raw, "poll_interval_seconds", context="vba", default=2)
            ),
            inputbox_timeout_seconds=float(
                _get(vba_raw, "inputbox_timeout_seconds", context="vba", default=60)
            ),
        ),
        excel=ExcelUiConfig(
            visible=_as_bool(_get(excel_raw, "visible", context="excel", default=True), "excel.visible"),
            display_alerts=_as_bool(
                _get(excel_raw, "display_alerts", context="excel", default=False),
                "excel.display_alerts",
            ),
        ),
        bloomberg=BloombergConfig(
            copy_range=str(_get(bbg_raw, "copy_range", context="bloomberg")),
            bbg_sheet_name=str(_get(bbg_raw, "bbg_sheet_name", context="bloomberg")),
            refresh_wait_seconds=float(_get(bbg_raw, "refresh_wait_seconds", context="bloomberg")),
            calculation_timeout_seconds=float(
                _get(bbg_raw, "calculation_timeout_seconds", context="bloomberg", default=120)
            ),
        ),
        chart=ChartConfig(
            forward_months=int(_get(chart_raw, "forward_months", context="chart")),
            engine=engine,
            image_width=int(_get(chart_raw, "image_width", context="chart", default=480)),
            image_height=int(_get(chart_raw, "image_height", context="chart", default=280)),
        ),
        holidays=load_holidays(holidays_raw.get("dates") or [], holidays_file),
        logging=LoggingConfig(
            level=str(_get(logging_raw, "level", context="logging", default="INFO")),
            file=(str(logging_raw["file"]).strip() or None) if logging_raw.get("file") else None,
        ),
    )
    return config


def validate_required_files(
    config: AppConfig,
    *,
    require_workbooks: bool = True,
    require_ref_workbook: bool | None = None,
    require_bbg_workbook: bool | None = None,
) -> None:
    """執行前檢查工作資料夾與來源工作簿是否存在；不存在就報錯退出。"""
    if require_ref_workbook is None:
        require_ref_workbook = require_workbooks
    if require_bbg_workbook is None:
        require_bbg_workbook = require_workbooks

    missing: list[str] = []
    if not config.paths.working_dir.is_dir():
        missing.append(f"工作資料夾不存在：{config.paths.working_dir}")
    if require_ref_workbook and not config.paths.ref_workbook.is_file():
        missing.append(f"參考工作簿不存在：{config.paths.ref_workbook}")
    if require_bbg_workbook and not config.paths.bbg_workbook.is_file():
        missing.append(f"Bloomberg 工作簿不存在：{config.paths.bbg_workbook}")
    if config.paths.holidays_file is not None and not config.paths.holidays_file.is_file():
        missing.append(f"公休日檔不存在：{config.paths.holidays_file}")
    if missing:
        raise ConfigError("執行前檔案檢查失敗：\n  " + "\n  ".join(missing))
    logger.info("檔案檢查通過：working_dir=%s", config.paths.working_dir)
    if require_ref_workbook:
        logger.info("參考工作簿：%s", config.paths.ref_workbook)
    if require_bbg_workbook:
        logger.info("BBG 工作簿：%s", config.paths.bbg_workbook)
