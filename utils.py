"""Compatibility facade for legacy imports."""

from __future__ import annotations

from file_io import find_file_path, read_file_with_appropriate_method, write_result_file
from matching import (
    _classify_order_amount,
    _detect_business_order_column,
    _extract_payment_fee,
    _find_exact_match,
    _find_fallback_match,
    _is_business_type_compatible,
    _matches_exact_order,
    _matches_hyphen_fallback,
    extract_p_number,
    match_orders_by_p_number,
    process_excel_files,
)
from sales_report import (
    add_sales_report_period,
    filter_unmarked_and_generate_report,
    get_year_month,
    parse_date,
    process_sales_report_workflow,
)

import logging
import shutil
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


def auto_backup(file_path: str) -> Path:
    source = Path(file_path)
    if not source.exists():
        return source
    backup_dir = source.parent / "backup"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{source.stem}_backup_{timestamp}{source.suffix}"
    shutil.copy2(source, backup_path)
    logger.info(f"已备份文件到: {backup_path}")
    return backup_path
