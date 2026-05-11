"""Application workflow/service layer for Excel merge entry points.

This module coordinates existing business functions from ``utils.py`` and
returns explicit result objects for CLI, interactive, and HTTP adapters.
It intentionally does not change matching or sales-report business rules.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional
import re

import pandas as pd

from utils import (
    add_sales_report_period,
    process_excel_files,
    process_sales_report_workflow,
    read_file_with_appropriate_method,
    write_result_file,
)


@dataclass
class WorkflowResult:
    """Result for CLI and interactive workflows."""

    output_file: str
    dataframe: Optional[pd.DataFrame] = None
    statistics: Optional[Dict[str, Any]] = None
    report_dataframe: Optional[pd.DataFrame] = None
    message: Optional[str] = None


@dataclass
class WorkflowError(Exception):
    """Normalized service-level error."""

    code: str
    message: str
    exit_code: Optional[int] = None
    original: Optional[BaseException] = None

    def __str__(self) -> str:
        return self.message


@dataclass
class ApiWorkflowResult:
    """Result for HTTP API workflows."""

    result_path: Path
    download_name: str
    download_url: str
    statistics: Dict[str, Any]
    files: Dict[str, str]
    report_dataframe: Optional[pd.DataFrame] = None


def build_match_statistics(df: pd.DataFrame) -> Dict[str, Any]:
    total_rows = len(df)
    matched_rows = int(df["支付手续费"].notna().sum()) if "支付手续费" in df.columns else 0
    match_rate = f"{(matched_rows / total_rows * 100):.2f}%" if total_rows > 0 else "0.00%"
    return {
        "total_rows": total_rows,
        "matched_rows": matched_rows,
        "match_rate": match_rate,
    }


def build_mark_statistics(df: pd.DataFrame) -> Dict[str, Any]:
    total_rows = len(df)
    marked_rows = int(df["销售报表账期"].notna().sum()) if "销售报表账期" in df.columns else 0
    return {
        "total_rows": total_rows,
        "marked_rows": marked_rows,
    }


def build_full_workflow_statistics(df: pd.DataFrame) -> Dict[str, Any]:
    stats = build_match_statistics(df)
    stats["marked_rows"] = build_mark_statistics(df)["marked_rows"]
    return stats


def build_api_report_statistics(updated_df: pd.DataFrame, report_df: pd.DataFrame) -> Dict[str, Any]:
    stats = build_full_workflow_statistics(updated_df)
    stats["report_rows"] = int(len(report_df))
    # API historically uses one decimal place.
    total_rows = stats["total_rows"]
    matched_rows = stats["matched_rows"]
    stats["match_rate"] = f"{matched_rows / total_rows * 100:.1f}%" if total_rows > 0 else "0%"
    return stats


def validate_target_month_value(target_month: Optional[str]) -> bool:
    if target_month is None:
        return False
    if not re.match(r"^\d{6}$", target_month):
        return False

    year = int(target_month[:4])
    month = int(target_month[4:])
    return 2020 <= year <= 2099 and 1 <= month <= 12


def _ensure_valid_target_month(target_month: Optional[str], label: str = "target_month") -> None:
    if not target_month:
        raise WorkflowError("usage_error", f"{label} is required.", exit_code=2)
    if not validate_target_month_value(target_month):
        raise WorkflowError(
            "usage_error",
            f"{label} must be a valid YYYYMM value between 202001 and 209912.",
            exit_code=2,
        )


def _ensure_file_exists(file_path: str | Path, label: str = "file") -> None:
    path = Path(file_path)
    if not path.exists():
        raise WorkflowError(
            code="file_not_found",
            message=f"{label} '{file_path}' does not exist.",
            exit_code=3,
        )


def _normalize_exception(exc: BaseException) -> WorkflowError:
    if isinstance(exc, WorkflowError):
        return exc
    if isinstance(exc, FileNotFoundError):
        return WorkflowError(
            code="file_not_found",
            message=str(exc),
            exit_code=3,
            original=exc,
        )
    if isinstance(exc, (ValueError, OSError)):
        return WorkflowError(
            code="processing_error",
            message=str(exc),
            exit_code=4,
            original=exc,
        )
    return WorkflowError(
        code="processing_error",
        message=str(exc),
        exit_code=4,
        original=exc,
    )


def _write_in_place(df: pd.DataFrame, output_file: str | Path) -> None:
    _write_result(df, output_file, f"无法写回订单文件 '{output_file}'")


def _write_result(df: pd.DataFrame, output_file: str | Path, message_prefix: str) -> None:
    try:
        write_result_file(df, Path(output_file))
    except Exception as exc:  # pragma: no cover - exercised via adapters/tests
        raise WorkflowError(
            code="processing_error",
            message=f"{message_prefix}: {exc}",
            exit_code=4,
            original=exc,
        ) from exc


def _ensure_result_folder(result_folder: Path) -> None:
    try:
        result_folder.mkdir(exist_ok=True)
    except Exception as exc:
        raise WorkflowError(
            code="processing_error",
            message=f"无法创建结果目录 '{result_folder}': {exc}",
            exit_code=4,
            original=exc,
        ) from exc


def run_match_only(
    order_file: str | Path,
    payment_file: str | Path,
    *,
    verbose: bool = False,
    write_back: bool = True,
) -> WorkflowResult:
    try:
        _ensure_file_exists(order_file, "Order file")
        _ensure_file_exists(payment_file, "Payment file")
        result_df = process_excel_files(str(order_file), str(payment_file), verbose=verbose)
        if write_back:
            _write_in_place(result_df, order_file)
        return WorkflowResult(
            output_file=str(order_file),
            dataframe=result_df,
            statistics=build_match_statistics(result_df),
        )
    except WorkflowError:
        raise
    except Exception as exc:
        raise _normalize_exception(exc) from exc


def run_mark_only(
    order_file: str | Path,
    *,
    verbose: bool = False,
    write_back: bool = True,
) -> WorkflowResult:
    try:
        _ensure_file_exists(order_file, "Order file")
        order_df = read_file_with_appropriate_method(str(order_file))
        marked_df = add_sales_report_period(order_df, verbose=verbose)
        if write_back:
            _write_in_place(marked_df, order_file)
        return WorkflowResult(
            output_file=str(order_file),
            dataframe=marked_df,
            statistics=build_mark_statistics(marked_df),
        )
    except WorkflowError:
        raise
    except Exception as exc:
        raise _normalize_exception(exc) from exc


def run_sales_report(
    order_file: str | Path,
    payment_file: str | Path,
    target_month: str,
    *,
    verbose: bool = False,
    write_back: bool = True,
) -> WorkflowResult:
    try:
        _ensure_valid_target_month(target_month)
        _ensure_file_exists(order_file, "Order file")
        _ensure_file_exists(payment_file, "Payment file")
        updated_df, report_df = process_sales_report_workflow(
            str(order_file), str(payment_file), target_month, verbose=verbose
        )
        if write_back:
            _write_in_place(updated_df, order_file)
        return WorkflowResult(
            output_file=str(order_file),
            dataframe=updated_df,
            report_dataframe=report_df,
            statistics=build_full_workflow_statistics(updated_df),
        )
    except WorkflowError:
        raise
    except Exception as exc:
        raise _normalize_exception(exc) from exc


def prepare_api_merge(
    order_path: str | Path,
    payment_path: str | Path,
    original_order_filename: str,
    result_folder: str | Path,
    *,
    original_payment_filename: Optional[str] = None,
    month: Optional[str] = None,
    session_id: Optional[str] = None,
    timestamp: Optional[str] = None,
    verbose: bool = False,
) -> ApiWorkflowResult:
    session_id = session_id or "session"
    timestamp = timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    result_folder = Path(result_folder)
    original_payment_filename = original_payment_filename or Path(payment_path).name

    try:
        _ensure_file_exists(order_path, "Order file")
        _ensure_file_exists(payment_path, "Payment file")
        _ensure_result_folder(result_folder)
        if month:
            _ensure_valid_target_month(month, "month")
            updated_df, report_df = process_sales_report_workflow(
                str(order_path), str(payment_path), month, verbose=verbose
            )
            if report_df.empty:
                raise WorkflowError(
                    "processing_error",
                    "Report generation produced no data",
                    exit_code=4,
                )

            result_filename = f"report_{month}_{session_id}.xlsx"
            result_path = result_folder / result_filename
            _write_result(report_df, result_path, f"无法保存结果文件 '{result_path}'")
            return ApiWorkflowResult(
                result_path=result_path,
                download_name=f"report_{month}.xlsx",
                download_url=f"/download/{result_filename}",
                statistics=build_api_report_statistics(updated_df, report_df),
                files={
                    "order": original_order_filename,
                    "payment": original_payment_filename,
                    "result": result_filename,
                },
                report_dataframe=report_df,
            )

        result_df = process_excel_files(str(order_path), str(payment_path), verbose=verbose)
        original_ext = Path(original_order_filename).suffix
        result_filename = f"merged_result_{timestamp}_{session_id}{original_ext}"
        result_path = result_folder / result_filename
        _write_result(result_df, result_path, f"无法保存结果文件 '{result_path}'")
        return ApiWorkflowResult(
            result_path=result_path,
            download_name=f"merged_{original_order_filename}",
            download_url=f"/download/{result_filename}",
            statistics=_api_match_statistics(result_df),
            files={
                "order": original_order_filename,
                "payment": original_payment_filename,
                "result": result_filename,
            },
        )
    except WorkflowError:
        raise
    except Exception as exc:
        raise _normalize_exception(exc) from exc


def _api_match_statistics(df: pd.DataFrame) -> Dict[str, Any]:
    stats = build_match_statistics(df)
    total_rows = stats["total_rows"]
    matched_rows = stats["matched_rows"]
    stats["match_rate"] = f"{matched_rows / total_rows * 100:.1f}%" if total_rows > 0 else "0%"
    return stats
