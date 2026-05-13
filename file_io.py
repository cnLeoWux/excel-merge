"""File I/O helpers for Excel Merge Tool."""

from __future__ import annotations

import logging
import zipfile
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


def _clean_string_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        col_str = str(col)
        if "订单" in col_str or "流水" in col_str:
            df[col] = df[col].fillna("").astype(str).str.strip('="\t ')
    return df


def _count_leading_comment_lines(file_path: str, encoding: str) -> int:
    skip_rows = 0
    with open(file_path, "r", encoding=encoding) as f:
        for line in f:
            normalized = line.lstrip("\ufeff").lstrip("ï»¿").strip()
            if normalized.startswith("#"):
                skip_rows += 1
            else:
                break
    return skip_rows


def _read_csv_with_fallback(file_path: str) -> pd.DataFrame:
    encodings = ["gbk", "utf-8", "gb2312", "latin-1", "utf-8-sig"]
    last_error = None

    for encoding in encodings:
        try:
            skip_rows = _count_leading_comment_lines(file_path, encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

        for sep in [",", ";", "\t"]:
            try:
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    skiprows=skip_rows,
                    header=0,
                    sep=sep,
                    engine="python",
                    on_bad_lines="warn",
                    dtype=str,
                )
                if df.shape[0] > 0 and df.shape[1] >= 2:
                    return _clean_string_columns(df)
            except (UnicodeDecodeError, pd.errors.ParserError) as exc:
                last_error = exc
                continue
            except Exception as exc:
                last_error = exc
                logger.warning(f"Error reading CSV file with encoding {encoding}: {exc}")
                continue

        try:
            df = pd.read_csv(
                file_path,
                encoding=encoding,
                skiprows=skip_rows,
                header=0,
                sep=None,
                engine="python",
                on_bad_lines="warn",
                skip_blank_lines=True,
                dtype=str,
            )
            if df.shape[0] > 0 and df.shape[1] >= 2:
                return _clean_string_columns(df)
        except (UnicodeDecodeError, pd.errors.ParserError) as exc:
            last_error = exc
            continue
        except Exception as exc:
            last_error = exc
            logger.warning(f"Error reading CSV file with encoding {encoding}: {exc}")

    raise ValueError(f"Unable to read CSV file '{file_path}' with encodings {encodings}: {last_error}")


def read_file_with_appropriate_method(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    ext = path.suffix.lower()

    if ext == ".csv":
        return _read_csv_with_fallback(str(file_path))

    if ext in [".xlsx", ".xls"]:
        if ext == ".xlsx":
            try:
                with zipfile.ZipFile(path, "r"):
                    engine = "openpyxl"
            except zipfile.BadZipFile:
                engine = "xlrd"
            except (ValueError, KeyError, OSError):
                engine = "openpyxl"
        else:
            engine = "xlrd"
        return pd.read_excel(file_path, dtype={"订单号": str, "商户订单号": str, "商务订单号": str}, engine=engine)

    try:
        return pd.read_excel(file_path, dtype={"订单号": str, "商户订单号": str, "商务订单号": str}, engine="openpyxl")
    except (ValueError, OSError, UnicodeDecodeError):
        for encoding in ["utf-8", "gbk", "gb2312", "latin-1"]:
            try:
                df = pd.read_csv(file_path, encoding=encoding, dtype=str)
                return _clean_string_columns(df)
            except UnicodeDecodeError:
                continue
        df = pd.read_csv(file_path, encoding="utf-8-sig", dtype=str)
        return _clean_string_columns(df)


def find_file_path(filename: str) -> Path:
    if Path(filename).exists():
        return Path(filename)
    excel_dir_path = Path("ExcelForHandel") / filename
    if excel_dir_path.exists():
        return excel_dir_path
    return Path(filename)


def write_result_file(df: pd.DataFrame, file_path: Path) -> None:
    if file_path.suffix.lower() == ".csv":
        df.to_csv(file_path, index=False, encoding="utf-8-sig")
    else:
        engine = "openpyxl" if file_path.suffix.lower() != ".xls" else None
        df.to_excel(file_path, index=False, engine=engine)
