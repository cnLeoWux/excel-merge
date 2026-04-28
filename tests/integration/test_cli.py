import json
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CLI_PATH = PROJECT_ROOT / "cli.py"


def run_cli(args, cwd=None):
    return subprocess.run(
        [sys.executable, str(CLI_PATH), *args],
        capture_output=True,
        text=True,
        cwd=str(cwd) if cwd else str(PROJECT_ROOT),
    )


@pytest.fixture
def workdir(sample_data_dir, tmp_path):
    order_src = sample_data_dir / "orders.xlsx"
    payment_src = sample_data_dir / "payments.csv"
    order_dst = tmp_path / "orders.xlsx"
    payment_dst = tmp_path / "payments.csv"
    order_dst.write_bytes(order_src.read_bytes())
    payment_dst.write_bytes(payment_src.read_bytes())
    return tmp_path, order_dst, payment_dst


def test_cli_basic_match_with_output_file(workdir):
    tmp_path, order, payment = workdir
    out = tmp_path / "result.xlsx"
    proc = run_cli([str(order), str(payment), "-o", str(out), "--quiet"], cwd=tmp_path)
    assert proc.returncode == 0, proc.stderr
    assert out.exists()
    df = pd.read_excel(out)
    assert "支付手续费" in df.columns


def test_cli_inplace_modification(workdir):
    tmp_path, order, payment = workdir
    proc = run_cli([str(order), str(payment), "--quiet"], cwd=tmp_path)
    assert proc.returncode == 0, proc.stderr
    df = pd.read_excel(order)
    assert "支付手续费" in df.columns


def test_cli_json_envelope_success(workdir):
    tmp_path, order, payment = workdir
    out = tmp_path / "result.xlsx"
    proc = run_cli(
        [str(order), str(payment), "-o", str(out), "--json", "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["ok"] is True
    assert payload["error"] is None
    assert "output_file" in payload["data"]
    assert "statistics" in payload["data"]
    stats = payload["data"]["statistics"]
    assert "total_rows" in stats
    assert "matched_rows" in stats
    assert "match_rate" in stats


def test_cli_file_not_found_exit_code(tmp_path):
    proc = run_cli(
        ["does_not_exist.xlsx", "also_missing.xlsx", "--json", "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 3
    payload = json.loads(proc.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "file_not_found"


def test_cli_invalid_arguments_exit_code(tmp_path):
    proc = run_cli([], cwd=tmp_path)
    assert proc.returncode == 2


def test_cli_sales_report_workflow(workdir):
    tmp_path, order, payment = workdir
    out = tmp_path / "updated.xlsx"
    report_dir = tmp_path / "reports"
    report_dir.mkdir()
    proc = run_cli(
        [
            str(order),
            str(payment),
            "-o",
            str(out),
            "--month",
            "202603",
            "--output-dir",
            str(report_dir),
            "--json",
            "--quiet",
        ],
        cwd=tmp_path,
    )
    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["ok"] is True
    assert "report_file" in payload["data"]
    assert "report_rows" in payload["data"]
    report_files = list(report_dir.glob("report_*.xlsx"))
    assert len(report_files) == 1
