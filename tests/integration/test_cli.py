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


def test_cli_inplace_modification(workdir):
    """All merge results write back to the order file in place."""
    tmp_path, order, payment = workdir
    proc = run_cli([str(order), str(payment), "--quiet"], cwd=tmp_path)
    assert proc.returncode == 0, proc.stderr
    df = pd.read_excel(order)
    assert "支付手续费" in df.columns


def test_cli_json_envelope_success(workdir):
    """JSON success envelope contains output_file + statistics, nothing else."""
    tmp_path, order, payment = workdir
    proc = run_cli(
        [str(order), str(payment), "--json", "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["ok"] is True
    assert payload["error"] is None
    data = payload["data"]
    assert data["output_file"] == str(order)
    assert set(data.keys()) == {"output_file", "statistics"}
    assert "report_file" not in data
    assert "report_rows" not in data
    assert "warnings" not in data
    stats = data["statistics"]
    assert set(stats.keys()) == {"total_rows", "matched_rows", "match_rate"}


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


def test_cli_rejects_removed_output_flag(workdir):
    """`-o`/`--output` was removed; passing it must fail with usage error (exit 2)."""
    tmp_path, order, payment = workdir
    proc = run_cli(
        [str(order), str(payment), "-o", str(tmp_path / "result.xlsx"), "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 2
    assert "unrecognized arguments" in proc.stderr or "-o" in proc.stderr


def test_cli_rejects_removed_output_dir_flag(workdir):
    """`--output-dir` was removed; passing it must fail with usage error (exit 2)."""
    tmp_path, order, payment = workdir
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    proc = run_cli(
        [str(order), str(payment), "--output-dir", str(out_dir), "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 2
    assert "unrecognized arguments" in proc.stderr or "--output-dir" in proc.stderr


def test_cli_sales_report_workflow_no_files_emitted(workdir):
    """`--month` writes back in place and produces no report_*.xlsx anywhere."""
    tmp_path, order, payment = workdir

    # Snapshot every file path under tmp_path before invocation
    snapshot_before = {p for p in tmp_path.rglob("*") if p.is_file()}

    proc = run_cli(
        [str(order), str(payment), "--month", "202603", "--json", "--quiet"],
        cwd=tmp_path,
    )
    assert proc.returncode == 0, proc.stderr

    payload = json.loads(proc.stdout)
    assert payload["ok"] is True
    data = payload["data"]
    assert data["output_file"] == str(order)
    assert set(data.keys()) == {"output_file", "statistics"}
    assert "report_file" not in data
    assert "report_rows" not in data
    assert "warnings" not in data

    # No report_*.xlsx written by THIS invocation (snapshot diff under cwd)
    snapshot_after = {p for p in tmp_path.rglob("*") if p.is_file()}
    new_files = snapshot_after - snapshot_before
    assert not any(f.name.startswith("report_") for f in new_files), new_files
    assert not list(tmp_path.rglob("report_*.xlsx"))

    # Order file was actually updated with the period column
    df = pd.read_excel(order)
    assert "支付手续费" in df.columns
    assert "销售报表账期" in df.columns


def test_cli_write_failure_returns_processing_error(workdir):
    """If the order file cannot be written, exit 4 / processing_error (no warnings path)."""
    tmp_path, order, payment = workdir

    # Make the order file read-only so write_result_file fails on overwrite
    import os
    import stat

    os.chmod(order, stat.S_IREAD)
    try:
        proc = run_cli(
            [str(order), str(payment), "--month", "202603", "--json", "--quiet"],
            cwd=tmp_path,
        )
    finally:
        os.chmod(order, stat.S_IREAD | stat.S_IWRITE)

    # On some filesystems chmod 0o400 still allows the owner to overwrite via
    # openpyxl/pandas; only assert the contract when the OS actually blocked us.
    if proc.returncode == 0:
        pytest.skip("Filesystem did not honor read-only flag for owner overwrite")

    assert proc.returncode == 4
    payload = json.loads(proc.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "processing_error"
