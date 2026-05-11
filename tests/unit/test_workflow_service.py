from pathlib import Path

import pandas as pd
import pytest

from workflow_service import (
    build_api_report_statistics,
    build_full_workflow_statistics,
    build_mark_statistics,
    WorkflowError,
    build_match_statistics,
    prepare_api_merge,
    run_mark_only,
    run_match_only,
    run_sales_report,
    validate_target_month_value,
)


def test_statistics_helpers():
    match_df = pd.DataFrame({"支付手续费": [-10.0, None, 0.0]})
    mark_df = pd.DataFrame({"销售报表账期": ["全退", None, "已取消"]})
    report_df = pd.DataFrame({"a": [1, 2]})

    assert build_match_statistics(match_df) == {
        "total_rows": 3,
        "matched_rows": 2,
        "match_rate": "66.67%",
    }
    assert build_mark_statistics(mark_df) == {"total_rows": 3, "marked_rows": 2}

    full_stats = build_full_workflow_statistics(mark_df)
    assert full_stats["total_rows"] == 3
    assert full_stats["matched_rows"] == 0
    assert full_stats["marked_rows"] == 2

    api_stats = build_api_report_statistics(match_df, report_df)
    assert api_stats["total_rows"] == 3
    assert api_stats["matched_rows"] == 2
    assert api_stats["match_rate"] == "66.7%"
    assert api_stats["report_rows"] == 2


def test_run_match_only_writes_in_place(sample_data_dir, tmp_path):
    order = tmp_path / "orders.xlsx"
    payment = tmp_path / "payments.csv"
    order.write_bytes((sample_data_dir / "orders.xlsx").read_bytes())
    payment.write_bytes((sample_data_dir / "payments.csv").read_bytes())

    result = run_match_only(order, payment)

    assert result.output_file == str(order)
    assert result.statistics is not None
    assert result.statistics["matched_rows"] == 6
    assert order.exists()
    assert "支付手续费" in pd.read_excel(order).columns


def test_run_mark_only_writes_in_place(sample_data_dir, tmp_path):
    order = tmp_path / "orders.xlsx"
    order.write_bytes((sample_data_dir / "orders.xlsx").read_bytes())

    result = run_mark_only(order)

    assert result.output_file == str(order)
    assert result.statistics is not None
    assert result.statistics["marked_rows"] == 3
    assert "销售报表账期" in pd.read_excel(order).columns


def test_run_sales_report_does_not_create_report_file(sample_data_dir, tmp_path):
    order = tmp_path / "orders.xlsx"
    payment = tmp_path / "payments.csv"
    order.write_bytes((sample_data_dir / "orders.xlsx").read_bytes())
    payment.write_bytes((sample_data_dir / "payments.csv").read_bytes())

    before = {p for p in tmp_path.rglob("*") if p.is_file()}
    result = run_sales_report(order, payment, "202603")
    after = {p for p in tmp_path.rglob("*") if p.is_file()}

    assert result.output_file == str(order)
    assert result.report_dataframe is not None
    assert result.statistics is not None
    assert result.statistics["marked_rows"] >= 0
    assert not any(p.name.startswith("report_") for p in after - before)
    assert not list(tmp_path.rglob("report_*.xlsx"))


def test_prepare_api_merge_metadata(sample_data_dir, tmp_path):
    result = prepare_api_merge(
        sample_data_dir / "orders.xlsx",
        sample_data_dir / "payments.csv",
        "orders.xlsx",
        tmp_path,
        original_payment_filename="payments.csv",
        month="202603",
        session_id="abc12345",
        timestamp="20260301_120000",
    )

    assert result.result_path == tmp_path / "report_202603_abc12345.xlsx"
    assert result.result_path.exists()
    assert result.download_url == "/download/report_202603_abc12345.xlsx"
    assert result.files == {
        "order": "orders.xlsx",
        "payment": "payments.csv",
        "result": "report_202603_abc12345.xlsx",
    }
    assert result.statistics["report_rows"] == 2


def test_missing_file_normalization(sample_data_dir, tmp_path):
    payment = tmp_path / "payments.csv"
    payment.write_bytes((sample_data_dir / "payments.csv").read_bytes())

    with pytest.raises(WorkflowError) as excinfo:
        run_match_only(tmp_path / "missing.xlsx", payment)

    assert excinfo.value.code == "file_not_found"
    assert excinfo.value.exit_code == 3


def test_invalid_month_normalization_and_short_circuit(monkeypatch, sample_data_dir, tmp_path):
    called = {"sales": False}

    def fail_if_called(*args, **kwargs):
        called["sales"] = True
        raise AssertionError("core workflow should not be called")

    monkeypatch.setattr("workflow_service.process_sales_report_workflow", fail_if_called)

    with pytest.raises(WorkflowError) as excinfo:
        run_sales_report(sample_data_dir / "orders.xlsx", sample_data_dir / "payments.csv", "202613")

    assert excinfo.value.code == "usage_error"
    assert excinfo.value.exit_code == 2
    assert called["sales"] is False

    called_api = {"sales": False}

    def fail_if_called_api(*args, **kwargs):
        called_api["sales"] = True
        raise AssertionError("core workflow should not be called")

    monkeypatch.setattr("workflow_service.process_sales_report_workflow", fail_if_called_api)
    with pytest.raises(WorkflowError) as excinfo_api:
        prepare_api_merge(
            sample_data_dir / "orders.xlsx",
            sample_data_dir / "payments.csv",
            "orders.xlsx",
            tmp_path,
            month="202613",
        )

    assert excinfo_api.value.code == "usage_error"
    assert excinfo_api.value.exit_code == 2
    assert called_api["sales"] is False


def test_write_failure_normalization(monkeypatch, sample_data_dir, tmp_path):
    df = pd.DataFrame({"订单号": ["1"], "支付手续费": [-1.0]})
    monkeypatch.setattr("workflow_service.process_excel_files", lambda *args, **kwargs: df)
    monkeypatch.setattr("workflow_service.write_result_file", lambda *args, **kwargs: (_ for _ in ()).throw(OSError("disk full")))

    with pytest.raises(WorkflowError) as excinfo:
        run_match_only(sample_data_dir / "orders.xlsx", sample_data_dir / "payments.csv")

    assert excinfo.value.code == "processing_error"
    assert excinfo.value.exit_code == 4


def test_prepare_api_merge_result_folder_failure_is_processing_error(sample_data_dir, tmp_path):
    result_folder = tmp_path / "missing_parent" / "results"

    with pytest.raises(WorkflowError) as excinfo:
        prepare_api_merge(
            sample_data_dir / "orders.xlsx",
            sample_data_dir / "payments.csv",
            "orders.xlsx",
            result_folder,
        )

    assert excinfo.value.code == "processing_error"
    assert excinfo.value.exit_code == 4


def test_prepare_api_merge_write_file_not_found_is_processing_error(monkeypatch, sample_data_dir, tmp_path):
    result_df = pd.DataFrame({"订单号": ["1"], "支付手续费": [-1.0]})
    monkeypatch.setattr("workflow_service.process_excel_files", lambda *args, **kwargs: result_df)
    monkeypatch.setattr(
        "workflow_service.write_result_file",
        lambda *args, **kwargs: (_ for _ in ()).throw(FileNotFoundError("missing output parent")),
    )

    with pytest.raises(WorkflowError) as excinfo:
        prepare_api_merge(
            sample_data_dir / "orders.xlsx",
            sample_data_dir / "payments.csv",
            "orders.xlsx",
            tmp_path,
        )

    assert excinfo.value.code == "processing_error"
    assert excinfo.value.exit_code == 4


def test_build_api_report_statistics_includes_marked_rows():
    updated_df = pd.DataFrame({
        "支付手续费": [-1.0, None],
        "销售报表账期": ["全退", None],
    })
    report_df = pd.DataFrame({"订单号": [1, 2, 3]})

    stats = build_api_report_statistics(updated_df, report_df)
    assert stats["matched_rows"] == 1
    assert stats["marked_rows"] == 1
    assert stats["report_rows"] == 3


def test_prepare_api_merge_match_only_metadata(monkeypatch, sample_data_dir, tmp_path):
    result_df = pd.DataFrame({"订单号": ["1"], "支付手续费": [-1.0]})
    monkeypatch.setattr("workflow_service.process_excel_files", lambda *args, **kwargs: result_df)
    monkeypatch.setattr("workflow_service.write_result_file", lambda *args, **kwargs: None)

    result = prepare_api_merge(
        sample_data_dir / "orders.xlsx",
        sample_data_dir / "payments.csv",
        "orders.xlsx",
        tmp_path,
        original_payment_filename="payments.csv",
        session_id="abc12345",
        timestamp="20260301_120000",
    )

    assert result.download_url == "/download/merged_result_20260301_120000_abc12345.xlsx"
    assert result.files == {
        "order": "orders.xlsx",
        "payment": "payments.csv",
        "result": "merged_result_20260301_120000_abc12345.xlsx",
    }
    assert result.statistics["matched_rows"] == 1


def test_validate_target_month_value():
    assert validate_target_month_value("202603") is True
    assert validate_target_month_value("202613") is False
    assert validate_target_month_value(None) is False
