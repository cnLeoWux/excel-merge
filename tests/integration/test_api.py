import io
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from excel_merge_api import app as flask_app


@pytest.fixture
def client(tmp_path, monkeypatch):
    import excel_merge_api as api_module

    upload = tmp_path / "uploads"
    result = tmp_path / "results"
    upload.mkdir()
    result.mkdir()
    monkeypatch.setattr(api_module, "UPLOAD_FOLDER", upload)
    monkeypatch.setattr(api_module, "RESULT_FOLDER", result)

    flask_app.config["TESTING"] = True
    with flask_app.test_client() as c:
        yield c


def _multipart(order_path, payment_path, month=None):
    data = {
        "order_file": (open(order_path, "rb"), order_path.name),
        "payment_file": (open(payment_path, "rb"), payment_path.name),
    }
    if month:
        data["month"] = month
    return data


def test_health_endpoint(client):
    resp = client.get("/health")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["status"] == "healthy"
    assert body["service"] == "excel-merge-api"


def test_index_page(client):
    resp = client.get("/")
    assert resp.status_code == 200
    assert b"Excel Merge Tool" in resp.data


def test_merge_endpoint_returns_file(client, sample_data_dir):
    order = sample_data_dir / "orders.xlsx"
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge",
        data=_multipart(order, payment),
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert len(resp.data) > 0
    cd = resp.headers.get("Content-Disposition", "")
    assert "attachment" in cd.lower()


def test_merge_endpoint_missing_order_file(client, sample_data_dir):
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge",
        data={"payment_file": (open(payment, "rb"), payment.name)},
        content_type="multipart/form-data",
    )
    assert resp.status_code == 400
    assert "order_file" in resp.get_json().get("error", "")


def test_merge_endpoint_invalid_extension(client, sample_data_dir, tmp_path):
    bad = tmp_path / "evil.txt"
    bad.write_text("not excel")
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge",
        data={
            "order_file": (open(bad, "rb"), bad.name),
            "payment_file": (open(payment, "rb"), payment.name),
        },
        content_type="multipart/form-data",
    )
    assert resp.status_code == 400


def test_merge_json_endpoint(client, sample_data_dir):
    order = sample_data_dir / "orders.xlsx"
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge/json",
        data=_multipart(order, payment),
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["success"] is True
    assert "download_url" in body
    assert body["download_url"].startswith("/download/")
    assert "statistics" in body
    assert "total_rows" in body["statistics"]


def test_merge_json_with_sales_report(client, sample_data_dir):
    order = sample_data_dir / "orders.xlsx"
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge/json",
        data=_multipart(order, payment, month="202603"),
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    body = resp.get_json()
    assert body["success"] is True
    assert "report_rows" in body["statistics"]


def test_download_missing_file_returns_404(client):
    resp = client.get("/download/nonexistent_file.xlsx")
    assert resp.status_code == 404


def test_merge_download_roundtrip(client, sample_data_dir):
    order = sample_data_dir / "orders.xlsx"
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge/json",
        data=_multipart(order, payment),
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200
    download_url = resp.get_json()["download_url"]
    dl = client.get(download_url)
    assert dl.status_code == 200
    assert len(dl.data) > 0

def test_merge_endpoint_with_sales_report(client, sample_data_dir):
    order = sample_data_dir / "orders.xlsx"
    payment = sample_data_dir / "payments.csv"
    resp = client.post(
        "/merge",
        data=_multipart(order, payment, month="202603"),
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200, resp.get_data(as_text=True)
    assert len(resp.data) > 0
    cd = resp.headers.get("Content-Disposition", "")
    assert "attachment" in cd
    assert "report_202603.xlsx" in cd

def test_merge_endpoint_empty_filename(client, tmp_path):
    import io
    resp = client.post(
        "/merge",
        data={
            "order_file": (io.BytesIO(b""), ""),
            "payment_file": (io.BytesIO(b""), "")
        },
        content_type="multipart/form-data",
    )
    assert resp.status_code == 400
    assert "No order file" in resp.get_json().get("error", "")

def test_merge_json_endpoint_empty_filename(client, tmp_path):
    import io
    resp = client.post(
        "/merge/json",
        data={
            "order_file": (io.BytesIO(b""), ""),
            "payment_file": (io.BytesIO(b""), "")
        },
        content_type="multipart/form-data",
    )
    assert resp.status_code == 400
    assert "Empty filename" in resp.get_json().get("error", "")
