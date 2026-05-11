from pathlib import Path


def test_cli_docs_use_positional_target_month_and_not_month_flag():
    agents = Path("AGENTS.md").read_text(encoding="utf-8")
    usage = Path("documents/USAGE_EXAMPLES.md").read_text(encoding="utf-8")

    assert "--month" not in agents
    assert "--month" not in usage
    assert "python cli.py order.xlsx payment.xlsx 202602" in agents
    assert "python cli.py order.xlsx payment.xlsx 202602" in usage
    assert "target_month" in agents
    assert "target_month" in usage
