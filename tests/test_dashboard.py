from pathlib import Path

from excelbench.results.dashboard import _build_dashboard


def test_dashboard_includes_best_adapter_by_workload_profile() -> None:
    fidelity = {
        "metadata": {"profile": "xlsx", "run_date": "2026-01-01T00:00:00Z"},
        "libraries": {
            "openpyxl": {"capabilities": ["read", "write"]},
            "xlsxwriter": {"capabilities": ["write"]},
        },
        "results": [],
    }
    perf = {
        "results": [
            {
                "feature": "cell_values_1k",
                "library": "openpyxl",
                "workload_size": "small",
                "perf": {
                    "read": {"op_count": 1000, "wall_ms": {"p50": 10.0}},
                    "write": {"op_count": 1000, "wall_ms": {"p50": 25.0}},
                },
            },
            {
                "feature": "cell_values_10k_bulk_write",
                "library": "xlsxwriter",
                "workload_size": "medium",
                "perf": {
                    "read": None,
                    "write": {"op_count": 10000, "wall_ms": {"p50": 20.0}},
                },
            },
            {
                "feature": "cell_values_100k_bulk_read",
                "library": "openpyxl",
                "workload_size": "large",
                "perf": {
                    "read": {"op_count": 100000, "wall_ms": {"p50": 250.0}},
                    "write": None,
                },
            },
        ]
    }

    lines = _build_dashboard(fidelity, perf)
    doc = "\n".join(lines)

    assert "## Best Adapter by Workload Profile" in doc
    assert "| small |" in doc
    assert "| medium |" in doc
    assert "| large |" in doc


def test_dashboard_filters_pyumya_and_shows_modify_column() -> None:
    fidelity = {
        "metadata": {"profile": "xlsx", "run_date": "2026-01-01T00:00:00Z"},
        "libraries": {
            "wolfxl": {"capabilities": ["read", "write"]},
            "openpyxl": {"capabilities": ["read", "write"]},
            "pyumya": {"capabilities": ["read", "write"]},
        },
        "results": [
            {
                "feature": "cell_values",
                "library": "wolfxl",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
            {
                "feature": "cell_values",
                "library": "openpyxl",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
            {
                "feature": "cell_values",
                "library": "pyumya",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
        ],
    }

    lines = _build_dashboard(fidelity, perf=None)
    doc = "\n".join(lines)

    assert "| Library | Caps | Modify | Green Features | Pass Rate | Best For |" in doc
    assert "| wolfxl | R+W | Patch |" in doc
    assert "| openpyxl | R+W | Rewrite |" in doc
    assert "pyumya" not in doc


def test_dashboard_includes_delta_since_last_run(tmp_path: Path) -> None:
    fidelity_history = tmp_path / "history.jsonl"
    fidelity_history.write_text(
        '{"scores":{"openpyxl":{"cell_values":{"read":2,"write":2}}}}\n'
        '{"scores":{"openpyxl":{"cell_values":{"read":3,"write":1}}}}\n'
    )
    perf_history = tmp_path / "perf_history.jsonl"
    perf_history.write_text(
        '{"p50_wall_ms":{"openpyxl":{"cell_values":{"read_p50":10,"write_p50":20}}}}\n'
        '{"p50_wall_ms":{"openpyxl":{"cell_values":{"read_p50":8,"write_p50":25}}}}\n'
    )
    fidelity = {
        "metadata": {},
        "libraries": {"openpyxl": {"capabilities": ["read", "write"]}},
        "results": [],
    }
    doc = "\n".join(
        _build_dashboard(
            fidelity,
            perf=None,
            fidelity_history_path=fidelity_history,
            perf_history_path=perf_history,
        )
    )
    assert "## Delta Since Last Run" in doc
    assert "Fidelity score changes" in doc
    assert "Median read throughput" in doc
