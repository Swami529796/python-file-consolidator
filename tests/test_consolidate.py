"""Tests for consolidate.py. Run with: python -m pytest"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import consolidate as fc  # noqa: E402

SAMPLE = Path(__file__).resolve().parents[1] / "sample-data"


def test_sample_data_matches_the_vba_version():
    result = fc.consolidate(SAMPLE)
    assert len(result.master) == 120
    assert result.loaded == 3
    assert result.skipped == 1
    skipped = [entry for entry in result.log if entry.status == "Skipped"][0]
    assert skipped.file == "Sales_West.xlsx"
    assert "expected 'Units', found 'Qty'" in skipped.message


def test_every_row_is_traceable():
    master = fc.consolidate(SAMPLE).master
    assert list(master.columns[-2:]) == ["Source File", "Source Sheet"]
    assert master["Source File"].notna().all()
    assert set(master["Source File"]) == {"Sales_East.csv", "Sales_North.xlsx", "Sales_South.xlsx"}


def test_dates_have_one_type_across_csv_and_excel():
    master = fc.consolidate(SAMPLE).master
    assert pd.api.types.is_datetime64_any_dtype(master["Date"])


def test_header_check_ignores_case_and_spaces():
    assert fc.check_headers(["Date", "Units"], [" date ", "UNITS"]) == ""
    assert "expected 2 columns, found 1" in fc.check_headers(["Date", "Units"], ["Date"])


def test_blank_rows_are_dropped_and_bad_files_are_logged(tmp_path):
    pd.DataFrame({"A": [1, None, 3], "B": ["x", None, "z"]}).to_csv(tmp_path / "1_good.csv", index=False)
    (tmp_path / "2_broken.xlsx").write_text("this is not a workbook")
    result = fc.consolidate(tmp_path)
    assert len(result.master) == 2
    assert [entry.status for entry in result.log] == ["Loaded", "Error"]


def test_output_workbook_has_master_and_log(tmp_path):
    output = tmp_path / "Master.xlsx"
    assert fc.main([str(SAMPLE), "--output", str(output)]) == 0
    book = load_workbook(output)
    assert book.sheetnames == ["Master", "Log"]
    assert book["Master"].max_row == 121  # header + 120 rows
    assert book["Log"].max_row == 5  # header + 4 files
