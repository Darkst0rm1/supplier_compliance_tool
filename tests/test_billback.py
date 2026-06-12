from src.compliance_engine import _billback_sheet_name


def test_sheet_name_prefixes_and_keeps_simple_name():
    used = set()
    assert _billback_sheet_name("BOB'S RED MILL", "1001", used) == "BB-BOB'S RED MILL"


def test_sheet_name_strips_illegal_chars():
    used = set()
    name = _billback_sheet_name("A/B:C*D[E]F?G\\H", "1", used)
    for ch in r":\/?*[]":
        assert ch not in name
    assert name.startswith("BB-")


def test_sheet_name_truncates_to_31_chars():
    used = set()
    name = _billback_sheet_name("X" * 60, "1", used)
    assert len(name) <= 31
    assert name.startswith("BB-")


def test_sheet_name_dedupes_collisions():
    used = set()
    first = _billback_sheet_name("SAME NAME", "1", used)
    second = _billback_sheet_name("SAME NAME", "2", used)
    assert first != second
    assert first in used and second in used
    assert len(second) <= 31


def test_sheet_name_falls_back_to_number_then_unknown():
    used = set()
    assert _billback_sheet_name("", "9999", used) == "BB-9999"
    assert _billback_sheet_name("", "", used) == "BB-Unknown Supplier"
