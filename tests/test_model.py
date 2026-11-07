import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import build_model
import model
from model import YEARS, FORECAST, PF_YEARS


@pytest.fixture(scope="module")
def res():
    return model.run()


def test_reported_equity_comes_out_right(res):
    # 2025 equity is the balancing item; Orrön reported 324.1 attributable to shareholders
    assert res["target"]["equity"][2025] == pytest.approx(324.1, abs=0.05)


def test_balance_sheet_balances_every_year(res):
    for y in YEARS:
        assert res["target"]["check"][y] == pytest.approx(0, abs=1e-9)


def test_cash_moves_with_the_cash_flow(res):
    t = res["target"]
    for prev, y in zip(YEARS, FORECAST):
        assert t["cash"][y] - t["cash"][prev] == pytest.approx(t["pre_financing"][y] + t["debt_change"][y])
        assert t["cash"][y] >= res["a"].min_cash - 1e-9


def test_rcf_stays_inside_the_facility(res):
    assert all(res["target"]["rcf_headroom"][y] > 0 for y in YEARS)


def test_offer_maths(res):
    a, d = res["a"], res["deal"]
    equity = a.orron_price_sek * 1.25 * a.orron_shares_m / a.eur_sek
    assert d["equity_value"] == pytest.approx(equity)
    shares = equity * a.eur_nok / a.cloud_price_nok
    assert d["new_shares_m"] == pytest.approx(shares)
    assert d["orron_holders_pct"] == pytest.approx(shares / (shares + a.cloud_shares_m))


def test_paying_below_book_writes_down_ppe_not_goodwill(res):
    d = res["deal"]
    assert d["equity_value"] < res["target"]["equity"][2025]
    assert d["goodwill"] == 0
    assert d["ppe_writedown"] == pytest.approx(res["target"]["equity"][2025] - d["equity_value"])


def test_pro_forma_balance_sheet_balances(res):
    assert res["bs"]["check"] == pytest.approx(0, abs=1e-9)


def test_2028_accretion_by_hand(res):
    a, t, d = res["a"], res["target"], res["deal"]
    y = 2028
    cloud = a.cloud_net_income_nok[y] / a.eur_nok
    syn = a.gna_cut * t["gna"][y] + a.om_saving_mwh * t["generation"][y] / 1000
    pf = cloud + t["net_income"][y] + (syn + d["da_saving"]) * (1 - a.cloud_tax)
    expected = (pf / d["pf_shares_m"]) / (cloud / a.cloud_shares_m) - 1
    assert res["pf"]["accretion"][y] == pytest.approx(expected)


def test_breakeven_price_is_eps_neutral():
    p = model.breakeven_price(2028)
    prices = {y: (model.Assumptions().price_mwh[y] if y < 2027 else p) for y in FORECAST}
    assert model.accretion(2028, price_mwh=prices) == pytest.approx(0, abs=1e-6)


def test_more_synergies_never_hurt():
    low = model.accretion(2028, gna_cut=0.2)
    high = model.accretion(2028, gna_cut=0.9)
    assert high > low


def _recalculate(src, out_dir):
    """Open the workbook in LibreOffice with 'always recalculate' and save it again."""
    profile = out_dir / "lo_profile" / "user"
    profile.mkdir(parents=True)
    (profile / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry">\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
        '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
        '</oor:items>\n')
    subprocess.run(
        ["soffice", f"-env:UserInstallation=file://{profile.parent}", "--headless",
         "--convert-to", "xlsx:Calc Office Open XML", "--outdir", str(out_dir / "out"), str(src)],
        check=True, capture_output=True, timeout=180)
    return out_dir / "out" / src.name


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_excel_formulas_match_python(tmp_path):
    path = tmp_path / "model.xlsx"
    cells = build_model.build(str(path))
    cached = openpyxl.load_workbook(path, data_only=True)
    recalced = openpyxl.load_workbook(_recalculate(path, tmp_path), data_only=True)

    wrong = []
    for sheet, cell in cells:
        expected = cached[sheet][cell].value
        got = recalced[sheet][cell].value
        if not isinstance(got, (int, float)) or abs(got - expected) > 1e-6 * max(1, abs(expected)):
            wrong.append((sheet, cell, expected, got))
    assert not wrong, wrong[:10]
    assert len(cells) > 300


def test_workbook_has_no_error_values(tmp_path):
    path = tmp_path / "model.xlsx"
    build_model.build(str(path))
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                assert not (isinstance(v, str) and v.startswith("#")), (ws.title, v)
