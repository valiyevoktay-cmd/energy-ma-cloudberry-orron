"""Cloudberry / Orrön merger model - the calculations.

Everything here is plain Python so the numbers can be checked without Excel.
build_model.py writes the same logic into a workbook as live formulas, and the
tests recalculate that workbook in LibreOffice and compare it with this file.

Units: EUR million unless the name says otherwise (_sek, _nok, _gwh, _mwh).
"""

from dataclasses import dataclass, field, replace

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]   # 2025 is reported, rest forecast
FORECAST = YEARS[1:]
PF_YEARS = [2027, 2028, 2029]                  # first full years after a 2027 close


def by_year(*values):
    return dict(zip(FORECAST, values))


@dataclass
class Assumptions:
    # market data (stockanalysis.com, early Nov 2026 for Orrön, 29 Sep 2026 for Cloudberry)
    orron_price_sek: float = 8.80
    orron_shares_m: float = 285.91
    cloud_price_nok: float = 11.76
    cloud_shares_m: float = 336.08
    eur_sek: float = 11.0
    eur_nok: float = 11.7

    # deal terms
    premium: float = 0.25
    stock_share: float = 1.0          # share of the equity price paid in Cloudberry shares
    new_debt_rate: float = 0.05       # on any cash part, funded with new debt
    fee_pct: float = 0.015            # advisory + other fees, paid from cash at close
    cloud_tax: float = 0.22           # Norway, used for deal adjustments

    # Orrön reported, FY2025 (consolidated)
    t25_revenue: float = 28.9         # 24.9 power + 4.0 project sales
    t25_ebitda: float = -10.3
    t25_da: float = 16.9
    t25_ebit: float = -27.2
    t25_net_income: float = -26.3
    t25_capex: float = 15.9
    t25_generation_gwh: float = 692.0  # 24.9m / EUR 36 per MWh

    # Orrön balance sheet 31 Dec 2025
    t_ppe: float = 278.3
    t_other_nca: float = 127.9        # JVs, deferred tax, other long-term assets
    t_current_assets: float = 34.1    # inventory, receivables, other (ex cash)
    t_cash: float = 15.9
    t_current_liabs: float = 9.6
    t_debt: float = 106.4
    t_other_ltl: float = 13.9
    t_nci: float = 2.2
    t_jv_book: float = 36.2           # inside other_nca, valued at book in the DCF

    # Orrön forecast drivers
    generation_gwh: dict = field(default_factory=lambda: by_year(760, 800, 800, 800, 800))
    price_mwh: dict = field(default_factory=lambda: by_year(38, 42, 44, 45, 46))
    opex_mwh: dict = field(default_factory=lambda: by_year(22, 21, 21, 21.5, 22))
    gna: dict = field(default_factory=lambda: by_year(9.0, 9.2, 9.4, 9.6, 9.8))
    legal: dict = field(default_factory=lambda: by_year(2.0, 0, 0, 0, 0))
    capex: dict = field(default_factory=lambda: by_year(12, 8, 4, 4, 4))
    da_rate: float = 0.06             # of opening PP&E; 16.9 / 281.3 in 2025
    debt_rate: float = 0.043          # RCF, floating; 2025 interest 4.1 on ~95 avg debt
    orron_tax: float = 0.206
    min_cash: float = 10.0
    rcf_limit: float = 170.0

    # Cloudberry standalone (NOK m) - own estimates, the fwd P/E points to ~165 for the next 12m
    cloud_net_income_nok: dict = field(default_factory=lambda: by_year(165, 190, 205, 215, 225))
    cloud_ebitda_nok: dict = field(default_factory=lambda: by_year(420, 460, 480, 495, 505))
    cloud_net_debt_nok: float = 2416.0
    cloud_cash_nok: float = 893.0
    cloud_debt_nok: float = 3308.0
    cloud_total_assets_nok: float = 9434.0
    cloud_equity_nok: float = 5427.0

    # synergies
    gna_cut: float = 0.60             # share of Orrön G&A removed (listing, board, overlap)
    om_saving_mwh: float = 1.0        # EUR per MWh from joint O&M / procurement
    phase_in: dict = field(default_factory=lambda: by_year(0, 0.5, 1, 1, 1))
    integration_cost: dict = field(default_factory=lambda: by_year(0, 4.0, 0, 0, 0))
    ppe_life_years: float = 20.0      # for D&A on any fair-value write-down

    # valuation
    wacc: float = 0.075
    tgr: float = 0.02


def target(a):
    """Orrön standalone: income statement, cash flow, balance sheet."""
    r = {k: {} for k in [
        "generation", "price", "revenue", "opex", "gna", "legal", "ebitda", "da", "ebit",
        "interest", "ebt", "tax", "net_income", "capex", "cfo", "pre_financing",
        "debt_change", "ppe", "other_nca", "current_assets", "cash", "total_assets",
        "current_liabs", "debt", "other_ltl", "equity", "nci", "total_le", "check",
        "net_debt", "rcf_headroom"]}
    y0 = YEARS[0]
    r["generation"][y0] = a.t25_generation_gwh
    r["revenue"][y0] = a.t25_revenue
    r["ebitda"][y0] = a.t25_ebitda
    r["da"][y0] = a.t25_da
    r["ebit"][y0] = a.t25_ebit
    r["net_income"][y0] = a.t25_net_income
    r["capex"][y0] = a.t25_capex
    r["ppe"][y0] = a.t_ppe
    r["other_nca"][y0] = a.t_other_nca
    r["current_assets"][y0] = a.t_current_assets
    r["cash"][y0] = a.t_cash
    r["current_liabs"][y0] = a.t_current_liabs
    r["debt"][y0] = a.t_debt
    r["other_ltl"][y0] = a.t_other_ltl
    r["nci"][y0] = a.t_nci
    r["equity"][y0] = (a.t_ppe + a.t_other_nca + a.t_current_assets + a.t_cash
                       - a.t_current_liabs - a.t_debt - a.t_other_ltl - a.t_nci)

    for prev, y in zip(YEARS, FORECAST):
        gen, price = a.generation_gwh[y], a.price_mwh[y]
        r["generation"][y] = gen
        r["price"][y] = price
        r["revenue"][y] = gen * price / 1000
        r["opex"][y] = gen * a.opex_mwh[y] / 1000
        r["gna"][y] = a.gna[y]
        r["legal"][y] = a.legal[y]
        r["ebitda"][y] = r["revenue"][y] - r["opex"][y] - a.gna[y] - a.legal[y]
        r["da"][y] = r["ppe"][prev] * a.da_rate
        r["ebit"][y] = r["ebitda"][y] - r["da"][y]
        r["interest"][y] = r["debt"][prev] * a.debt_rate
        r["ebt"][y] = r["ebit"][y] - r["interest"][y]
        # losses are not carried forward here - keeps it simple and slightly conservative
        r["tax"][y] = max(0.0, r["ebt"][y]) * a.orron_tax
        r["net_income"][y] = r["ebt"][y] - r["tax"][y]

        r["capex"][y] = a.capex[y]
        r["cfo"][y] = r["net_income"][y] + r["da"][y]          # working capital held flat
        r["pre_financing"][y] = r["cfo"][y] - r["capex"][y]
        cash_before = r["cash"][prev] + r["pre_financing"][y]
        # draw on the RCF to keep minimum cash, repay with anything above it
        r["debt_change"][y] = max(a.min_cash - cash_before, -r["debt"][prev])

        r["ppe"][y] = r["ppe"][prev] + r["capex"][y] - r["da"][y]
        r["other_nca"][y] = r["other_nca"][prev]
        r["current_assets"][y] = r["current_assets"][prev]
        r["cash"][y] = cash_before + r["debt_change"][y]
        r["current_liabs"][y] = r["current_liabs"][prev]
        r["debt"][y] = r["debt"][prev] + r["debt_change"][y]
        r["other_ltl"][y] = r["other_ltl"][prev]
        r["nci"][y] = r["nci"][prev]
        r["equity"][y] = r["equity"][prev] + r["net_income"][y]

    for y in YEARS:
        r["total_assets"][y] = r["ppe"][y] + r["other_nca"][y] + r["current_assets"][y] + r["cash"][y]
        r["total_le"][y] = (r["current_liabs"][y] + r["debt"][y] + r["other_ltl"][y]
                            + r["equity"][y] + r["nci"][y])
        r["check"][y] = r["total_assets"][y] - r["total_le"][y]
        r["net_debt"][y] = r["debt"][y] - r["cash"][y]
        r["rcf_headroom"][y] = a.rcf_limit - r["debt"][y]
    return r


def dcf(a, t):
    """Unlevered DCF of Orrön's existing business, valued at 31 Dec 2025."""
    out = {"nopat_tax": {}, "fcff": {}, "df": {}, "pv": {}}
    for i, y in enumerate(FORECAST, start=1):
        out["nopat_tax"][y] = max(0.0, t["ebit"][y]) * a.orron_tax
        out["fcff"][y] = t["ebitda"][y] - out["nopat_tax"][y] - t["capex"][y]
        out["df"][y] = 1 / (1 + a.wacc) ** i
        out["pv"][y] = out["fcff"][y] * out["df"][y]
    last = FORECAST[-1]
    out["sum_pv"] = sum(out["pv"].values())
    out["tv"] = out["fcff"][last] * (1 + a.tgr) / (a.wacc - a.tgr)
    out["pv_tv"] = out["tv"] * out["df"][last]
    out["ev"] = out["sum_pv"] + out["pv_tv"]
    out["equity"] = out["ev"] + a.t_jv_book - t["net_debt"][2025] - a.t_nci
    out["per_share_sek"] = out["equity"] * a.eur_sek / a.orron_shares_m
    return out


def deal(a, t):
    d = {}
    d["offer_price_sek"] = a.orron_price_sek * (1 + a.premium)
    d["equity_value"] = d["offer_price_sek"] * a.orron_shares_m / a.eur_sek
    d["net_debt"] = t["net_debt"][2025]
    d["ev"] = d["equity_value"] + d["net_debt"] + a.t_nci
    d["ev_ebitda_2027"] = d["ev"] / t["ebitda"][2027]
    d["ev_ebitda_2028"] = d["ev"] / t["ebitda"][2028]
    d["price_to_book"] = d["equity_value"] / t["equity"][2025]
    d["stock_paid"] = d["equity_value"] * a.stock_share
    d["cash_paid"] = d["equity_value"] - d["stock_paid"]
    d["new_debt"] = d["cash_paid"]
    d["fees"] = d["equity_value"] * a.fee_pct
    d["new_shares_m"] = d["stock_paid"] * a.eur_nok / a.cloud_price_nok
    d["exchange_ratio"] = d["new_shares_m"] / a.orron_shares_m
    d["pf_shares_m"] = a.cloud_shares_m + d["new_shares_m"]
    d["orron_holders_pct"] = d["new_shares_m"] / d["pf_shares_m"]
    # paying less than book value: write PP&E down instead of booking goodwill
    excess = d["equity_value"] - t["equity"][2025]
    d["goodwill"] = max(0.0, excess)
    d["ppe_writedown"] = max(0.0, -excess)
    d["da_saving"] = d["ppe_writedown"] / a.ppe_life_years
    return d


def pro_forma(a, t, d):
    p = {k: {} for k in [
        "cloud_ni", "cloud_eps", "orron_ni", "synergies", "integration", "new_interest",
        "da_saving", "adjustments_after_tax", "pf_ni", "pf_eps", "accretion",
        "pf_ebitda", "pf_net_debt", "leverage"]}
    for y in FORECAST:
        p["cloud_ni"][y] = a.cloud_net_income_nok[y] / a.eur_nok
        p["cloud_eps"][y] = p["cloud_ni"][y] / a.cloud_shares_m
        p["orron_ni"][y] = t["net_income"][y]
        p["synergies"][y] = a.phase_in[y] * (a.gna_cut * t["gna"][y]
                                             + a.om_saving_mwh * t["generation"][y] / 1000)
        p["integration"][y] = a.integration_cost[y]
        p["new_interest"][y] = d["new_debt"] * a.new_debt_rate if y >= PF_YEARS[0] else 0.0
        p["da_saving"][y] = d["da_saving"] if y >= PF_YEARS[0] else 0.0
        pre_tax = p["synergies"][y] - p["integration"][y] - p["new_interest"][y] + p["da_saving"][y]
        p["adjustments_after_tax"][y] = pre_tax * (1 - a.cloud_tax)
        p["pf_ni"][y] = p["cloud_ni"][y] + p["orron_ni"][y] + p["adjustments_after_tax"][y]
        p["pf_eps"][y] = p["pf_ni"][y] / d["pf_shares_m"]
        p["accretion"][y] = p["pf_eps"][y] / p["cloud_eps"][y] - 1
        p["pf_ebitda"][y] = (a.cloud_ebitda_nok[y] / a.eur_nok + t["ebitda"][y]
                             + p["synergies"][y] - p["integration"][y])

    close_nd = (a.cloud_net_debt_nok / a.eur_nok + t["net_debt"][2025]
                + d["new_debt"] + d["fees"])
    for y in PF_YEARS:
        p["pf_net_debt"][y] = close_nd
        p["leverage"][y] = close_nd / p["pf_ebitda"][y]
    return p


def pro_forma_balance_sheet(a, t, d):
    """Combination on the latest reported balance sheets (31 Dec 2025), as if closed then."""
    c_cash = a.cloud_cash_nok / a.eur_nok
    c_debt = a.cloud_debt_nok / a.eur_nok
    c_assets = a.cloud_total_assets_nok / a.eur_nok
    c_equity = a.cloud_equity_nok / a.eur_nok
    c_other_assets = c_assets - c_cash
    c_other_liabs = c_assets - c_equity - c_debt
    o = t
    y = 2025
    bs = {
        "cash": c_cash + o["cash"][y] + d["new_debt"] - d["cash_paid"] - d["fees"],
        "ppe_and_other": (c_other_assets + o["ppe"][y] + o["other_nca"][y]
                          + o["current_assets"][y] - d["ppe_writedown"]),
        "goodwill": d["goodwill"],
        "debt": c_debt + o["debt"][y] + d["new_debt"],
        "other_liabs": c_other_liabs + o["current_liabs"][y] + o["other_ltl"][y],
        # Orrön's equity is eliminated and replaced by the new shares issued
        "equity": c_equity + d["stock_paid"] - d["fees"],
        "nci": o["nci"][y],
    }
    bs["total_assets"] = bs["cash"] + bs["ppe_and_other"] + bs["goodwill"]
    bs["total_le"] = bs["debt"] + bs["other_liabs"] + bs["equity"] + bs["nci"]
    bs["check"] = bs["total_assets"] - bs["total_le"]
    return bs


def synergy_value(a, p):
    """PV of after-tax synergies net of integration costs, same WACC and growth as the DCF."""
    pv = 0.0
    for i, y in enumerate(FORECAST, start=1):
        pv += (p["synergies"][y] - p["integration"][y]) * (1 - a.cloud_tax) / (1 + a.wacc) ** i
    last = FORECAST[-1]
    run_rate = p["synergies"][last] * (1 - a.cloud_tax)
    pv += run_rate * (1 + a.tgr) / (a.wacc - a.tgr) / (1 + a.wacc) ** len(FORECAST)
    return pv


def run(a=None):
    a = a or Assumptions()
    t = target(a)
    v = dcf(a, t)
    d = deal(a, t)
    p = pro_forma(a, t, d)
    bs = pro_forma_balance_sheet(a, t, d)
    syn = synergy_value(a, p)
    summary = {
        "dcf_ev": v["ev"],
        "synergy_pv": syn,
        "deal_ev": d["ev"],
        "fees": d["fees"],
        # what the price implies for everything the DCF leaves out (mainly the pipeline)
        "jv_book": a.t_jv_book,
        "gap": d["ev"] + d["fees"] - v["ev"] - a.t_jv_book - syn,
    }
    return {"a": a, "target": t, "dcf": v, "deal": d, "pf": p, "bs": bs, "summary": summary}


def accretion(year, **changes):
    res = run(replace(Assumptions(), **changes))
    return res["pf"]["accretion"][year]


def breakeven_price(year=2028, **changes):
    """Flat power price from 2027 on that makes `year` EPS-neutral (bisection)."""
    def acc(p):
        prices = {y: (Assumptions().price_mwh[y] if y < 2027 else p) for y in FORECAST}
        return accretion(year, price_mwh=prices, **changes)
    lo, hi = 0.0, 300.0
    if acc(hi) < 0:
        return None
    for _ in range(60):
        mid = (lo + hi) / 2
        if acc(mid) < 0:
            lo = mid
        else:
            hi = mid
    return hi


if __name__ == "__main__":
    res = run()
    t, d, p, s = res["target"], res["deal"], res["pf"], res["summary"]
    print("Orrön EBITDA  ", {y: round(t["ebitda"][y], 1) for y in YEARS})
    print("Orrön NI      ", {y: round(t["net_income"][y], 1) for y in YEARS})
    print("BS check      ", {y: round(t["check"][y], 6) for y in YEARS})
    print("debt          ", {y: round(t["debt"][y], 1) for y in YEARS})
    print(f"offer {d['offer_price_sek']:.2f} SEK, equity {d['equity_value']:.1f}, EV {d['ev']:.1f}, "
          f"EV/EBITDA 27E {d['ev_ebitda_2027']:.1f}x, P/B {d['price_to_book']:.2f}")
    print(f"new shares {d['new_shares_m']:.1f}m, Orrön holders {d['orron_holders_pct']:.1%}, "
          f"writedown {d['ppe_writedown']:.1f}, goodwill {d['goodwill']:.1f}")
    print("accretion     ", {y: f"{p['accretion'][y]:.1%}" for y in PF_YEARS})
    print("leverage      ", {y: round(p["leverage"][y], 1) for y in PF_YEARS})
    print("DCF EV", round(res["dcf"]["ev"], 1), "per share SEK", round(res["dcf"]["per_share_sek"], 2))
    print("summary", {k: round(v, 1) for k, v in s.items()})
    print("PF BS check", round(res["bs"]["check"], 6))
    print("breakeven price 2028", breakeven_price())
