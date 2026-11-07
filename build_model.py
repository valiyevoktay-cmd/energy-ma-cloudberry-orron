"""Writes Cloudberry_Orron_MA_Model.xlsx.

Every calculated cell is a live Excel formula that points back to the Inputs
sheet, so changing an input in Excel flows through the whole workbook. The
cached values come from model.py, which is what the tests check against.

    python build_model.py
"""

import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

import model
from model import YEARS, FORECAST, PF_YEARS

OUT = "Cloudberry_Orron_MA_Model.xlsx"
FIRST_YEAR_COL = 2          # column C holds 2025
YEAR_COL = {y: FIRST_YEAR_COL + i for i, y in enumerate(YEARS)}


class Book:
    """Small helper: remembers where every named value lives and turns
    templates like "{revenue}-{opex}" or "{ppe@-1}" into cell references."""

    def __init__(self, path):
        self.wb = xlsxwriter.Workbook(path)
        self.names = {}       # name -> ("scalar", sheet, cell) or ("row", sheet, row)
        self.formulas = []    # (sheet, cell) of every formula, for the tests
        self.pending = []     # formulas are written at close(), once every name is known
        base = {"font_name": "Arial", "font_size": 10}
        f = lambda **kw: self.wb.add_format({**base, **kw})
        self.fmt = {
            "title": f(bold=True, font_size=13, font_color="#1F3864"),
            "section": f(bold=True, bottom=1),
            "label": f(),
            "bold": f(bold=True),
            "note": f(italic=True, font_color="#595959"),
            "year": f(bold=True, align="right", bottom=1),
        }
        for kind, colour in [("in", "#0000FF"), ("calc", "#000000"), ("link", "#008000")]:
            for num, nf in [("num", "#,##0.0;(#,##0.0);–"), ("pct", "0.0%;(0.0%);–"),
                            ("x", '0.0"x"'), ("eps", "0.0000;(0.0000)"), ("int", "#,##0"),
                            ("px", "0.00")]:
                self.fmt[f"{kind}_{num}"] = f(font_color=colour, num_format=nf)
                self.fmt[f"{kind}_{num}_total"] = f(font_color=colour, num_format=nf,
                                                    bold=True, top=1)

    # --- references -------------------------------------------------------
    def ref(self, name, sheet, col, shift=0):
        kind, where_sheet, where = self.names[name]
        prefix = "" if where_sheet == sheet else f"'{where_sheet}'!"
        if kind == "scalar":
            return prefix + where
        return prefix + xl_rowcol_to_cell(where, col + shift)

    def resolve(self, template, sheet, col):
        def sub(m):
            return self.ref(m.group(1), sheet, col, -1 if m.group(2) else 0)
        return "=" + re.sub(r"\{(\w+)(@-1)?\}", sub, template)

    def is_link(self, template, sheet):
        m = re.fullmatch(r"\{(\w+)\}", template)
        return bool(m) and self.names[m.group(1)][1] != sheet

    # --- writers ----------------------------------------------------------
    def sheet(self, name, title, note=None):
        ws = self.wb.add_worksheet(name)
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 46)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 8, 11)
        ws.write(0, 0, title, self.fmt["title"])
        if note:
            ws.write(1, 0, note, self.fmt["note"])
        return ws

    def years_header(self, ws, row, years=YEARS):
        for y in years:
            label = f"{y}A" if y == YEARS[0] else f"{y}E"
            ws.write(row, YEAR_COL[y], label, self.fmt["year"])

    def scalar_input(self, ws, row, name, label, value, unit="", num="num"):
        ws.write(row, 0, label, self.fmt["label"])
        ws.write(row, 1, unit, self.fmt["note"])
        ws.write_number(row, 2, value, self.fmt[f"in_{num}"])
        self.names[name] = ("scalar", ws.name, xl_rowcol_to_cell(row, 2, True, True))

    def scalar_formula(self, ws, row, name, label, template, value, unit="", num="num",
                       total=False):
        ws.write(row, 0, label, self.fmt["bold" if total else "label"])
        ws.write(row, 1, unit, self.fmt["note"])
        self.names[name] = ("scalar", ws.name, xl_rowcol_to_cell(row, 2, True, True))
        self.formula(ws, row, 2, template, value, num, total)

    def row(self, ws, row, name, label, cells, unit="", num="num", total=False):
        """cells: {year: number | (template, cached_value)}"""
        ws.write(row, 0, label, self.fmt["bold" if total else "label"])
        ws.write(row, 1, unit, self.fmt["note"])
        self.names[name] = ("row", ws.name, row)
        suffix = "_total" if total else ""
        for y, cell in cells.items():
            col = YEAR_COL[y]
            if isinstance(cell, tuple):
                template, value = cell
                self.formula(ws, row, col, template, value, num, total)
            else:
                ws.write_number(row, col, cell, self.fmt[f"in_{num}{suffix}"])

    def formula(self, ws, row, col, template, value, num="num", total=False):
        def write():
            kind = "link" if self.is_link(template, ws.name) else "calc"
            fmt = self.fmt[f"{kind}_{num}" + ("_total" if total else "")]
            ws.write_formula(row, col, self.resolve(template, ws.name, col), fmt, value)
        self.pending.append(write)
        self.formulas.append((ws.name, xl_rowcol_to_cell(row, col)))

    def close(self):
        for write in self.pending:
            write()
        self.wb.close()

    def section(self, ws, row, text):
        ws.write(row, 0, text, self.fmt["section"])


def fc(template, values, years=FORECAST):
    """Same formula template across a set of years."""
    return {y: (template, values[y]) for y in years}


def build(path=OUT):
    res = model.run()
    a, t, v, d, p, bs, s = (res[k] for k in ["a", "target", "dcf", "deal", "pf", "bs", "summary"])
    b = Book(path)

    summary_ws = b.sheet("Summary", "Cloudberry / Orrön Energy – merger case (hypothetical)",
                         "Inputs in blue, formulas in black, links to other sheets in green. EUR m unless stated.")

    # ------------------------------------------------------------------ Inputs
    ws = b.sheet("Inputs", "Inputs", "Reported figures: company FY2025 reports. Forecasts: my assumptions.")
    r = 3
    b.section(ws, r, "Market data"); r += 1
    for name, label, unit, num in [
        ("orron_price_sek", "Orrön share price (4 Nov 2026)", "SEK", "px"),
        ("orron_shares_m", "Orrön shares outstanding", "m", "num"),
        ("cloud_price_nok", "Cloudberry share price (29 Sep 2026)", "NOK", "px"),
        ("cloud_shares_m", "Cloudberry shares outstanding", "m", "num"),
        ("eur_sek", "EUR/SEK", "", "px"),
        ("eur_nok", "EUR/NOK", "", "px"),
    ]:
        b.scalar_input(ws, r, name, label, getattr(a, name), unit, num); r += 1

    r += 1; b.section(ws, r, "Deal terms"); r += 1
    for name, label, unit, num in [
        ("premium", "Premium to Orrön share price", "%", "pct"),
        ("stock_share", "Paid in Cloudberry shares", "%", "pct"),
        ("new_debt_rate", "Interest on new acquisition debt", "%", "pct"),
        ("fee_pct", "Transaction fees, % of equity value", "%", "pct"),
        ("cloud_tax", "Tax rate on deal adjustments (Norway)", "%", "pct"),
    ]:
        b.scalar_input(ws, r, name, label, getattr(a, name), unit, num); r += 1

    r += 1; b.section(ws, r, "Orrön FY2025 reported (consolidated)"); r += 1
    for name, label, unit in [
        ("t25_generation_gwh", "Power generation (24.9m / EUR 36 per MWh)", "GWh"),
        ("t25_revenue", "Revenue (power 24.9 + project sales 4.0)", "EUR m"),
        ("t25_ebitda", "EBITDA", "EUR m"),
        ("t25_da", "D&A", "EUR m"),
        ("t25_ebit", "EBIT", "EUR m"),
        ("t25_net_income", "Net result", "EUR m"),
        ("t25_capex", "Capex", "EUR m"),
        ("t_ppe", "PP&E", "EUR m"),
        ("t_other_nca", "Other non-current assets (JVs, deferred tax, other)", "EUR m"),
        ("t_jv_book", "  of which investments in JVs", "EUR m"),
        ("t_current_assets", "Current assets excl. cash", "EUR m"),
        ("t_cash", "Cash", "EUR m"),
        ("t_current_liabs", "Current liabilities excl. debt", "EUR m"),
        ("t_debt", "Borrowings (RCF)", "EUR m"),
        ("t_other_ltl", "Other long-term liabilities", "EUR m"),
        ("t_nci", "Non-controlling interests", "EUR m"),
    ]:
        b.scalar_input(ws, r, name, label, getattr(a, name), unit); r += 1

    r += 1; b.section(ws, r, "Orrön operating assumptions"); r += 1
    for name, label, unit, num in [
        ("da_rate", "D&A, % of opening PP&E", "%", "pct"),
        ("debt_rate", "Interest on RCF", "%", "pct"),
        ("orron_tax", "Tax rate (Sweden)", "%", "pct"),
        ("min_cash", "Minimum cash", "EUR m", "num"),
        ("rcf_limit", "RCF size", "EUR m", "num"),
    ]:
        b.scalar_input(ws, r, name, label, getattr(a, name), unit, num); r += 1
    r += 1
    b.years_header(ws, r, FORECAST); r += 1
    for name, label, unit, src, num in [
        ("in_generation", "Generation", "GWh", a.generation_gwh, "int"),
        ("in_price", "Achieved power price", "EUR/MWh", a.price_mwh, "num"),
        ("in_opex_mwh", "Operating cost", "EUR/MWh", a.opex_mwh, "num"),
        ("in_gna", "G&A", "EUR m", a.gna, "num"),
        ("in_legal", "Legal costs (Sudan case)", "EUR m", a.legal, "num"),
        ("in_capex", "Capex", "EUR m", a.capex, "num"),
    ]:
        b.row(ws, r, name, label, dict(src), unit, num); r += 1

    r += 1; b.section(ws, r, "Cloudberry"); r += 1
    for name, label in [
        ("cloud_net_debt_nok", "Net interest-bearing debt, 31 Dec 2025"),
        ("cloud_cash_nok", "Cash, 31 Dec 2025"),
        ("cloud_debt_nok", "Interest-bearing debt, 31 Dec 2025"),
        ("cloud_total_assets_nok", "Total assets, 31 Dec 2025"),
        ("cloud_equity_nok", "Total equity, 31 Dec 2025"),
    ]:
        b.scalar_input(ws, r, name, label, getattr(a, name), "NOK m"); r += 1
    b.years_header(ws, r, FORECAST); r += 1
    b.row(ws, r, "in_cloud_ni", "Net income (own estimate)", dict(a.cloud_net_income_nok), "NOK m"); r += 1
    b.row(ws, r, "in_cloud_ebitda", "EBITDA (own estimate)", dict(a.cloud_ebitda_nok), "NOK m"); r += 1

    r += 1; b.section(ws, r, "Synergies"); r += 1
    b.scalar_input(ws, r, "gna_cut", "Orrön G&A removed", getattr(a, "gna_cut"), "%", "pct"); r += 1
    b.scalar_input(ws, r, "om_saving_mwh", "O&M / procurement saving", a.om_saving_mwh, "EUR/MWh"); r += 1
    b.scalar_input(ws, r, "ppe_life_years", "Remaining life for PP&E write-down", a.ppe_life_years, "years"); r += 1
    b.years_header(ws, r, FORECAST); r += 1
    b.row(ws, r, "in_phase", "Phase-in", dict(a.phase_in), "%", "pct"); r += 1
    b.row(ws, r, "in_integration", "One-off integration cost", dict(a.integration_cost), "EUR m"); r += 1

    r += 1; b.section(ws, r, "Valuation"); r += 1
    b.scalar_input(ws, r, "wacc", "WACC", a.wacc, "%", "pct"); r += 1
    b.scalar_input(ws, r, "tgr", "Terminal growth", a.tgr, "%", "pct"); r += 1

    # ------------------------------------------------------------------ Orrön
    ws = b.sheet("Orron", "Orrön Energy – standalone three statements",
                 "2025A as reported. Working capital held flat; tax losses not carried forward.")
    r = 3
    b.years_header(ws, r); r += 1
    b.section(ws, r, "Income statement"); r += 1
    b.row(ws, r, "generation", "Generation", {2025: ("{t25_generation_gwh}", t["generation"][2025]),
                                              **fc("{in_generation}", t["generation"])}, "GWh", "int"); r += 1
    b.row(ws, r, "price", "Achieved price", fc("{in_price}", t["price"]), "EUR/MWh"); r += 1
    b.row(ws, r, "revenue", "Revenue", {2025: ("{t25_revenue}", t["revenue"][2025]),
                                        **fc("{generation}*{price}/1000", t["revenue"])}, "EUR m", total=True); r += 1
    b.row(ws, r, "opex", "Operating costs", fc("{generation}*{in_opex_mwh}/1000", t["opex"]), "EUR m"); r += 1
    b.row(ws, r, "gna", "G&A", fc("{in_gna}", t["gna"]), "EUR m"); r += 1
    b.row(ws, r, "legal", "Legal costs", fc("{in_legal}", t["legal"]), "EUR m"); r += 1
    b.row(ws, r, "ebitda", "EBITDA", {2025: ("{t25_ebitda}", t["ebitda"][2025]),
                                      **fc("{revenue}-{opex}-{gna}-{legal}", t["ebitda"])}, "EUR m", total=True); r += 1
    b.row(ws, r, "da", "D&A", {2025: ("{t25_da}", t["da"][2025]),
                               **fc("{ppe@-1}*{da_rate}", t["da"])}, "EUR m"); r += 1
    b.row(ws, r, "ebit", "EBIT", {2025: ("{t25_ebit}", t["ebit"][2025]),
                                  **fc("{ebitda}-{da}", t["ebit"])}, "EUR m", total=True); r += 1
    b.row(ws, r, "interest", "Interest on RCF", fc("{debt@-1}*{debt_rate}", t["interest"]), "EUR m"); r += 1
    b.row(ws, r, "ebt", "Profit before tax", fc("{ebit}-{interest}", t["ebt"]), "EUR m"); r += 1
    b.row(ws, r, "tax", "Tax", fc("MAX(0,{ebt})*{orron_tax}", t["tax"]), "EUR m"); r += 1
    b.row(ws, r, "net_income", "Net result", {2025: ("{t25_net_income}", t["net_income"][2025]),
                                              **fc("{ebt}-{tax}", t["net_income"])}, "EUR m", total=True); r += 2

    b.section(ws, r, "Cash flow"); r += 1
    b.row(ws, r, "cfo", "Net result + D&A", fc("{net_income}+{da}", t["cfo"]), "EUR m"); r += 1
    b.row(ws, r, "capex", "Capex", {2025: ("{t25_capex}", t["capex"][2025]),
                                    **fc("{in_capex}", t["capex"])}, "EUR m"); r += 1
    b.row(ws, r, "pre_financing", "Cash flow before financing", fc("{cfo}-{capex}", t["pre_financing"]),
          "EUR m", total=True); r += 1
    b.row(ws, r, "debt_change", "RCF drawn / (repaid)",
          fc("MAX({min_cash}-({cash@-1}+{pre_financing}),-{debt@-1})", t["debt_change"]), "EUR m"); r += 2

    # balance sheet rows reference each other, so reserve the rows first
    b.section(ws, r, "Balance sheet"); r += 1
    rows = ["ppe", "other_nca", "current_assets", "cash", "total_assets", None,
            "current_liabs", "debt", "other_ltl", "equity", "nci", "total_le", "check", None,
            "net_debt", "rcf_headroom"]
    at = {}
    for name in rows:
        if name:
            at[name] = r
            b.names[name] = ("row", ws.name, r)
        r += 1

    def bs_row(name, label, first, template, total=False):
        cells = {2025: (first, t[name][2025]), **fc(template, t[name])}
        b.row(ws, at[name], name, label, cells, "EUR m", total=total)

    bs_row("ppe", "PP&E", "{t_ppe}", "{ppe@-1}+{capex}-{da}")
    bs_row("other_nca", "Other non-current assets", "{t_other_nca}", "{other_nca@-1}")
    bs_row("current_assets", "Current assets excl. cash", "{t_current_assets}", "{current_assets@-1}")
    bs_row("cash", "Cash", "{t_cash}", "{cash@-1}+{pre_financing}+{debt_change}")
    bs_row("total_assets", "Total assets", "{ppe}+{other_nca}+{current_assets}+{cash}",
           "{ppe}+{other_nca}+{current_assets}+{cash}", total=True)
    bs_row("current_liabs", "Current liabilities excl. debt", "{t_current_liabs}", "{current_liabs@-1}")
    bs_row("debt", "Borrowings (RCF)", "{t_debt}", "{debt@-1}+{debt_change}")
    bs_row("other_ltl", "Other long-term liabilities", "{t_other_ltl}", "{other_ltl@-1}")
    bs_row("equity", "Shareholders' equity",
           "{total_assets}-{current_liabs}-{debt}-{other_ltl}-{nci}", "{equity@-1}+{net_income}")
    bs_row("nci", "Non-controlling interests", "{t_nci}", "{nci@-1}")
    bs_row("total_le", "Total liabilities and equity",
           "{current_liabs}+{debt}+{other_ltl}+{equity}+{nci}",
           "{current_liabs}+{debt}+{other_ltl}+{equity}+{nci}", total=True)
    bs_row("check", "Check (should be zero)", "ROUND({total_assets}-{total_le},6)",
           "ROUND({total_assets}-{total_le},6)")
    bs_row("net_debt", "Net debt", "{debt}-{cash}", "{debt}-{cash}")
    bs_row("rcf_headroom", "RCF headroom", "{rcf_limit}-{debt}", "{rcf_limit}-{debt}")

    # ------------------------------------------------------------------ DCF
    ws = b.sheet("DCF", "Orrön – DCF of the existing business",
                 "Valued at 31 Dec 2025. Development pipeline not included - that is the point of the gap on Summary.")
    r = 3
    b.years_header(ws, r, FORECAST); r += 1
    b.row(ws, r, "d_ebitda", "EBITDA", fc("{ebitda}", t["ebitda"]), "EUR m"); r += 1
    b.row(ws, r, "d_tax", "Tax on EBIT", fc("MAX(0,{ebit})*{orron_tax}", v["nopat_tax"]), "EUR m"); r += 1
    b.row(ws, r, "d_capex", "Capex", fc("{capex}", t["capex"]), "EUR m"); r += 1
    b.row(ws, r, "fcff", "Unlevered free cash flow", fc("{d_ebitda}-{d_tax}-{d_capex}", v["fcff"]),
          "EUR m", total=True); r += 1
    b.row(ws, r, "df", "Discount factor",
          {y: (f"1/(1+{{wacc}})^{i}", v["df"][y]) for i, y in enumerate(FORECAST, 1)}, "", "eps"); r += 1
    b.row(ws, r, "pv", "Present value", fc("{fcff}*{df}", v["pv"]), "EUR m"); r += 2

    first, last = xl_col_to_name(YEAR_COL[FORECAST[0]]), xl_col_to_name(YEAR_COL[FORECAST[-1]])
    pv_row = b.names["pv"][2] + 1
    fcff_last = f"{last}{b.names['fcff'][2] + 1}"
    df_last = f"{last}{b.names['df'][2] + 1}"
    b.scalar_formula(ws, r, "sum_pv", "Sum of PV, 2026-30", f"SUM({first}{pv_row}:{last}{pv_row})",
                     v["sum_pv"], "EUR m"); r += 1
    b.scalar_formula(ws, r, "tv", "Terminal value (Gordon growth)",
                     f"{fcff_last}*(1+{{tgr}})/({{wacc}}-{{tgr}})", v["tv"], "EUR m"); r += 1
    b.scalar_formula(ws, r, "pv_tv", "PV of terminal value", f"{{tv}}*{df_last}", v["pv_tv"], "EUR m"); r += 1
    b.scalar_formula(ws, r, "dcf_ev", "Enterprise value", "{sum_pv}+{pv_tv}", v["ev"], "EUR m", total=True); r += 1
    b.scalar_formula(ws, r, "dcf_jv", "+ JVs at book", "{t_jv_book}", a.t_jv_book, "EUR m"); r += 1
    b.scalar_formula(ws, r, "dcf_nd", "– Net debt, 31 Dec 2025", f"'Orron'!C{at['net_debt'] + 1}",
                     t["net_debt"][2025], "EUR m"); r += 1
    b.scalar_formula(ws, r, "dcf_nci", "– Non-controlling interests", "{t_nci}", a.t_nci, "EUR m"); r += 1
    b.scalar_formula(ws, r, "dcf_equity", "Equity value", "{dcf_ev}+{dcf_jv}-{dcf_nd}-{dcf_nci}",
                     v["equity"], "EUR m", total=True); r += 1
    b.scalar_formula(ws, r, "dcf_per_share", "Per share", "{dcf_equity}*{eur_sek}/{orron_shares_m}",
                     v["per_share_sek"], "SEK", "px"); r += 1

    # ------------------------------------------------------------------ Deal
    ws = b.sheet("Deal", "Offer, consideration and accounting")
    r = 3
    nd25 = f"'Orron'!C{at['net_debt'] + 1}"
    eq25 = f"'Orron'!C{at['equity'] + 1}"
    eb27 = f"'Orron'!{xl_col_to_name(YEAR_COL[2027])}{b.names['ebitda'][2] + 1}"
    eb28 = f"'Orron'!{xl_col_to_name(YEAR_COL[2028])}{b.names['ebitda'][2] + 1}"
    for name, label, template, value, unit, num, total in [
        ("offer_price", "Offer price per Orrön share", "{orron_price_sek}*(1+{premium})", d["offer_price_sek"], "SEK", "px", False),
        ("equity_value", "Equity value", "{offer_price}*{orron_shares_m}/{eur_sek}", d["equity_value"], "EUR m", "num", True),
        ("deal_nd", "+ Net debt assumed", nd25, d["net_debt"], "EUR m", "num", False),
        ("deal_nci", "+ Non-controlling interests", "{t_nci}", a.t_nci, "EUR m", "num", False),
        ("deal_ev", "Transaction enterprise value", "{equity_value}+{deal_nd}+{deal_nci}", d["ev"], "EUR m", "num", True),
        ("ev_ebitda_27", "EV / EBITDA 2027E", f"{{deal_ev}}/{eb27}", d["ev_ebitda_2027"], "x", "x", False),
        ("ev_ebitda_28", "EV / EBITDA 2028E", f"{{deal_ev}}/{eb28}", d["ev_ebitda_2028"], "x", "x", False),
        ("p_b", "Price / book equity", f"{{equity_value}}/{eq25}", d["price_to_book"], "x", "x", False),
    ]:
        b.scalar_formula(ws, r, name, label, template, value, unit, num, total); r += 1
    r += 1; b.section(ws, r, "Consideration"); r += 1
    for name, label, template, value, unit, num, total in [
        ("stock_paid", "Paid in Cloudberry shares", "{equity_value}*{stock_share}", d["stock_paid"], "EUR m", "num", False),
        ("cash_paid", "Paid in cash", "{equity_value}-{stock_paid}", d["cash_paid"], "EUR m", "num", False),
        ("new_debt", "New debt raised for the cash part", "{cash_paid}", d["new_debt"], "EUR m", "num", False),
        ("fees", "Transaction fees (from cash)", "{equity_value}*{fee_pct}", d["fees"], "EUR m", "num", False),
        ("new_shares", "New Cloudberry shares issued", "{stock_paid}*{eur_nok}/{cloud_price_nok}", d["new_shares_m"], "m", "num", False),
        ("exchange_ratio", "Cloudberry shares per Orrön share", "{new_shares}/{orron_shares_m}", d["exchange_ratio"], "", "eps", False),
        ("pf_shares", "Pro forma shares outstanding", "{cloud_shares_m}+{new_shares}", d["pf_shares_m"], "m", "num", True),
        ("orron_pct", "Orrön shareholders' stake in the combined company", "{new_shares}/{pf_shares}", d["orron_holders_pct"], "%", "pct", False),
    ]:
        b.scalar_formula(ws, r, name, label, template, value, unit, num, total); r += 1
    r += 1; b.section(ws, r, "Purchase price allocation (simplified)"); r += 1
    for name, label, template, value, unit in [
        ("goodwill", "Goodwill (price above book equity)", f"MAX(0,{{equity_value}}-{eq25})", d["goodwill"], "EUR m"),
        ("writedown", "PP&E write-down (price below book equity)", f"MAX(0,{eq25}-{{equity_value}})", d["ppe_writedown"], "EUR m"),
        ("da_saving", "Lower D&A per year from the write-down", "{writedown}/{ppe_life_years}", d["da_saving"], "EUR m"),
    ]:
        b.scalar_formula(ws, r, name, label, template, value, unit); r += 1

    # ------------------------------------------------------------------ Pro forma
    ws = b.sheet("Pro forma", "Accretion / dilution for Cloudberry shareholders",
                 "Assumes the deal closes at the start of 2027.")
    r = 3
    b.section(ws, r, "Synergies"); r += 1
    b.years_header(ws, r, FORECAST); r += 1
    b.row(ws, r, "p_syn", "Synergies (pre-tax)",
          fc("{in_phase}*({gna_cut}*{gna}+{om_saving_mwh}*{generation}/1000)", p["synergies"]), "EUR m"); r += 1
    b.row(ws, r, "p_integ", "Integration costs", fc("{in_integration}", p["integration"]), "EUR m"); r += 1
    net = {y: (p["synergies"][y] - p["integration"][y]) * (1 - a.cloud_tax) for y in FORECAST}
    b.row(ws, r, "p_syn_net", "Net of integration, after tax", fc("({p_syn}-{p_integ})*(1-{cloud_tax})", net),
          "EUR m"); r += 1
    b.row(ws, r, "p_syn_pv", "Present value",
          {y: (f"{{p_syn_net}}/(1+{{wacc}})^{i}", net[y] / (1 + a.wacc) ** i) for i, y in enumerate(FORECAST, 1)},
          "EUR m"); r += 1
    first, last = xl_col_to_name(YEAR_COL[FORECAST[0]]), xl_col_to_name(YEAR_COL[FORECAST[-1]])
    pv_r, syn_r = b.names["p_syn_pv"][2] + 1, b.names["p_syn"][2] + 1
    b.scalar_formula(ws, r, "synergy_pv", "PV of synergies incl. terminal value",
                     f"SUM({first}{pv_r}:{last}{pv_r})+{last}{syn_r}*(1-{{cloud_tax}})*(1+{{tgr}})/({{wacc}}-{{tgr}})"
                     f"/(1+{{wacc}})^{len(FORECAST)}", s["synergy_pv"], "EUR m", total=True); r += 2

    b.section(ws, r, "Earnings"); r += 1
    b.years_header(ws, r, PF_YEARS); r += 1
    yrs = PF_YEARS
    pf = lambda tmpl, key, **kw: fc(tmpl, p[key], yrs)
    for name, label, template, key, unit, num, total in [
        ("p_cloud_ni", "Cloudberry net income", "{in_cloud_ni}/{eur_nok}", "cloud_ni", "EUR m", "num", False),
        ("p_orron_ni", "Orrön net result", "{net_income}", "orron_ni", "EUR m", "num", False),
        ("p_syn2", "Synergies (pre-tax)", "{p_syn}", "synergies", "EUR m", "num", False),
        ("p_integ2", "Integration costs", "{p_integ}", "integration", "EUR m", "num", False),
        ("p_int", "Interest on new debt", "{new_debt}*{new_debt_rate}", "new_interest", "EUR m", "num", False),
        ("p_da", "Lower D&A from write-down", "{da_saving}", "da_saving", "EUR m", "num", False),
        ("p_adj", "Adjustments after tax", "({p_syn2}-{p_integ2}-{p_int}+{p_da})*(1-{cloud_tax})", "adjustments_after_tax", "EUR m", "num", False),
        ("p_ni", "Pro forma net income", "{p_cloud_ni}+{p_orron_ni}+{p_adj}", "pf_ni", "EUR m", "num", True),
    ]:
        b.row(ws, r, name, label, pf(template, key), unit, num, total); r += 1
    r += 1
    b.row(ws, r, "p_sa_eps", "Cloudberry standalone EPS", pf("{p_cloud_ni}/{cloud_shares_m}", "cloud_eps"), "EUR", "eps"); r += 1
    b.row(ws, r, "p_eps", "Pro forma EPS", pf("{p_ni}/{pf_shares}", "pf_eps"), "EUR", "eps"); r += 1
    b.row(ws, r, "p_acc", "Accretion / (dilution)", pf("{p_eps}/{p_sa_eps}-1", "accretion"), "%", "pct", True); r += 2

    b.section(ws, r, "Leverage"); r += 1
    b.scalar_formula(ws, r, "pf_nd", "Pro forma net debt at close", "{cloud_net_debt_nok}/{eur_nok}+{deal_nd}+{new_debt}+{fees}",
                     p["pf_net_debt"][2027], "EUR m"); r += 1
    b.years_header(ws, r, PF_YEARS); r += 1
    b.row(ws, r, "p_ebitda", "Pro forma EBITDA",
          pf("{in_cloud_ebitda}/{eur_nok}+{ebitda}+{p_syn}-{p_integ}", "pf_ebitda"), "EUR m"); r += 1
    b.row(ws, r, "p_lev", "Net debt / EBITDA", pf("{pf_nd}/{p_ebitda}", "leverage"), "x", "x"); r += 1

    # ------------------------------------------------------------------ PF balance sheet
    ws = b.sheet("PF balance sheet", "Pro forma balance sheet on the 31 Dec 2025 balance sheets",
                 "Illustrative: as if the deal had closed on the latest reported balance sheet date.")
    for c, h in enumerate(["Cloudberry", "Orrön", "Deal adj.", "Pro forma"], start=2):
        ws.write(3, c, h, b.fmt["year"])
    o = lambda key: f"'Orron'!C{at[key] + 1}"
    c_cash = a.cloud_cash_nok / a.eur_nok
    c_assets = a.cloud_total_assets_nok / a.eur_nok
    c_debt = a.cloud_debt_nok / a.eur_nok
    c_eq = a.cloud_equity_nok / a.eur_nok
    lines = [
        # label, cloudberry formula/value, orrön formula/value, adjustment formula/value
        ("Cash", ("{cloud_cash_nok}/{eur_nok}", c_cash), (o("cash"), t["cash"][2025]),
         ("{new_debt}-{cash_paid}-{fees}", d["new_debt"] - d["cash_paid"] - d["fees"])),
        ("PP&E and other assets", ("({cloud_total_assets_nok}-{cloud_cash_nok})/{eur_nok}", c_assets - c_cash),
         (f"{o('total_assets')}-{o('cash')}", t["total_assets"][2025] - t["cash"][2025]),
         ("-{writedown}", -d["ppe_writedown"])),
        ("Goodwill", ("0", 0.0), ("0", 0.0), ("{goodwill}", d["goodwill"])),
        ("Total assets", None, None, None),
        ("Interest-bearing debt", ("{cloud_debt_nok}/{eur_nok}", c_debt), (o("debt"), t["debt"][2025]),
         ("{new_debt}", d["new_debt"])),
        ("Other liabilities", ("({cloud_total_assets_nok}-{cloud_equity_nok}-{cloud_debt_nok})/{eur_nok}", c_assets - c_eq - c_debt),
         (f"{o('current_liabs')}+{o('other_ltl')}", t["current_liabs"][2025] + t["other_ltl"][2025]), ("0", 0.0)),
        ("Equity", ("{cloud_equity_nok}/{eur_nok}", c_eq), (o("equity"), t["equity"][2025]),
         (f"{{stock_paid}}-{{fees}}-{o('equity')}", d["stock_paid"] - d["fees"] - t["equity"][2025])),
        ("Non-controlling interests", ("0", 0.0), (o("nci"), t["nci"][2025]), ("0", 0.0)),
        ("Total liabilities and equity", None, None, None),
        ("Check (should be zero)", None, None, None),
    ]
    r = 4
    vals = {}
    for i, (label, *cols) in enumerate(lines):
        row = r + i
        total = label.startswith("Total")
        ws.write(row, 0, label, b.fmt["bold" if total else "label"])
        for j in range(4):
            col = 2 + j
            letter = xl_col_to_name(col)
            if label == "Total assets":
                f, val = f"SUM({letter}{r + 1}:{letter}{r + 3})", sum(vals[(k, j)] for k in range(3))
            elif label == "Total liabilities and equity":
                f, val = f"SUM({letter}{r + 5}:{letter}{r + 8})", sum(vals[(k, j)] for k in range(4, 8))
            elif label.startswith("Check"):
                f, val = f"ROUND({letter}{r + 4}-{letter}{r + 9},6)", round(vals[(3, j)] - vals[(8, j)], 6)
            elif j == 3:
                f, val = f"SUM(C{row + 1}:E{row + 1})", sum(vals[(i, k)] for k in range(3))
            else:
                f, val = cols[j]
            vals[(i, j)] = val
            b.formula(ws, row, col, f, val, "num", total)

    # ------------------------------------------------------------------ Sensitivities
    ws = b.sheet("Sensitivities", "Sensitivities",
                 "Static values written by build_model.py (each cell is a full model rerun). Rerun the script after changing inputs.")
    r = 3

    def grid(r, title, row_label, col_label, rows, cols, fn, row_fmt, col_fmt, cell_fmt):
        ws.write(r, 0, title, b.fmt["section"]); r += 1
        ws.write(r, 0, f"{row_label} ↓   {col_label} →", b.fmt["note"])
        for j, cv in enumerate(cols):
            ws.write_number(r, 2 + j, cv, b.fmt[f"in_{col_fmt}"])
        r += 1
        for rv in rows:
            ws.write_number(r, 1, rv, b.fmt[f"in_{row_fmt}"])
            for j, cv in enumerate(cols):
                ws.write_number(r, 2 + j, fn(rv, cv), b.fmt[f"calc_{cell_fmt}"])
            r += 1
        return r + 1

    def flat_price(p):
        return {y: (model.Assumptions().price_mwh[y] if y < 2027 else p) for y in FORECAST}

    r = grid(r, "2028E EPS accretion: premium vs share of price paid in stock", "Premium", "Stock share",
             [0.0, 0.10, 0.20, 0.25, 0.30, 0.40], [1.0, 0.75, 0.5, 0.25],
             lambda pr, st: model.accretion(2028, premium=pr, stock_share=st), "pct", "pct", "pct")
    r = grid(r, "2028E EPS accretion: power price from 2027 vs Orrön G&A removed", "EUR/MWh", "G&A cut",
             [35, 40, 45, 50, 60, 70], [0.0, 0.3, 0.6, 0.9],
             lambda pr, g: model.accretion(2028, price_mwh=flat_price(pr), gna_cut=g), "num", "pct", "pct")
    r = grid(r, "DCF enterprise value of Orrön's existing business (EUR m)", "WACC", "Terminal growth",
             [0.065, 0.07, 0.075, 0.08, 0.085, 0.09], [0.01, 0.015, 0.02, 0.025],
             lambda w, g: model.run(model.Assumptions(wacc=w, tgr=g))["dcf"]["ev"], "pct", "pct", "num")
    r = grid(r, "DCF enterprise value (EUR m): power price from 2027 vs operating cost", "EUR/MWh", "Opex EUR/MWh",
             [35, 40, 45, 50, 60, 70], [18, 20, 22, 24],
             lambda pr, ox: model.run(model.Assumptions(price_mwh=flat_price(pr),
                                                        opex_mwh={y: (22 if y == 2026 else ox) for y in FORECAST}))["dcf"]["ev"],
             "num", "num", "num")
    be = model.breakeven_price()
    ws.write(r, 0, "Flat power price from 2027 that makes 2028E EPS-neutral", b.fmt["bold"])
    ws.write(r, 1, "EUR/MWh", b.fmt["note"])
    ws.write_number(r, 2, be, b.fmt["calc_num"])
    b.names["breakeven"] = ("scalar", ws.name, xl_rowcol_to_cell(r, 2, True, True))

    # ------------------------------------------------------------------ Summary
    ws = summary_ws
    r = 3
    b.section(ws, r, "Offer"); r += 1
    for label, template, value, unit, num in [
        ("Offer price per Orrön share", "{offer_price}", d["offer_price_sek"], "SEK", "px"),
        ("Premium", "{premium}", a.premium, "%", "pct"),
        ("Equity value", "{equity_value}", d["equity_value"], "EUR m", "num"),
        ("Transaction EV", "{deal_ev}", d["ev"], "EUR m", "num"),
        ("EV / EBITDA 2028E", "{ev_ebitda_28}", d["ev_ebitda_2028"], "x", "x"),
        ("Price / book", "{p_b}", d["price_to_book"], "x", "x"),
        ("Orrön holders' stake in combined company", "{orron_pct}", d["orron_holders_pct"], "%", "pct"),
    ]:
        b.scalar_formula(ws, r, f"s_{r}", label, template, value, unit, num); r += 1
    r += 1; b.section(ws, r, "What the price pays for"); r += 1
    for name, label, template, value, total in [
        ("s_dcf", "DCF of Orrön's existing business", "{dcf_ev}", s["dcf_ev"], False),
        ("s_jv", "JVs at book", "{t_jv_book}", s["jv_book"], False),
        ("s_syn", "PV of synergies, net of integration cost", "{synergy_pv}", s["synergy_pv"], False),
        ("s_dealcost", "Transaction EV + fees", "{deal_ev}+{fees}", s["deal_ev"] + s["fees"], False),
        ("s_gap", "Left to be justified by the development pipeline", "{s_dealcost}-{s_dcf}-{s_jv}-{s_syn}", s["gap"], True),
    ]:
        b.scalar_formula(ws, r, name, label, template, value, "EUR m", "num", total); r += 1
    r += 1; b.section(ws, r, "Cloudberry shareholders"); r += 1
    b.years_header(ws, r, PF_YEARS); r += 1
    b.row(ws, r, "s_acc", "EPS accretion / (dilution)", fc("{p_acc}", p["accretion"], PF_YEARS), "%", "pct"); r += 1
    b.row(ws, r, "s_lev", "Net debt / EBITDA", fc("{p_lev}", p["leverage"], PF_YEARS), "x", "x"); r += 2
    ws.write(r, 0, "Power price for EPS-neutral 2028E", b.fmt["label"])
    ws.write(r, 1, "EUR/MWh", b.fmt["note"])
    b.formula(ws, r, 2, "{breakeven}", be)

    b.close()
    return b.formulas


if __name__ == "__main__":
    n = len(build())
    print(f"wrote {OUT} ({n} formulas)")
