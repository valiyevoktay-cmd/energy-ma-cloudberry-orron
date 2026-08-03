import xlsxwriter

workbook = xlsxwriter.Workbook('Cloudberry_Orron_MA_Model_v3.xlsx')

font_name = 'Arial'
font_size = 10

# Formats (IB Conventions)
fmt_header = workbook.add_format({'bold': True, 'bottom': 1, 'align': 'center', 'font_name': font_name, 'font_size': font_size})
fmt_label = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': font_size})
fmt_sublabel = workbook.add_format({'indent': 1, 'font_name': font_name, 'font_size': font_size})

# Inputs: Blue
fmt_input_num = workbook.add_format({'font_color': 'blue', 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})
fmt_input_pct = workbook.add_format({'font_color': 'blue', 'num_format': '0.0%', 'font_name': font_name, 'font_size': font_size})

# Calculations: Black
fmt_calc_num = workbook.add_format({'font_color': 'black', 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})
fmt_calc_pct = workbook.add_format({'font_color': 'black', 'num_format': '0.0%', 'font_name': font_name, 'font_size': font_size})

# Links: Green
fmt_link_num = workbook.add_format({'font_color': 'green', 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})
fmt_link_pct = workbook.add_format({'font_color': 'green', 'num_format': '0.0%', 'font_name': font_name, 'font_size': font_size})

# Borders
fmt_border_top = workbook.add_format({'top': 1, 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})
fmt_border_bottom = workbook.add_format({'bottom': 1, 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})
fmt_border_both = workbook.add_format({'top': 1, 'bottom': 1, 'bold': True, 'num_format': '#,##0.0', 'font_name': font_name, 'font_size': font_size})

# Title
fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'font_name': font_name, 'font_color': '#002060'})

# --- 0. Executive Summary Tab ---
ws_sum = workbook.add_worksheet('Executive Summary')
ws_sum.hide_gridlines(2)
ws_sum.set_column('A:A', 50)
ws_sum.set_column('B:C', 15)

ws_sum.write('A1', 'Cloudberry / Orrön Energy - Strategic Acquisition Summary', fmt_title)

ws_sum.write('A3', 'Strategic Rationale', fmt_label)
ws_sum.write('A4', '• Accelerates Cloudberry\'s Nordic renewable footprint (hydro + wind) scale.', fmt_sublabel)
ws_sum.write('A5', '• Revenue diversification reducing volatility from local price areas (NOA).', fmt_sublabel)
ws_sum.write('A6', '• Synergies derived from grid connection sharing and operational scaling.', fmt_sublabel)

ws_sum.write('A8', 'Key Deal Metrics', fmt_label)
ws_sum.write('A9', 'Implied Target Enterprise Value (EV) €M', fmt_sublabel)
ws_sum.write_formula('B9', "='DCF'!C24", fmt_link_num)

ws_sum.write('A10', 'Pro Forma 2026E EPS Accretion / (Dilution) %', fmt_sublabel)
ws_sum.write_formula('B10', "='Accretion-Dilution'!D13", fmt_link_pct)


# --- 1. Inputs Tab ---
ws_in = workbook.add_worksheet('Inputs')
ws_in.hide_gridlines(2)
ws_in.set_column('A:B', 40)
ws_in.set_column('C:C', 15)

ws_in.write('A1', 'Orrön Energy (Target) Assumptions', fmt_label)
ws_in.write('A2', 'Annual Revenue Growth (2026E+)', fmt_sublabel)
ws_in.write('C2', 0.05, fmt_input_pct)
ws_in.write('A3', 'EBITDA Margin', fmt_sublabel)
ws_in.write('C3', 0.35, fmt_input_pct)
ws_in.write('A4', 'D&A (% of Revenue)', fmt_sublabel)
ws_in.write('C4', 0.15, fmt_input_pct)
ws_in.write('A5', 'Tax Rate', fmt_sublabel)
ws_in.write('C5', 0.22, fmt_input_pct)
ws_in.write('A6', 'Capex (% of Revenue)', fmt_sublabel)
ws_in.write('C6', 0.20, fmt_input_pct)
ws_in.write('A7', 'NWC (% of Revenue)', fmt_sublabel)
ws_in.write('C7', 0.10, fmt_input_pct)

ws_in.write('A9', 'Valuation & Deal Assumptions', fmt_label)
ws_in.write('A10', 'WACC', fmt_sublabel)
ws_in.write('C10', 0.085, fmt_input_pct)
ws_in.write('A11', 'Terminal Growth Rate (PGR)', fmt_sublabel)
ws_in.write('C11', 0.02, fmt_input_pct)
ws_in.write('A12', 'Cloudberry Standalone Net Income (2026E, €M)', fmt_sublabel)
ws_in.write('C12', 30.0, fmt_input_num)
ws_in.write('A13', 'Cloudberry Shares Outstanding (M)', fmt_sublabel)
ws_in.write('C13', 290.0, fmt_input_num)
ws_in.write('A14', 'New Shares Issued to Target (M)', fmt_sublabel) # Deal structure: 27% of enlarged company
ws_in.write('C14', 107.0, fmt_input_num)
ws_in.write('A15', 'Cost of Debt (after-tax)', fmt_sublabel)
ws_in.write('C15', 0.045, fmt_input_pct)
ws_in.write('A16', 'New Debt Raised for Cash Portion (€M)', fmt_sublabel)
ws_in.write('C16', 4.2, fmt_input_num)
ws_in.write('A17', 'Post-tax Synergies (€M)', fmt_sublabel)
ws_in.write('C17', 5.0, fmt_input_num)

# --- 2. 3-Statement Model ---
ws_3s = workbook.add_worksheet('3-Statement')
ws_3s.hide_gridlines(2)
ws_3s.set_column('A:B', 30)
ws_3s.set_column('C:I', 12)

years = ['2024A', '2025A', '2026E', '2027E', '2028E', '2029E', '2030E']
for col, year in enumerate(years, 2):
    ws_3s.write(1, col, year, fmt_header)

ws_3s.write('A3', 'INCOME STATEMENT', fmt_label)
ws_3s.write('B4', 'Revenue')
ws_3s.write_row('C4', [35.0, 36.0], fmt_input_num) # 24A, 25A
# Projections 2026E-2030E
for c in range(4, 9):
    # E4 = D4 * (1 + Inputs!C2)
    ws_3s.write_formula(3, c, f'={chr(65+c-1)}4*(1+Inputs!$C$2)', fmt_link_num)

ws_3s.write('B5', 'COGS & Opex')
ws_3s.write_row('C5', [25.0, 40.5], fmt_input_num)
for c in range(4, 9):
    # EBITDA margin = Inputs!C3 -> COGS = Rev * (1 - Margin)
    ws_3s.write_formula(4, c, f'={chr(65+c)}4*(1-Inputs!$C$3)', fmt_link_num)

ws_3s.write('B6', 'EBITDA', fmt_label)
for c in range(2, 9):
    ws_3s.write_formula(5, c, f'={chr(65+c)}4-{chr(65+c)}5', fmt_border_top)

ws_3s.write('B7', 'D&A')
ws_3s.write_row('C7', [5.0, 5.5], fmt_input_num)
for c in range(4, 9):
    # D&A = Rev * Inputs!C4
    ws_3s.write_formula(6, c, f'={chr(65+c)}4*Inputs!$C$4', fmt_link_num)

ws_3s.write('B8', 'EBIT', fmt_label)
for c in range(2, 9):
    ws_3s.write_formula(7, c, f'={chr(65+c)}6-{chr(65+c)}7', fmt_border_top)

ws_3s.write('B9', 'Interest Expense')
ws_3s.write_row('C9', [2.0, 2.5], fmt_input_num)
for c in range(4, 9):
    ws_3s.write(8, c, 2.5, fmt_calc_num) # Flat for simplicity

ws_3s.write('B10', 'EBT', fmt_label)
for c in range(2, 9):
    ws_3s.write_formula(9, c, f'={chr(65+c)}8-{chr(65+c)}9', fmt_border_top)

ws_3s.write('B11', 'Taxes')
for c in range(2, 9):
    # Max(EBT * TaxRate, 0)
    ws_3s.write_formula(10, c, f'=MAX({chr(65+c)}10*Inputs!$C$5, 0)', fmt_link_num)

ws_3s.write('B12', 'Net Income', fmt_label)
for c in range(2, 9):
    ws_3s.write_formula(11, c, f'={chr(65+c)}10-{chr(65+c)}11', fmt_border_both)

# NWC and Capex sections for DCF linking
ws_3s.write('A14', 'CASH FLOW DRIVERS', fmt_label)
ws_3s.write('B15', 'NWC')
for c in range(2, 9):
    ws_3s.write_formula(14, c, f'={chr(65+c)}4*Inputs!$C$7', fmt_link_num)

ws_3s.write('B16', 'Change in NWC')
ws_3s.write(15, 2, 0, fmt_calc_num) # 24A
for c in range(3, 9):
    ws_3s.write_formula(15, c, f'={chr(65+c)}15-{chr(65+c-1)}15', fmt_calc_num)

ws_3s.write('B17', 'Capital Expenditures (Capex)')
ws_3s.write_row('C17', [10.0, 8.0], fmt_input_num)
for c in range(4, 9):
    ws_3s.write_formula(16, c, f'={chr(65+c)}4*Inputs!$C$6', fmt_link_num)

# --- 3. DCF ---
ws_dcf = workbook.add_worksheet('DCF')
ws_dcf.hide_gridlines(2)
ws_dcf.set_column('A:B', 30)
ws_dcf.set_column('C:I', 12)

for col, year in enumerate(years[2:], 4): # Start 2026E at E
    ws_dcf.write(1, col, year, fmt_header)

ws_dcf.write('A3', 'UNLEVERED FREE CASH FLOW (FCFF)', fmt_label)
ws_dcf.write('B4', 'EBIT')
for c in range(4, 9):
    ws_dcf.write_formula(3, c, f"='3-Statement'!{chr(65+c)}8", fmt_link_num)

ws_dcf.write('B5', 'Less: Taxes on EBIT')
for c in range(4, 9):
    ws_dcf.write_formula(4, c, f'={chr(65+c)}4*Inputs!$C$5', fmt_link_num)

ws_dcf.write('B6', 'NOPAT', fmt_label)
for c in range(4, 9):
    ws_dcf.write_formula(5, c, f'={chr(65+c)}4-{chr(65+c)}5', fmt_border_top)

ws_dcf.write('B7', 'Plus: D&A')
for c in range(4, 9):
    ws_dcf.write_formula(6, c, f"='3-Statement'!{chr(65+c)}7", fmt_link_num)

ws_dcf.write('B8', 'Less: Capex')
for c in range(4, 9):
    ws_dcf.write_formula(7, c, f"='3-Statement'!{chr(65+c)}17", fmt_link_num)

ws_dcf.write('B9', 'Less: Change in NWC')
for c in range(4, 9):
    ws_dcf.write_formula(8, c, f"='3-Statement'!{chr(65+c)}16", fmt_link_num)

ws_dcf.write('B10', 'FCFF', fmt_label)
for c in range(4, 9):
    ws_dcf.write_formula(9, c, f'={chr(65+c)}6+{chr(65+c)}7-{chr(65+c)}8-{chr(65+c)}9', fmt_border_both)

ws_dcf.write('B11', 'Discount Factor')
for c in range(4, 9):
    period = c - 3 # E=1, F=2, etc.
    ws_dcf.write_formula(10, c, f'=1/((1+Inputs!$C$10)^{period})', fmt_link_num)

ws_dcf.write('B12', 'PV of FCFF', fmt_label)
for c in range(4, 9):
    ws_dcf.write_formula(11, c, f'={chr(65+c)}10*{chr(65+c)}11', fmt_border_top)

ws_dcf.write('A14', 'TERMINAL VALUE', fmt_label)
ws_dcf.write('B15', '2030E FCFF')
ws_dcf.write_formula('C15', '=I10', fmt_link_num)
ws_dcf.write('B16', 'Terminal Growth Rate')
ws_dcf.write_formula('C16', '=Inputs!C11', fmt_link_pct)
ws_dcf.write('B17', 'WACC')
ws_dcf.write_formula('C17', '=Inputs!C10', fmt_link_pct)
ws_dcf.write('B18', 'Terminal Value')
ws_dcf.write_formula('C18', '=IF(C17<=C16, "Error", C15*(1+C16)/(C17-C16))', fmt_border_top)
ws_dcf.write('B19', 'PV of Terminal Value')
ws_dcf.write_formula('C19', '=C18*I11', fmt_border_top)

ws_dcf.write('A21', 'ENTERPRISE VALUE', fmt_label)
ws_dcf.write('B22', 'Sum of PV of FCFF')
ws_dcf.write_formula('C22', '=SUM(E12:I12)', fmt_calc_num)
ws_dcf.write('B23', 'PV of Terminal Value')
ws_dcf.write_formula('C23', '=C19', fmt_calc_num)
ws_dcf.write('B24', 'Implied Enterprise Value (EV)', fmt_label)
ws_dcf.write_formula('C24', '=C22+C23', fmt_border_both)

# --- 4. Accretion / Dilution ---
ws_acc = workbook.add_worksheet('Accretion-Dilution')
ws_acc.hide_gridlines(2)
ws_acc.set_column('A:B', 35)
ws_acc.set_column('C:D', 15)

ws_acc.write('A1', 'Accretion / Dilution Analysis (2026E)', fmt_label)
ws_acc.write('B2', 'Cloudberry', fmt_header)
ws_acc.write('C2', 'Orrön', fmt_header)
ws_acc.write('D2', 'Pro Forma', fmt_header)

ws_acc.write('A3', 'Net Income (€M)')
ws_acc.write_formula('B3', '=Inputs!C12', fmt_link_num)
ws_acc.write_formula('C3', "='3-Statement'!E12", fmt_link_num) # 2026E Net income

ws_acc.write('A4', 'Plus: Post-Tax Synergies')
ws_acc.write_formula('D4', '=Inputs!C17', fmt_link_num)

ws_acc.write('A5', 'Less: New Interest Expense (Post-Tax)')
ws_acc.write_formula('D5', '=-Inputs!C16*Inputs!C15', fmt_link_num)

ws_acc.write('A6', 'Pro Forma Net Income', fmt_label)
ws_acc.write_formula('D6', '=B3+C3+D4+D5', fmt_border_top)

ws_acc.write('A8', 'Shares Outstanding (M)')
ws_acc.write_formula('B8', '=Inputs!C13', fmt_link_num)
ws_acc.write_formula('D8', '=Inputs!C13+Inputs!C14', fmt_link_num)

ws_acc.write('A10', 'EPS (€)')
ws_acc.write_formula('B10', '=IF(B8<=0, "Error", B3/B8)', fmt_calc_num)
ws_acc.write_formula('D10', '=IF(D8<=0, "Error", D6/D8)', fmt_calc_num)

ws_acc.write('A12', 'Accretion / (Dilution) €')
ws_acc.write_formula('D12', '=D10-B10', fmt_border_top)
ws_acc.write('A13', 'Accretion / (Dilution) %')
ws_acc.write_formula('D13', '=D12/B10', fmt_calc_pct)

# --- 5. Breaking Points ---
ws_bp = workbook.add_worksheet('Breaking Points')
ws_bp.hide_gridlines(2)
ws_bp.set_column('A:A', 40)
ws_bp.set_column('B:G', 12)

ws_bp.write('A1', 'Downside Stress Tests & Breaking Points', fmt_title)

# Table 1: The Valuation Collapse
ws_bp.write('A3', 'Table 1: At What Cost of Capital Does Valuation Collapse?', fmt_label)
ws_bp.write('A4', 'Implied Enterprise Value (EV) vs. WACC and Terminal Growth', fmt_sublabel)

# Link the output metric to the top-left corner
ws_bp.write_formula('B6', "='DCF'!C24", fmt_link_num)

# Row input: Terminal Growth Rate (0% to 4%)
tgr_range = [0.00, 0.01, 0.02, 0.03, 0.04]
ws_bp.write_row('C6', tgr_range, fmt_header)
ws_bp.write('B5', 'Terminal Growth ->')

# Col input: WACC (6.5% to 12.5%)
wacc_range = [0.065, 0.08, 0.095, 0.110, 0.125]
ws_bp.write_column('B7', wacc_range, fmt_header)
ws_bp.write('A7', 'WACC v')

# Data Table for Table 1
ws_bp.write_array_formula('C7:G11', '{=TABLE(Inputs!C11, Inputs!C10)}')

# Conditional Formatting for Table 1 (Red for EV < 0, else Gray)
fmt_neutral = workbook.add_format({'bg_color': '#E8E8E8', 'font_color': 'black', 'num_format': '#,##0.0'})
fmt_fail = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white', 'bold': True, 'num_format': '#,##0.0'})

ws_bp.conditional_format('C7:G11', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_fail})
ws_bp.conditional_format('C7:G11', {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt_neutral})

# Table 2: The Operational Disaster
ws_bp.write('A14', 'Table 2: When Does the Deal Destroy EPS?', fmt_label)
ws_bp.write('A15', 'Pro Forma EPS Accretion / (Dilution) % vs. Target Revenue Growth and Synergies', fmt_sublabel)

ws_bp.write_formula('B17', "='Accretion-Dilution'!D13", fmt_link_pct)

# Row input: Target Revenue Growth (-10% to +10%)
rev_growth_range = [-0.10, -0.05, 0.00, 0.05, 0.10]
ws_bp.write_row('C17', rev_growth_range, fmt_header)
ws_bp.write('B16', 'Target Rev Growth ->')

# Col input: Synergies (0 to 10)
syn_range = [0.0, 2.5, 5.0, 7.5, 10.0]
ws_bp.write_column('B18', syn_range, fmt_header)
ws_bp.write('A18', 'Post-tax Synergies v')

# Data Table for Table 2
ws_bp.write_array_formula('C18:G22', '{=TABLE(Inputs!C2, Inputs!C17)}')

fmt_fail_pct = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white', 'bold': True, 'num_format': '0.0%'})
fmt_neutral_pct = workbook.add_format({'bg_color': '#E8E8E8', 'font_color': 'black', 'num_format': '0.0%'})

ws_bp.conditional_format('C18:G22', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_fail_pct})
ws_bp.conditional_format('C18:G22', {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt_neutral_pct})


workbook.close()
print("Excel model successfully generated with IB conventions, Summary Tab, and Breaking Points.")
