import openpyxl
import pytest

def test_excel_formulas_and_integrity():
    wb = openpyxl.load_workbook('output/Cloudberry_Orron_MA_Model_v3.xlsx', data_only=False)
    
    # Ensure key sheets exist
    expected_sheets = ['Executive Summary', 'Inputs', '3-Statement', 'DCF', 'Accretion-Dilution', 'Breaking Points']
    for sheet in expected_sheets:
        assert sheet in wb.sheetnames, f"Missing sheet: {sheet}"
        
    # Check for broken references in formulas
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str) and cell.startswith('='):
                    assert '#REF!' not in cell, f"Broken reference (#REF!) found in {sheetname}"
                    assert '#DIV/0!' not in cell, f"Division by zero error formula in {sheetname}"
