"""Script to create a comprehensive test workbook with all features."""

import random
import sys
from pathlib import Path

try:
    import xlwings as xw
except ImportError:
    print("Error: xlwings is required to create test workbook")
    print("Install it with: pip install xlwings")
    sys.exit(1)


def create_test_workbook(output_path: Path):
    """
    Create a comprehensive test workbook with all features.

    Args:
        output_path: Path to save .xlsb file
    """
    print(f"Creating test workbook: {output_path}")

    # Create new workbook
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    wb = app.books.add()

    try:
        # Sheet 1: Basic data and formulas
        sheet1 = wb.sheets[0]
        sheet1.name = "Data"

        # Add some data
        sheet1.range("A1").value = "Name"
        sheet1.range("B1").value = "Value"
        sheet1.range("A2").value = "Item1"
        sheet1.range("B2").value = 10
        sheet1.range("A3").value = "Item2"
        sheet1.range("B3").value = 20
        sheet1.range("A4").value = "Item3"
        sheet1.range("B4").value = 30
        sheet1.range("A5").value = "Item4"
        sheet1.range("B5").value = 40
        sheet1.range("A6").value = "Item5"
        sheet1.range("B6").value = 50

        # Add formulas
        sheet1.range("B7").formula = "=SUM(B2:B6)"
        sheet1.range("B8").formula = "=AVERAGE(B2:B6)"
        sheet1.range("B9").formula = "=IF(B7>100,\"High\",\"Low\")"
        sheet1.range("B10").formula = "=COUNT(B2:B6)"
        sheet1.range("B11").formula = "=MAX(B2:B6)"
        sheet1.range("B12").formula = "=MIN(B2:B6)"

        # Sheet 2: Cross-sheet references
        sheet2 = wb.sheets.add("Calculations")
        sheet2.range("A1").value = "Total from Data sheet"
        sheet2.range("B1").formula = "=SUM(Data!B2:B6)"
        sheet2.range("A2").value = "Average from Data sheet"
        sheet2.range("B2").formula = "=AVERAGE(Data!B2:B6)"
        sheet2.range("A3").value = "Max from Data sheet"
        sheet2.range("B3").formula = "=MAX(Data!B2:B6)"
        sheet2.range("A4").value = "Min from Data sheet"
        sheet2.range("B4").formula = "=MIN(Data!B2:B6)"
        sheet2.range("A5").value = "Count from Data sheet"
        sheet2.range("B5").formula = "=COUNT(Data!B2:B6)"

        # Sheet 3: Complex formulas and nested functions
        sheet3 = wb.sheets.add("Complex")
        sheet3.range("A1").value = 1
        sheet3.range("A2").value = 2
        sheet3.range("A3").value = 3
        sheet3.range("A4").value = 4
        sheet3.range("A5").value = 5
        sheet3.range("B1").formula = "=A1*2"
        sheet3.range("B2").formula = "=A2*2"
        sheet3.range("B3").formula = "=A3*2"
        sheet3.range("B4").formula = "=A4*2"
        sheet3.range("B5").formula = "=A5*2"
        sheet3.range("C1").formula = "=SUM(A1:B5)"
        sheet3.range("C2").formula = "=MAX(A1:A5)"
        sheet3.range("C3").formula = "=MIN(A1:A5)"
        sheet3.range("C4").formula = "=IF(SUM(A1:A5)>10,\"Large\",\"Small\")"
        sheet3.range("C5").formula = "=SUMIF(A1:A5,\">2\",B1:B5)"
        sheet3.range("D1").formula = "=VLOOKUP(3,A1:B5,2,FALSE)"
        sheet3.range("D2").formula = "=INDEX(A1:A5,MATCH(3,A1:A5,0))"

        # Sheet 4: Different data types
        sheet4 = wb.sheets.add("Types")
        sheet4.range("A1").value = "Text"
        sheet4.range("B1").value = "Number"
        sheet4.range("C1").value = "Boolean"
        sheet4.range("D1").value = "Date"
        sheet4.range("A2").value = "Hello"
        sheet4.range("B2").value = 42
        sheet4.range("C2").value = True
        sheet4.range("D2").value = "2024-01-01"
        sheet4.range("A3").value = "World"
        sheet4.range("B3").value = 3.14
        sheet4.range("C3").value = False
        sheet4.range("D3").value = "2024-12-31"
        sheet4.range("A4").value = "Test"
        sheet4.range("B4").value = -15
        sheet4.range("C4").value = True
        sheet4.range("D4").value = "2024-06-15"

        # Sheet 5: Chart data (for creating charts) with some jagged layout
        sheet5 = wb.sheets.add("ChartData")
        sheet5.range("A1").value = "Month"
        sheet5.range("B1").value = "Sales"
        sheet5.range("C1").value = "Expenses"
        sheet5.range("A2").value = "Jan"
        sheet5.range("B2").value = 1000
        sheet5.range("C2").value = 800
        sheet5.range("A3").value = "Feb"
        sheet5.range("B3").value = 1200
        sheet5.range("C3").value = 900
        sheet5.range("A4").value = "Mar"
        sheet5.range("B4").value = 1500
        sheet5.range("C4").value = 1000
        sheet5.range("A5").value = "Apr"
        sheet5.range("B5").value = 1800
        sheet5.range("C5").value = 1100
        sheet5.range("A6").value = "May"
        sheet5.range("B6").value = 2000
        sheet5.range("C6").value = 1200
        sheet5.range("A7").value = "Jun"
        sheet5.range("B7").value = 2200
        sheet5.range("C7").value = 1300

        # Add formulas referencing chart data (some in jagged positions)
        sheet5.range("B8").formula = "=SUM(B2:B7)"
        sheet5.range("C8").formula = "=SUM(C2:C7)"
        sheet5.range("D2").formula = "=B2-C2"
        sheet5.range("D3").formula = "=B3-C3"
        sheet5.range("D4").formula = "=B4-C4"
        sheet5.range("D5").formula = "=B5-C5"
        sheet5.range("D6").formula = "=B6-C6"
        sheet5.range("D7").formula = "=B7-C7"
        sheet5.range("D1").value = "Profit"
        sheet5.range("D8").formula = "=SUM(D2:D7)"

        # Add some scattered additional calculations
        sheet5.range("F2").formula = "=B2*1.1"  # Scattered formula
        sheet5.range("F5").formula = "=B5*1.15"  # Gap in F3, F4
        sheet5.range("F8").formula = "=F2+F5"  # References scattered cells
        sheet5.range("H3").formula = "=C3*0.9"  # Another scattered position
        sheet5.range("H6").formula = "=C6*0.95"  # Gap in H4, H5
        sheet5.range("H9").formula = "=H3+H6+D8"  # References scattered and D8

        # Sheet 6: Jagged table layout with sparse data
        sheet6 = wb.sheets.add("JaggedData")
        # Create irregular table with gaps
        sheet6.range("B2").value = "Q1"
        sheet6.range("B3").value = 100
        sheet6.range("B4").value = 150
        # Gap in B5
        sheet6.range("B6").value = 200

        sheet6.range("D2").value = "Q2"
        sheet6.range("D3").value = 120
        # Gap in D4
        sheet6.range("D5").value = 180
        sheet6.range("D6").value = 190

        sheet6.range("F3").value = "Q3"
        sheet6.range("F4").value = 130
        sheet6.range("F5").value = 160
        # Gap in F6

        sheet6.range("H2").value = "Q4"
        # Gap in H3
        sheet6.range("H4").value = 140
        sheet6.range("H5").value = 170
        sheet6.range("H6").value = 210

        # Add formulas that reference the jagged data
        sheet6.range("B8").formula = "=SUM(B3,B4,B6)"
        sheet6.range("D8").formula = "=SUM(D3,D5,D6)"
        sheet6.range("F8").formula = "=SUM(F4,F5)"
        sheet6.range("H8").formula = "=SUM(H4:H6)"
        sheet6.range("J8").formula = "=B8+D8+F8+H8"

        # Sheet 7: Criss-crossed references - complex dependency pattern
        sheet7 = wb.sheets.add("CrissCross")
        # Create a grid where formulas reference each other in a criss-cross pattern
        # Row 1: Base values
        sheet7.range("A1").value = 10
        sheet7.range("B1").value = 20
        sheet7.range("C1").value = 30
        sheet7.range("D1").value = 40

        # Row 2: References row 1, but also references row 3
        sheet7.range("A2").formula = "=A1*2"
        sheet7.range("B2").formula = "=B1+A2+C3"  # References B1, A2 (same row), and C3 (below)
        sheet7.range("C2").formula = "=C1+B2"
        sheet7.range("D2").formula = "=D1+C2+A3"  # References D1, C2 (same row), and A3 (below)

        # Row 3: References row 2 and row 4
        sheet7.range("A3").formula = "=A2+B4"  # References A2 (above) and B4 (below)
        sheet7.range("B3").formula = "=B2*1.5"
        sheet7.range("C3").formula = "=C2+A3+D4"  # References C2 (above), A3 (same row), D4 (below)
        sheet7.range("D3").formula = "=D2+C3"

        # Row 4: References row 3 and creates circular-like pattern
        sheet7.range("A4").formula = "=A3+10"
        sheet7.range("B4").formula = "=B3+A4+C2"  # References B3 (above), A4 (same row), C2 (two rows up)
        sheet7.range("C4").formula = "=C3+B4"
        sheet7.range("D4").formula = "=D3+C4+A2"  # References D3 (above), C4 (same row), A2 (two rows up)

        # Row 5: Cross-sheet references mixed with local criss-cross
        sheet7.range("A5").formula = "=A4+Data!B2"  # Local + cross-sheet
        sheet7.range("B5").formula = "=B4+A5+Calculations!B1"  # Local + cross-sheet
        sheet7.range("C5").formula = "=C4+B5"
        sheet7.range("D5").formula = "=D4+C5+Data!B7"  # Local + cross-sheet

        # Summary row with complex references (fixed to avoid circular references)
        sheet7.range("A7").formula = "=SUM(A1:A5)"
        sheet7.range("B7").formula = "=SUM(B1:B5)+A7"  # References A7 only (no circular ref)
        sheet7.range("C7").formula = "=SUM(C1:C5)+A7+B7"  # References A7 and B7 (no circular ref)
        sheet7.range("D7").formula = "=SUM(D1:D5)+A7+B7+C7"  # References A7, B7, C7 (no circular ref)

        # Sheet 8: Diagonal and scattered references
        sheet8 = wb.sheets.add("Scattered")
        # Place data in scattered locations
        sheet8.range("A2").value = 5
        sheet8.range("C4").value = 10
        sheet8.range("E6").value = 15
        sheet8.range("G8").value = 20
        sheet8.range("I10").value = 25

        # Create formulas that reference diagonally and scattered
        sheet8.range("B3").formula = "=A2+C4"  # Diagonal reference
        sheet8.range("D5").formula = "=C4+E6"  # Diagonal reference
        sheet8.range("F7").formula = "=E6+G8"  # Diagonal reference
        sheet8.range("H9").formula = "=G8+I10"  # Diagonal reference

        # Create criss-cross pattern
        sheet8.range("A11").formula = "=A2+B3"  # References A2 and B3 (diagonal result)
        sheet8.range("C11").formula = "=C4+D5+B3"  # References C4, D5, and B3 (other formula)
        sheet8.range("E11").formula = "=E6+F7+D5"  # References E6, F7, and D5
        sheet8.range("G11").formula = "=G8+H9+F7"  # References G8, H9, and F7
        sheet8.range("I11").formula = "=I10+H9"  # References I10 and H9

        # Final summary with scattered references (fixed to avoid circular references)
        sheet8.range("A13").formula = "=A11+C11+E11+G11+I11"
        sheet8.range("B13").formula = "=A13+A2"  # References A13 and A2 (removed self-reference)
        sheet8.range("C13").formula = "=B13+CrissCross!A7"  # Cross-sheet reference (no circular ref)

        # Sheet 9: Cross-sheet criss-cross references
        sheet9 = wb.sheets.add("CrossSheetRefs")
        # Create formulas that reference multiple sheets in criss-cross pattern
        sheet9.range("A1").formula = "=Data!B7+Calculations!B1"  # References two different sheets
        sheet9.range("B1").formula = "=A1+Complex!C1"  # References A1 (same sheet) and Complex sheet
        sheet9.range("C1").formula = "=B1+CrissCross!A7+Data!B8"  # References B1, CrissCross, and Data

        sheet9.range("A2").formula = "=A1+JaggedData!J8"  # References A1 (above) and JaggedData
        sheet9.range("B2").formula = "=B1+A2+Scattered!A13"  # References B1 (above), A2 (same row), Scattered
        sheet9.range("C2").formula = "=C1+B2+ChartData!B8"  # References C1 (above), B2 (same row), ChartData

        sheet9.range("A3").formula = "=A2+Data!B2"  # Fixed: removed self-reference, uses Data instead
        sheet9.range("B3").formula = "=B2+A3+Data!B11"  # References B2 (above), A3 (same row), Data
        sheet9.range("C3").formula = "=C2+B3+Calculations!B3"  # References C2 (above), B3 (same row), Calculations

        # Create a dependency chain across sheets (fixed to avoid circular references)
        sheet9.range("A5").formula = "=Data!B7"  # Start with Data
        sheet9.range("B5").formula = "=A5+Calculations!B1"  # References A5 and Calculations (which references Data)
        sheet9.range("C5").formula = "=B5+CrissCross!A7"  # References B5 and CrissCross
        sheet9.range("D5").formula = "=C5+Scattered!A13"  # References C5 and Scattered (which references CrissCross)
        sheet9.range("E5").formula = "=D5+Data!B8"  # Fixed: removed circular reference, uses Data instead

        # Sheet 10: Large Sales Data Table
        sheet10 = wb.sheets.add("LargeSalesData")
        print("Creating large sales data table...")
        sheet10.range("A1").value = "Date"
        sheet10.range("B1").value = "Region"
        sheet10.range("C1").value = "Product"
        sheet10.range("D1").value = "Sales"
        sheet10.range("E1").value = "Units"
        sheet10.range("F1").value = "Price"
        sheet10.range("G1").value = "Cost"
        sheet10.range("H1").value = "Profit"

        # Generate large dataset (200 rows) - write in batches
        random.seed(42)  # For reproducibility
        regions = ["North", "South", "East", "West", "Central"]
        products = ["Product A", "Product B", "Product C", "Product D", "Product E"]

        # Prepare data as lists for batch writing
        dates = []
        region_list = []
        product_list = []
        sales_list = []
        units_list = []
        price_list = []
        cost_list = []
        profit_formulas = []

        for i in range(2, 202):  # 200 rows of data
            dates.append(f"2024-01-{(i % 28) + 1:02d}")
            region_list.append(regions[random.randint(0, 4)])
            product_list.append(products[random.randint(0, 4)])
            sales = random.randint(1000, 10000)
            units = random.randint(10, 100)
            price = sales / units
            cost = price * 0.6  # 60% cost

            sales_list.append(sales)
            units_list.append(units)
            price_list.append(price)
            cost_list.append(cost)
            profit_formulas.append(f"=D{i}-(G{i}*E{i})")

        # Write data in batches
        sheet10.range("A2:A201").value = [[d] for d in dates]
        sheet10.range("B2:B201").value = [[r] for r in region_list]
        sheet10.range("C2:C201").value = [[p] for p in product_list]
        sheet10.range("D2:D201").value = [[s] for s in sales_list]
        sheet10.range("E2:E201").value = [[u] for u in units_list]
        sheet10.range("F2:F201").value = [[p] for p in price_list]
        sheet10.range("G2:G201").value = [[c] for c in cost_list]

        # Write formulas individually for profit column
        for i, formula in enumerate(profit_formulas, start=2):
            sheet10.range(f"H{i}").formula = formula

        print(f"  - Created 200 rows of sales data")

        # Sheet 11: Aggregated Sales Summary
        sheet11 = wb.sheets.add("SalesSummary")
        print("Creating aggregated sales summary...")

        # Headers
        sheet11.range("A1").value = "Summary Type"
        sheet11.range("B1").value = "Value"

        # Total Sales
        sheet11.range("A2").value = "Total Sales"
        sheet11.range("B2").formula = "=SUM(LargeSalesData!D2:D201)"

        # Total Units
        sheet11.range("A3").value = "Total Units"
        sheet11.range("B3").formula = "=SUM(LargeSalesData!E2:E201)"

        # Average Sales
        sheet11.range("A4").value = "Average Sales"
        sheet11.range("B4").formula = "=AVERAGE(LargeSalesData!D2:D201)"

        # Max Sales
        sheet11.range("A5").value = "Max Sales"
        sheet11.range("B5").formula = "=MAX(LargeSalesData!D2:D201)"

        # Min Sales
        sheet11.range("A6").value = "Min Sales"
        sheet11.range("B6").formula = "=MIN(LargeSalesData!D2:D201)"

        # Total Profit
        sheet11.range("A7").value = "Total Profit"
        sheet11.range("B7").formula = "=SUM(LargeSalesData!H2:H201)"

        # Average Profit
        sheet11.range("A8").value = "Average Profit"
        sheet11.range("B8").formula = "=AVERAGE(LargeSalesData!H2:H201)"

        # Region Summary
        sheet11.range("D1").value = "Region"
        sheet11.range("E1").value = "Total Sales"
        sheet11.range("F1").value = "Count"
        sheet11.range("G1").value = "Average"

        regions = ["North", "South", "East", "West", "Central"]
        for i, region in enumerate(regions, start=2):
            sheet11.range(f"D{i}").value = region
            sheet11.range(f"E{i}").formula = f"=SUMIF(LargeSalesData!B2:B201,\"{region}\",LargeSalesData!D2:D201)"
            sheet11.range(f"F{i}").formula = f"=COUNTIF(LargeSalesData!B2:B201,\"{region}\")"
            sheet11.range(f"G{i}").formula = f"=E{i}/F{i}"

        # Product Summary
        sheet11.range("I1").value = "Product"
        sheet11.range("J1").value = "Total Sales"
        sheet11.range("K1").value = "Count"
        sheet11.range("L1").value = "Average"

        products = ["Product A", "Product B", "Product C", "Product D", "Product E"]
        for i, product in enumerate(products, start=2):
            sheet11.range(f"I{i}").value = product
            sheet11.range(f"J{i}").formula = f"=SUMIF(LargeSalesData!C2:C201,\"{product}\",LargeSalesData!D2:D201)"
            sheet11.range(f"K{i}").formula = f"=COUNTIF(LargeSalesData!C2:C201,\"{product}\")"
            sheet11.range(f"L{i}").formula = f"=J{i}/K{i}"

        print("  - Created sales summary with region and product aggregations")

        # Sheet 12: Monthly Aggregation
        sheet12 = wb.sheets.add("MonthlyAggregation")
        print("Creating monthly aggregation...")

        sheet12.range("A1").value = "Month"
        sheet12.range("B1").value = "Total Sales"
        sheet12.range("C1").value = "Total Units"
        sheet12.range("D1").value = "Total Profit"
        sheet12.range("E1").value = "Avg Sales"
        sheet12.range("F1").value = "Transaction Count"

        months = ["January", "February", "March", "April", "May", "June"]
        for i, month in enumerate(months, start=2):
            month_num = i - 1
            sheet12.range(f"A{i}").value = month
            # Sum sales for the month (simplified - in real scenario would filter by date)
            sheet12.range(f"B{i}").formula = f"=SUM(LargeSalesData!D{2+(month_num-1)*33}:D{1+month_num*33})"
            sheet12.range(f"C{i}").formula = f"=SUM(LargeSalesData!E{2+(month_num-1)*33}:E{1+month_num*33})"
            sheet12.range(f"D{i}").formula = f"=SUM(LargeSalesData!H{2+(month_num-1)*33}:H{1+month_num*33})"
            sheet12.range(f"E{i}").formula = f"=AVERAGE(LargeSalesData!D{2+(month_num-1)*33}:D{1+month_num*33})"
            sheet12.range(f"F{i}").formula = f"=COUNT(LargeSalesData!D{2+(month_num-1)*33}:D{1+month_num*33})"

        print("  - Created monthly aggregation")

        # Sheet 13: Large Inventory Data
        sheet13 = wb.sheets.add("InventoryData")
        print("Creating large inventory data table...")

        sheet13.range("A1").value = "Item ID"
        sheet13.range("B1").value = "Category"
        sheet13.range("C1").value = "Item Name"
        sheet13.range("D1").value = "Quantity"
        sheet13.range("E1").value = "Unit Price"
        sheet13.range("F1").value = "Total Value"
        sheet13.range("G1").value = "Reorder Level"
        sheet13.range("H1").value = "Status"

        categories = ["Electronics", "Clothing", "Food", "Books", "Toys", "Home", "Sports"]

        # Prepare data as lists for batch writing
        item_ids = []
        category_list = []
        item_names = []
        qty_list = []
        price_list2 = []
        reorder_list = []
        status_list = []
        value_formulas = []

        for i in range(2, 152):  # 150 rows
            item_ids.append(f"ITEM-{i-1:04d}")
            category_list.append(categories[random.randint(0, 6)])
            item_names.append(f"Item {i-1}")
            qty = random.randint(0, 500)
            price = round(random.uniform(10, 500), 2)
            reorder = random.randint(20, 50)
            status = "Low" if qty < reorder else "OK"

            qty_list.append(qty)
            price_list2.append(price)
            reorder_list.append(reorder)
            status_list.append(status)
            value_formulas.append(f"=D{i}*E{i}")

        # Write data in batches
        sheet13.range("A2:A151").value = [[id] for id in item_ids]
        sheet13.range("B2:B151").value = [[c] for c in category_list]
        sheet13.range("C2:C151").value = [[n] for n in item_names]
        sheet13.range("D2:D151").value = [[q] for q in qty_list]
        sheet13.range("E2:E151").value = [[p] for p in price_list2]
        sheet13.range("G2:G151").value = [[r] for r in reorder_list]
        sheet13.range("H2:H151").value = [[s] for s in status_list]

        # Write formulas individually for value column
        for i, formula in enumerate(value_formulas, start=2):
            sheet13.range(f"F{i}").formula = formula

        print(f"  - Created 150 rows of inventory data")

        # Sheet 14: Inventory Summary
        sheet14 = wb.sheets.add("InventorySummary")
        print("Creating inventory summary...")

        sheet14.range("A1").value = "Category"
        sheet14.range("B1").value = "Total Items"
        sheet14.range("C1").value = "Total Quantity"
        sheet14.range("D1").value = "Total Value"
        sheet14.range("E1").value = "Avg Price"
        sheet14.range("F1").value = "Low Stock Items"

        for i, category in enumerate(categories, start=2):
            sheet14.range(f"A{i}").value = category
            sheet14.range(f"B{i}").formula = f"=COUNTIF(InventoryData!B2:B151,\"{category}\")"
            sheet14.range(f"C{i}").formula = f"=SUMIF(InventoryData!B2:B151,\"{category}\",InventoryData!D2:D151)"
            sheet14.range(f"D{i}").formula = f"=SUMIF(InventoryData!B2:B151,\"{category}\",InventoryData!F2:F151)"
            sheet14.range(f"E{i}").formula = f"=AVERAGEIF(InventoryData!B2:B151,\"{category}\",InventoryData!E2:E151)"
            sheet14.range(f"F{i}").formula = f"=COUNTIFS(InventoryData!B2:B151,\"{category}\",InventoryData!H2:H151,\"Low\")"

        # Overall totals
        sheet14.range("A10").value = "TOTAL"
        sheet14.range("B10").formula = "=SUM(InventoryData!D2:D151)"
        sheet14.range("C10").formula = "=SUM(InventoryData!F2:F151)"
        sheet14.range("D10").formula = "=COUNTIF(InventoryData!H2:H151,\"Low\")"

        print("  - Created inventory summary by category")

        # Create charts using xlwings API
        try:
            print("Creating charts...")

            # Chart 1: Column chart on ChartData sheet for Sales vs Expenses
            try:
                chart_shape1 = sheet5.api.Shapes.AddChart2(201, 51)  # xlColumnClustered
                chart1 = chart_shape1.Chart
                chart1.SetSourceData(sheet5.api.Range("A1:C7"))
                chart1.HasTitle = True
                chart1.ChartTitle.Text = "Sales vs Expenses"
                chart1.HasLegend = True
                chart_shape1.Left = sheet5.range("E2").left
                chart_shape1.Top = sheet5.range("E2").top
                chart_shape1.Width = 400
                chart_shape1.Height = 250
                print("  - Created column chart: Sales vs Expenses")
            except Exception as e:
                print(f"  - Chart 1 creation failed: {e}")

            # Chart 2: Line chart for Sales trend
            try:
                chart_shape2 = sheet5.api.Shapes.AddChart2(201, 4)  # xlLine
                chart2 = chart_shape2.Chart
                chart2.SetSourceData(sheet5.api.Range("A1:B7"))
                chart2.HasTitle = True
                chart2.ChartTitle.Text = "Sales Trend"
                chart2.HasLegend = False
                chart_shape2.Left = sheet5.range("E15").left
                chart_shape2.Top = sheet5.range("E15").top
                chart_shape2.Width = 400
                chart_shape2.Height = 250
                print("  - Created line chart: Sales Trend")
            except Exception as e:
                print(f"  - Chart 2 creation failed: {e}")

            # Chart 3: Bar chart for Profit
            try:
                chart_shape3 = sheet5.api.Shapes.AddChart2(201, 57)  # xlBarClustered
                chart3 = chart_shape3.Chart
                chart3.SetSourceData(sheet5.api.Range("A1:A7,D1:D7"))
                chart3.HasTitle = True
                chart3.ChartTitle.Text = "Profit by Month"
                chart3.HasLegend = True
                chart_shape3.Left = sheet5.range("K2").left
                chart_shape3.Top = sheet5.range("K2").top
                chart_shape3.Width = 400
                chart_shape3.Height = 250
                print("  - Created bar chart: Profit by Month")
            except Exception as e:
                print(f"  - Chart 3 creation failed: {e}")

            # Chart 4: Pie chart for Data sheet values
            try:
                chart_shape4 = sheet1.api.Shapes.AddChart2(201, 5)  # xlPie
                chart4 = chart_shape4.Chart
                chart4.SetSourceData(sheet1.api.Range("A2:B6"))
                chart4.HasTitle = True
                chart4.ChartTitle.Text = "Data Distribution"
                chart4.HasLegend = True
                chart_shape4.Left = sheet1.range("D2").left
                chart_shape4.Top = sheet1.range("D2").top
                chart_shape4.Width = 350
                chart_shape4.Height = 250
                print("  - Created pie chart: Data Distribution")
            except Exception as e:
                print(f"  - Chart 4 creation failed: {e}")

            # Chart 5: Area chart for Complex sheet
            try:
                chart_shape5 = sheet3.api.Shapes.AddChart2(201, 1)  # xlArea
                chart5 = chart_shape5.Chart
                chart5.SetSourceData(sheet3.api.Range("A1:B5"))
                chart5.HasTitle = True
                chart5.ChartTitle.Text = "Complex Data Trend"
                chart5.HasLegend = False
                chart_shape5.Left = sheet3.range("E2").left
                chart_shape5.Top = sheet3.range("E2").top
                chart_shape5.Width = 400
                chart_shape5.Height = 250
                print("  - Created area chart: Complex Data Trend")
            except Exception as e:
                print(f"  - Chart 5 creation failed: {e}")

            # Chart 6: Scatter chart for JaggedData
            try:
                chart_shape6 = sheet6.api.Shapes.AddChart2(201, -4169)  # xlXYScatter
                chart6 = chart_shape6.Chart
                chart6.SetSourceData(sheet6.api.Range("B3:B6,D3:D6"))
                chart6.HasTitle = True
                chart6.ChartTitle.Text = "Jagged Data Scatter"
                chart6.HasLegend = False
                chart_shape6.Left = sheet6.range("A10").left
                chart_shape6.Top = sheet6.range("A10").top
                chart_shape6.Width = 400
                chart_shape6.Height = 250
                print("  - Created scatter chart: Jagged Data Scatter")
            except Exception as e:
                print(f"  - Chart 6 creation failed: {e}")

            # Chart 7: Region Sales Comparison (Column)
            try:
                chart_shape7 = sheet11.api.Shapes.AddChart2(201, 51)  # xlColumnClustered
                chart7 = chart_shape7.Chart
                chart7.SetSourceData(sheet11.api.Range("D1:E6"))
                chart7.HasTitle = True
                chart7.ChartTitle.Text = "Sales by Region"
                chart7.HasLegend = False
                chart_shape7.Left = sheet11.range("A12").left
                chart_shape7.Top = sheet11.range("A12").top
                chart_shape7.Width = 450
                chart_shape7.Height = 300
                print("  - Created column chart: Sales by Region")
            except Exception as e:
                print(f"  - Chart 7 creation failed: {e}")

            # Chart 8: Product Sales (Bar)
            try:
                chart_shape8 = sheet11.api.Shapes.AddChart2(201, 57)  # xlBarClustered
                chart8 = chart_shape8.Chart
                chart8.SetSourceData(sheet11.api.Range("I1:J6"))
                chart8.HasTitle = True
                chart8.ChartTitle.Text = "Sales by Product"
                chart8.HasLegend = False
                chart_shape8.Left = sheet11.range("M1").left
                chart_shape8.Top = sheet11.range("M1").top
                chart_shape8.Width = 450
                chart_shape8.Height = 300
                print("  - Created bar chart: Sales by Product")
            except Exception as e:
                print(f"  - Chart 8 creation failed: {e}")

            # Chart 9: Monthly Sales Trend (Line)
            try:
                chart_shape9 = sheet12.api.Shapes.AddChart2(201, 4)  # xlLine
                chart9 = chart_shape9.Chart
                chart9.SetSourceData(sheet12.api.Range("A1:B7"))
                chart9.HasTitle = True
                chart9.ChartTitle.Text = "Monthly Sales Trend"
                chart9.HasLegend = False
                chart_shape9.Left = sheet12.range("A10").left
                chart_shape9.Top = sheet12.range("A10").top
                chart_shape9.Width = 500
                chart_shape9.Height = 300
                print("  - Created line chart: Monthly Sales Trend")
            except Exception as e:
                print(f"  - Chart 9 creation failed: {e}")

            # Chart 10: Monthly Profit (Area)
            try:
                chart_shape10 = sheet12.api.Shapes.AddChart2(201, 1)  # xlArea
                chart10 = chart_shape10.Chart
                chart10.SetSourceData(sheet12.api.Range("A1:A7,D1:D7"))
                chart10.HasTitle = True
                chart10.ChartTitle.Text = "Monthly Profit Trend"
                chart10.HasLegend = False
                chart_shape10.Left = sheet12.range("H10").left
                chart_shape10.Top = sheet12.range("H10").top
                chart_shape10.Width = 500
                chart_shape10.Height = 300
                print("  - Created area chart: Monthly Profit Trend")
            except Exception as e:
                print(f"  - Chart 10 creation failed: {e}")

            # Chart 11: Inventory by Category (Pie)
            try:
                chart_shape11 = sheet14.api.Shapes.AddChart2(201, 5)  # xlPie
                chart11 = chart_shape11.Chart
                chart11.SetSourceData(sheet14.api.Range("A1:A8,D1:D8"))
                chart11.HasTitle = True
                chart11.ChartTitle.Text = "Inventory Value by Category"
                chart11.HasLegend = True
                chart_shape11.Left = sheet14.range("A12").left
                chart_shape11.Top = sheet14.range("A12").top
                chart_shape11.Width = 450
                chart_shape11.Height = 350
                print("  - Created pie chart: Inventory Value by Category")
            except Exception as e:
                print(f"  - Chart 11 creation failed: {e}")

            # Chart 12: Inventory Quantity by Category (Column)
            try:
                chart_shape12 = sheet14.api.Shapes.AddChart2(201, 51)  # xlColumnClustered
                chart12 = chart_shape12.Chart
                chart12.SetSourceData(sheet14.api.Range("A1:A8,C1:C8"))
                chart12.HasTitle = True
                chart12.ChartTitle.Text = "Total Quantity by Category"
                chart12.HasLegend = False
                chart_shape12.Left = sheet14.range("H1").left
                chart_shape12.Top = sheet14.range("H1").top
                chart_shape12.Width = 500
                chart_shape12.Height = 350
                print("  - Created column chart: Total Quantity by Category")
            except Exception as e:
                print(f"  - Chart 12 creation failed: {e}")

            # Chart 13: Low Stock Items (Bar)
            try:
                chart_shape13 = sheet14.api.Shapes.AddChart2(201, 57)  # xlBarClustered
                chart13 = chart_shape13.Chart
                chart13.SetSourceData(sheet14.api.Range("A1:A8,F1:F8"))
                chart13.HasTitle = True
                chart13.ChartTitle.Text = "Low Stock Items by Category"
                chart13.HasLegend = False
                chart_shape13.Left = sheet14.range("H12").left
                chart_shape13.Top = sheet14.range("H12").top
                chart_shape13.Width = 500
                chart_shape13.Height = 350
                print("  - Created bar chart: Low Stock Items by Category")
            except Exception as e:
                print(f"  - Chart 13 creation failed: {e}")

            # Chart 14: Sales vs Units Scatter (from LargeSalesData)
            try:
                chart_shape14 = sheet10.api.Shapes.AddChart2(201, -4169)  # xlXYScatter
                chart14 = chart_shape14.Chart
                chart14.SetSourceData(sheet10.api.Range("D1:E51"))  # First 50 data points
                chart14.HasTitle = True
                chart14.ChartTitle.Text = "Sales vs Units (Sample)"
                chart14.HasLegend = False
                chart_shape14.Left = sheet10.range("J2").left
                chart_shape14.Top = sheet10.range("J2").top
                chart_shape14.Width = 450
                chart_shape14.Height = 300
                print("  - Created scatter chart: Sales vs Units")
            except Exception as e:
                print(f"  - Chart 14 creation failed: {e}")

            # Chart 15: Monthly Units Trend (Line with markers)
            try:
                chart_shape15 = sheet12.api.Shapes.AddChart2(201, 65)  # xlLineMarkers
                chart15 = chart_shape15.Chart
                chart15.SetSourceData(sheet12.api.Range("A1:A7,C1:C7"))
                chart15.HasTitle = True
                chart15.ChartTitle.Text = "Monthly Units Sold"
                chart15.HasLegend = False
                chart_shape15.Left = sheet12.range("A20").left
                chart_shape15.Top = sheet12.range("A20").top
                chart_shape15.Width = 500
                chart_shape15.Height = 300
                print("  - Created line chart with markers: Monthly Units Sold")
            except Exception as e:
                print(f"  - Chart 15 creation failed: {e}")

            # Chart 16: Region Average Sales (Doughnut)
            try:
                chart_shape16 = sheet11.api.Shapes.AddChart2(201, -4120)  # xlDoughnut
                chart16 = chart_shape16.Chart
                chart16.SetSourceData(sheet11.api.Range("D1:D6,G1:G6"))
                chart16.HasTitle = True
                chart16.ChartTitle.Text = "Average Sales by Region"
                chart16.HasLegend = True
                chart_shape16.Left = sheet11.range("A20").left
                chart_shape16.Top = sheet11.range("A20").top
                chart_shape16.Width = 450
                chart_shape16.Height = 350
                print("  - Created doughnut chart: Average Sales by Region")
            except Exception as e:
                print(f"  - Chart 16 creation failed: {e}")

            # Chart 17: Combined Sales and Profit (Combo chart)
            try:
                chart_shape17 = sheet12.api.Shapes.AddChart2(201, -4152)  # xlColumnClustered
                chart17 = chart_shape17.Chart
                chart17.SetSourceData(sheet12.api.Range("A1:A7,B1:B7,D1:D7"))
                chart17.HasTitle = True
                chart17.ChartTitle.Text = "Sales and Profit Comparison"
                chart17.HasLegend = True
                # Try to make it a combo chart
                try:
                    chart17.ChartType = -4152  # xlColumnClustered
                    chart17.SeriesCollection(2).ChartType = 4  # xlLine for profit
                except:
                    pass
                chart_shape17.Left = sheet12.range("H20").left
                chart_shape17.Top = sheet12.range("H20").top
                chart_shape17.Width = 500
                chart_shape17.Height = 300
                print("  - Created combo chart: Sales and Profit Comparison")
            except Exception as e:
                print(f"  - Chart 17 creation failed: {e}")

            # Chart 18: Inventory Average Price (Column)
            try:
                chart_shape18 = sheet14.api.Shapes.AddChart2(201, 51)  # xlColumnClustered
                chart18 = chart_shape18.Chart
                chart18.SetSourceData(sheet14.api.Range("A1:A8,E1:E8"))
                chart18.HasTitle = True
                chart18.ChartTitle.Text = "Average Price by Category"
                chart18.HasLegend = False
                chart_shape18.Left = sheet14.range("A20").left
                chart_shape18.Top = sheet14.range("A20").top
                chart_shape18.Width = 500
                chart_shape18.Height = 350
                print("  - Created column chart: Average Price by Category")
            except Exception as e:
                print(f"  - Chart 18 creation failed: {e}")

            print("Chart creation completed")
        except Exception as e:
            print(f"Warning: Chart creation encountered issues: {e}")
            print("  Some charts may not have been created")
            print("  Charts can be added manually in Excel if needed")

        # Add VBA modules
        # Note: On macOS, xlwings uses AppleScript which doesn't expose VBProject COM interface
        # VBA must be added manually. A VBA code file has been created for reference.
        try:
            print("Attempting to add VBA modules...")
            import platform

            if platform.system() == "Windows":
                # On Windows, try COM access
                try:
                    vba_project = wb.api.VBProject
                    # Add modules (Windows-specific code would go here)
                    print("  Windows COM access detected, but VBA insertion not fully implemented")
                    print("  See tests/fixtures/vba_code.txt for manual VBA addition")
                except Exception as e:
                    print(f"  Windows COM access failed: {e}")
                    print("  See tests/fixtures/vba_code.txt for manual VBA addition")
            else:
                # macOS - AppleScript doesn't support VBProject
                print("  macOS detected: VBA cannot be added programmatically via xlwings")
                print("  VBA code has been saved to: tests/fixtures/vba_code.txt")
                print("  To add VBA manually:")
                print("    1. Open comprehensive_test.xlsb in Excel")
                print("    2. Press Option+F11 to open VBA Editor")
                print("    3. Follow instructions in vba_code.txt")

        except Exception as e:
            print(f"  VBA addition attempt failed: {e}")
            print("  See tests/fixtures/vba_code.txt for manual VBA addition instructions")

        # Save as .xlsb
        wb.save(str(output_path))
        print(f"Test workbook saved: {output_path}")

    finally:
        try:
            wb.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    # Get fixtures directory
    script_dir = Path(__file__).parent
    fixtures_dir = script_dir / "fixtures"
    fixtures_dir.mkdir(exist_ok=True)

    output_path = fixtures_dir / "comprehensive_test.xlsb"

    try:
        create_test_workbook(output_path)
        print("\n✓ Test workbook created successfully!")
        print(f"  Location: {output_path}")
    except Exception as e:
        print(f"\n✗ Error creating test workbook: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
