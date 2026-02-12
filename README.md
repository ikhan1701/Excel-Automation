# Excel-Automation
Automating mundane excel tasks using Python.
import xlwings as xw
import os
import time     
import pandas as pd
import re
import pdfplumber
from datetime import datetime
import sys
import os

def extractPDF(folder_path):
    output_file = os.path.join(folder_path, "extractedData.xlsx")

    VALID_CODES = {"100", "200", "300", "400", "500", "700"}
    CODE_TO_SHEET = {
        "100": "100_200_500", "200": "100_200_500", "500": "100_200_500",
        "300": "300_400_700", "400": "300_400_700", "700": "300_400_700",
    }
    HEADER_PREFIX = re.compile(
        r'\*{4}\s+.*?(?:RECEIVED|DELIVERED)\s+BY\s+YOU\s*\*{4}\s*',
        re.IGNORECASE
    )
    STAR_STRIP = re.compile(r'^\*+\s*|\s*\*+$')

    SUMMARY_NOISE = re.compile(
        r'SUB\s+TOTAL|GRAND\s+BANK|BANK\s+SUMMARY|COMPOSITE'
        r'|^TOTAL\s*='
        r'|^TOTAL\s+(RECEIVED|DELIVERED)'
        r'|^NET\s*'
        r'|^CITY\s+NET'
        r'|^\(?\d[\d,]*\.?\d*\)?$'          
        r'|^GROUP\s+[AB]\s*\(',            
        re.IGNORECASE )

    PATTERN = re.compile(
        r'(?:'
            r'(?P<code_a>\d{4})\s+(?P<name_a>.+?)\s+GROUP\s+[AB]\s*=\s*(?P<items_a>\d+)\s+(?P<amount_a>[\d,]+\.\d{2})'
        r'|'
            r'(?P<name_c>\D.*?)(?P<code_c>\d{4})\s+(?P<amount_c>[\d,]+\.\d{2})\s+(?P<items_c>\d+)\s+GROUP\s+[AB]\s*='
        r'|'
            r'(?P<name_b>\D.*?)(?P<code_b>\d{4})\s+(?P<items_b>\d+)\s+(?P<amount_b>[\d,]+\.\d{2})\s+GROUP\s+[AB]\s*='
        r')')

    extracted_rows = []
    warnings       = []

    for f in sorted(os.listdir(folder_path)):
        if not (f.upper().endswith(".PDF") and f.upper().startswith("BKTS107")):
            continue

        match = re.search(r"BKTS107(\d{3})", f.upper())
        if not match:
            warnings.append(f"[SKIP] Could not parse code from filename: {f}")
            continue

        file_code = match.group(1)
        if file_code not in VALID_CODES:
            warnings.append(f"[SKIP] Unrecognised code '{file_code}' in: {f}")
            continue

        sheet_name = CODE_TO_SHEET[file_code]
        pdf_file   = os.path.join(folder_path, f)

        current_clearing_type = None
        rows_in_file          = 0

        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text(x_tolerance=5, y_tolerance=3)
                if not text:
                    continue

                text = re.sub(r'(?<!\n)(\*{4})', r'\n\1', text)

                for line in text.replace('\xa0', ' ').split('\n'):
                    line       = line.strip()
                    line_upper = line.upper()

                    line_clearing_type = None
                    if "RECEIVED BY YOU" in line_upper:
                        line_clearing_type = "Received by you"
                        current_clearing_type = "Received by you"
                    elif "DELIVERED BY YOU" in line_upper:
                        line_clearing_type = "Delivered by you"
                        current_clearing_type = "Delivered by you"

                    cleaned = HEADER_PREFIX.sub('', line).strip()
                    cleaned = STAR_STRIP.sub('',  cleaned).strip()

                    matched_any = False
                    for m in PATTERN.finditer(cleaned):
                        matched_any = True

                        clearing_to_use = line_clearing_type if line_clearing_type else current_clearing_type
                        
                        if clearing_to_use is None:
                            warnings.append(f"[WARN] No clearing type before data in: {f} p{page_num}")
                            clearing_to_use = "Unknown"

                        g = m.groupdict()
                        code   = g['code_a']   or g['code_b']   or g['code_c']
                        name   = (g['name_a']  or g['name_b']   or g['name_c']  or '').strip()
                        items  = g['items_a']  or g['items_b']  or g['items_c']
                        amount = g['amount_a'] or g['amount_b'] or g['amount_c']

                        extracted_rows.append({
                            "NIFT Code":     code,
                            "Br Name":       name,
                            "Clearing Type": clearing_to_use,
                            "Items":         int(items),
                            "Amount":        float(amount.replace(",", "")),
                            "Sheet":         sheet_name,
                            "Source File":   f,
                        })
                        rows_in_file += 1

                    if not matched_any and "GROUP" in line_upper and not SUMMARY_NOISE.search(cleaned):
                        warnings.append(f"[UNMATCHED GROUP] {f} p{page_num} → {cleaned}")

        if rows_in_file == 0:
            warnings.append(f"[WARN] Zero rows extracted from: {f}")

    # ── DEDUPLICATE & WRITE ─────────────────────────────────────────────
    df = pd.DataFrame(extracted_rows)

    if df.empty:
        print("No data extracted. Check folder path and PDF format.")
    else:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            for sheet, data in df.groupby("Sheet"):
                # Drop Source File and Sheet columns before writing to Excel
                data_output = data.drop(columns=['Source File', 'Sheet'])
                data_output.to_excel(writer, sheet_name=sheet, index=False, float_format="%.2f")

        print(f"Extraction complete — {len(df)} rows written to:\n{output_file}")

def createBulkSheet(folder_path):
    inward_csv = os.path.join(folder_path, "INWARD.CSV")
    outward_csv = os.path.join(folder_path, "OUTWARD.CSV")

    df_inward = pd.read_csv(inward_csv, header=None, names = ["","Debit Account (Br GL)","Amount","Credit account" ])
    df_inward.insert(loc=1, column= "Branch Code", value=df_inward["Debit Account (Br GL)"].astype(str).str[-4:])


    df_outward = pd.read_csv(outward_csv, header=None, names = ["","Credit Account (Br GL)","Amount", "Debit Account (State Bank)"])
    df_outward.insert(loc=1, column= "Branch Code", value=df_outward["Debit Account (State Bank)"].astype(str).str[-4:])

    os.makedirs(folder_path, exist_ok=True)

    output_file = os.path.join(folder_path, f"Bulk Downloaded Sheets.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_inward.to_excel(writer, sheet_name="Inward", index=False, header=True)
        df_outward.to_excel(writer, sheet_name="Outward", index=False, header=True)

    print("Bulk downloaded sheets created successfully.")
    time.sleep(1)

def OP_Balance_(folder_path):
    clearing_file = os.path.join(folder_path, r'Inward & Outward Clearing GL.xlsx')

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(clearing_file)
        ws = wb.sheets["Outward"]
            # Find last row in column B (starting from row 3)
        last_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row

        # Read values from column H (row 3 to last row)
        values = ws.range(f"H3:H{last_row}").value

        # Write values to column D (row 3 to last row)
        ws.range(f"D3:D{last_row}").options(transpose=True).value = values

        # Loop through both sheets
        for sheet_name in ["Outward", "Inward"]:
            sheet = wb.sheets[sheet_name]

            # Find last row with data in column B
            last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row

            # Apply formula row by row
            for row in range(3, last_row + 1):  
                formula = f"=E{row}+F{row}+G{row}"
                sheet.range(f"H{row}").formula = formula

        wb.save()

    finally:
        wb.close()
        app.quit()

def clearingBulk(folder_path):
    main_file = os.path.join(folder_path, "Inward & Outward Clearing GL.xlsx")
    csv_file = os.path.join(folder_path, "INWARD.CSV")

    # Open main workbook
    wb_main = xw.Book(main_file, update_links=False)
    sh_main = wb_main.sheets['Inward']

    # Open CSV workbook
    csv_wb = xw.Book(csv_file)
    csv_sh = csv_wb.sheets[0]   # usually "INWARD" or "Sheet1"

    # Get names for formula reference
    csv_name = os.path.basename(csv_file)   # "INWARD.CSV"
    csv_sheet_name = csv_sh.name            # actual sheet name

    
    # Find last row in column B
    last_row = sh_main.range('B' + str(sh_main.cells.last_cell.row)).end('up').row
    

    # Build formula string
    formula = (
        f"=IFERROR("
        f"VLOOKUP($E$2 & B3,'[{csv_name}]{csv_sheet_name}'!$B:$C,2,FALSE),0)"
    )

    # Apply formula to range J3:J{last_row}
    target_range = sh_main.range(f"J3:J{last_row}")
    target_range.formula = formula

    # Save and close
    wb_main.save()
    wb_main.close()
    csv_wb.close()
    print("Clearing Bulk - Column J populated successfully!")

def clearingBulkD(folder_path):
    main_file = os.path.join(folder_path, "Inward & Outward Clearing GL.xlsx")
    csv_file = os.path.join(folder_path, "OUTWARD.CSV")

    # Open main workbook
    wb_main = xw.Book(main_file, update_links=False)
    sh_main = wb_main.sheets['Outward']

    # Open CSV workbook
    csv_wb = xw.Book(csv_file)
    csv_sh = csv_wb.sheets[0]   # usually "OUTWARD" or "Sheet1"

    # Get names for formula reference
    csv_name = os.path.basename(csv_file)   # "OUTWARD.CSV"
    csv_sheet_name = csv_sh.name            # actual sheet name

    
    # Find last row in column B
    last_row = sh_main.range('B' + str(sh_main.cells.last_cell.row)).end('up').row
    

    # Build formula string
    formula = (
        f"=IFERROR("
        f"VLOOKUP($E$2 & B3,'[{csv_name}]{csv_sheet_name}'!$B:$C,2,FALSE),0)"
    )

    # Apply formula to range K3:K{last_row}
    target_range = sh_main.range(f"K3:K{last_row}")
    target_range.formula = formula

    # Save and close
    wb_main.save()
    wb_main.close()
    csv_wb.close()
    print("Clearing Bulk - Column K populated successfully!")

def combineEMACC(folder_path):
    output_file = os.path.join(folder_path, f"GL_File.xlsx")
# Find all CSV files that start with 'EM_ACCT_BAL_TODAY'
    matching_files = []
    for file in os.listdir(folder_path):
        if file.startswith('EM_ACCT_BAL_TODAY') and file.upper().endswith('.CSV'):
            matching_files.append(file)

    all_data = []

    for file in sorted(matching_files):
        file_path = os.path.join(folder_path, file)
        
        try:
            # Read CSV, skipping first 2 rows, using row 3 as header (skiprows=[0,1])
            df = pd.read_csv(file_path, skiprows=[0, 1])
            all_data.append(df)
            
        except Exception as e:
            print(f"    ✗ Error reading {file}: {e}")
            continue

    if len(all_data) == 0:
        exit()

    combined_df = pd.concat(all_data, ignore_index=True)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, sheet_name="EM_ACCT_BAL_DATA", index=False)

    print(f"  ✓ Saved to: {output_file}")

def normalInward(folder_path):

    clearing_file = os.path.join(folder_path, f"Inward & Outward Clearing GL.xlsx")
    gl_file = os.path.join(folder_path, f"GL_File.xlsx")

    df = pd.read_excel(gl_file)

    # Columns to convert
    cols_to_float = ['Ledger Balance', 'Cleared Balance', 'Working Balance']

    for col in cols_to_float:
        # Remove commas, single quotes, spaces, then convert to float
        df[col] = (
            df[col]
            .astype(str)           # ensure everything is a string
            .str.replace(",", "", regex=False)  # remove commas
            .str.replace("'", "", regex=False)  # remove single quotes
            .str.strip()           # remove leading/trailing spaces
        )
        # Convert to float, set invalid entries to NaN
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df.to_excel(gl_file, index=False)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(clearing_file)
        sheet = wb.sheets["Inward"]

        wb2 = app.books.open(gl_file)
        gl_sheet = wb2.sheets[0]
        gl_sheet_name = gl_sheet.name
        
        # Get just the workbook name for Excel external reference
        gl_workbook_name = os.path.basename(gl_file)

        last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
        last_row_gl = gl_sheet.range("A" + str(gl_sheet.cells.last_cell.row)).end("up").row

        # Write Normal Inward formula once in E3
        formula = (
            f'=IFERROR(VLOOKUP($E$2&TEXT(B3,"0000"),'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)"
        )
        sheet.range("E3").formula = formula
        sheet.range("E3").autofill(sheet.range(f"E3:E{last_row}"))

        #sameday inward formula
        sd_formula =  (
            f'=IFERROR(VLOOKUP($F$2&TEXT(B3,"0000"),'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)" )
        sheet.range("F3").formula = sd_formula
        sheet.range("F3").autofill(sheet.range(f"F3:F{last_row}"))

        #intercity inward formula
        ic_formula =  (
            f'=IFERROR(VLOOKUP($G$2&TEXT(B3,"0000"),'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)" )

        sheet.range("G3").formula = ic_formula
        sheet.range("G3").autofill(sheet.range(f"G3:G{last_row}"))
        
        wb.save()
        time.sleep(1)  # Ensure file is saved before closing  


    finally:
        wb.close()
        wb2.close()
        app.quit()

def normalOutward(folder_path):
    clearing_file = os.path.join(folder_path, f"Inward & Outward Clearing GL.xlsx")
    gl_file = os.path.join(folder_path, f"GL_File.xlsx")   


    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(clearing_file)
        sheet = wb.sheets["Outward"]

        wb2 = app.books.open(gl_file)
        gl_sheet = wb2.sheets[0]
        gl_sheet_name = gl_sheet.name
        
        # Get just the workbook name for Excel external reference
        gl_workbook_name = os.path.basename(gl_file)

        last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
        last_row_gl = gl_sheet.range("A" + str(gl_sheet.cells.last_cell.row)).end("up").row

        # Write Normal Outward formula once in E3
        formula = (
            f'=IFERROR(VLOOKUP($E$2&B4,'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)"
        )
        sheet.range("E3").formula = formula
        sheet.range("E3").autofill(sheet.range(f"E3:E{last_row}"))

        #sameday outward formula
        sd_formula =  (
            f'=IFERROR(VLOOKUP($F$2&B4,'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)" )
        sheet.range("F3").formula = sd_formula
        sheet.range("F3").autofill(sheet.range(f"F3:F{last_row}"))

        #intercity outward formula
        ic_formula =  (
            f'=IFERROR(VLOOKUP($G$2&B4,'
            f"'[{gl_workbook_name}]{gl_sheet_name}'!$A$2:$K${last_row_gl},8,FALSE),0)" )

        sheet.range("G3").formula = ic_formula
        sheet.range("G3").autofill(sheet.range(f"G3:G{last_row}"))

        wb.save()
        
    finally:
        wb.close()
        wb2.close()
        app.quit()
        print("Normal outward done")

def create_pivot_tables(folder_path):
    extracted_data_path = os.path.join(folder_path, f"ExtractedData.xlsx")
    app = None

    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        extracted_data_wb = app.books.open(extracted_data_path, update_links=False)

        data_sheet_100 = extracted_data_wb.sheets['100_200_500']
        data_sheet_300 = extracted_data_wb.sheets['300_400_700']

        # --- Delete pivot sheets if they already exist ---
        for sheet_name in ['pivot_100_200_500', 'pivot_300_400_700']:
            if sheet_name in [s.name for s in extracted_data_wb.sheets]:
                extracted_data_wb.sheets[sheet_name].delete()

        pivot_sheet_100 = extracted_data_wb.sheets.add('pivot_100_200_500')
        pivot_sheet_300 = extracted_data_wb.sheets.add('pivot_300_400_700')

        source_range_100 = data_sheet_100.used_range
        source_range_300 = data_sheet_300.used_range

        # ================= Pivot 100 =================
        pivot100_cache = extracted_data_wb.api.PivotCaches().Create(
            SourceType=1,
            SourceData=source_range_100.api)

        pivot100_table = pivot100_cache.CreatePivotTable(
            TableDestination=pivot_sheet_100.range("A3").api,
            TableName="PivotTable100")

        pivot100_table.PivotFields("NIFT Code").Orientation = 1
        pivot100_table.PivotFields("NIFT Code").Position = 1

        value_field_100 = pivot100_table.PivotFields("Amount")
        value_field_100.Orientation = 4
        value_field_100.Function = -4157
        value_field_100.Name = "Sum of Amount"

        pivot100_table.PivotFields("Clearing Type").Orientation = 3
        pivot100_table.RefreshTable()

        # ================= Pivot 300 =================
        pivot300_cache = extracted_data_wb.api.PivotCaches().Create(
            SourceType=1,
            SourceData=source_range_300.api)

        pivot300_table = pivot300_cache.CreatePivotTable(
            TableDestination=pivot_sheet_300.range("A3").api,
            TableName="PivotTable300")

        pivot300_table.PivotFields("NIFT Code").Orientation = 1
        pivot300_table.PivotFields("NIFT Code").Position = 1

        value_field_300 = pivot300_table.PivotFields("Amount")
        value_field_300.Orientation = 4
        value_field_300.Function = -4157
        value_field_300.Name = "Sum of Amount"

        pivot300_table.PivotFields("Clearing Type").Orientation = 3
        pivot300_table.RefreshTable()

        extracted_data_wb.save()

    finally:
        if app:
            app.quit()
        print("Pivot tables created")

def nift_summary(folder_path):
    extracted_data_path = os.path.join(folder_path, r"extractedData.xlsx" ) 
    inward_wb_path = os.path.join(folder_path, r"Inward & Outward Clearing GL.xlsx")
        
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        extracted_wb = app.books.open(extracted_data_path, update_links=False)

        # --- Pivot sheets ---
        pivot_sheet_100 = extracted_wb.sheets['pivot_100_200_500']
        pivot_sheet_300 = extracted_wb.sheets['pivot_300_400_700']

        # --- Apply Clearing Type filters ---
        pivot_sheet_100.api.PivotTables(1) \
            .PivotFields("Clearing Type").CurrentPage = "Received by you"

        pivot_sheet_300.api.PivotTables(1) \
            .PivotFields("Clearing Type").CurrentPage = "Delivered by you"

        # Force refresh 
        pivot_sheet_100.api.PivotTables(1).RefreshTable()
        pivot_sheet_300.api.PivotTables(1).RefreshTable()

        # READ pivot_100_200_500 → DataFrame
        pt_100 = pivot_sheet_100.api.PivotTables(1)
        rng_100 = pivot_sheet_100.range(pt_100.TableRange1.Address)
        values_100 = rng_100.value

        df_received = pd.DataFrame(values_100[1:], columns=values_100[0])
        df_received.columns = ["nift_code", "amount"]

        df_received["nift_code"] = df_received["nift_code"].astype(str).str.strip()
        df_received["amount"] = pd.to_numeric(df_received["amount"], errors="coerce")

        df_received = df_received.dropna()
        df_received = df_received[df_received["amount"] != 0]
        df_received = df_received[df_received["nift_code"].str.lower() != "grand total"]

        # READ pivot_300_400_700 → DataFrame
        pt_300 = pivot_sheet_300.api.PivotTables(1)
        rng_300 = pivot_sheet_300.range(pt_300.TableRange1.Address)
        values_300 = rng_300.value

        df_delivered = pd.DataFrame(values_300[1:], columns=values_300[0])
        df_delivered.columns = ["nift_code", "amount"]

        df_delivered["nift_code"] = df_delivered["nift_code"].astype(str).str.strip()
        df_delivered["amount"] = pd.to_numeric(df_delivered["amount"], errors="coerce")

        df_delivered = df_delivered.dropna()
        df_delivered = df_delivered[df_delivered["amount"] != 0]
        df_delivered = df_delivered[df_delivered["nift_code"].str.lower() != "grand total"]

        # RECONCILIATION 
        df_recon = df_received.merge(
            df_delivered,
            on="nift_code",
            how="inner",
            suffixes=("_received", "_delivered")
        )

        df_recon["difference"] = (
            df_recon["amount_received"] - df_recon["amount_delivered"]
        )

        # Preserve leading zeros 
        df_recon["nift_code"] = df_recon["nift_code"].str.zfill(4)

        # --- Create lookup dictionary ---
        lookup = dict(zip(df_recon["nift_code"], df_recon["difference"]))

        # --- Open Outward sheet directly ---
        wb_inward = xw.Book(inward_wb_path)
        sheet_inward = wb_inward.sheets["Inward"]

        # --- Read NIFT codes from column A ---
        last_row = sheet_inward.cells.last_cell.row
        nift_codes = sheet_inward.range(f"A2:A{last_row}").value
        if isinstance(nift_codes, (str, int, float)):
            nift_codes = [nift_codes]

        # Convert all codes to 4-digit strings
        nift_codes = [
            str(code).strip().zfill(4) if code is not None and str(code).strip() != "" else None
            for code in nift_codes
        ]

        # --- Write differences safely in column I ---
        for i, code_str in enumerate(nift_codes, start=2):
            if code_str is None or code_str.lower() == "grand total":
                continue
            if code_str in lookup:
                sheet_inward.range(f"I{i}").value = lookup[code_str]

        wb_inward.save()
        wb_inward.close()


    finally:
        extracted_wb.close()
        app.quit()
        print("NIft summary inward complete")

def nift_summary_outward(folder_path):   
    pivot_wb_path = os.path.join(folder_path, r"extractedData.xlsx" ) 
    outward_wb_path = os.path.join(folder_path, r"Inward & Outward Clearing GL.xlsx")  

    try:
        # --- Launch Excel ---
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        # --- Open pivot workbook ---
        pivot_wb = app.books.open(pivot_wb_path, update_links=False)
        pivot_sheet_100 = pivot_wb.sheets['pivot_100_200_500']
        pivot_sheet_300 = pivot_wb.sheets['pivot_300_400_700']

        # --- Apply clearing type filters ---
        pivot_sheet_100.api.PivotTables(1).PivotFields("Clearing Type").CurrentPage = "Delivered by you"
        pivot_sheet_300.api.PivotTables(1).PivotFields("Clearing Type").CurrentPage = "Received by you"

        # --- Refresh pivots ---
        pivot_sheet_100.api.PivotTables(1).RefreshTable()
        pivot_sheet_300.api.PivotTables(1).RefreshTable()

        # --- Read pivot_100_200_500 (Delivered) ---
        pt_100 = pivot_sheet_100.api.PivotTables(1)
        values_100 = pivot_sheet_100.range(pt_100.TableRange1.Address).value
        df_delivered = pd.DataFrame(values_100[1:], columns=values_100[0])
        df_delivered.columns = ["nift_code", "amount"]
        df_delivered["nift_code"] = df_delivered["nift_code"].astype(str).str.strip().str.zfill(4)
        df_delivered["amount"] = pd.to_numeric(df_delivered["amount"], errors="coerce")
        df_delivered = df_delivered.dropna(subset=["nift_code", "amount"])
        df_delivered = df_delivered[df_delivered["amount"] != 0]
        df_delivered = df_delivered[df_delivered["nift_code"].str.lower() != "grand total"]

        # --- Read pivot_300_400_700 (Received) ---
        pt_300 = pivot_sheet_300.api.PivotTables(1)
        values_300 = pivot_sheet_300.range(pt_300.TableRange1.Address).value
        df_received = pd.DataFrame(values_300[1:], columns=values_300[0])
        df_received.columns = ["nift_code", "amount"]
        df_received["nift_code"] = df_received["nift_code"].astype(str).str.strip().str.zfill(4)
        df_received["amount"] = pd.to_numeric(df_received["amount"], errors="coerce")
        df_received = df_received.dropna(subset=["nift_code", "amount"])
        df_received = df_received[df_received["amount"] != 0]
        df_received = df_received[df_received["nift_code"].str.lower() != "grand total"]

        # --- Aggregate by NIFT code to handle duplicates ---
        df_delivered_agg = df_delivered.groupby("nift_code", as_index=False)["amount"].sum()
        df_received_agg = df_received.groupby("nift_code", as_index=False)["amount"].sum()

        # --- Reconciliation (only codes present in both) ---
        df_recon = pd.merge(
            df_received_agg,
            df_delivered_agg,
            on="nift_code",
            how="inner",
            suffixes=("_received", "_delivered")
        )
        df_recon["difference"] = df_recon["amount_delivered"] - df_recon["amount_received"]

        # --- Create lookup dictionary ---
        lookup = dict(zip(df_recon["nift_code"], df_recon["difference"]))

        # --- Open Outward sheet directly ---
        wb_outward = xw.Book(outward_wb_path)
        sheet_outward = wb_outward.sheets["Outward"]

        # --- Read NIFT codes from column A ---
        last_row = sheet_outward.cells.last_cell.row
        nift_codes = sheet_outward.range(f"A2:A{last_row}").value
        if isinstance(nift_codes, (str, int, float)):
            nift_codes = [nift_codes]

        # Convert all codes to 4-digit strings
        nift_codes = [
            str(code).strip().zfill(4) if code is not None and str(code).strip() != "" else None
            for code in nift_codes
        ]

        # --- Write differences safely in column I ---
        for i, code_str in enumerate(nift_codes, start=2):
            if code_str is None or code_str.lower() == "grand total":
                continue
            if code_str in lookup:
                sheet_outward.range(f"I{i}").value = lookup[code_str]

        wb_outward.save()
        wb_outward.close()

    finally:
        pivot_wb.close()
        app.quit()
        print("nift summary outward complete")

def todays_lodgement(folderPath): 
    clearing_file = os.path.join(folderPath, r"Inward & Outward Clearing GL.xlsx")  
    pivot_wb_path = os.path.join(folderPath, r"extractedData.xlsx")    

    # Launch Excel
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # Open clearing workbook
        wb_clearing = app.books.open(clearing_file)
        sheet_outward = wb_clearing.sheets["Outward"]

        # Open pivot workbook
        pivot_wb = app.books.open(pivot_wb_path, update_links=False)
        pivot_sheet_100 = pivot_wb.sheets['pivot_100_200_500']
        pivot_sheet_300 = pivot_wb.sheets['pivot_300_400_700']

        # Apply filters
        pivot_sheet_100.api.PivotTables(1).PivotFields("Clearing Type").CurrentPage = "Delivered by you"
        pivot_sheet_300.api.PivotTables(1).PivotFields("Clearing Type").CurrentPage = "Delivered by you"

        # Refresh pivots
        pivot_sheet_100.api.PivotTables(1).RefreshTable()
        pivot_sheet_300.api.PivotTables(1).RefreshTable()

        # Function to read pivot into pandas
        def read_pivot(sheet):
            pt = sheet.api.PivotTables(1)
            values = sheet.range(pt.TableRange1.Address).value
            df = pd.DataFrame(values[1:], columns=values[0])
            df.columns = ["nift_code", "amount"]
            df["nift_code"] = df["nift_code"].astype(str).str.strip().str.zfill(4)
            df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
            df = df.dropna(subset=["nift_code", "amount"])
            df = df[df["amount"] != 0]
            df = df[df["nift_code"].str.lower() != "grand total"]
            return df

        df_100 = read_pivot(pivot_sheet_100)
        df_300 = read_pivot(pivot_sheet_300)

        # Outer merge to include all codes
        df_merged = pd.merge(df_100, df_300, on="nift_code", how="outer", suffixes=("_100", "_300"))

        # Add amounts together, filling missing with 0
        df_merged["final_amount"] = df_merged[["amount_100", "amount_300"]].fillna(0).sum(axis=1)

        # Create lookup dictionary
        lookup = dict(zip(df_merged["nift_code"], df_merged["final_amount"]))

        # Read NIFT codes from clearing file column A
        last_row = sheet_outward.range("A" + str(sheet_outward.cells.last_cell.row)).end("up").row
        nift_codes = sheet_outward.range(f"A2:A{last_row}").value

        # Ensure always a list
        if isinstance(nift_codes, (str, int, float)):
            nift_codes = [nift_codes]

        # Convert to 4-digit strings
        nift_codes = [
            str(code).strip().zfill(4) if code is not None and str(code).strip() != "" else None
            for code in nift_codes
        ]

        # Write final amounts to column M
        for i, code in enumerate(nift_codes, start=2):
            if code:
                sheet_outward.range(f"M{i}").value = float(lookup.get(code, 0))

        # Save clearing file
        wb_clearing.save()

    finally:
        wb_clearing.close()
        pivot_wb.close()
        app.quit()
        print("Todays Lodgement done")
    
def differenceColumn(folderPath):
    clearing_file = os.path.join(folderPath, r'Inward & Outward Clearing GL.xlsx')

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(clearing_file)

        sheet = wb.sheets["Inward"]

        # Find last row with data in column H
        last_row = sheet.range("H" + str(sheet.cells.last_cell.row)).end("up").row

        # Apply formula row by row
        for row in range(3, last_row + 1):  
            formula = f"=H{row}-I{row}-J{row}+K{row}"
            sheet.range(f"L{row}").formula = formula

        sheet1 = wb.sheets["Outward"]

        # Find last row with data in column H
        last_row = sheet1.range("H" + str(sheet1.cells.last_cell.row)).end("up").row

        # Apply formula row by row
        for row in range(3, last_row + 1):  
            formula = f"=H{row}-(I{row}+J{row}+K{row}+L{row}-M{row})"
            sheet1.range(f"N{row}").formula = formula


        wb.save()
        print("Difference column formulas applied successfully!")

    finally:
        wb.close()
        app.quit()

if __name__ == "__main__":    
    extractPDF(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    createBulkSheet(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    OP_Balance_(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    clearingBulk(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    clearingBulkD(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess") 
    combineEMACC(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    normalInward(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    normalOutward(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    create_pivot_tables(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    nift_summary(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    nift_summary_outward(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    todays_lodgement(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")
    differenceColumn(r"C:\Users\khani\Downloads\NIFT Recon\ReadyToProcess")

