"""Build Sensors_BOM.xlsx (+ .csv) in the AirZzen BOM style.

Each ITEM has its own header row and its own columns. Every value is copied from
the manufacturer's product page or datasheet (see the link columns); anything
not stated there is "Unknown".
Run:  python build_bom.py
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
UNK = "Unknown"
NA = "Not linked on product page"

# ---------------------------------------------------------------- data
# Link columns hold (display text, target). target None = plain text.
SX_PDF = "salenity/SensorsX SX-CCFNA/SensorsX_SX-CCFNA_Datasheet_RevA_ffed563b-795e-462b-b9ea-870c98f6074c.pdf"

ITEMS = [
    {
        "title": "1   ITEM 1 — SALINITY UNIT  (Temperature · Conductivity · Salinity)",
        "columns": ["#", "Parameter", "Sensor (Part No.)", "Manufacturer", "Unit Price (USD)", "Include in Cost?",
                    "Conductivity Range", "Salinity Range", "Temperature Range", "Operating Voltage",
                    "Comm. Protocol / Output", "Wiring", "Dimensions", "Buy Link", "Datasheet",
                    "Factory Calibrated?", "How to Calibrate"],
        "rows": [
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "SX-CCFNA",
             "Manufacturer": ("SensorsX ↗", "https://sensorsx.com/products/sx-ccfna-series"),
             "Unit Price (USD)": UNK,
             "Conductivity Range": "0.001 – 5 mS/cm or\n0.001 – 100 mS/cm (order option)",
             "Salinity Range": "0 – 2.5 ppt or\n0 – 80 ppt (order option)",
             "Temperature Range": "0 – 50 °C (operating)",
             "Operating Voltage": "9 – 24 V DC, >500 mA; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus)",
             "Wiring": UNK, "Dimensions": "Φ22 × 175.5 mm",
             "Buy Link": ("sensorsx.com", "https://sensorsx.com/products/sx-ccfna-series"),
             "Datasheet": ("Datasheet Rev A (local PDF)", SX_PDF),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Factory-calibrated (built-in calibration parameters). Datasheet lists 1-point or 2-point calibration; procedure not given."},
            {"Parameter": "Conductivity · Temp", "Sensor (Part No.)": "YEX-S2-EC-S",
             "Manufacturer": ("Yexsensor ↗", "https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html"),
             "Unit Price (USD)": UNK,
             "Conductivity Range": "0 – 5,000 µS/cm", "Salinity Range": UNK,
             "Temperature Range": "Working temp 0 – 65 °C (datasheet); web page says 0 – 60 °C",
             "Operating Voltage": "12 – 24 V DC; 0.2 W @ 12 V",
             "Comm. Protocol / Output": "RS-485 (Modbus/RTU), 9600 8N1, addr 1; 4–20 mA optional",
             "Wiring": "Red = 12–24 V DC, Black = GND, Blue = 485A, Green = 485B, Yellow = current output",
             "Dimensions": "Ø30 × 185 mm",
             "Buy Link": ("yexsensor.com", "https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html"),
             "Datasheet": ("yexsensor.com", "https://www.yexsensor.com/static/upload/file/20260507/1778143536308368.pdf"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Two-point. Zero: rinse with distilled water, blot dry, hold vertical in air ~3 min, write zero register 0x1000. Slope: in standard solution (10 % FS – FS), ≥2 cm from walls, write register 0x1004. Re-calibrate after a sensor reset."},
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "Y523-A",
             "Manufacturer": ("Yosemitech ↗", "https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html"),
             "Unit Price (USD)": UNK,
             "Conductivity Range": "0 – 200 mS/cm\n(0–2, 2–20, 20–200 mS/cm)", "Salinity Range": "0 – 175 ppt",
             "Temperature Range": "0 – 50 °C (not stated if measuring or operating)",
             "Operating Voltage": "9 – 24 V DC, >500 mA; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus)", "Wiring": UNK, "Dimensions": "Ø22 × 175.5 mm",
             "Buy Link": ("e.yosemitech.com", "https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Calibrated before delivery (built-in calibration parameters). Supports one-point or two-point calibration; procedure not given on page."},
            {"Parameter": "Conductivity · Salinity · Temp · Depth (CTD)", "Sensor (Part No.)": "WS-OM-MUL-17A",
             "Manufacturer": ("WizSensor ↗", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
             "Unit Price (USD)": UNK,
             "Conductivity Range": "0 – 100 mS/cm", "Salinity Range": "0 – 80 ppt (calculated)",
             "Temperature Range": "-5 to 50 °C",
             "Operating Voltage": "6 – 36 V DC; ~20 mA @ 12 V",
             "Comm. Protocol / Output": "RS-485 (Modbus RTU), 9600 8N1, addr 1",
             "Wiring": "Red = VCC, Black = GND, White = RS485 B, Green = RS485 A",
             "Dimensions": "Ø22 × 220.5 mm",
             "Buy Link": ("wizsensor.com", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
             "Datasheet": ("On request from manufacturer", None),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Temp & conductivity: one- or two-point. Depth: zero-, one- or two-point via SmartPC or Modbus. Calibration is stored in the probe."},
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "SEN0707",
             "Manufacturer": ("DFRobot ↗", "https://www.dfrobot.com/product-3099.html"),
             "Unit Price (USD)": 159.00,
             "Conductivity Range": "10 – 20,000 µS/cm (K=10)", "Salinity Range": "0 – 11,476 ppm",
             "Temperature Range": "-20 to 60 °C (operating -20 to 80 °C)",
             "Operating Voltage": "10 – 30 V DC; 0.4 W",
             "Comm. Protocol / Output": "RS-485 (Modbus-RTU), 4800 default, addr 1",
             "Wiring": "Brown = VCC, Black = GND, Yellow = 485-A, Blue = 485-B",
             "Dimensions": "Ø30 × 212 (266 overall), units not labelled",
             "Buy Link": ("dfrobot.com", "https://www.dfrobot.com/product-3099.html"),
             "Datasheet": ("wiki.dfrobot.com", "https://wiki.dfrobot.com/sen0707/"),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Factory-calibrated, ready to use. 12,880 µS/cm calibration solution (100 ml) included; registers allow re-optimization after long use."},
        ],
    },
    {
        "title": "2   ITEM 2 — CURRENT UNIT  (Water Current Speed · Direction)",
        "columns": ["#", "Parameter", "Sensor (Part No.)", "Manufacturer", "Unit Price (USD)", "Include in Cost?",
                    "Speed Range", "Speed Accuracy", "Direction Range / Accuracy", "Other Sensors",
                    "Operating Voltage", "Comm. Protocol / Output", "Dimensions", "Buy Link", "Datasheet",
                    "Factory Calibrated?", "How to Calibrate"],
        "rows": [
            {"Parameter": "Current Speed · Direction (acoustic)", "Sensor (Part No.)": "OCEAN-SPCM",
             "Manufacturer": ("Oceantek ↗", "https://oceanadcp.com/products/single-point-current-meter/"),
             "Unit Price (USD)": UNK,
             "Speed Range": "±5.0 m/s (default)\n±20.0 m/s (max)",
             "Speed Accuracy": "±0.3 % ± 3 mm/s; resolution 0.1 mm/s",
             "Direction Range / Accuracy": "0 – 360°, ±2°; resolution 0.01°",
             "Other Sensors": "Temp -5 to 45 °C (±0.1 °C); pressure ±0.25 %; roll/pitch ±85°, heading 0–360°",
             "Operating Voltage": "20 – 50 V, 5 A (DR); 504 Wh Li battery (SC)",
             "Comm. Protocol / Output": "RS-232 or RS-422; PD0 binary, 1 Hz",
             "Dimensions": "Φ148 mm; DR 150–167 mm / SC 350–367 mm long; depth 1000/3000/6000 m",
             "Buy Link": ("oceanadcp.com", "https://oceanadcp.com/products/single-point-current-meter/"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Not stated", "How to Calibrate": UNK},
            {"Parameter": "Current Speed · Direction · Depth (rotor)", "Sensor (Part No.)": "EMCON Current Meter",
             "Manufacturer": ("EMCON India ↗", "https://emconindia.com/products/water-current-meter-with-direction-for-river-oceans-or-estuaries/"),
             "Unit Price (USD)": UNK,
             "Speed Range": "0 – 5 m/s",
             "Speed Accuracy": "0.02 m/s; threshold < 0.04 m/s",
             "Direction Range / Accuracy": "Unknown (Hall-effect sensor, no spec given)",
             "Other Sensors": "Depth (hydrostatic), 20 m, ±1 %",
             "Operating Voltage": "Unknown (handheld reader has internal battery)",
             "Comm. Protocol / Output": "Unknown (reads m/s on handheld LCD)",
             "Dimensions": "Unknown; 20 m cable (longer on demand)",
             "Buy Link": ("emconindia.com", "https://emconindia.com/products/water-current-meter-with-direction-for-river-oceans-or-estuaries/"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Not stated", "How to Calibrate": UNK},
        ],
    },
]

NOTES = [
    "Notes",
    "• Every value comes from the manufacturer's product page or datasheet linked in each row. Anything not stated there is marked \"Unknown\".",
    "• Sensors within an item are alternatives. Set \"Include in Cost?\" to Yes for the one you choose; subtotals only count Yes rows with a numeric price.",
    "• Prices: only SEN0707 lists a price (USD 159; 155 for 3+, 150 for 5+, 146 for 10+). All others are quote-on-request.",
    "• SX-CCFNA and Y523-A share the same size, power, cable (10 m) and housing (Ni + PEEK + Ti) in their documents.",
]

# ---------------------------------------------------------------- styles
NAVY, BLUE, PALE, ROWALT, BORDER = "22449A", "2589CE", "BCD3EA", "EAF3FB", "BCD3EA"
TXT = "1B2B4B"
F = lambda **k: Font(name="Arial", **{"size": 10, "color": TXT, **k})
thin = lambda c: Side(style="thin", color=c)
BOX = Border(left=thin(BORDER), right=thin(BORDER), top=thin(BORDER), bottom=thin(BORDER))
HBOX = Border(left=thin("4A66B0"), right=thin("4A66B0"), top=thin("4A66B0"), bottom=thin("4A66B0"))
fill = lambda c: PatternFill("solid", fgColor=c)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LINK = F(color="0563C1", underline="single")
USD = '"$"#,##0.00'

WIDTHS = [5, 24, 18, 16, 12, 11, 24, 22, 24, 26, 24, 26, 26, 17, 20, 13, 60]
LEFT_COLS = {"Parameter", "Sensor (Part No.)", "Manufacturer", "How to Calibrate", "Wiring", "Other Sensors"}
NCOL = len(WIDTHS)


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.row_dimensions[1].height = 7.5
    ws["B2"] = "Alex-Mon — Sensors BOM"
    ws["B2"].font = Font(name="Arial", size=20, bold=True, color=NAVY)
    ws.row_dimensions[2].height = 31.5
    ws["B3"] = "SSTM  |  Simply Smart"
    ws["B3"].font = Font(name="Arial", size=10, bold=True, color="7F93AD")
    ws.row_dimensions[4].height = 3.75

    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False,
                        errorTitle="Include in Cost?", error="Choose Yes or No")
    ws.add_data_validation(dv)
    subtotals, csv_rows = [], [["Alex-Mon — Sensors BOM"], []]
    r, n = 6, 0
    for item in ITEMS:
        cols = item["columns"]
        ci = {c: i + 1 for i, c in enumerate(cols)}
        # item band
        for c in range(1, NCOL + 1):
            ws.cell(r, c).fill = fill(BLUE)
        ws.cell(r, 1, item["title"]).font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 21.75
        csv_rows.append([item["title"]])
        r += 1
        # header
        for name, c in ci.items():
            cell = ws.cell(r, c, name)
            cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            cell.fill, cell.alignment, cell.border = fill(NAVY), CENTER, HBOX
        ws.row_dimensions[r].height = 37.5
        csv_rows.append(cols)
        r += 1
        first = r
        for k, row in enumerate(item["rows"]):
            n += 1
            bg = fill("FFFFFF" if k % 2 == 0 else ROWALT)
            row = {"#": n, "Include in Cost?": "No", **row}
            csv_line, longest = [], 0
            for name, c in ci.items():
                val = row.get(name, UNK)
                target = None
                if isinstance(val, tuple):
                    val, target = val
                cell = ws.cell(r, c, val)
                cell.fill, cell.border = bg, BOX
                cell.alignment = LEFT if name in LEFT_COLS else CENTER
                cell.font = F()
                if name in ("Parameter", "Include in Cost?", "Factory Calibrated?"):
                    cell.font = F(bold=True)
                if name == "Sensor (Part No.)":
                    cell.font = F(bold=True, color=NAVY)
                if name == "Factory Calibrated?" and str(val).startswith("Yes"):
                    cell.font = F(bold=True, color="1E7B34")
                if name == "Unit Price (USD)" and isinstance(val, (int, float)):
                    cell.number_format = USD
                if str(val).startswith(UNK) or val in (NA, "Not stated"):
                    cell.font = F(italic=True, color="7F7F7F")
                if target:
                    cell.hyperlink = target
                    cell.font = LINK
                if name == "Include in Cost?":
                    dv.add(cell)
                csv_line.append(f"{val} ({target})" if target else val)
                longest = max(longest, len(str(val)) / max(WIDTHS[c - 1], 1))
            ws.row_dimensions[r].height = max(30, min(110, 13 * (longest + 1.2)))
            csv_rows.append(csv_line)
            r += 1
        last = r - 1
        inc = f"F{first}:F{last}"
        price = f"E{first}:E{last}"
        for rule, font_col, bg_col in (('="Yes"', "1E7B34", "E2F0D9"), ('="No"', "7F7F7F", "F2F2F2")):
            ws.conditional_formatting.add(inc, FormulaRule(formula=[f"F{first}{rule}"],
                                          font=Font(color=font_col), fill=fill(bg_col)))
        # subtotal
        idx = item["title"].split()[2]
        for c in range(2, NCOL + 1):
            ws.cell(r, c).fill, ws.cell(r, c).border = fill(PALE), BOX
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 2, f"Item {idx} SUBTOTAL (USD) — included items").font = Font(name="Arial", size=11, bold=True, color=TXT)
        ws.cell(r, 2).alignment = Alignment(vertical="center")
        sub = ws.cell(r, 5, f'=SUMIFS({price},{inc},"Yes")')
        sub.font, sub.alignment, sub.number_format = Font(name="Arial", size=11, bold=True, color=NAVY), CENTER, USD
        warn = ws.cell(r, 7, f'=IF(SUMPRODUCT(({inc}="Yes")*NOT(ISNUMBER({price})))>0,"⚠ "&SUMPRODUCT(({inc}="Yes")*NOT(ISNUMBER({price})))&" included item(s) with unknown price — not in subtotal","")')
        warn.font, warn.alignment = Font(name="Arial", size=10, bold=True, color="B26A00"), Alignment(vertical="center")
        ws.row_dimensions[r].height = 24
        subtotals.append(f"E{r}")
        csv_rows += [[f"Item {idx} subtotal: see xlsx (Include in Cost? = Yes)"], []]
        r += 2

    # grand total
    for c in range(2, NCOL + 1):
        ws.cell(r, c).fill = fill(NAVY)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.cell(r, 2, "GRAND TOTAL — ALL ITEMS (USD)").font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    ws.cell(r, 2).alignment = Alignment(vertical="center")
    g = ws.cell(r, 5, "=" + "+".join(subtotals))
    g.font, g.alignment, g.number_format = Font(name="Arial", size=12, bold=True, color="FFFFFF"), CENTER, USD
    ws.row_dimensions[r].height = 26
    r += 2

    for k, note in enumerate(NOTES):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NCOL)
        cell = ws.cell(r, 2, note)
        cell.font = Font(name="Arial", size=11 if k == 0 else 10, bold=k == 0, color=NAVY if k == 0 else TXT)
        cell.alignment = LEFT
        ws.row_dimensions[r].height = 21.75 if k == 0 else 18
        csv_rows.append([note])
        r += 1

    ws.freeze_panes = "D5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(HERE / "Sensors_BOM.xlsx")
    with open(HERE / "Sensors_BOM.csv", "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(csv_rows)
    print("saved", r - 1, "rows")


if __name__ == "__main__":
    main()
