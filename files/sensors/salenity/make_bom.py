"""Generate the salinity-sensor BOM (xlsx + csv).

Every value below is copied from the manufacturer's product page or datasheet
listed in SOURCES. Anything not stated there is "unknown".
Run:  python make_bom.py
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

U = "unknown"
HERE = Path(__file__).parent

SENSORS = ["SX-CCFNA", "YEX-S2-EC-S", "Y523-A", "WS-OM-MUL-17A"]

# (spec, {sensor: value})
ROWS = [
    ("Item #", {"SX-CCFNA": "1", "YEX-S2-EC-S": "2", "Y523-A": "3", "WS-OM-MUL-17A": "4"}),
    ("Manufacturer / brand", {
        "SX-CCFNA": "SensorsX",
        "YEX-S2-EC-S": "Changsha Yex Technology Co., Ltd. (Yexsensor)",
        "Y523-A": "Yosemitech",
        "WS-OM-MUL-17A": "WizSensor"}),
    ("Product name", {
        "SX-CCFNA": "SX-CCFNA series conductivity sensor",
        "YEX-S2-EC-S": "Online Conductivity Sensor",
        "Y523-A": "Online Conductivity (Salinity) Sensor",
        "WS-OM-MUL-17A": "CTD Sensor (conductivity, salinity, temperature, depth)"}),
    ("Quantity", {s: "1" for s in SENSORS}),
    ("Unit price", {s: U for s in SENSORS}),
    ("Price note", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Quote on request (\"Get quote\" link)",
        "Y523-A": "Quote on request (price-list inquiry form)",
        "WS-OM-MUL-17A": "Quote on request; pricing and lead time confirmed by quotation"}),
    ("Measured parameters", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Conductivity + temperature",
        "Y523-A": "Conductivity, salinity, TDS, temperature",
        "WS-OM-MUL-17A": "Temperature, conductivity, salinity (calculated from conductivity), depth"}),
    ("Measuring principle / electrode", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": U,
        "Y523-A": "Four-electrode",
        "WS-OM-MUL-17A": U}),
    ("Conductivity range", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "0 ~ 5000 µS/cm",
        "Y523-A": "0 - 200 mS/cm (sub-ranges 0-2, 2-20, 20-200 mS/cm)",
        "WS-OM-MUL-17A": "0 - 100 mS/cm"}),
    ("Conductivity accuracy", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "±1.5 % F.S.",
        "Y523-A": "±1 % or 0.01 mS/cm (whichever is bigger)",
        "WS-OM-MUL-17A": "±1 % or 0.01 mS/cm (whichever is greater)"}),
    ("Conductivity resolution", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "1 µS/cm",
        "Y523-A": U,
        "WS-OM-MUL-17A": "0.1 µS/cm"}),
    ("Salinity range", {
        "SX-CCFNA": U, "YEX-S2-EC-S": U,
        "Y523-A": "0 - 175 ppt",
        "WS-OM-MUL-17A": "0 - 80 ppt"}),
    ("Salinity accuracy", {
        "SX-CCFNA": U, "YEX-S2-EC-S": U,
        "Y523-A": "±1 ppt",
        "WS-OM-MUL-17A": "±1 ppt"}),
    ("Salinity resolution", {s: U for s in SENSORS}),
    ("TDS range", {
        "SX-CCFNA": U, "YEX-S2-EC-S": U,
        "Y523-A": "0 - 128000 mg/L",
        "WS-OM-MUL-17A": U}),
    ("Temperature range", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": U,
        "Y523-A": "0 - 50 °C (page does not say if measuring or operating range)",
        "WS-OM-MUL-17A": "-5 to 50 °C"}),
    ("Temperature accuracy", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "±0.3 °C",
        "Y523-A": U,
        "WS-OM-MUL-17A": "±0.05 °C"}),
    ("Temperature resolution", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "0.1 °C",
        "Y523-A": U,
        "WS-OM-MUL-17A": "0.005 °C"}),
    ("Depth range", {
        "SX-CCFNA": U, "YEX-S2-EC-S": U, "Y523-A": U,
        "WS-OM-MUL-17A": "0 - 200 m (accuracy ±0.1 % F.S., resolution 0.01 m)"}),
    ("Sensitivity", {s: U for s in SENSORS}),
    ("Response time", {s: U for s in SENSORS}),
    ("Operating / supply voltage", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "12 ~ 24 VDC",
        "Y523-A": "DC 9 - 24 V",
        "WS-OM-MUL-17A": "6 - 36 VDC"}),
    ("Power / current", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "0.2 W @ 12 V",
        "Y523-A": "0.2 W (listed as \"0.2W(DC 9-24V,>500mA)\")",
        "WS-OM-MUL-17A": "Approx. 20 mA @ 12 VDC"}),
    ("Output / communication", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "RS-485 (Modbus/RTU); 4-20 mA optional",
        "Y523-A": "RS-485, Modbus protocol",
        "WS-OM-MUL-17A": "RS485 Modbus RTU"}),
    ("Default serial settings", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "9600 baud, 8 data bits, no parity, 1 stop bit; address 1 (settable 1-255)",
        "Y523-A": U,
        "WS-OM-MUL-17A": "9600 baud, 8N1, address 1"}),
    ("Wiring", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Red = 12-24 VDC, Black = GND, Blue = 485A, Green = 485B, Yellow = current output",
        "Y523-A": U,
        "WS-OM-MUL-17A": "Red = VCC, Black = GND, White = RS485 B, Green = RS485 A"}),
    ("Temperature compensation", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Automatic (Pt1000)",
        "Y523-A": "Automatic (built-in temperature sensor)",
        "WS-OM-MUL-17A": U}),
    ("Working temperature", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "0 ~ 65 °C (datasheet); web page says 0 ~ 60 °C",
        "Y523-A": U,
        "WS-OM-MUL-17A": U}),
    ("Working pressure / depth rating", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "< 0.6 MPa",
        "Y523-A": U,
        "WS-OM-MUL-17A": "0 - 20 bar (200 m depth range)"}),
    ("Protection rating", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "IP68",
        "Y523-A": "IP68",
        "WS-OM-MUL-17A": U}),
    ("Housing / wetted material", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "ABS, POM and 316L stainless steel",
        "Y523-A": "Ni + PEEK + Ti",
        "WS-OM-MUL-17A": "316L stainless steel"}),
    ("Dimensions", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Ø30 x 185 mm (datasheet drawing)",
        "Y523-A": "Ø22 x 175.5 mm",
        "WS-OM-MUL-17A": "Ø22 x 220.5 mm"}),
    ("Weight", {s: U for s in SENSORS}),
    ("Cable", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "5 m (customizable); 5-core shielded twisted pair; M16 5-pin waterproof connector",
        "Y523-A": "10 m default, customizable",
        "WS-OM-MUL-17A": "30 m standard, custom available; 4-core shielded, AWG 24/26, OD 5.5 mm"}),
    ("Mounting", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Submersion mount, 3/4\" NPT thread; keep ≥2 cm from bottom and walls",
        "Y523-A": U,
        "WS-OM-MUL-17A": "Do not use the signal cable to lift or hang the sensor"}),
    ("Calibration", {
        "SX-CCFNA": U,
        "YEX-S2-EC-S": "Two-point (zero in air, slope in standard solution)",
        "Y523-A": "One- or two-point; factory-calibrated before delivery",
        "WS-OM-MUL-17A": "Temp & conductivity: one/two-point; depth: zero/one/two-point; stored in probe"}),
    ("Other notes", {
        "SX-CCFNA": "Product page could not be read (HTTP 429 rate limit); all specs pending",
        "YEX-S2-EC-S": "72-hour aging and stability test before shipment; package = sensor, manual, certificate",
        "Y523-A": "No datasheet linked on the page",
        "WS-OM-MUL-17A": "Typical cleaning interval 4-8 weeks; datasheet/manual/Modbus map on request; logger, battery and telemetry not included"}),
]

DFR = "SEN0707"
SENSORS.append(DFR)
DFR_VALUES = {
    "Item #": "5",
    "Manufacturer / brand": "DFRobot",
    "Product name": "RS485 Industrial-Grade Water Electrical Conductivity Sensor (K=10, EC / TDS / Salinity)",
    "Quantity": "1",
    "Unit price": "USD 159.00",
    "Price note": "Volume: USD 155.00 (3+), 150.00 (5+), 146.00 (10+); stock status not stated",
    "Measured parameters": "Conductivity, TDS, salinity, temperature",
    "Measuring principle / electrode": "Electrode constant K=10",
    "Conductivity range": "10 - 20000 µS/cm",
    "Conductivity accuracy": "±1 % FS",
    "Conductivity resolution": "1 µS/cm",
    "Salinity range": "0 - 11476 ppm",
    "Salinity accuracy": U,
    "Salinity resolution": U,
    "TDS range": "0 - 13400 ppm",
    "Temperature range": "-20 to +60 °C",
    "Temperature accuracy": "±0.5 °C",
    "Temperature resolution": "0.1 °C",
    "Depth range": U,
    "Sensitivity": U,
    "Response time": "< 1 s",
    "Operating / supply voltage": "10 ~ 30 V DC",
    "Power / current": "0.4 W",
    "Output / communication": "RS485, standard Modbus-RTU",
    "Default serial settings": "4800 bit/s default (1200-115200 selectable); address 1",
    "Wiring": "Brown = VCC (10-30 VDC), Black = GND, Yellow = 485-A, Blue = 485-B",
    "Temperature compensation": "Automatic, built-in sensor; -20 to +60 °C; default coefficient 0.02",
    "Working temperature": "-20 to +80 °C",
    "Working pressure / depth rating": "0.6 MPa",
    "Protection rating": "IP68",
    "Housing / wetted material": "Corrosion-resistant plastic",
    "Dimensions": "Ø30, 212 body / 266 overall incl. connector (wiki drawing, units not labelled)",
    "Weight": U,
    "Cable": "5 m",
    "Mounting": "3/4\" NPT thread at both ends (wiki drawing)",
    "Calibration": "Factory-calibrated; 12880 µS/cm calibration solution (100 ml) included",
    "Other notes": "Box: sensor, 5 m cable, calibration solution, protective mesh, 2 cable ties; wiki advises 120 Ω termination for long runs",
}
for _spec, _vals in ROWS:
    _vals[DFR] = DFR_VALUES[_spec]

# SX-CCFNA: from the SensorsX datasheet (Rev A, local PDF in "SensorsX SX-CCFNA/").
SX_DATASHEET = "SensorsX SX-CCFNA/SensorsX_SX-CCFNA_Datasheet_RevA_ffed563b-795e-462b-b9ea-870c98f6074c.pdf"
SX_VALUES = {
    "Manufacturer / brand": "SensorsX Technology Inc.",
    "Product name": "SX-CCFNA Series Four-Electrode Conductivity (Salinity) Sensor",
    "Price note": "Not stated in datasheet",
    "Measured parameters": "Conductivity, salinity, TDS, temperature",
    "Measuring principle / electrode": "Four-electrode",
    "Conductivity range": "0.001-5 mS/cm or 0.001-100 mS/cm (order option)",
    "Conductivity accuracy": "±1 % or 0.01 mS/cm (whichever is greater)",
    "Salinity range": "0-2.5 ppt or 0-80 ppt (order option)",
    "Salinity accuracy": "±0.05 ppt (0-2.5 ppt) or ±1 ppt (0-80 ppt)",
    "TDS range": "0-3,200 mg/L or 0-64,000 mg/L (order option)",
    "Temperature range": "0-50 °C (operating)",
    "Operating / supply voltage": "DC 9-24 V, >500 mA",
    "Power / current": "0.2 W",
    "Output / communication": "RS-485, Modbus",
    "Temperature compensation": "Automatic (built-in temperature sensor)",
    "Working temperature": "0-50 °C (32-122 °F)",
    "Protection rating": "IP68",
    "Housing / wetted material": "Nickel + PEEK + Titanium alloy",
    "Dimensions": "Φ22 mm x 175.5 mm",
    "Cable": "10 m default, customizable",
    "Calibration": "1-point or 2-point; factory-calibrated",
    "Other notes": "Order code SX-CCFNA-11111 selects ranges; specs based on lab standard-solution conditions",
}
for _spec, _vals in ROWS:
    if _spec in SX_VALUES:
        _vals["SX-CCFNA"] = SX_VALUES[_spec]

# Columns shown on the BOM sheet (chosen by the user), plus link columns.
KEEP = ["Manufacturer / brand", "Product name", "Unit price", "Price note", "Measured parameters",
        "Conductivity range", "Salinity range", "Temperature range", "Operating / supply voltage",
        "Output / communication", "Wiring", "Dimensions"]
LINKS = {  # sensor: (product page, datasheet)
    "SX-CCFNA": ("https://sensorsx.com/products/sx-ccfna-series", SX_DATASHEET),
    "YEX-S2-EC-S": ("https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html",
                    "https://www.yexsensor.com/static/upload/file/20260507/1778143536308368.pdf"),
    "Y523-A": ("https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html", "Not linked on product page"),
    "WS-OM-MUL-17A": ("https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/", "On request from manufacturer"),
    "SEN0707": ("https://www.dfrobot.com/product-3099.html", "https://wiki.dfrobot.com/sen0707/"),
}

SOURCES = [
    ("SX-CCFNA", "Product page", "https://sensorsx.com/products/sx-ccfna-series"),
    ("SX-CCFNA", "Datasheet Rev A (local PDF)", SX_DATASHEET),
    ("YEX-S2-EC-S", "Product page", "https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html"),
    ("YEX-S2-EC-S", "User manual / datasheet (PDF)", "https://www.yexsensor.com/static/upload/file/20260507/1778143536308368.pdf"),
    ("Y523-A", "Product page", "https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html"),
    ("WS-OM-MUL-17A", "Product page", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
    ("SEN0707", "Product page", "https://www.dfrobot.com/product-3099.html"),
    ("SEN0707", "Product wiki", "https://wiki.dfrobot.com/sen0707/"),
]


def transposed():
    """Header = ["Sensor", spec1, spec2, ...]; one row per sensor."""
    rows = dict(ROWS)
    header = ["Sensor"] + KEEP + ["Product page link", "Datasheet link"]
    table = [[s] + [rows[k][s] for k in KEEP] + list(LINKS[s]) for s in SENSORS]
    return header, table


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Salinity BOM"
    head_fill = PatternFill("solid", fgColor="0B4F7A")
    spec_fill = PatternFill("solid", fgColor="E6F4F8")
    unk_font = Font(color="9A9A9A", italic=True)
    thin = Side(style="thin", color="B0C4CC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    # One row per sensor, one column per specification.
    header, table = transposed()
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 32
    for row in table:
        ws.append(row)
        for i, c in enumerate(ws[ws.max_row]):
            c.alignment = wrap
            c.border = border
            if i == 0:
                c.font = Font(bold=True)
                c.fill = spec_fill
            elif c.value == U:
                c.font = unk_font
            elif i >= len(header) - 2 and (c.value.startswith("http") or c.value.endswith(".pdf")):
                c.hyperlink = c.value
                c.font = Font(color="0563C1", underline="single")
    ws.column_dimensions["A"].width = 18
    for i, name in enumerate(header[1:], start=2):
        longest = max(len(str(r[i - 1])) for r in table)
        ws.column_dimensions[get_column_letter(i)].width = min(40, max(12, len(name) + 2, longest * 0.9))
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    src = wb.create_sheet("Sources")
    src.append(["Sensor", "Document", "URL"])
    for c in src[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    for row in SOURCES:
        src.append(row)
        src.cell(src.max_row, 3).hyperlink = row[2]
        src.cell(src.max_row, 3).style = "Hyperlink"
    src.append([])
    src.append(["Rule: every value comes from the listed manufacturer page or datasheet; anything not stated there is \"unknown\"."])
    src.column_dimensions["A"].width = 18
    src.column_dimensions["B"].width = 32
    src.column_dimensions["C"].width = 100

    wb.save(HERE / "Salinity_Sensors_BOM.xlsx")
    with open(HERE / "Salinity_Sensors_BOM.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(table)
        w.writerow([])
        w.writerow(["Sources"])
        for row in SOURCES:
            w.writerow(row)


if __name__ == "__main__":
    main()
