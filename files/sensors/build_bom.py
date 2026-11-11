"""Build Sensors_BOM.xlsx (+ .csv) in the AirZzen BOM style.

Each ITEM has its own header row and its own columns. Spec values are copied
from the manufacturer's product page or datasheet (see the link columns);
anything not stated there is "Unknown". Estimated prices and DIY entries come
from "Underwater Site Monitoring – Sensor Options_1.xlsx" and are labelled so
in the Notes column.
Run:  python build_bom.py
"""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
UNK = "Unknown"
NA = "Not linked on product page"
OPT = "Sensor Options sheet"  # the user's options workbook, used for estimates / DIY entries

SX_PDF = "salenity/SensorsX SX-CCFNA/SensorsX_SX-CCFNA_Datasheet_RevA_ffed563b-795e-462b-b9ea-870c98f6074c.pdf"

# Columns every item starts and ends with.
HEAD = ["#", "Parameter", "Sensor (Part No.)", "Manufacturer", "Unit Price (USD)", "Include in Cost?", "Seawater OK?"]
TAIL = ["Buy Link", "Datasheet", "Factory Calibrated?", "How to Calibrate", "Notes"]

# ---------------------------------------------------------------- data
# Link-type values are (display text, target); target None = plain text.
ITEMS = [
    {
        "title": "1   ITEM 1 — SALINITY UNIT  (Temperature · Conductivity · Salinity)  ·  Priority: Essential",
        "columns": HEAD + ["Conductivity Range", "Salinity Range", "Temperature Range", "Operating Voltage",
                           "Comm. Protocol / Output", "Wiring", "Dimensions"] + TAIL,
        "rows": [
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "SX-CCFNA",
             "Manufacturer": ("SensorsX ↗", "https://sensorsx.com/products/sx-ccfna-series"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Conductivity Range": "0.001 – 5 mS/cm or\n0.001 – 100 mS/cm (order option)",
             "Salinity Range": "0 – 2.5 ppt or\n0 – 80 ppt (order option)",
             "Temperature Range": "0 – 50 °C (operating)",
             "Operating Voltage": "9 – 24 V DC, >500 mA; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus)",
             "Wiring": UNK, "Dimensions": "Φ22 × 175.5 mm; 10 m cable; IP68",
             "Buy Link": ("sensorsx.com", "https://sensorsx.com/products/sx-ccfna-series"),
             "Datasheet": ("Datasheet Rev A (local PDF)", SX_PDF),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Factory-calibrated (built-in calibration parameters). Datasheet lists 1-point or 2-point calibration; procedure not given.",
             "Notes": "Order the 0 – 100 mS/cm / 0 – 80 ppt range for seawater. Same size, power, cable and housing as Y523-A."},
            {"Parameter": "Conductivity · Temp", "Sensor (Part No.)": "YEX-S2-EC-S",
             "Manufacturer": ("Yexsensor ↗", "https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html"),
             "Unit Price (USD)": UNK, "Seawater OK?": "No",
             "Conductivity Range": "0 – 5,000 µS/cm", "Salinity Range": UNK,
             "Temperature Range": "Working temp 0 – 65 °C (datasheet); web page says 0 – 60 °C",
             "Operating Voltage": "12 – 24 V DC; 0.2 W @ 12 V",
             "Comm. Protocol / Output": "RS-485 (Modbus/RTU), 9600 8N1, addr 1; 4–20 mA optional",
             "Wiring": "Red = 12–24 V DC, Black = GND, Blue = 485A, Green = 485B, Yellow = current output",
             "Dimensions": "Ø30 × 185 mm; 5 m cable; IP68; < 0.6 MPa",
             "Buy Link": ("yexsensor.com", "https://www.yexsensor.com/cn/s2-series-13/CN-YEX-S2-EC-S-High-Precision-Online-Conductivity.html"),
             "Datasheet": ("yexsensor.com", "https://www.yexsensor.com/static/upload/file/20260507/1778143536308368.pdf"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Two-point. Zero: rinse with distilled water, blot dry, hold vertical in air ~3 min, write zero register 0x1000. Slope: in standard solution (10 % FS – FS), ≥2 cm from walls, write register 0x1004. Re-calibrate after a sensor reset.",
             "Notes": "Max 5 mS/cm is far below seawater (~50,000 – 60,000 µS/cm per Sensor Options sheet)."},
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "Y523-A",
             "Manufacturer": ("Yosemitech ↗", "https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Conductivity Range": "0 – 200 mS/cm\n(0–2, 2–20, 20–200 mS/cm)", "Salinity Range": "0 – 175 ppt",
             "Temperature Range": "0 – 50 °C (not stated if measuring or operating)",
             "Operating Voltage": "9 – 24 V DC, >500 mA; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus)", "Wiring": UNK,
             "Dimensions": "Ø22 × 175.5 mm; 10 m cable; IP68",
             "Buy Link": ("e.yosemitech.com", "https://e.yosemitech.com/CT/Y523-A-Online-Conductivity-Salinity-Sensor.html"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Calibrated before delivery (built-in calibration parameters). Supports one-point or two-point calibration; procedure not given on page.",
             "Notes": "Price on request."},
            {"Parameter": "Conductivity · Salinity · Temp · Depth (CTD)", "Sensor (Part No.)": "WS-OM-MUL-17A",
             "Manufacturer": ("WizSensor ↗", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Conductivity Range": "0 – 100 mS/cm", "Salinity Range": "0 – 80 ppt (calculated)",
             "Temperature Range": "-5 to 50 °C (±0.05 °C)",
             "Operating Voltage": "6 – 36 V DC; ~20 mA @ 12 V",
             "Comm. Protocol / Output": "RS-485 (Modbus RTU), 9600 8N1, addr 1",
             "Wiring": "Red = VCC, Black = GND, White = RS485 B, Green = RS485 A",
             "Dimensions": "Ø22 × 220.5 mm; 30 m cable; depth 0 – 200 m",
             "Buy Link": ("wizsensor.com", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
             "Datasheet": ("On request from manufacturer", None),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Temp & conductivity: one- or two-point. Depth: zero-, one- or two-point via SmartPC or Modbus. Calibration is stored in the probe.",
             "Notes": "Also gives depth. Typical cleaning every 4 – 8 weeks. Price on request."},
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "SEN0707",
             "Manufacturer": ("DFRobot ↗", "https://www.dfrobot.com/product-3099.html"),
             "Unit Price (USD)": 159.00, "Seawater OK?": "No",
             "Conductivity Range": "10 – 20,000 µS/cm (K=10)", "Salinity Range": "0 – 11,476 ppm",
             "Temperature Range": "-20 to 60 °C (operating -20 to 80 °C)",
             "Operating Voltage": "10 – 30 V DC; 0.4 W",
             "Comm. Protocol / Output": "RS-485 (Modbus-RTU), 4800 default, addr 1",
             "Wiring": "Brown = VCC, Black = GND, Yellow = 485-A, Blue = 485-B",
             "Dimensions": "Ø30 × 212 (266 overall), units not labelled; 5 m cable; IP68; 0.6 MPa",
             "Buy Link": ("dfrobot.com", "https://www.dfrobot.com/product-3099.html"),
             "Datasheet": ("wiki.dfrobot.com", "https://wiki.dfrobot.com/sen0707/"),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Factory-calibrated, ready to use. 12,880 µS/cm calibration solution (100 ml) included; registers allow re-optimization after long use.",
             "Notes": "Seawater (~50,000 – 60,000 µS/cm) is above its range, so it reads at full scale — not suitable (Sensor Options sheet)."},
            {"Parameter": "Conductivity · Salinity · Temp", "Sensor (Part No.)": "DR-EC200",
             "Manufacturer": ("Dragino ↗", "https://docs.dragino.com/docs/Non-Wireless-Sensors/water-quality/dr-eck1.0-dr-eck10.0-ec200-ec-digital-electrode-sensor"),
             "Unit Price (USD)": 375.00, "Include in Cost?": "Yes", "Seawater OK?": "Yes",
             "Conductivity Range": "1 – 200,000 µS/cm (±1 % FS; 1 µS/cm)",
             "Salinity Range": "0 – 70 PSU (0.1 PSU)",
             "Temperature Range": "-5 to 80 °C (±0.5 °C); ambient 0 – 60 °C",
             "Operating Voltage": "7 – 30 V DC; < 0.5 W",
             "Comm. Protocol / Output": "RS-485 (Modbus-RTU), 9600 8N1",
             "Wiring": UNK,
             "Dimensions": "Unknown; 5 m cable; IP68 (except connector); 0.6 MPa",
             "Buy Link": ("antratek.com", "https://www.antratek.com/dr-ec200-water-ec-sensor-for-wqs"),
             "Datasheet": ("docs.dragino.com", "https://docs.dragino.com/docs/Non-Wireless-Sensors/water-quality/dr-eck1.0-dr-eck10.0-ec200-ec-digital-electrode-sensor"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Two-point via Modbus registers: first 1,413 µS/cm, then 111,310 µS/cm standard solution. Calibrate before each use; every 3 months for long-term use.",
             "Notes": "Four-electrode; Dragino says it suits seawater. Price is an estimate from Sensor Options sheet (€325.61 at Antratek; page could not be checked)."},
            {"Parameter": "Conductivity · Salinity · TDS · Temp", "Sensor (Part No.)": "C4E",
             "Manufacturer": ("Aqualabo (Ponsel) ↗", "https://ictinternational.com/wp-content/uploads/2023/06/Datasheet-Conductivity-digital-sensor-C4E.pdf"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Conductivity Range": "0 – 200.0 µS/cm, 0 – 2,000 µS/cm, 0 – 20.00 mS/cm, 0 – 200.0 mS/cm (±1 % of full range)",
             "Salinity Range": "5 – 60 g/kg",
             "Temperature Range": "Working 0 – 50 °C (NTC compensation)",
             "Operating Voltage": "5 – 12 V; standby 25 µA, avg 6.3 mA (RS-485, 1 meas/s), pulse 500 mA",
             "Comm. Protocol / Output": "Modbus RS-485 (option SDI-12); refresh < 1 s",
             "Wiring": "Red = V+, Black = V−, Blue = SDI-12, Green = RS-485 B, White = RS-485 A, Green/yellow = shield (cable ≤ 15 m)",
             "Dimensions": "Ø27 × 157 mm (219 overall); 350 g with 3 m cable; IP68; 5 bar",
             "Buy Link": ("ictinternational.com", "https://ictinternational.com/product/awq-c4e-digital-conductivity-salinity-sensor/"),
             "Datasheet": ("ictinternational.com", "https://ictinternational.com/wp-content/uploads/2023/06/Datasheet-Conductivity-digital-sensor-C4E.pdf"),
             "Factory Calibrated?": "Yes (stored in sensor)",
             "How to Calibrate": "Sensor stores calibration and history data (\"plug and play\" without re-calibration). Procedure not given in datasheet.",
             "Notes": "Very low standby power — good for a battery logger. Price on request. Buy link returned 404 when checked."},
            {"Parameter": "Conductivity · Salinity", "Sensor (Part No.)": "ENV-40-EC-K10 + EZO-EC",
             "Manufacturer": ("Atlas Scientific ↗", "https://atlas-scientific.com/probes/conductivity-probe-k-10/"),
             "Unit Price (USD)": 267.98, "Seawater OK?": "Yes",
             "Conductivity Range": "Probe 10 µS/cm – 1 S/cm (±2 %); EZO-EC 0.07 – 500,000+ µS/cm (±2 %)",
             "Salinity Range": "0.00 – 42.00 PSU (EZO-EC)",
             "Temperature Range": "Probe 1 – 110 °C; no internal temp sensor",
             "Operating Voltage": "EZO-EC 3.3 – 5 V",
             "Comm. Protocol / Output": "UART / I2C (EZO-EC, default I2C 0x64, ASCII)",
             "Wiring": "Probe: male SMA (BNC optional)",
             "Dimensions": "Probe: 1 m cable, 200 PSI / 141 m; EZO-EC board 13.97 × 20.16 mm",
             "Buy Link": ("atlas-scientific.com", "https://atlas-scientific.com/probes/conductivity-probe-k-10/"),
             "Datasheet": ("atlas-scientific.com", "https://files.atlas-scientific.com/EC_K_10_probe.pdf"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "EZO-EC: 2- or 3-point calibration with temperature compensation. Probe: calibrate at installation; ~10 years before recalibration.",
             "Notes": "Price = probe $199.99 + EZO-EC circuit $67.99. EZO board must sit in your own housing; needs a separate temperature sensor."},
        ],
    },
    {
        "title": "2   ITEM 2 — CURRENT UNIT  (Water Current Speed · Direction)  ·  Priority: Essential",
        "columns": HEAD + ["Speed Range", "Speed Accuracy", "Direction Range / Accuracy", "Other Sensors",
                           "Operating Voltage", "Comm. Protocol / Output", "Dimensions"] + TAIL,
        "rows": [
            {"Parameter": "Current Speed · Direction (acoustic)", "Sensor (Part No.)": "OCEAN-SPCM",
             "Manufacturer": ("Oceantek ↗", "https://oceanadcp.com/products/single-point-current-meter/"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Speed Range": "±5.0 m/s (default)\n±20.0 m/s (max)",
             "Speed Accuracy": "±0.3 % ± 3 mm/s; resolution 0.1 mm/s",
             "Direction Range / Accuracy": "0 – 360°, ±2°; resolution 0.01°",
             "Other Sensors": "Temp -5 to 45 °C (±0.1 °C); pressure ±0.25 %; roll/pitch ±85°, heading 0–360°",
             "Operating Voltage": "20 – 50 V, 5 A (DR); 504 Wh Li battery (SC)",
             "Comm. Protocol / Output": "RS-232 or RS-422; PD0 binary, 1 Hz",
             "Dimensions": "Φ148 mm; DR 150–167 mm / SC 350–367 mm long; depth 1000/3000/6000 m",
             "Buy Link": ("oceanadcp.com", "https://oceanadcp.com/products/single-point-current-meter/"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Not stated", "How to Calibrate": UNK,
             "Notes": "Titanium housing. Quote on request."},
            {"Parameter": "Current Speed · Direction · Depth (rotor)", "Sensor (Part No.)": "EMCON Current Meter",
             "Manufacturer": ("EMCON India ↗", "https://emconindia.com/products/water-current-meter-with-direction-for-river-oceans-or-estuaries/"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Speed Range": "0 – 5 m/s",
             "Speed Accuracy": "0.02 m/s; threshold < 0.04 m/s",
             "Direction Range / Accuracy": "Unknown (Hall-effect sensor, no spec given)",
             "Other Sensors": "Depth (hydrostatic), 20 m, ±1 %",
             "Operating Voltage": "Unknown (handheld reader has internal battery)",
             "Comm. Protocol / Output": "Unknown (reads m/s on handheld LCD)",
             "Dimensions": "Unknown; 20 m cable (longer on demand)",
             "Buy Link": ("emconindia.com", "https://emconindia.com/products/water-current-meter-with-direction-for-river-oceans-or-estuaries/"),
             "Datasheet": (NA, None),
             "Factory Calibrated?": "Not stated", "How to Calibrate": UNK,
             "Notes": "Page says designed for marine use. Handheld readout, no logging stated."},
            {"Parameter": "Current Speed · Direction (tilt)", "Sensor (Part No.)": "DIY tilt current meter",
             "Manufacturer": ("Own build ↗", "https://hackaday.io/project/158083-cheap-underwater-tilt-current-sensor"),
             "Unit Price (USD)": 200.00, "Include in Cost?": "Yes", "Seawater OK?": "Yes",
             "Speed Range": "Unknown — set by own calibration (Hackaday build: ~10° tilt at ~0.2 kt, aimed at up to 1 kt)",
             "Speed Accuracy": UNK,
             "Direction Range / Accuracy": "Unknown — needs a tilt-compensated compass",
             "Other Sensors": "Accelerometer / IMU (Sensor Options sheet: Adafruit BNO085)",
             "Operating Voltage": "Battery (Hackaday build: 18650, ~1 month logging)",
             "Comm. Protocol / Output": "Own MCU + SD logger",
             "Dimensions": "Hackaday build: 2\" PVC housing, tested dry at 60 – 70 ft",
             "Buy Link": ("hackaday.io", "https://hackaday.io/project/158083-cheap-underwater-tilt-current-sensor"),
             "Datasheet": ("None (DIY)", None),
             "Factory Calibrated?": "No",
             "How to Calibrate": "Convert tilt to speed with your own calibration, e.g. tow at known GPS speeds (Sensor Options sheet). Keep steel away from the compass.",
             "Notes": "Price is an estimate (Sensor Options sheet). Accuracy and feasibility not yet known — validate against a TCM-1."},
            {"Parameter": "Current Speed · Direction · Temp (tilt)", "Sensor (Part No.)": "TCM-1",
             "Manufacturer": ("Lowell Instruments ↗", "https://hobodataloggers.com/products/data-loggers/tcm-1"),
             "Unit Price (USD)": 1295.00, "Seawater OK?": "Yes",
             "Speed Range": "0 – 40 cm/s (low range, ballast washer)\n0 – 80 cm/s (high range)",
             "Speed Accuracy": "Low: 2 cm/s + 3 % of reading; High: 3 cm/s + 3 % of reading; resolution 0.1 cm/s",
             "Direction Range / Accuracy": "0 – 360°, 5° (for speed > 5 cm/s); resolution 0.1°",
             "Other Sensors": "Temp -5 to 30 °C (0.1 °C); 3-axis accelerometer + magnetometer",
             "Operating Voltage": "Internal lithium battery (months to years, depends on rate)",
             "Comm. Protocol / Output": "Internal logger (8 GB microSD); USB 2.0 micro-B offload; 64 Hz – 1/h sampling",
             "Dimensions": "Ø2.7 × 73 cm; 340 g; PVC housing; 300 m depth",
             "Buy Link": ("hobodataloggers.com", "https://hobodataloggers.com/products/data-loggers/tcm-1"),
             "Datasheet": ("User guide (PDF)", "https://dev-d9.onsetcomp.com/sites/default/files/2025-07/Universal_User_Guide_Lowell_CurrentMeters_Loggers.pdf"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Not given on page (uses a calibrated ballast weight and calibration coefficients).",
             "Notes": "Ready-made; also a good reference to validate the DIY build. USD 1,295 excl. tax (MSRP)."},
            {"Parameter": "Current Speed · Direction (drag-tilt)", "Sensor (Part No.)": "Marotte HS",
             "Manufacturer": ("JCU Marine Geophysics Lab ↗", "https://www.jcu.edu.au/marine-geophysics-laboratory"),
             "Unit Price (USD)": UNK, "Seawater OK?": "Yes",
             "Speed Range": UNK, "Speed Accuracy": UNK,
             "Direction Range / Accuracy": "Unknown (magnetometer gives direction without knowing deployment orientation)",
             "Other Sensors": UNK, "Operating Voltage": UNK, "Comm. Protocol / Output": "Internal logger",
             "Dimensions": UNK,
             "Buy Link": ("jcu.edu.au", "https://www.jcu.edu.au/marine-geophysics-laboratory"),
             "Datasheet": ("Not available", None),
             "Factory Calibrated?": "Not stated", "How to Calibrate": UNK,
             "Notes": "JCU page could not be loaded to check specs; description from Sensor Options sheet. Price on request from JCU."},
        ],
    },
    {
        "title": "3   ITEM 3 — TEMPERATURE & LIGHT UNIT  (Temperature · Light)  ·  Priority: Essential",
        "columns": HEAD + ["Temperature Range / Accuracy", "Light Range / Accuracy", "Logging",
                           "Power / Battery", "Comm. / Data Offload", "Depth Rating", "Dimensions"] + TAIL,
        "rows": [
            {"Parameter": "Temperature · Light", "Sensor (Part No.)": "HOBO Pendant MX2202",
             "Manufacturer": ("Onset (HOBO) ↗", "https://www.instrumart.com/products/44580/hobo-mx2200-series-data-loggers"),
             "Unit Price (USD)": 100.00, "Include in Cost?": "Yes", "Seawater OK?": "Yes",
             "Temperature Range / Accuracy": "-20 to 50 °C in water (-20 to 70 °C in air); ±0.5 °C; resolution 0.04 °C",
             "Light Range / Accuracy": "0 – 167,731 lux; ±10 % typical (direct sunlight)",
             "Logging": "1 s – 18 h interval; 96,000 measurements",
             "Power / Battery": "CR2032 3 V lithium, user-replaceable; ~1 yr at 1-min logging (2 yr with Bluetooth Always Off)",
             "Comm. / Data Offload": "Bluetooth Low Energy (~30.5 m line of sight); HOBOmobile app",
             "Depth Rating": "Waterproof to 30.5 m (100 ft)",
             "Dimensions": "3.35 × 5.64 × 1.6 cm; 12.75 g; polypropylene case",
             "Buy Link": ("instrumart.com", "https://www.instrumart.com/products/44580/hobo-mx2200-series-data-loggers"),
             "Datasheet": ("instrumart.com", "https://www.instrumart.com/assets/MX2201-MX2202-Datasheet.pdf"),
             "Factory Calibrated?": "Yes (light sensor)",
             "How to Calibrate": "Light sensor factory-calibrated for enclosure attenuation. Temperature calibration procedure not given.",
             "Notes": "Manual says it works in fresh and salt water. Price is an estimate (Sensor Options sheet: ~$96 – 109); Instrumart lists the series from $90 but no MX2202 price."},
        ],
    },
    {
        "title": "4   ITEM 4 — SEDIMENTATION UNIT  (Sediment · Turbidity)  ·  Priority: Possible addition",
        "columns": HEAD + ["Measuring Range", "Accuracy / Resolution", "Response Time", "Operating Voltage",
                           "Comm. Protocol / Output", "Wiring", "Depth / Protection"] + TAIL,
        "rows": [
            {"Parameter": "Sediment accumulation", "Sensor (Part No.)": "DIY sediment traps",
             "Manufacturer": ("Own build", None),
             "Unit Price (USD)": 50.00, "Include in Cost?": "Yes", "Seawater OK?": "Yes",
             "Measuring Range": "Sediment mass per period (weighed in lab)",
             "Accuracy / Resolution": "n/a", "Response Time": "n/a", "Operating Voltage": "None",
             "Comm. Protocol / Output": "None", "Wiring": "n/a", "Depth / Protection": "n/a",
             "Buy Link": ("None (DIY)", None), "Datasheet": ("None (DIY)", None),
             "Factory Calibrated?": "n/a", "How to Calibrate": "n/a",
             "Notes": "Open PVC tubes on a frame; divers swap them each visit. Price is an estimate (Sensor Options sheet)."},
            {"Parameter": "Turbidity · Temp", "Sensor (Part No.)": "SEN0710",
             "Manufacturer": ("DFRobot ↗", "https://www.dfrobot.com/product-3104.html"),
             "Unit Price (USD)": 229.00, "Seawater OK?": "Check",
             "Measuring Range": "0 – 1000 NTU; temp 0 – 60 °C",
             "Accuracy / Resolution": "±5 % FS (25 °C), 0.1 NTU; temp ±0.5 °C, 0.1 °C",
             "Response Time": "≤ 30 s",
             "Operating Voltage": "10 – 30 V DC; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus-RTU), 4800 default",
             "Wiring": "Brown = VCC, Black = GND, Yellow = 485-A, Blue = 485-B",
             "Depth / Protection": "0.6 MPa; IP68; operating 0 – 40 °C; 5 m cable",
             "Buy Link": ("dfrobot.com", "https://www.dfrobot.com/product-3104.html"),
             "Datasheet": ("wiki.dfrobot.com", "https://wiki.dfrobot.com/sen0710/"),
             "Factory Calibrated?": "Yes",
             "How to Calibrate": "Two-point via registers: deionized water (0 NTU) and 1000 NTU standard; recalibrate every 90 days. Remove the black rubber cover before deployment.",
             "Notes": "90° light scattering. Seawater use not stated (marketed for drinking water, wastewater, rivers). Optical window fouls — plan cleaning or a wiper. USD 210 at 10+."},
        ],
    },
    {
        "title": "5   ITEM 5 — POLLUTION UNIT  (Dissolved Oxygen)  ·  Priority: If available",
        "columns": HEAD + ["Measuring Range", "Accuracy / Resolution", "Response Time", "Operating Voltage",
                           "Comm. Protocol / Output", "Wiring", "Depth / Protection"] + TAIL,
        "rows": [
            {"Parameter": "Dissolved Oxygen · Temp", "Sensor (Part No.)": "SEN0681",
             "Manufacturer": ("DFRobot ↗", "https://www.dfrobot.com/product-3009.html"),
             "Unit Price (USD)": 209.00, "Seawater OK?": "Yes",
             "Measuring Range": "0 – 20 mg/L; 0 – 200 % saturation",
             "Accuracy / Resolution": "±3 % FS; 0.01 mg/L, 0.1 %; temp ±0.5 °C, 0.1 °C",
             "Response Time": "≤ 60 s",
             "Operating Voltage": "10 – 30 V DC; 0.2 W",
             "Comm. Protocol / Output": "RS-485 (Modbus-RTU), addr 1",
             "Wiring": "Brown = VCC, Black = GND, Yellow = 485-A, Blue = 485-B",
             "Depth / Protection": "0.6 MPa; IP68; operating 0 – 40 °C; 5 m cable",
             "Buy Link": ("dfrobot.com", "https://www.dfrobot.com/product-3009.html"),
             "Datasheet": ("wiki.dfrobot.com", "https://wiki.dfrobot.com/sen0681/"),
             "Factory Calibrated?": "Not stated",
             "How to Calibrate": "Not given on product page or wiki summary (wiki Reference page has calibration instructions). Salinity setting defaults to 30 ‰ for seawater.",
             "Notes": "Fluorescence method, no electrolyte; film lasts ~1 year. Seawater version. A drop in oxygen with a turbidity spike can flag sewage reaching the site (Sensor Options sheet). USD 193 at 10+."},
        ],
    },
    {
        "title": "6   ITEM 6 — SUPPORTING  (Housing for logger & battery)",
        "columns": HEAD + ["Size", "Materials", "Depth Rating", "End Caps / Penetrators", "What's Included"]
                   + ["Buy Link", "Datasheet", "Notes"],
        "rows": [
            {"Parameter": "Watertight enclosure", "Sensor (Part No.)": "WTE-VP (2\")",
             "Manufacturer": ("Blue Robotics ↗", "https://bluerobotics.com/store/watertight-enclosures/wte-vp/"),
             "Unit Price (USD)": 61.00, "Include in Cost?": "Yes", "Seawater OK?": "Yes",
             "Size": "2\": aluminium ID 50.0 / OD 58.0 mm; lengths 100, 150, 300 mm",
             "Materials": "Aluminium 6061-T6 (type III anodized) or cast acrylic tube",
             "Depth Rating": "Aluminium tube 1000 m; acrylic 225 / 130 / 100 m (100 / 150 / 300 mm), short-term < 2 weeks",
             "End Caps / Penetrators": "2\" flange caps with 0, 2 or 4 M10 holes",
             "What's Included": "Tube, 2 caps + flanges, vent & plug or relief valve, O-ring pick; penetrators not included",
             "Buy Link": ("bluerobotics.com", "https://bluerobotics.com/store/watertight-enclosures/wte-vp/"),
             "Datasheet": ("bluerobotics.com", "https://bluerobotics.com/wp-content/uploads/2026/05/WTE-DATASHEET-RevC.5-MAY2026.pdf"),
             "Notes": "Price is the page's \"From $61.00\" (configuration not stated). Holds the RS-485 logger and the DIY tilt-meter electronics."},
        ],
    },
]

NOTES = [
    "Notes",
    "• Spec values come from the manufacturer's product page or datasheet linked in each row; anything not stated there is \"Unknown\".",
    "• Estimated prices and DIY entries come from the \"Underwater Site Monitoring – Sensor Options\" sheet and are marked in each row's Notes.",
    "• Sensors within an item are alternatives. \"Include in Cost?\" = Yes marks the starter set; subtotals only count Yes rows with a numeric price.",
    "• \"Seawater OK?\" compares the stated range with seawater (~50,000 – 60,000 µS/cm). \"Check\" = seawater use not stated by the maker.",
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

# Preferred width per column name; a sheet column gets the widest preference among items.
PREF = {"#": 5, "Parameter": 24, "Sensor (Part No.)": 18, "Manufacturer": 17, "Unit Price (USD)": 12,
        "Include in Cost?": 11, "Seawater OK?": 11, "Buy Link": 17, "Datasheet": 20,
        "Factory Calibrated?": 13, "How to Calibrate": 46, "Notes": 46}
LEFT_COLS = {"Parameter", "Sensor (Part No.)", "Manufacturer", "How to Calibrate", "Wiring", "Other Sensors", "Notes",
             "What's Included"}
NCOL = max(len(i["columns"]) for i in ITEMS)
WIDTHS = [max(PREF.get(item["columns"][c], 24) for item in ITEMS if c < len(item["columns"])) for c in range(NCOL)]
VERDICT = {"Yes": "1E7B34", "No": "C0392B", "Check": "B26A00"}


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
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
                if name == "Seawater OK?":
                    cell.font = F(bold=True, color=VERDICT.get(val, TXT))
                if name == "Unit Price (USD)" and isinstance(val, (int, float)):
                    cell.number_format = USD
                if str(val).startswith(UNK) or val in (NA, "Not stated", "n/a", "Not available"):
                    cell.font = F(italic=True, color="7F7F7F")
                if target:
                    cell.hyperlink = target
                    cell.font = LINK
                if name == "Include in Cost?":
                    dv.add(cell)
                csv_line.append(f"{val} ({target})" if target else val)
                longest = max(longest, len(str(val)) / max(WIDTHS[c - 1], 1))
            ws.row_dimensions[r].height = max(30, min(120, 13 * (longest + 1.2)))
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
