"""Build Site_Water_Parameters.xlsx: measured seawater ranges at the four ALEX-MON sites.

Every site value comes from the published source codes in [brackets] (see the
Sources sheet). The "Suggested sensor range" column is derived from those
measurements plus a safety margin; it is not a measured value.
Run:  python build_water_data.py
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
NF = "Not found"

SITES = ["Qaitbay · Pharos\n(6–8 m, 0–150 m from citadel)",
         "Eastern Harbour · Antirhodos\n(2–8 m, ≈200 m)",
         "Canopus East\n(4–5 m, ≈2 km off Abu Qir)",
         "Thonis-Heracleion\n(5–7 m, ≈6–7 km in Abu Qir Bay)"]

# (parameter, unit, qaitbay, eastern harbour, canopus, heracleion, region reference, suggested range, notes)
ROWS = [
    ("Water temperature", "°C",
     "17.1 – 31.5\n(min Feb, max Aug 2009; annual mean 24.9) [R1]",
     "16.0 – 31.5\nsurface 16.0–30.0, bottom 17.0–28.5; coldest Jan 1986 = 16.0 [R1][R4][R5][R6]",
     "15.5 – 29.5\nJan 1990 15.5–16.9; Jul 1990 29.5 near coast [A5]\nwinter mean 16.1, summer 26.4 (2009–10) [A1]\nSW bay up to 31.5 [A4]",
     "15.5 – 29.5\nJan 1990 15.5–16.9; Jul 1990 27.0 outside bay [A5]\nW-headland station: 16.3 winter, 26.3 summer [A1]",
     "Monthly SST 16.5 (Feb min) – 28.3 (Aug max) [C1]\nWarming +0.4 °C/decade (1948–2018) [C2]",
     "10 – 35 °C",
     "Never near 0 °C: coldest measured anywhere was 15.5 °C (Jan). Stratification only in top 5 m in Apr and Aug [R5]."),
    ("Salinity", "PSU (‰)",
     "35.2 – 38.9\nmean 37.3–37.4 (2009) [R1]",
     "34.8 – 39.5\nsurface means 37.2 (2009), 38.2 (2022); bottom up to 39.5 [R1][R3][R4][R6]",
     "37.4 – 40.5 at W-headland station (39.9 summer, 40.5 winter) [A1]\nInner SW bay drops to 28.7 [A1], worst 22.6 [A2] near El-Tabia / El-Maadia outlets",
     "37.5 – 40.5\nJan 1990 western bay 37.5; Nov max 39.9 [A5]\nsalinity rises offshore [A1]",
     "Open sea 38.6 – 39.3 [C3]\nEl-Mex open-sea max 40.09 [C4]",
     "20 – 42 PSU\n(≈ 30 – 63 mS/cm conductivity, calculated)",
     "Lower salinity marks sewage / lake-water inflow. Abu Qir values swing most because of El-Maadia and El-Tabia discharges."),
    ("pH", "–",
     "7.6 – 8.3\nmean 7.9–7.95 (2009) [R1]",
     "7.03 – 8.58\nrecent: 7.8–8.1 (2009), 8.16 (2022) [R1][R3][R4][R6]",
     "7.34 – 8.07 (2009–10) [A1]\nSW bay 7.57–8.26 [A2]",
     "7.34 – 8.07 (bay-wide, 2009–10) [A1]\nW-headland station: 7.98 summer, 7.40 winter",
     "Open coast 8.09 – 8.25 (El-Mex 2008) [C4]",
     "6.5 – 9.0",
     ""),
    ("Dissolved oxygen", "mg/L",
     "3.5 – 10.1\nmean 5.4–5.5 (2009); saturation mostly 36–64 % [R1]",
     "Means 5.3–6.9 (2009), 8.0 (2022) [R1][R3]\nNear-bottom summer down to 0.1 ml/L (≈0.14 mg/L, 1985–86) [R5]",
     "2.15 – 10.49 ml/L (≈ 3.1 – 15.0 mg/L) [A1]\nSW bay 0.6 – 9.9 mg/L [A4]",
     "2.15 – 10.49 ml/L (≈ 3.1 – 15.0 mg/L, bay-wide) [A1]\nW-headland: 4.21–5.59 ml/L",
     "El-Mex open sea surface up to 8.45 mg/L [C4]",
     "0 – 20 mg/L",
     "ml/L converted with ×1.429. Summer bottom hypoxia recorded in the harbour."),
    ("Transparency (Secchi depth)", "m",
     NF,
     "0.5 – 4.0\n2018–19 mean 2.1; 96 % of readings 1–3 m (\"medium turbidity\") [R2][R4][R6]",
     "0.3 – 4.0\nSW-bay mean 1.24 (1998–99) [A3][A4]",
     NF + " (only SW-bay data)",
     "",
     "Turbidity 0 – 1000 NTU",
     "No study reports turbidity in NTU or TSS in mg/L for any of the four sites."),
    ("Chlorophyll-a", "µg/L",
     "annual average 2.0 (2009) [R1]",
     "avg 2.4 (2009); summer peaks up to 16.75 (2018–19) [R1][R2]",
     "2.06 – 52.64 monthly averages, SW bay (1998–99) [A3]",
     NF + " (only SW-bay data)",
     "",
     "—",
     "High values = eutrophication from sewage; relevant to biological growth on the artefacts."),
    ("Water current speed", "cm/s",
     "Not measured at the site.\nHarbour mouth area: 0 – 43.2 (1996) [R7]",
     "0 – 43.2 (1996) [R7]\n3 – 70 measured near surface and bottom [C8]\nmodel mean 8.6 [C8]",
     "Inner bay up to 50 (NW, March); outer bay 5 [A1]\nmodel surface 10 – 15 [A5]",
     "Outside the bay 15 – 17 toward NE (1969) [A5]\nouter bay 5 [A1]",
     "Offshore geostrophic current 2 – 5, eastward [C9]",
     "0 – 100 cm/s (0 – 1 m/s)",
     "Max measured anywhere: 70 cm/s (Eastern Harbour)."),
    ("Current direction", "°",
     "Not measured at the site",
     "In through the Diamond pass (W), out through El-Silsila (E); main component toward NE, driven by NW wind [C8][R7]",
     "Wind-driven; NW in inner bay (March) [A1][A5]",
     "NE outside the bay [A5]",
     "Prevailing wind NW (≈ 319°) [C10]; occasional reversal with NE waves [C9]",
     "0 – 360°",
     "Project guide: current reverses at times when winds come from the east (guide p.5)."),
    ("Wave height", "m",
     "Offshore max Hs 4.5 from NW; 85 % of waves ≤ 0.6 [R8]",
     "Measured max 0.75 – 1.23 inside harbour [C8]\nmodel up to ≈ 2 in NW storm [R8]",
     "Mean 0.94 (T 6.5 s), 42 % from NW; max 4.19 (Nov 1986) [A6]",
     "Mean 0.94, max 4.19; gales ≈ 4 m [A6]",
     "1-yr return ≈ 4 m, 100-yr ≈ 8 m offshore [R8]",
     "—",
     "Winter storms (Nov–Mar) are the design case for mounting the device."),
    ("Tidal range", "cm",
     "Spring 16 – 23, neap 4 (Alexandria gauge) [C6][C7]",
     "Spring 16 – 23, neap 4; total sea-level range incl. surge ≈ 53 [C6][C7][R7]",
     "Spring 6, neap 2 (Abu Qir gauge) [C6]",
     "Spring 6, neap 2 (Abu Qir gauge) [C6]",
     "Mixed, mainly semidiurnal [C6]",
     "—",
     "Small tides: depth (pressure) sensors see mostly wind surge, not tide."),
    ("Sea-level rise / subsidence", "mm/yr",
     "+2.2 (Alexandria, 1974–2006) [C6]",
     "+2.2 (Alexandria, 1974–2006) [C6]",
     "+6.4 (Abu Qir, 1990–2010) [C6]; land subsidence 1 – 2.5 [A7]",
     "+6.4 (Abu Qir, 1990–2010) [C6]; land subsidence 1 – 2.5 [A7]",
     "Egyptian coast average 3.4 [C6]",
     "—",
     ""),
    ("Sedimentation rate", "cm/yr",
     NF,
     "≈ 6, only in the sink area downstream of a sewage outfall (1996–99) [R9]",
     NF,
     NF,
     "",
     "—",
     "Project guide says sedimentation increased after the breakwater was built; no measured value exists for the site."),
    ("BOD (organic pollution)", "mg/L",
     NF,
     "means 1.3 – 1.9 ml O₂/L (1990–91) [R4]",
     "1.98 – 3.71 (2009–10) [A1]",
     "1.98 – 3.71 (bay-wide, 2009–10) [A1]",
     "",
     "—",
     "Fecal coliform counts not found for any site."),
]

MONTHLY = [  # month, min, avg, max  [C1]
    ("Jan", 16.8, 18.2, 19.7), ("Feb", 16.5, 17.3, 18.2), ("Mar", 16.6, 17.4, 18.2),
    ("Apr", 17.0, 18.1, 19.3), ("May", 18.7, 20.6, 22.6), ("Jun", 22.1, 23.9, 25.7),
    ("Jul", 24.4, 26.3, 28.2), ("Aug", 26.5, 27.4, 28.3), ("Sep", 26.1, 27.0, 28.0),
    ("Oct", 23.5, 25.3, 27.2), ("Nov", 20.3, 22.3, 24.3), ("Dec", 18.4, 20.1, 21.8),
]

GAPS = [
    ("All sites", "Turbidity in NTU and total suspended solids (mg/L)", "Only Secchi depth is published."),
    ("All sites", "Fecal coliform counts", "EEAA reports mention exceedances but give no accessible numbers."),
    ("Qaitbay · Pharos", "Any measurement at the site itself", "Nearest data: Kayet Bey stations St-1/St-2 at the fort, next to the Al-Anfoushi sewage outfall [R1]."),
    ("Qaitbay · Pharos", "Current speed and direction", "Only Eastern Harbour values exist."),
    ("Canopus / Heracleion", "Station data at the sites", "Closest is El-Gohary 2011 Station 1 off the Abu Qir headland [A1]."),
    ("Canopus / Heracleion", "Sedimentation rate", "—"),
    ("All sites", "Underwater visibility measured by divers", "Only travel / dive blogs; not used."),
    ("Note", "Site distances", "Brocard et al. 2024 place Canopus ≈1.5 km and Heracleion ≈4 km offshore at 4–6 m [A7], closer than the slide values."),
]

SOURCES = [
    ("R1", "Khairy, Hussein, Faragallah & Dorgham (2014). Phytoplankton communities in two eutrophic areas on the Alexandria coast. Rev. Biol. Mar. Oceanogr. 49(2):267–277. Monthly 2009, Eastern Harbour + Kayet Bey.", "https://www.redalyc.org/pdf/479/47931776007.pdf"),
    ("R2", "Labib et al. (2023). Water quality indices as tools for assessment of the Eastern Harbor's water status. SN Applied Sciences 5:89. Weekly 2018–19.", "https://doi.org/10.1007/s42452-023-05304-z"),
    ("R3", "Ismail, Diab & ElSheekh (2023). Trophic status determination of the Egyptian Eastern Mediterranean. Environ. Monit. Assess. 195:1040. Early summer 2022.", "https://doi.org/10.1007/s10661-023-11690-z"),
    ("R4", "Abdalla, Zaghloul & Hussein (1995). Statistical modelling of phytoplankton eutrophication in the Eastern Harbour. Bull. NIOF 21(1):125–146. 1990–91.", "https://niof-eg.com/wp-content/uploads/2022/07/A-STATISTICAL-MODELLING-OF-PHYTOPLANKTON.pdf"),
    ("R5", "Maiyza & Said (1989). Time variation of temperature, salinity and oxygen in the Eastern Harbour. Bull. NIOF 15(2):39–45. 1985–86.", "https://niof-eg.com/wp-content/uploads/2022/06/THE-TIME-VARIATION-OF-TEMPERATURE-SALINITY-AND-OXYGEN-CONTENT-IN-THE-EASTERN-HARBOUR-OF.-ALEXANDRIA.pdf"),
    ("R6", "Dowidar, El-Nady, Aboul-Kassim & Abdel-Moati (1987). Impact of sewage discharge on the Eastern Harbour. Bull. NIOF 13(2):121–132. 1985–86.", "https://niof-eg.com/wp-content/uploads/2022/06/Impact-of-sewage-discharge-on-the-water-quality-of-the-Eastern.pdf"),
    ("R7", "Khidr, Abd al-Rahman & Alam El-Din (2018). Currents and sea level variability of Alexandria coast. J. King Abdulaziz Univ. Marine Sci. 28(2):27–42. (abstract)", "https://search.emarefa.net/en/detail/BIM-896826-currents-and-sea-level-variability-of-alexandria-coast-in-as"),
    ("R8", "Abdellah (2006). Simulation of wave propagation in the Eastern Harbor of Alexandria. Egyptian J. Aquatic Research 32(1):48–60. (model)", "https://niof-eg.com/wp-content/uploads/2022/07/SIMULATION-OF-WAVE-PROPAGATION-IN-THE-EASTERN.pdf"),
    ("R9", "Ismael, Khadr & Halim (2001). Dinoflagellate cysts from sediments of the Eastern Harbour. Rapp. Comm. int. Mer Médit. 36:392.", "https://ciesm.org/online/archives/abstracts/pdf/36/CIESM_Congress_2001_Monaco_article_0392.htm"),
    ("A1", "El-Gohary, El-Naggar & Zaki (2011). Water quality along the coastal zone of Abu-Qir Bay. Egyptian J. Aquatic Research 37(3):209–219. Summer 2009, winter 2010.", "https://niof-eg.com/wp-content/uploads/2022/07/Study-on-water-quality-along-the-coastal-zone-of-Abu-Qir-Bay-Mediterranean-Sea-Alexandria-Egypt.pdf"),
    ("A2", "Nessim, Masoud & Maximous (2005). Water characteristics of Alexandria hot spots. Egyptian J. Aquatic Research 31:25–37. 2003–04.", "https://niof-eg.com/wp-content/uploads/2022/07/WATER-CHARACTERISTICS-OF-ALEXANDRIA-HOT-SPOTS.pdf"),
    ("A3", "Abdel-Aziz (2001). Hydrography, nutrients and plankton in the hot spot of Abu Qir Bay. Mediterranean Marine Science 2:17–32. 1998–99.", "https://epublishing.ekt.gr/en/8021/Mediterranean-Marine-Science/8617"),
    ("A4", "Shams-El-Din & Dorgham (2007). Phytoplankton community in Abu Qir Bay as a hot spot. Egyptian J. Aquatic Research 33(1):163–182. 1998–99.", "https://niof-eg.com/wp-content/uploads/2022/07/PHYTOPLANKTON-COMMUNITY-IN-ABU-QIR-BAY-AS-A-HOT.pdf"),
    ("A5", "Alam El-Din & Al-Hogaraty (2001). Modeling the seasonal variability of density current in Abu-Qir Bay. Bull. NIOF 27:275–292. 1990 survey.", "https://niof-eg.com/wp-content/uploads/2022/07/MODELING-THE-SEASONAL-VARIABILITY-OF-DENSITY.pdf"),
    ("A6", "Abdallah, Sharaf El-Din & Shereet (2006). Analysis of wave observations and wave transformations in Abu-Qir Bay. Egyptian J. Aquatic Research 32(1):22–33. 1986 data.", "https://niof-eg.com/wp-content/uploads/2022/07/ANALYSIS-OF-WAVE-OBSERVATIONS-AND-WAVE.pdf"),
    ("A7", "Brocard, Goiran & Robinson (2024). Fate of two cities built on sinking ground: Thonis-Heracleion and Canopus. Book chapter.", "https://ancientportsantiques.com/wp-content/uploads/Documents/PLACES/Egypt-Libya/NileDelta/CanopeThonis-Brocard2024.pdf"),
    ("C1", "Seatemperature.org — Alexandria monthly sea temperature (NOAA satellite data, multi-year).", "https://www.seatemperature.org/africa/egypt/alexandria.htm"),
    ("C2", "El-Geziry (2021). Long-term changes in SST within the southern Levantine Basin. Acta Oceanologica Sinica 40(3):27–33. 1948–2018.", "https://castjournals.cast.org.cn/joweb/aos/EN/PDF/10.1007/s13131-021-1709-2"),
    ("C3", "Said, Maiyza, Hussain & Radwan (2007). Acta Adriatica 48(2):145–159. Winter 1989/1990 cruises.", "https://hrcak.srce.hr/file/29268"),
    ("C4", "Okbah et al. (2012). El-Mex Bay water quality. Environ. Monit. Assess. 185:3103–3115. 2008.", "https://pmc.ncbi.nlm.nih.gov/articles/PMC3586065/"),
    ("C6", "El-Geziry & Said (2020). Spatial variations of sea level along the Egyptian Mediterranean coast. Athens J. Mediterranean Studies 6(2):141–154.", "https://www.athensjournals.gr/mediterranean/2020-6-2-3-El-Geziry.pdf"),
    ("C7", "Moursy (1992). Tidal levels in Alexandria, Egypt. Bull. NIOF 18:1–9. 1976.", "https://niof-eg.com/wp-content/uploads/2022/06/TIDAL-LEVELS-IN-ALEXANDRIA-EGYPT.pdf"),
    ("C8", "Millet & Goiran (2007). Impacts of Alexandria's Heptastadion on coastal hydro-sedimentary dynamics. Int. J. Nautical Archaeology 36(1):167–176 (cites Lotfy & Badr 1999 measurements).", "https://ancientportsantiques.com/wp-content/uploads/Documents/PLACES/Egypt-Libya/Alexandria/Alexandria-MilletGoiran2007.pdf"),
    ("C9", "Kamel (1999). Seasonal variability of surface current off the Nile Delta. Bull. NIOF 25:51–67.", "https://niof-eg.com/wp-content/uploads/2022/07/SEASONAL-VARIABILITY-OF-SURFACE-CURRENT.pdf"),
    ("C10", "Windfinder — Alexandria wind statistics (2009–2026).", "https://www.windfinder.com/windstatistics/alexandria"),
]

# ---------------------------------------------------------------- styles (match Sensors_BOM)
NAVY, BLUE, PALE, ROWALT, BORDER, TXT = "22449A", "2589CE", "BCD3EA", "EAF3FB", "BCD3EA", "1B2B4B"
thin = Side(style="thin", color=BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", fgColor=c)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")


def title(ws, text, sub):
    ws.sheet_view.showGridLines = False
    ws["B2"] = text
    ws["B2"].font = Font(name="Arial", size=18, bold=True, color=NAVY)
    ws["B3"] = sub
    ws["B3"].font = Font(name="Arial", size=10, bold=True, color="7F93AD")


def header(ws, r, names, c0=2):
    for i, n in enumerate(names):
        c = ws.cell(r, c0 + i, n)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill, c.alignment, c.border = fill(NAVY), CENTER, BOX


def body(ws, r, values, alt, c0=2, bold_first=True):
    longest = 0
    for i, v in enumerate(values):
        c = ws.cell(r, c0 + i, v)
        c.fill = fill("FFFFFF" if not alt else ROWALT)
        c.border, c.alignment = BOX, WRAP
        c.font = Font(name="Arial", size=9, bold=(i == 0 and bold_first), color=TXT)
        if isinstance(v, str) and v.startswith(NF):
            c.font = Font(name="Arial", size=9, italic=True, color="7F7F7F")
        width = ws.column_dimensions[get_column_letter(c0 + i)].width or 12
        lines = sum(max(1, -(-len(s) // int(width * 0.95))) for s in str(v).split("\n"))
        longest = max(longest, lines)
    ws.row_dimensions[r].height = max(18, 13 * longest + 6)


def main():
    wb = Workbook()

    # Sheet 1: site ranges
    ws = wb.active
    ws.title = "Site ranges"
    title(ws, "ALEX-MON — Measured water parameters at the four sites",
          "SSTM  |  Simply Smart   ·   values from published studies; codes in [ ] link to the Sources sheet")
    cols = ["Parameter", "Unit"] + SITES + ["Alexandria region\n(reference)", "Suggested sensor range\n(derived, not measured)", "Notes"]
    widths = [20, 9, 30, 34, 34, 32, 28, 18, 34]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(2 + i)].width = w
    ws.column_dimensions["A"].width = 2
    header(ws, 5, cols)
    ws.row_dimensions[5].height = 56
    for k, row in enumerate(ROWS):
        body(ws, 6 + k, row, k % 2)
        sug = ws.cell(6 + k, 9)
        sug.font = Font(name="Arial", size=9, bold=True, color="1E7B34")
    note_r = 6 + len(ROWS) + 1
    for k, t in enumerate([
        "How to read: each site cell gives the measured min – max, then the detail and the source code(s).",
        "Canopus and Heracleion share Abu Qir Bay data: no study sampled exactly at either site. Station 1 off the Abu Qir headland [A1] is the closest.",
        "Suggested sensor range = widest measured value plus a safety margin; use it to check sensor specs in Sensors_BOM.xlsx.",
        "Conductivity in the salinity row is calculated from salinity (PSS-78, 25 °C), not measured.",
    ]):
        c = ws.cell(note_r + k, 2, t)
        c.font = Font(name="Arial", size=9, italic=(k > 0), bold=(k == 0), color=TXT)
    ws.freeze_panes = "D6"

    # Sheet 2: monthly temperature
    m = wb.create_sheet("Monthly temperature")
    title(m, "Alexandria monthly sea-surface temperature", "Source [C1] seatemperature.org (NOAA satellite, multi-year)")
    for col, w in zip("BCDE", (10, 12, 12, 12)):
        m.column_dimensions[col].width = w
    header(m, 5, ["Month", "Min °C", "Average °C", "Max °C"])
    for k, row in enumerate(MONTHLY):
        body(m, 6 + k, row, k % 2)
        for j in range(3, 6):
            m.cell(6 + k, j).alignment = CENTER
            m.cell(6 + k, j).number_format = "0.0"
    r = 6 + len(MONTHLY) + 1
    for t in ["Coldest month: February (min 16.5 °C). Warmest: August (max 28.3 °C).",
              "Measured in the water at the sites: lowest 15.5 °C (Abu Qir, Jan 1990) [A5], 16.0 °C (Eastern Harbour, Jan 1986) [R5];",
              "highest 31.5 °C (Eastern Harbour / Kayet Bey, Aug 2009) [R1]. Basin warming ≈ +0.4 °C per decade [C2]."]:
        m.cell(r, 2, t).font = Font(name="Arial", size=9, color=TXT)
        r += 1

    # Sheet 3: data gaps
    g = wb.create_sheet("Data gaps")
    title(g, "What the research could not find", "Measure these with the ALEX-MON device")
    for col, w in zip("BCD", (22, 42, 70)):
        g.column_dimensions[col].width = w
    header(g, 5, ["Site", "Missing data", "Comment"])
    for k, row in enumerate(GAPS):
        body(g, 6 + k, row, k % 2)

    # Sheet 4: sources
    s = wb.create_sheet("Sources")
    title(s, "Sources", "R = Eastern Harbour / Qaitbay · A = Abu Qir Bay · C = Alexandria region")
    for col, w in zip("BCD", (7, 95, 60)):
        s.column_dimensions[col].width = w
    header(s, 5, ["Code", "Reference", "Link"])
    for k, row in enumerate(SOURCES):
        body(s, 6 + k, row, k % 2)
        link = s.cell(6 + k, 4)
        link.hyperlink = row[2]
        link.font = Font(name="Arial", size=9, color="0563C1", underline="single")

    for sheet in wb.worksheets:
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    out = HERE / "Site_Water_Parameters.xlsx"
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
