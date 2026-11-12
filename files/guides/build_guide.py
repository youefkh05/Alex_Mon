"""Build the ALEX-MON Selected Sensors Guide (HTML -> PDF) in the style of Water_Current_Sensors_Guide.pdf.

Specs and prices come from Sensors_BOM.xlsx / Under_Water_BoM.xlsx (maker datasheets and team supplier
prices); site values come from files/site_water_data/Site_Water_Parameters.xlsx.
Run:  python build_guide.py      (needs Google Chrome for the PDF step)
"""
import html
import subprocess
from pathlib import Path

HERE = Path(__file__).parent
OUT_HTML = HERE / "ALEX-MON_Selected_Sensors_Guide.html"
OUT_PDF = HERE / "ALEX-MON_Selected_Sensors_Guide.pdf"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DATE = "نوفمبر 2026"
RUNNING = "دليل الحساسات المختارة لمشروع ALEX-MON"

e = html.escape


def ltr(s):
    """Isolate Latin/number runs inside Arabic text."""
    return f'<span class="ltr">{e(s)}</span>'


# ---------------------------------------------------------------- site water data (short)
SITES = [
    ("قايتباي · فاروس", "Qaitbay · Pharos", "6–8 م", "0–150 م من سور القلعة"),
    ("الميناء الشرقي · أنتيرودوس", "Eastern Harbour", "2–8 م", "≈ 200 م من الكورنيش"),
    ("كانوب الشرقية", "Canopus East", "4–5 م", "≈ 2 كم من شاطئ أبو قير"),
    ("ثونيس-هيراكليون", "Thonis-Heracleion", "5–7 م", "≈ 6–7 كم داخل خليج أبو قير"),
]
WATER = [  # parameter, unit, 4 site values
    ("درجة الحرارة", "°C", "17 – 31.5", "16 – 31.5", "15.5 – 29.5", "15.5 – 29.5"),
    ("الملوحة", "PSU", "35 – 39", "35 – 39.5", "37 – 40.5", "37.5 – 40.5"),
    ("سرعة التيار", "m/s", "لم يُقَس بالموقع", "0 – 0.7", "حتى 0.5", "0.05 – 0.17"),
    ("الأكسجين المذاب", "mg/L", "3.5 – 10", "5 – 8 (متوسط)", "3 – 15", "3 – 15"),
    ("الأس الهيدروجيني pH", "", "7.6 – 8.3", "7.0 – 8.6", "7.3 – 8.1", "7.3 – 8.1"),
]
TAKEAWAYS = [
    "الحرارة لا تقترب أبداً من الصفر: أقل قراءة مسجلة 15.5 °C (يناير) وأعلى قراءة 31.5 °C (أغسطس).",
    "الملوحة عالية (37 – 40.5 PSU، أي موصلية حتى ≈ 60 mS/cm)، وتنخفض قرب مصبات الصرف وبحيرة إدكو في أبو قير.",
    "التيارات هادئة: أعلى قياس 0.7 م/ث داخل الميناء الشرقي، والمد صغير جداً (6 – 23 سم).",
    "لا توجد قياسات منشورة للعكارة (NTU) ولا لمعدل الترسيب في أي موقع — وهذا ما سيقيسه جهازنا لأول مرة.",
]
REQUIRED = [("الحرارة", "10 – 35 °C"), ("الموصلية", "30 – 63 mS/cm"), ("الملوحة", "20 – 42 PSU"),
            ("سرعة التيار", "0 – 1 m/s"), ("الاتجاه", "0 – 360°"), ("عمق التشغيل", "≥ 10 م")]

# ---------------------------------------------------------------- sensors
DIY_SVG = """<svg viewBox="0 0 520 300" xmlns="http://www.w3.org/2000/svg" font-family="Cairo, sans-serif">
<defs><linearGradient id="w" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#e6f4f7"/><stop offset="1" stop-color="#bfe3ea"/></linearGradient></defs>
<rect x="0" y="0" width="520" height="300" fill="url(#w)"/>
<rect x="0" y="262" width="520" height="38" fill="#d9c9a3"/>
<rect x="225" y="244" width="70" height="20" rx="3" fill="#9aa4ad"/>
<text x="260" y="290" text-anchor="middle" font-size="13" fill="#5b4a2a">قاعدة خرسانية</text>
<line x1="260" y1="246" x2="330" y2="120" stroke="#3b3b3b" stroke-width="3"/>
<g transform="rotate(-29 335 92)"><rect x="315" y="40" width="40" height="110" rx="18" fill="#0b2a4a"/>
<rect x="322" y="70" width="26" height="30" rx="4" fill="#14b8a6"/><text x="335" y="90" text-anchor="middle" font-size="9" fill="#fff">IMU</text></g>
<path d="M300 92 A70 70 0 0 1 260 175" fill="none" stroke="#c0392b" stroke-width="2" stroke-dasharray="5 4"/>
<line x1="260" y1="246" x2="260" y2="60" stroke="#7f93ad" stroke-width="1.5" stroke-dasharray="4 4"/>
<text x="290" y="158" font-size="15" fill="#c0392b" font-weight="700">θ</text>
<g fill="#0b6b7a"><path d="M40 120 h120 v-10 l26 20 -26 20 v-10 h-120z"/><path d="M40 190 h90 v-8 l22 16 -22 16 v-8 h-90z"/></g>
<text x="95" y="100" text-anchor="middle" font-size="15" fill="#0b2a4a" font-weight="700">اتجاه التيار</text>
<text x="420" y="40" text-anchor="middle" font-size="13" fill="#0b2a4a">عوامة + غلاف 2"</text>
<text x="420" y="58" text-anchor="middle" font-size="12" fill="#0b2a4a">زاوية الميل θ ← السرعة</text>
</svg>"""

SENSORS = [
    {
        "brand": "SensorsX", "model": "SX-CCFNA", "sub": "حساس موصلية / ملوحة رباعي الأقطاب",
        "tags": ["ملوحة وموصلية", "4 أقطاب"], "img": "img/sx-ccfna.webp",
        "desc": "حساس رقمي يقيس الموصلية والملوحة والمواد الذائبة (TDS) ودرجة الحرارة معاً بتقنية الأقطاب الأربعة، ويرسل القراءات عبر RS-485 (Modbus).",
        "tiles": [("0.001 – 100 mS/cm", "مدى الموصلية (خيار المدى العالي)"), ("0 – 80 ppt", "مدى الملوحة"),
                  ("±1 %", "الدقة (أو 0.01 mS/cm)"), ("RS-485", "الاتصال (Modbus)"),
                  ("9 – 24 V", "التغذية (0.2 W)"), ("IP68", "الحماية · Φ22 × 175.5 mm")],
        "price": "$473", "price_note": "سعر الموقع الرسمي sensorsx.com (من مراجعة الفريق)",
        "steps": ["اطلب خيار المدى العالي 0–100 mS/cm / 0–80 ppt (خيار 0–5 mS/cm لا يصلح لمياه البحر).",
                  "ثبّته على الإطار في القاع ووصّل كابل RS-485 بوحدة التجميع.",
                  "يقرأ الموصلية والملوحة والحرارة بشكل مستمر عبر Modbus.",
                  "نظّفه دورياً من الحشف الحيوي وأعد معايرته (نقطة أو نقطتين)."],
        "pros": ["يقيس 4 قيم في جهاز واحد", "جسم Ni + PEEK + تيتانيوم مناسب لمياه البحر", "استهلاك منخفض 0.2 W", "معاير من المصنع"],
        "cons": ["لازم طلب خيار المدى العالي", "الكابل القياسي 10 م فقط", "عمق التشغيل غير مذكور في الداتاشيت", "طريقة المعايرة غير موضحة"],
        "fit": ("مناسب (OK range)", "ok"), "team": ("▲ مشروط", "cond"),
        "links": [("Product Page", "https://sensorsx.com/products/sx-ccfna-series"),
                  ("Datasheet (PDF)", "Rev A — files/sensors/salenity/SensorsX SX-CCFNA/")],
    },
    {
        "brand": "WizSensor", "model": "WS-OM-MUL-17A", "sub": "حساس CTD: موصلية · ملوحة · حرارة · عمق",
        "tags": ["CTD", "ملوحة وموصلية"], "img": "img/ws-om-mul-17a.webp",
        "desc": "حساس CTD يقيس الموصلية والحرارة والعمق ويحسب الملوحة، بمخرج RS-485 Modbus RTU، ويضيف قياس منسوب المياه فوق الموقع.",
        "tiles": [("0 – 100 mS/cm", "مدى الموصلية (±1 %)"), ("0 – 80 ppt", "الملوحة (محسوبة، ±1 ppt)"),
                  ("±0.05 °C", "دقة الحرارة (−5 إلى 50 °C)"), ("0 – 200 m", "مدى العمق (±0.1 % FS)"),
                  ("RS-485", "Modbus RTU · 9600 8N1"), ("6 – 36 V", "التغذية (≈ 20 mA عند 12 V)")],
        "price": "عند الطلب", "price_note": "في انتظار عرض سعر من WizSensor (لم يرد على البريد بعد)",
        "steps": ["اطلب عرض السعر وأكّد طول الكابل (30 م قياسي) وتصنيف الضغط ≥ 10 م.",
                  "ثبّته على الإطار، ولا تستخدم كابل الإشارة لتعليق الجهاز.",
                  "يرسل الحرارة والموصلية والملوحة والعمق لحظياً عبر Modbus.",
                  "نظّفه كل 4–8 أسابيع وأعد المعايرة (الحرارة والموصلية بنقطة أو نقطتين)."],
        "pros": ["أفضل توافق فني مع ملوحة وحرارة الموقع", "يقيس العمق ومنسوب المياه أيضاً", "استهلاك منخفض يناسب العوامة الشمسية", "دقة حرارة عالية ±0.05 °C"],
        "cons": ["السعر غير معلن", "جسم 316L — تأكد من ملاءمته للغمر الطويل", "الداتاشيت يُرسل عند الطلب فقط", "دقة الملوحة ±1 ppt"],
        "fit": ("مناسب (OK range)", "ok"), "team": ("✔ مناسب", "yes"),
        "links": [("Product Page", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/"),
                  ("Datasheet (PDF)", "On request from WizSensor")],
    },
    {
        "brand": "Dragino", "model": "DR-EC200", "sub": "حساس موصلية رباعي الأقطاب لمياه البحر",
        "tags": ["ملوحة وموصلية", "4 أقطاب"], "img": "img/dr-ec200-drawing.png", "img_class": "drawing",
        "img_caption": "رسم الأبعاد من مستندات Dragino (عائلة DR-EC)",
        "desc": "حساس موصلية رقمي بأربعة أقطاب مصمم للمياه عالية الموصلية مثل مياه البحر، يقيس الموصلية والملوحة والحرارة ويرسلها عبر RS-485.",
        "tiles": [("1 – 200,000 µS/cm", "مدى الموصلية (±1 % FS)"), ("0 – 70 PSU", "مدى الملوحة (دقة 0.1)"),
                  ("±0.5 °C", "دقة الحرارة (−5 إلى 80 °C)"), ("RS-485", "Modbus-RTU · 9600"),
                  ("7 – 30 V", "التغذية (< 0.5 W)"), ("0.6 MPa", "الضغط · IP68 (عدا الموصّل)")],
        "price": "$375", "price_note": "€325.61 لدى Antratek — القيمة بالدولار تقريبية",
        "steps": ["مدّد الكابل (5 م قياسي) ليصل إلى السطح.",
                  "عاير بنقطتين عبر Modbus: محلول 1,413 µS/cm ثم 111,310 µS/cm.",
                  "يقرأ الموصلية والملوحة والحرارة بشكل مستمر.",
                  "أعد المعايرة كل 3 شهور ونظّفه من الحشف الحيوي."],
        "pros": ["أرخص حساس ملوحة مناسب للموقع", "مصمم لمياه البحر", "خطوات المعايرة موضحة بالتفصيل", "مدخل تغذية مرن 7–30 V"],
        "cons": ["المدى أكبر من المطلوب بثلاث مرات (Overpowered)", "الكابل 5 م فقط", "يحتاج معايرة كل 3 شهور", "أبعاد الجسم غير مذكورة نصياً"],
        "fit": ("أكبر من المطلوب (Overpowered)", "over"), "team": ("✔ مناسب", "yes"),
        "links": [("Product Page", "https://www.antratek.com/dr-ec200-water-ec-sensor-for-wqs"),
                  ("Datasheet", "https://docs.dragino.com/docs/Non-Wireless-Sensors/water-quality/dr-eck1.0-dr-eck10.0-ec200-ec-digital-electrode-sensor")],
    },
    {
        "brand": "تصنيع ذاتي", "model": "DIY Tilt Current Meter", "sub": "مقياس تيار بزاوية الميل (BNO085)",
        "tags": ["ميل (Tilt)", "تصنيع ذاتي"], "svg": DIY_SVG,
        "desc": "عوامة مربوطة بالقاع تميل مع التيار؛ حساس حركة (IMU) داخل غلاف مقاوم للماء يقيس زاوية الميل واتجاهها، ثم تُحوَّل الزاوية إلى سرعة بالمعايرة.",
        "tiles": [("حسب المعايرة", "مدى السرعة (المطلوب 0–1 م/ث)"), ("0 – 360°", "الاتجاه (بوصلة معوّضة للميل)"),
                  ("BNO085", "حساس الحركة (Adafruit)"), ("18650", "البطارية (≈ شهر تسجيل)"),
                  ("60–70 ft", "عمق مجرّب بدون تسرب (≈ 20 م)"), ("SD", "التسجيل (أضف RS-485 للبث اللحظي)")],
        "price": "$200 (تقديري)", "price_note": "BNO085 ‏$29.50 + غلاف Blue Robotics من $61 + معالج وبطارية وعوامة وحبل",
        "steps": ["ركّب الـIMU والمعالج والبطارية داخل غلاف 2 بوصة.",
                  "اربط العوامة بحبل إلى القاعدة الخرسانية في القاع.",
                  "عاير الجهاز بسحبه بسرعات معروفة (GPS) لتحويل الزاوية إلى سرعة.",
                  "أبعد الحديد عن البوصلة، وأضف RS-485 لإرسال البيانات لحظياً."],
        "pros": ["أرخص خيار بفارق كبير", "قطع متوفرة وسهلة الشراء", "يمكن تعديله حسب احتياج المشروع", "لا يحتوي أجزاء دوّارة"],
        "cons": ["الدقة غير معروفة قبل المعايرة", "لا يرسل بيانات لحظية بدون تعديل", "نمو الكائنات على العوامة يغيّر القراءة", "لم يُختبر في الموقع بعد"],
        "fit": ("غير معروف (Unknown)", "unk"), "team": ("▲ مشروط", "cond"),
        "links": [("Project Page", "https://hackaday.io/project/158083-cheap-underwater-tilt-current-sensor"),
                  ("Enclosure", "https://bluerobotics.com/store/watertight-enclosures/wte-vp/")],
    },
    {
        "brand": "Nortek", "model": "Signature 500", "sub": "بروفايلر ADCP خماسي الأشعة · 500 kHz",
        "tags": ["دوبلر صوتي", "بروفايلر"], "img": "img/signature500.png",
        "desc": "بروفايلر دوبلر صوتي بخمس أشعة يقيس التيار في طبقات عمود الماء، مع حساسات حرارة وضغط وبوصلة، ويرسل عبر Ethernet أو RS-232/422.",
        "tiles": [("1 – 5 m/s", "مدى السرعة (قابل للضبط)"), ("0.3 % ±0.3 cm/s", "الدقة"),
                  ("60 – 70 m", "مدى البروفايل (خلايا 0.5–4 م)"), ("300 m", "أقصى عمق"),
                  ("16 – 128 GB", "الذاكرة"), ("12 – 48 V", "التغذية (8 W عند 1 Hz)")],
        "price": "غير معلن", "price_note": "السعر عند الطلب من Nortek",
        "steps": ["اضبط مدى السرعة على 1 م/ث وحجم الخلية 0.5 م.",
                  "ثبّته على إطار في القاع متجهاً لأعلى.",
                  "يسجل ذاتياً، أو يوصل بكابل للقراءة اللحظية (Ethernet / RS-422).",
                  "حلل البيانات ببرنامج Nortek بعد الرفع أو لحظياً."],
        "pros": ["دقة عالية جداً", "يقيس كل طبقات عمود الماء", "حرارة وضغط وبوصلة مدمجة", "ذاكرة كبيرة حتى 128 GB"],
        "cons": ["مصمم للمياه العميقة (مدى 60–70 م)", "منطقة ميتة 0.5 م أمام الحساس", "كبير وثقيل (Ø228 مم، 6.4 كجم)", "سعر مرتفع غير معلن"],
        "fit": ("مناسب (OK range)", "ok"), "team": ("لم تتم مراجعته", "unk"),
        "note": "بديل أنسب للعمق 2–8 م: Signature 1000 (مدى 25–30 م، خلايا من 0.2 م، وضع HR من 0.1 م).",
        "links": [("Product Page", "https://www.nortekgroup.com/products/signature-500"),
                  ("Datasheet (PDF)", "https://www.nortekgroup.com/products/signature-500/pdf")],
    },
    {
        "brand": "Nortek", "model": "AWAC 600 kHz", "sub": "بروفايلر تيار وأمواج · الجيل الثاني",
        "tags": ["دوبلر صوتي", "تيار + أمواج"], "img": "img/awac600.png",
        "desc": "بروفايلر دوبلر صوتي يقيس التيار على طبقات، ويقيس الأمواج أيضاً بتقنية AST، مع حساسات حرارة وضغط.",
        "tiles": [("1 – 5 m/s", "مدى السرعة (قابل للضبط)"), ("±1 % ±0.5 cm/s", "الدقة"),
                  ("50 m", "مدى البروفايل (خلايا 0.5–8 م)"), ("300 m", "أقصى عمق"),
                  ("حتى 60 م", "قياس الأمواج"), ("12 – 48 V", "التغذية (بطاريات خارجية)")],
        "price": "غير معلن", "price_note": "السعر عند الطلب من Nortek",
        "steps": ["اضبط مدى السرعة على 1 م/ث واختر أصغر خلية (0.5 م).",
                  "ثبّته على إطار في القاع مع علبة البطارية الخارجية (540 / 1800 Wh).",
                  "يسجل التيار والأمواج، أو يوصل بكابل Ethernet / RS-422 للقراءة اللحظية.",
                  "حلل بيانات التيار والأمواج ببرنامج Nortek."],
        "pros": ["يقيس التيار والأمواج معاً", "جهاز قوي ومجرب في السواحل", "حرارة وضغط مدمجة", "اتصال Ethernet"],
        "cons": ["مدى 50 م أكبر بكثير من عمق الموقع", "خلايا 0.5–8 م خشنة لعمق 2–8 م", "يحتاج علبة بطاريات خارجية", "سعر مرتفع غير معلن"],
        "fit": ("مناسب (OK range)", "ok"), "team": ("لم تتم مراجعته", "unk"),
        "note": "بديل أنسب للعمق 2–8 م: AWAC 1 MHz (مدى 30 م، خلايا من 0.25 م).",
        "links": [("Product Page", "https://www.nortekgroup.com/products/awac2-600-khz"),
                  ("Datasheet (PDF)", "https://www.nortekgroup.com/products/awac2-600-khz/pdf")],
    },
    {
        "brand": "DFRobot", "model": "SEN0710", "sub": "حساس العكارة (Turbidity) RS485 لجودة المياه",
        "tags": ["عكارة وترسيب", "تشتت 90°"], "img": "img/sen0710.webp",
        "desc": "حساس رقمي يقيس عكارة المياه (الجسيمات العالقة والرواسب) بطريقة تشتت الضوء بزاوية 90° بالأشعة تحت الحمراء، مع قياس درجة الحرارة، ويرسل القراءات عبر RS485 Modbus-RTU.",
        "tiles": [("0 – 1000 NTU", "مدى العكارة (دقة 0.1 NTU)"), ("±5 % FS", "الدقة (عند 25 °C)"),
                  ("0 – 60 °C", "قياس الحرارة (±0.5 °C)"), ("RS485", "Modbus-RTU · 4800 افتراضياً"),
                  ("10 – 30 V", "التغذية (0.2 W)"), ("IP68", "0.6 MPa · كابل 5 م")],
        "price": "غير معلن", "price_note": "لم يُنشر سعر على صفحة DFRobot الرسمية (wiki)",
        "steps": ["ثبّته على الإطار بحيث يكون رأس القياس بعيداً عن القاع وعن الجدران القريبة.",
                  "وصّل الأسلاك: بني VCC، أسود GND، أصفر 485-A، أزرق 485-B بوحدة التجميع.",
                  "يقرأ العكارة والحرارة بشكل مستمر عبر Modbus-RTU (زمن الاستجابة ≤ 30 ث).",
                  "نظّف النافذة البصرية من الحشف الحيوي وأعد المعايرة بنقطتين عند الحاجة."],
        "pros": ["يقيس الرواسب العالقة (العكارة) مباشرة", "معاير من المصنع وبالإمكان معايرة نقطتين", "نفس واجهة RS485 للحساسات الأخرى", "استهلاك منخفض 0.2 W"],
        "cons": ["التوافق مع مياه البحر غير مذكور (الجسم بلاستيك)", "لا يوجد عمق تشغيل أقصى (0.6 MPa فقط)", "الحرارة التشغيلية 0–40 °C", "لا توجد قياسات عكارة منشورة للمواقع للمقارنة"],
        "fit": ("غير معروف (Unknown)", "unk"), "team": ("لم تتم مراجعته", "unk"),
        "note": "تحقق من DFRobot من ملاءمة الجسم البلاستيكي للغمر الطويل في مياه البحر قبل الشراء.",
        "links": [("Wiki", "https://wiki.dfrobot.com/sen0710/"),
                  ("Modbus Reference", "https://wiki.dfrobot.com/sen0710/docs/23385")],
    },
    {
        "brand": "DFRobot", "model": "SEN0681", "sub": "حساس الأكسجين المذاب الفلورسنتي RS485 (مياه البحر)",
        "tags": ["أكسجين مذاب", "فلورسنت"], "img": "img/sen0681.webp",
        "desc": "حساس أكسجين مذاب بتقنية الفلورسنت (بدون تحريك أو استهلاك أكسجين)، نسخة مخصصة لمياه البحر بضبط ملوحة افتراضي 30‰، مع قياس الحرارة وإخراج RS485 Modbus-RTU.",
        "tiles": [("0 – 20 mg/L", "الأكسجين المذاب (0–200 % تشبع)"), ("±3 % FS", "الدقة (±0.5 °C عند 25 °C)"),
                  ("0.01 mg/L", "الدقة الرقمية (Resolution)"), ("RS485", "Modbus-RTU"),
                  ("10 – 30 V", "التغذية (0.2 W)"), ("IP68", "0.6 MPa · كابل 5 م")],
        "price": "غير معلن", "price_note": "لم يُنشر سعر (صفحة المتجر الرسمية لم تُفتح)",
        "steps": ["اضبط الملوحة على قيمة الموقع (36–40 ‰ تقريباً) بدل الافتراضي 30‰.",
                  "ثبّته على الإطار والغشاء الفلورسنتي مواجه للماء الحر دون فقاعات.",
                  "يقرأ الأكسجين (mg/L و% تشبع) والحرارة عبر Modbus-RTU (زمن الاستجابة ≤ 60 ث).",
                  "بدّل غشاء الفلورسنت سنوياً ونظّفه من الحشف الحيوي وعاير حسب دليل DFRobot."],
        "pros": ["نسخة مخصصة لمياه البحر", "تقنية فلورسنت: لا يستهلك أكسجين ولا يحتاج تحريكاً", "المدى 0–20 mg/L يغطي 3–15 mg/L في المواقع", "نفس واجهة RS485 للحساسات الأخرى"],
        "cons": ["عمر الغشاء الفلورسنتي سنة واحدة", "الدقة ±3 % FS", "الجسم بلاستيك — تأكد من الغمر الطويل", "لا يوجد عمق تشغيل أقصى (0.6 MPa فقط)"],
        "fit": ("مناسب (OK range)", "ok"), "team": ("لم تتم مراجعته", "unk"),
        "links": [("Wiki", "https://wiki.dfrobot.com/SEN0681"),
                  ("Dimension Drawing", "https://dfimg.dfrobot.com/wiki/21095/SEN0681_seawater-dissolved-oxygen-sensor_dimension_V1.0.jpg")],
    },
]

# ---------------------------------------------------------------- HTML
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
@page { size: A4; margin: 0; }
:root{--navy:#0b2a4a;--teal:#14a6a0;--teal2:#0e7c86;--ink:#1b2b4b;--muted:#5b6b82;--line:#dbe4ee;--pale:#f1f5f9;--green:#1e7b34;--red:#b03a2e;--amber:#b26a00}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Cairo','Segoe UI',Tahoma,sans-serif;color:var(--ink);background:#fff}
.ltr{direction:ltr;unicode-bidi:isolate;display:inline-block}
.page{width:210mm;height:297mm;position:relative;overflow:hidden;page-break-after:always;padding:16mm 14mm 14mm}
.page:last-child{page-break-after:auto}
.run{position:absolute;top:7mm;left:14mm;right:14mm;display:flex;justify-content:space-between;font-size:9pt;color:var(--muted)}
.foot{position:absolute;bottom:7mm;left:0;right:0;text-align:center;font-size:9pt;color:var(--muted)}
.kicker{color:var(--teal);font-weight:700;font-size:10.5pt;margin-top:4mm}
h1.title{font-size:22pt;font-weight:800;color:var(--navy);margin:1mm 0 4mm}
p.lead{font-size:10.5pt;line-height:1.75;color:var(--ink);margin-bottom:5mm}
/* cover */
.cover{background:var(--navy);color:#fff;padding:0}
.cover .inner{position:absolute;top:58mm;right:22mm;left:22mm;direction:ltr;text-align:left}
.cover .k{color:#2ec4b6;font-weight:700;font-size:11pt;letter-spacing:.3px}
.cover h1{font-size:38pt;font-weight:800;line-height:1.12;margin:7mm 0 6mm}
.cover p{font-size:12.5pt;line-height:1.6;color:#d6e2ee;max-width:150mm}
.cover .bar{width:30mm;height:2.2mm;background:#2ec4b6;margin:12mm 0 10mm}
.chips{display:flex;flex-wrap:wrap;gap:3mm}
.chip{border:1px solid #6b86a6;border-radius:20px;padding:1.4mm 4.5mm;font-size:10pt;color:#e6eef6}
.waves{position:absolute;bottom:0;left:0;width:100%}
/* tables */
table{width:100%;border-collapse:collapse;font-size:10pt}
th{background:var(--navy);color:#fff;font-weight:700;padding:3mm 3mm;text-align:right}
td{padding:2.6mm 3mm;border-bottom:1px solid var(--line);vertical-align:middle}
tr:nth-child(even) td{background:var(--pale)}
td.num{color:var(--teal2);font-weight:800;text-align:center;width:10mm}
td a{color:#1d4ed8;font-size:8.6pt;word-break:break-all;direction:ltr;unicode-bidi:isolate}
.brand{color:var(--teal);font-weight:800}
/* water page */
.sites{display:grid;grid-template-columns:repeat(4,1fr);gap:3mm;margin-bottom:5mm}
.site{background:var(--pale);border-radius:3mm;padding:3mm;border-top:1.2mm solid var(--teal)}
.site b{display:block;font-size:10.5pt;color:var(--navy)}
.site .en{font-size:8.5pt;color:var(--muted);direction:ltr}
.site .d{font-size:9.5pt;margin-top:1.5mm}
.wtable td{text-align:center;font-size:9.6pt}.wtable td:first-child{text-align:right;font-weight:700}
.wtable th{text-align:center;font-size:9.4pt}
.takeaways{margin:5mm 0 4mm;padding:0;list-style:none}
.takeaways li{font-size:10pt;line-height:1.65;padding:1.2mm 6mm 1.2mm 0;position:relative}
.takeaways li:before{content:'';position:absolute;right:0;top:3.6mm;width:2.4mm;height:2.4mm;border-radius:50%;background:var(--teal)}
.req{background:#eaf6f5;border-right:1.5mm solid var(--teal);border-radius:2mm;padding:3.5mm 5mm}
.req h3{font-size:11.5pt;color:var(--navy);margin-bottom:2mm}
.req .g{display:grid;grid-template-columns:repeat(3,1fr);gap:2mm 5mm;font-size:10pt}
.req .g b{color:var(--teal2)}
.src{font-size:8.3pt;color:var(--muted);margin-top:4mm;line-height:1.6}
/* sensor page */
.head{display:flex;align-items:center;gap:4mm;margin-top:3mm}
.numbox{width:15mm;height:15mm;border-radius:3mm;background:var(--navy);color:#fff;font-size:18pt;font-weight:800;display:flex;align-items:center;justify-content:center;flex:none}
.head .t{flex:1}
.head .t .b{color:var(--teal);font-weight:800;font-size:11pt}
.head .t h2{font-size:19pt;font-weight:800;color:var(--navy);line-height:1.2;direction:ltr;text-align:right}
.head .t .s{font-size:9.6pt;color:var(--muted)}
.tags{display:flex;flex-direction:column;gap:1.5mm;align-items:flex-start}
.tag{border-radius:20px;padding:.8mm 4mm;font-size:9pt;font-weight:700;border:1.2px solid var(--teal2);color:var(--teal2)}
.tag:first-child{background:var(--teal2);color:#fff}
.imgbox{border:1px solid var(--line);border-radius:4mm;height:46mm;margin:3mm 0 2mm;display:flex;align-items:center;justify-content:center;overflow:hidden;background:#fff;position:relative}
.imgbox img{max-width:100%;max-height:44mm;object-fit:contain}
.imgbox img.drawing{max-height:42mm}
.imgbox svg{width:100%;height:100%}
.imgbox .cap{position:absolute;bottom:1.5mm;left:3mm;font-size:8pt;color:var(--muted)}
.desc{border-right:1.4mm solid var(--teal2);padding:1mm 4mm;font-size:10.6pt;line-height:1.7;margin-bottom:3mm}
.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:2.2mm}
.tile{background:var(--pale);border-radius:2.5mm;padding:2mm 2mm;text-align:center}
.tile b{display:block;font-size:11pt;font-weight:800;color:var(--navy);direction:ltr;unicode-bidi:isolate}
.tile span{font-size:8.4pt;color:var(--muted)}
.pricebar{display:flex;align-items:center;gap:4mm;margin:3mm 0 2mm;background:var(--navy);color:#fff;border-radius:3mm;padding:2.4mm 5mm}
.pricebar .l{font-size:9.5pt;color:#9fb7d0}
.pricebar .v{font-size:15pt;font-weight:800;color:#2ec4b6;direction:ltr;unicode-bidi:isolate}
.pricebar .n{font-size:8.8pt;color:#d6e2ee;flex:1}
.fitrow{display:flex;gap:3mm;margin-bottom:2mm;font-size:8.8pt}
.pill{border-radius:20px;padding:.6mm 3.5mm;font-weight:700}
.ok,.yes{background:#e2f0d9;color:var(--green)}.over,.cond{background:#fff2cc;color:var(--amber)}.no{background:#fadbd8;color:var(--red)}.unk{background:#eef1f4;color:#6b7280}
h3.sec{font-size:12pt;color:var(--navy);border-right:1.2mm solid var(--teal);padding-right:2.5mm;margin:2mm 0 2mm}
.steps{display:grid;grid-template-columns:1fr 1fr;gap:1.6mm 6mm;margin-bottom:3mm}
.step{display:flex;gap:2mm;font-size:9.2pt;line-height:1.55;align-items:flex-start}
.step i{font-style:normal;flex:none;width:5.6mm;height:5.6mm;border-radius:50%;background:var(--teal2);color:#fff;font-weight:800;font-size:8.6pt;display:flex;align-items:center;justify-content:center;margin-top:.4mm}
.pc{display:grid;grid-template-columns:1fr 1fr;gap:4mm;margin-bottom:3mm}
.box{border-radius:3mm;padding:2.6mm 4mm;border-top:1.2mm solid}
.box h4{font-size:11pt;margin-bottom:1mm}
.box ul{list-style:none;font-size:9.2pt;line-height:1.6}
.box li:before{content:'• '}
.pros{background:#eef7ee;border-color:var(--green)}.pros h4{color:var(--green)}
.cons{background:#fbeeee;border-color:var(--red)}.cons h4{color:var(--red)}
.note{font-size:9pt;background:#fff8e6;border-radius:2mm;padding:1.8mm 3.5mm;margin-bottom:2.5mm;color:#6b4a00}
.links td{font-size:8.6pt;padding:1.6mm 3mm;background:var(--pale)}
.links td.k{background:var(--navy);color:#fff;font-weight:700;width:34mm;direction:ltr;text-align:left}
.links{border-collapse:separate;border-spacing:0 1.2mm}
"""

WAVES = """<svg class="waves" viewBox="0 0 800 260" preserveAspectRatio="none"><path d="M0 60 C200 0 330 70 520 40 C650 20 740 30 800 50 V260 H0Z" fill="#0f3a63"/>
<path d="M0 120 C180 70 320 140 520 105 C660 80 740 90 800 110 V260 H0Z" fill="#0f5a73"/>
<path d="M0 175 C200 140 360 200 560 165 C680 145 750 150 800 165 V260 H0Z" fill="#0e7c86"/></svg>"""


def page(body, n, cls=""):
    run = f'<div class="run"><span>{DATE}</span><span>{RUNNING}</span></div>' if n > 1 else ""
    foot = f'<div class="foot">{n}</div>' if n > 1 else ""
    return f'<section class="page {cls}">{run}{body}{foot}</section>'


def cover():
    chips = "".join(f'<span class="chip">{c}</span>' for c in ["SensorsX", "WizSensor", "Dragino", "Nortek", "DFRobot", "DIY build", "4 Alexandria sites"])
    return page(f"""<div class="inner"><div class="k">TECHNICAL REPORT · ALEX-MON UNDERWATER HERITAGE MONITORING</div>
<h1>Selected Sensors<br>Guide</h1>
<p>Water conditions at the four Alexandria underwater sites, and eight shortlisted sensors for salinity, temperature, water current, turbidity and dissolved oxygen: key specifications, price, advantages and limitations.</p>
<div class="bar"></div><div class="chips">{chips}</div></div>{WAVES}""", 1, "cover")


def price_cell(p):
    return f'<b class="ltr">{e(p)}</b>' if p.startswith("$") else e(p)


def contents():
    rows = "".join(
        f'<tr><td class="num">{i}</td><td><span class="brand">{e(s["brand"])}</span> <b class="ltr">{e(s["model"])}</b></td>'
        f'<td>{e(s["tags"][0])}</td><td>{price_cell(s["price"])}</td>'
        f'<td style="text-align:center;font-weight:700">{i + 3}</td></tr>'
        for i, s in enumerate(SENSORS, 1))
    return page(f"""<div class="kicker">Contents</div><h1 class="title">المحتويات</h1>
<table><tr><th>#</th><th>القسم</th><th></th><th style="text-align:center">صفحة</th></tr>
<tr><td class="num">—</td><td><b>بيانات المياه في المواقع الأربعة</b></td><td>الخلفية</td><td style="text-align:center;font-weight:700">3</td></tr></table>
<h3 class="sec" style="margin-top:7mm">قائمة الحساسات</h3>
<table><tr><th>#</th><th>الحساس</th><th>القياس</th><th>السعر</th><th style="text-align:center">صفحة</th></tr>{rows}</table>
<p class="src" style="margin-top:6mm">الأسعار من مراجعة الفريق للموردين (Under_Water_BoM.xlsx). "عند الطلب" و"غير معلن" تعني أن المورد لم ينشر سعراً بعد.
المواصفات منقولة من الداتاشيت الرسمية لكل مصنّع، والمرجع الكامل في Sensors_BOM.xlsx.</p>""", 2)


def water():
    sites = "".join(f'<div class="site"><b>{a}</b><div class="en">{en}</div><div class="d">العمق: {d}</div><div class="d">{dist}</div></div>'
                    for a, en, d, dist in SITES)
    head = "".join(f"<th>{a}</th>" for a, *_ in SITES)
    rows = "".join(f'<tr><td>{p}{f" <span class=ltr>({u})</span>" if u else ""}</td>' + "".join(f"<td>{ltr(v) if v[0].isdigit() else e(v)}</td>" for v in vals) + "</tr>"
                   for p, u, *vals in WATER)
    tk = "".join(f"<li>{e(t)}</li>" for t in TAKEAWAYS)
    req = "".join(f"<div>{k}: <b class=ltr>{e(v)}</b></div>" for k, v in REQUIRED)
    return page(f"""<div class="kicker">الخلفية · بيانات المياه</div><h1 class="title">ظروف المياه في المواقع الأربعة</h1>
<p class="lead">جمعنا القياسات المنشورة لمياه مواقع الآثار الغارقة الأربعة في الإسكندرية، لتحديد النطاق الذي يجب أن تغطيه الحساسات. الأرقام أدناه مختصرة؛ القيم التفصيلية ومصادرها في ملف <span class="ltr">Site_Water_Parameters.xlsx</span>.</p>
<div class="sites">{sites}</div>
<table class="wtable"><tr><th>القياس</th>{head}</tr>{rows}</table>
<ul class="takeaways">{tk}</ul>
<div class="req"><h3>النطاقات المطلوبة للحساسات (مشتقة من القياسات + هامش أمان)</h3><div class="g">{req}</div></div>
<p class="src">المصادر الرئيسية: NIOF Bulletin، Egyptian Journal of Aquatic Research، Rev. Biol. Mar. Oceanogr. (2014)، SN Applied Sciences (2023)، seatemperature.org — القائمة الكاملة (24 مصدراً) في ملف الإكسل.
كانوب وهيراكليون يتشاركان بيانات خليج أبو قير لأنه لا توجد قياسات منشورة عند الموقعين نفسيهما.</p>""", 3)


def sensor_page(i, s):
    tags = "".join(f'<span class="tag">{e(t)}</span>' for t in s["tags"])
    img = s["svg"] if "svg" in s else f'<img class="{s.get("img_class", "")}" src="{s["img"]}" alt="{e(s["model"])}">'
    cap = f'<span class="cap">{e(s["img_caption"])}</span>' if s.get("img_caption") else ""
    tiles = "".join(f'<div class="tile"><b>{e(v)}</b><span>{e(k)}</span></div>' for v, k in s["tiles"])
    steps = "".join(f'<div class="step"><i>{k}</i><div>{e(t)}</div></div>' for k, t in enumerate(s["steps"], 1))
    pros = "".join(f"<li>{e(t)}</li>" for t in s["pros"])
    cons = "".join(f"<li>{e(t)}</li>" for t in s["cons"])
    links = "".join(
        f'<tr><td class="k">{e(k)}</td><td>' + (f'<a href="{e(v)}">{e(v.replace("https://", ""))}</a>' if v.startswith("http") else f'<span class="ltr">{e(v)}</span>') + "</td></tr>"
        for k, v in s["links"])
    note = f'<div class="note">{e(s["note"])}</div>' if s.get("note") else ""
    fit, fcls = s["fit"]
    team, tcls = s["team"]
    return page(f"""<div class="head"><div class="numbox">{i}</div><div class="t"><div class="b">{e(s["brand"])}</div><h2>{e(s["model"])}</h2><div class="s">{e(s["sub"])}</div></div><div class="tags">{tags}</div></div>
<div class="imgbox">{img}{cap}</div>
<div class="desc">{e(s["desc"])}</div>
<div class="tiles">{tiles}</div>
<div class="pricebar"><span class="l">السعر</span><span class="v">{e(s["price"])}</span><span class="n">{e(s["price_note"])}</span></div>
<div class="fitrow"><span>ملاءمة مدى القياس:</span><span class="pill {fcls}">{e(fit)}</span><span style="margin-right:4mm">مراجعة الفريق:</span><span class="pill {tcls}">{e(team)}</span></div>
{note}<h3 class="sec">طريقة التشغيل</h3><div class="steps">{steps}</div>
<div class="pc"><div class="box pros"><h4>المميزات</h4><ul>{pros}</ul></div><div class="box cons"><h4>العيوب</h4><ul>{cons}</ul></div></div>
<table class="links">{links}</table>""", i + 3)


def main():
    pages = [cover(), contents(), water()] + [sensor_page(i, s) for i, s in enumerate(SENSORS, 1)]
    OUT_HTML.write_text(f"""<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<title>ALEX-MON Selected Sensors Guide</title><style>{CSS}</style></head><body>{''.join(pages)}</body></html>""", encoding="utf-8")
    print("wrote", OUT_HTML.name)
    if Path(CHROME).exists():
        subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=8000",
                        f"--print-to-pdf={OUT_PDF}", OUT_HTML.resolve().as_uri()], check=True, capture_output=True)
        print("wrote", OUT_PDF.name)


if __name__ == "__main__":
    main()
