"""
Kurs ishi - 5-variant
Cisco Packet Tracer: 78.0.0.0/8 tarmog'i
Word (.docx) fayl yaratuvchi script
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_border(cell):
    """Yacheykaga chegara qo'yish"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{border}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        tcBorders.append(b)
    tcPr.append(tcBorders)


def add_heading_centered(doc, text, size=14, bold=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    return p


def add_paragraph_justified(doc, text, size=14, bold=False, first_line_indent=1.25):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(first_line_indent)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = bold
    return p


def add_section_title(doc, text, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.bold = True
    return p


def add_code_block(doc, code):
    """CLI buyruqlari uchun monoshrift blok"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(code)
    run.font.name = 'Courier New'
    run.font.size = Pt(11)
    return p


# ==================================================================
# HUJJAT YARATISH
# ==================================================================

doc = Document()

# Sahifa o'lchamlarini sozlash
section = doc.sections[0]
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)
section.left_margin = Cm(3)
section.right_margin = Cm(1.5)

# Default style
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(14)

# ==================================================================
# TITUL VARAQ
# ==================================================================

add_heading_centered(doc, "O'ZBEKISTON RESPUBLIKASI TRANSPORT VAZIRLIGI", 14, True)
add_heading_centered(doc, "TOSHKENT DAVLAT TRANSPORT UNIVERSITETI", 14, True)

for _ in range(3):
    doc.add_paragraph()

add_heading_centered(doc, "„TRANSPORTDA AXBOROT TIZIMLARI VA", 14, True)
add_heading_centered(doc, "TEXNOLOGIYALARI\" KAFEDRASI", 14, True)

for _ in range(2):
    doc.add_paragraph()

add_heading_centered(doc, "\"Kompyuter tarmoqlari va tarmoq texnologiyalari\"", 14, True)
add_heading_centered(doc, "fanidan", 14, False)
add_heading_centered(doc, "5-variant asosida", 14, False)

doc.add_paragraph()

add_heading_centered(doc, "KURS ISHI", 18, True)

for _ in range(3):
    doc.add_paragraph()

# Imzolar jadvali
table = doc.add_table(rows=3, cols=4)
table.alignment = WD_ALIGN_PARAGRAPH.CENTER
headers = ['Guruh: AT-8', 'F.I.O', 'SANA', 'IMZO']
for i, h in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = h
    for run in cell.paragraphs[0].runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.bold = True
    set_cell_border(cell)

table.rows[1].cells[0].text = 'Topshirdi:'
table.rows[1].cells[1].text = '_______________'
table.rows[2].cells[0].text = 'Tekshirdi:'
table.rows[2].cells[1].text = 'Jumabayev F.X.'

for row in table.rows:
    for cell in row.cells:
        set_cell_border(cell)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

doc.add_page_break()

# ==================================================================
# MUNDARIJA
# ==================================================================

add_heading_centered(doc, "MUNDARIJA", 14, True)
doc.add_paragraph()

mundarija = [
    ("KIRISH", "3"),
    ("I-BOB. Nazariy qism", "5"),
    ("1.1. Tarmoq va uning turlari. Tarmoq topologiyasini qurish", "5"),
    ("1.2. Marshrutlashtirish. Marshrutlashtirish protokollari (RIP, OSPF, BGP)", "7"),
    ("1.3. DHCP va DNS sozlamalari", "9"),
    ("1.4. VLAN texnologiyasi va inkapsulatsiya jarayoni", "11"),
    ("II-BOB. Amaliy qism", "13"),
    ("2.1. Topologiyani tuzish", "13"),
    ("2.2. Qurilmalarga IP manzillarni belgilash", "14"),
    ("2.3. Kommutatorlarda VLAN 1 va VLAN 2 yaratish", "16"),
    ("2.4. OSPF protokolini sozlash", "18"),
    ("2.5. DHCP va DNS serverlarni sozlash", "19"),
    ("2.6. Inkapsulatsiya jarayonini tahlil qilish", "21"),
    ("2.7. Marshrutlash jadvalini qurish", "22"),
    ("Xulosa", "23"),
    ("Foydalanilgan adabiyotlar ro'yxati", "24"),
]

for item, page in mundarija:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(item)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    # Nuqtalar bilan to'ldirish
    dots = "." * (90 - len(item) - len(page))
    run2 = p.add_run(dots)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(14)
    run3 = p.add_run(page)
    run3.font.name = 'Times New Roman'
    run3.font.size = Pt(14)

doc.add_page_break()

# ==================================================================
# KIRISH
# ==================================================================

add_heading_centered(doc, "KIRISH", 14, True)
doc.add_paragraph()

kirish_text = [
    "Hozirgi zamonaviy axborot texnologiyalari rivojlanishining asosiy yo'nalishlaridan biri bu kompyuter tarmoqlari va tarmoq texnologiyalaridir. Kompyuter tarmoqlari turli xildagi qurilmalarni o'zaro bog'lash, ma'lumot almashish va resurslardan birgalikda foydalanish imkonini beruvchi murakkab tizimlardir. Tarmoqlarsiz zamonaviy biznes, ta'lim, ishlab chiqarish va ko'pgina sohalar samarali ishlay olmaydi.",
    
    "Kompyuter tarmoqlarini ko'plab belgilar, xususan hududiy ta'minlanishi jihatidan tasniflash mumkin. Bunga ko'ra global, mintaqaviy va lokal (mahalliy) tarmoqlar farqlanadi. Lokal tarmoqlar (LAN) bir bino yoki ofis ichida joylashgan kompyuterlarni birlashtirsa, global tarmoqlar (WAN) butun dunyoni qamrab olishi mumkin. Internet — bu eng katta global tarmoqning yorqin misolidir.",

    "Tarmoq qurilmalari orasida router, switch, hub, server va boshqa qurilmalar muhim o'rin egallaydi. Routerlar turli tarmoqlar o'rtasida ma'lumot paketlarini yo'naltirish vazifasini bajaradi. Switchlar bitta tarmoq segmenti ichidagi qurilmalarni bog'laydi. Hublar esa eng oddiy ulanish qurilmalari hisoblanadi. Serverlar foydalanuvchilarga turli xizmatlarni (DHCP, DNS, HTTP va boshqalar) taqdim etadi.",

    "Tarmoqlarni boshqarish uchun maxsus protokollar mavjud. Marshrutlashtirish protokollari (RIP, OSPF, BGP) ma'lumot paketlarini eng qisqa va samarali yo'l orqali manzilga yetkazadi. DHCP protokoli tarmoqdagi qurilmalarga avtomatik tarzda IP manzil va boshqa parametrlarni taqsimlaydi. DNS xizmati esa domen nomlarini IP manzillarga aylantiradi.",

    "VLAN (Virtual Local Area Network) texnologiyasi bir jismoniy tarmoqni mantiqiy jihatdan bir nechta alohida tarmoqlarga bo'lish imkonini beradi. Bu tarmoqning xavfsizligini, samaradorligini va boshqaruvini sezilarli darajada oshiradi. VLAN orqali turli bo'limlar yoki guruhlar uchun alohida tarmoq segmentlari yaratish mumkin.",

    "Ushbu kurs ishida 5-variant bo'yicha 78.0.0.0/8 tarmog'iga asoslangan kompyuter tarmog'i loyihalandi. Cisco Packet Tracer dasturi yordamida bir nechta kompyuter, switch, hub, server va router qurilmalaridan iborat topologiya tuzildi. Tarmoqda VLAN sozlamalari amalga oshirildi, OSPF marshrutlashtirish protokoli ishga tushirildi, DHCP va DNS xizmatlari sozlandi. Shuningdek, ma'lumot paketlarining inkapsulatsiya jarayoni o'rganildi va tahlil qilindi.",

    "Kurs ishini bajarish jarayonida kompyuter tarmoqlari sohasidagi nazariy bilimlar amaliy ko'nikmalar bilan mustahkamlanadi. Tarmoq qurilmalarini sozlash, IP manzillarni rejalashtirish, marshrutlashtirish va xizmatlarni boshqarish bo'yicha tajriba orttiriladi. Bu bilim va ko'nikmalar kelajakdagi kasbiy faoliyat uchun zarur poydevor bo'lib xizmat qiladi.",

    "Kurs ishining asosiy maqsadi — Cisco Packet Tracer dasturi yordamida 5-variant shartlariga muvofiq kompyuter tarmog'ini loyihalash, sozlash va sinovdan o'tkazishdir. Vazifalar sirasiga tarmoq topologiyasini qurish, qurilmalarga IP manzil berish, VLAN yaratish, marshrutlashtirish protokolini ishga tushirish, DHCP va DNS sozlash hamda inkapsulatsiya jarayonini tahlil qilish kiradi.",
]

for txt in kirish_text:
    add_paragraph_justified(doc, txt)

doc.add_page_break()

# ==================================================================
# I-BOB. NAZARIY QISM
# ==================================================================

add_heading_centered(doc, "I-BOB. NAZARIY QISM", 14, True)
doc.add_paragraph()

# 1.1
add_section_title(doc, "1.1. Tarmoq va uning turlari. Tarmoq topologiyasini qurish")

bob_1_1 = [
    "Tarmoq (network) — bu ikki yoki undan ortiq kompyuterlar, serverlar, printerlar va boshqa qurilmalarni o'zaro bog'lab, ular o'rtasida ma'lumot almashish, resurslardan birgalikda foydalanish va aloqa o'rnatish imkonini beruvchi tizimdir. Kompyuter tarmoqlari zamonaviy axborot texnologiyalarining asosiy qismi bo'lib, ular orqali foydalanuvchilar tezkor ravishda ma'lumot yuborishi va qabul qilishi, internetdan foydalanishi hamda turli xizmatlarga ulanib ishlashi mumkin.",

    "Tarmoqning asosiy vazifalari ma'lumot almashish, resurslardan umumiy foydalanish, markazlashgan boshqaruvni ta'minlash va masofaviy ishlash imkoniyatini yaratishdan iborat. Masalan, bir ofisda bir nechta kompyuterlar bitta printerga ulanib, undan birgalikda foydalanishi mumkin yoki server orqali barcha foydalanuvchilar bir xil ma'lumotlarga kirish imkoniga ega bo'ladi.",

    "Kompyuter tarmoqlari bir nechta mezonlarga ko'ra turlarga bo'linadi. Eng asosiy tasnif hududiy qamroviga ko'ra amalga oshiriladi: PAN, LAN, MAN va WAN.",

    "PAN (Personal Area Network) — bu juda kichik hududni, ya'ni bir necha metr atrofini qamrab oladigan shaxsiy tarmoq hisoblanadi. PAN odatda bir foydalanuvchiga tegishli qurilmalarni bog'lash uchun ishlatiladi. Masalan, telefon va quloqchin, telefon va smartwatch yoki noutbuk va sichqoncha o'rtasidagi Bluetooth aloqasi PAN tarmog'iga misol bo'la oladi.",

    "LAN (Local Area Network) — lokal tarmoq bo'lib, u bir bino, ofis, maktab yoki uy kabi kichik hududni qamrab oladi. LAN tarmoqlari eng ko'p ishlatiladigan tarmoqlardan biri hisoblanadi. U yuqori tezlikda ma'lumot uzatadi va nisbatan arzon bo'ladi. LAN tarmog'i orqali bir nechta kompyuterlar bir-biriga ulanib, umumiy printer, server yoki internetdan foydalanishi mumkin.",

    "MAN (Metropolitan Area Network) — bu shahar miqyosidagi tarmoq bo'lib, bir nechta LAN tarmoqlarini birlashtiradi. MAN tarmoqlari katta hududni qamrab oladi, lekin WAN dan kichikroq hisoblanadi. U ko'pincha shahar ichidagi universitetlar, banklar yoki davlat tashkilotlari tarmoqlarini bog'lash uchun ishlatiladi.",

    "WAN (Wide Area Network) — eng keng hududli tarmoq bo'lib, u mamlakatlar, qit'alar va butun dunyoni qamrab olishi mumkin. WAN tarmoqlari turli shaharlar va davlatlardagi tarmoqlarni bir-biri bilan bog'laydi. Internet eng katta WAN tarmog'ining yorqin misolidir.",

    "Tarmoq topologiyasi tarmoqdagi qurilmalarning qanday joylashganini va ular o'zaro qanday ulanib turganini ifodalaydi. Topologiya tarmoqning fizik yoki mantiqiy tuzilishini ko'rsatadi va u tarmoqning ishlash samaradorligiga katta ta'sir qiladi.",

    "Bus (shina) topologiyasida barcha kompyuterlar bitta umumiy kabelga ulanadi. Ma'lumot shu bitta kabel orqali uzatiladi. Bu topologiya sodda va arzon bo'lsa-da, asosiy kabelda muammo yuzaga kelsa, butun tarmoq ishdan chiqadi.",

    "Star (yulduz) topologiyasida barcha kompyuterlar markaziy qurilma — switch yoki hubga ulanadi. Bu topologiya eng ko'p ishlatiladi, chunki uni boshqarish oson va bir kompyuter nosoz bo'lsa ham tarmoq ishlashda davom etadi. Ushbu kurs ishida ham star topologiyasi asos qilib olingan.",

    "Ring (halqa) topologiyasida kompyuterlar yopiq halqa shaklida ulanadi. Mesh topologiyasida har bir qurilma boshqa barcha qurilmalar bilan to'g'ridan-to'g'ri ulanadi — bu eng ishonchli, lekin qimmat topologiya hisoblanadi. Tree (daraxt) topologiyasi esa bir nechta star tarmoqlarning ierarxik birlashmasidan iborat.",
]

for txt in bob_1_1:
    add_paragraph_justified(doc, txt)

doc.add_paragraph()

# 1.2
add_section_title(doc, "1.2. Marshrutlashtirish. Marshrutlashtirish protokollari (RIP, OSPF, BGP)")

bob_1_2 = [
    "Marshrutlashtirish (routing) — bu tarmoqda ma'lumot paketlarini bir qurilmadan (manbadan) boshqa qurilmaga (manzilga) eng to'g'ri va samarali yo'l orqali uzatish jarayonidir. Marshrutlashtirish odatda router (yo'naltiruvchi qurilma) yordamida amalga oshiriladi. Router tarmoqdagi turli yo'llarni tahlil qilib, ma'lumot uchun eng yaxshi yo'lni tanlaydi va uni manzilga yetkazadi.",

    "Marshrutlashtirishning asosiy vazifasi — ma'lumotlarni to'g'ri, tez va ishonchli yetkazishdir. Katta tarmoqlarda (masalan internetda) ma'lumot bir nechta routerlar orqali o'tadi va har bir router paketni keyingi eng yaxshi yo'nalishga uzatadi.",

    "Marshrutlashtirish ikki asosiy turga bo'linadi: statik marshrutlashtirish va dinamik marshrutlashtirish. Statik marshrutlashtirishda yo'llar administrator tomonidan qo'lda sozlanadi. Dinamik marshrutlashtirishda esa routerlar o'zaro axborot almashib, eng yaxshi yo'llarni avtomatik aniqlaydi.",
]

for txt in bob_1_2:
    add_paragraph_justified(doc, txt)

add_section_title(doc, "1. RIP (Routing Information Protocol)", 13)
add_paragraph_justified(doc, "RIP — eng oddiy va eski marshrutlashtirish protokollaridan biri hisoblanadi. U distance-vector (masofa-vektor) prinsipida ishlaydi. RIP tarmoqda eng yaxshi yo'lni aniqlashda \"hop count\" (sakrashlar soni) dan foydalanadi. Ya'ni ma'lumot nechta routerdan o'tishini hisoblaydi. Eng kam hop count bo'lgan yo'l eng yaxshi yo'l deb tanlanadi.")

add_paragraph_justified(doc, "RIP ning asosiy xususiyatlari: maksimal 15 hopgacha ishlaydi (16 — cheksiz, ya'ni yo'l yo'q); har 30 soniyada yangilanadi; kichik tarmoqlar uchun mos; sodda sozlanadi.")

add_section_title(doc, "2. OSPF (Open Shortest Path First)", 13)
add_paragraph_justified(doc, "OSPF — zamonaviy va kuchli marshrutlashtirish protokoli bo'lib, link-state (aloqa holati) asosida ishlaydi. OSPF tarmoqdagi barcha yo'llarni xarita sifatida o'rganadi va Dijkstra algoritmi yordamida eng qisqa yo'lni hisoblab chiqadi.")

add_paragraph_justified(doc, "OSPF xususiyatlari: katta tarmoqlar uchun mos; tez ishlaydi; tarmoqni hududlarga (area) bo'lib boshqaradi; metrik sifatida cost (xarajat) dan foydalanadi; hop count cheklovi yo'q. Ushbu kurs ishida aynan OSPF protokoli qo'llanildi, chunki u kichik va o'rta tarmoqlar uchun ham juda samarali hisoblanadi.")

add_section_title(doc, "3. BGP (Border Gateway Protocol)", 13)
add_paragraph_justified(doc, "BGP — eng katta va eng muhim marshrutlashtirish protokoli bo'lib, internet darajasida ishlatiladi. U path-vector (yo'l vektori) prinsipiga asoslanadi. BGP asosan turli internet provayderlar va yirik tarmoqlar (autonomous system) o'rtasida ma'lumot almashish uchun ishlatiladi.")

add_paragraph_justified(doc, "BGP xususiyatlari: internetning asosiy protokoli hisoblanadi; turli tarmoqlarni bog'laydi; siyosat (policy) asosida yo'l tanlaydi; juda murakkab sozlamalarga ega.")

doc.add_paragraph()

# 1.3
add_section_title(doc, "1.3. DHCP va DNS sozlamalari")

bob_1_3 = [
    "DHCP va DNS sozlamalari kompyuter tarmoqlarida juda muhim rol o'ynaydi, chunki ular tarmoqdagi qurilmalar o'rtasida avtomatik aloqa o'rnatish va internet resurslariga qulay kirishni ta'minlaydi. Bu ikki xizmat tarmoqni boshqarishni osonlashtiradi va foydalanuvchilarga qo'lda murakkab sozlamalar kiritish zaruratini kamaytiradi.",

    "DHCP (Dynamic Host Configuration Protocol) — bu tarmoqdagi qurilmalarga avtomatik ravishda IP manzil, subnet mask, gateway va DNS server kabi tarmoq parametrlarini beruvchi protokoldir. DHCP ning asosiy vazifasi — tarmoqdagi har bir qurilmaga qo'lda IP berish o'rniga, avtomatik tarzda kerakli sozlamalarni taqsimlashdir.",

    "DHCP to'rt asosiy bosqichda ishlaydi (DORA jarayoni): Discover — qurilma tarmoqqa ulanadi va DHCP serverni qidiradi; Offer — server mavjud IP manzilni taklif qiladi; Request — qurilma taklif qilingan IP ni qabul qilishini bildiradi; Acknowledge — server IP ni tasdiqlaydi va qurilmaga biriktiradi.",

    "DHCP faqat IP berish bilan cheklanmaydi, balki quyidagi ma'lumotlarni ham avtomatik beradi: IP manzil, Subnet mask, Default gateway, DNS server manzili va Lease time (IP muddatidan foydalanish vaqti).",

    "DHCP ning asosiy afzalliklari shundan iboratki, u tarmoqni boshqarishni soddalashtiradi, IP manzillarni samarali taqsimlaydi va xatoliklarni kamaytiradi. Ayniqsa katta tarmoqlarda qo'lda IP berish juda qiyin bo'lgani uchun DHCP juda keng qo'llaniladi.",

    "DNS (Domain Name System) — bu domen nomlarini IP manzillarga aylantiruvchi tizimdir. Oddiy qilib aytganda, DNS internetdagi \"telefon kitobi\"ga o'xshaydi. Foydalanuvchi brauzerga sayt nomini (masalan, google.com) yozganda, DNS uni mos IP manzilga aylantiradi va kompyuterni kerakli serverga ulaydi.",

    "DNS tizimi ierarxik tuzilishga ega bo'lib, bir nechta darajalardan iborat: root serverlar, top-level domain (TLD) serverlar (.com, .org, .uz) va authoritative serverlar. Har bir so'rov ushbu bosqichlar orqali o'tib, kerakli IP manzil topiladi.",

    "DNS sozlamalarida odatda quyidagi parametrlar kiritiladi: primary DNS server va secondary DNS server. Primary DNS asosiy so'rovlarni bajaradi, secondary esa zaxira sifatida ishlaydi. Ushbu kurs ishida Server0 qurilmasi DNS server vazifasini bajaradi va uning IP manzili 78.0.0.100 dir.",
]

for txt in bob_1_3:
    add_paragraph_justified(doc, txt)

doc.add_paragraph()

# 1.4
add_section_title(doc, "1.4. VLAN texnologiyasi va inkapsulatsiya jarayoni")

bob_1_4 = [
    "VLAN (Virtual Local Area Network) — bu virtual lokal tarmoq bo'lib, bir jismoniy switch ichida bir nechta mantiqiy tarmoqlar yaratish imkonini beradi. VLAN texnologiyasi yordamida turli bo'limlar yoki guruhlardagi kompyuterlar bir-biridan alohida ishlay oladi, garchi ular bitta switchga ulangan bo'lsa ham.",

    "VLAN ning asosiy afzalliklari: tarmoqning xavfsizligini oshiradi, chunki turli VLAN dagi kompyuterlar bir-biri bilan to'g'ridan-to'g'ri muloqot qila olmaydi; broadcast trafikni kamaytiradi va tarmoq samaradorligini oshiradi; tarmoqni mantiqiy bo'limlarga ajratish imkonini beradi; boshqaruvni soddalashtiradi.",

    "VLAN da har bir port aniq VLAN ga biriktiriladi (access mode) yoki bir nechta VLAN ma'lumotlarini bir vaqtda uzatadigan trunk port sifatida sozlanadi. Trunk portlar odatda switchlar yoki switch va router o'rtasida ishlatiladi.",

    "Inkapsulatsiya (encapsulation) — bu ma'lumotlarni tarmoq bo'ylab uzatish uchun OSI modelining har bir qatlamida sarlavhalar (header) va trailerlar qo'shish jarayonidir. OSI modeli 7 qatlamdan iborat: Application (7), Presentation (6), Session (5), Transport (4), Network (3), Data Link (2) va Physical (1).",

    "Ma'lumot foydalanuvchidan tarmoq orqali yuborilganda, har bir qatlamda quyidagi inkapsulatsiya jarayoni amalga oshiriladi: Application qatlamida foydalanuvchi ma'lumoti shakllantiriladi; Transport qatlamida TCP yoki UDP sarlavhasi qo'shiladi (ma'lumot segmentga aylanadi); Network qatlamida IP sarlavhasi qo'shiladi (paket); Data Link qatlamida MAC manzillar qo'shiladi (freym); Physical qatlamda esa freym bitlar oqimiga aylantiriladi va simli yoki simsiz muhit orqali uzatiladi.",

    "Qabul qiluvchi tomonda esa teskari jarayon — dekapsulatsiya (decapsulation) amalga oshiriladi: har bir qatlamda mos sarlavhalar olib tashlanadi va asl ma'lumot foydalanuvchiga yetkaziladi. Cisco Packet Tracer dasturi Simulation rejimida har bir qatlamdagi inkapsulatsiya jarayonini ko'rsatish imkonini beradi (PDU Details oynasi orqali).",
]

for txt in bob_1_4:
    add_paragraph_justified(doc, txt)

doc.add_page_break()

# ==================================================================
# II-BOB. AMALIY QISM
# ==================================================================

add_heading_centered(doc, "II-BOB. AMALIY QISM", 14, True)
doc.add_paragraph()

# 2.1
add_section_title(doc, "2.1. Topologiyani tuzish")

bob_2_1 = [
    "Ushbu kurs ishida 5-variant shartlari asosida 78.0.0.0/8 tarmog'i bo'yicha kompyuter tarmog'i loyihalandi. Tarmoq topologiyasi Cisco Packet Tracer dasturida quyidagi qurilmalardan iborat:",

    "• 1 ta Router (Router1) — asosiy yo'naltiruvchi qurilma;\n• 3 ta Switch (Switch0, Switch1, Switch2) — kompyuterlarni birlashtiruvchi kommutatorlar;\n• 2 ta Hub (Hub0, Hub1) — switchlar va Router orasidagi ulanish;\n• 7 ta kompyuter (PC0, PC1, PC2, PC3, PC4, PC5, PC6) — foydalanuvchi qurilmalari;\n• 1 ta Server (Server0) — DHCP va DNS xizmatlarini taqdim etuvchi server;\n• 1 ta Access Point (Access Point0) — simsiz aloqani ta'minlovchi qurilma;\n• 1 ta Laptop (Laptop0) va 1 ta Tablet PC (Tablet PC0) — simsiz qurilmalar.",

    "Tarmoq topologiyasi quyidagi ulanish sxemasiga ega: PC3 va PC4 kompyuterlari Switch0 ga ulangan; PC0, PC1, PC2 kompyuterlari Switch1 ga ulangan; PC5 va PC6 kompyuterlari Switch2 ga ulangan. Switch0 va Switch1 kommutatorlari Hub0 orqali bog'langan. Switch2, Server0, Access Point0 va Router1 esa Hub1 orqali bog'langan. Hub0 va Hub1 o'zaro ulangan. Access Point0 ga simsiz Laptop0 va Tablet PC0 qurilmalari ulangan.",

    "Bunday topologiya kombinatsiyalashgan (gibrid) topologiya hisoblanadi, chunki unda star (yulduz) va bus (shina) topologiyalarining elementlari birga ishlatilgan. Switchlar atrofida star topologiyasi, hublar orqali esa bus topologiyasi tashkil qilingan.",

    "Tarmoqning asosiy parametrlari quyidagicha belgilangan:\nTarmoq IP: 78.0.0.0\nMaksimal IP: 78.255.255.254\nMinimal IP: 78.0.0.1\nTarmoq maskasi: 255.0.0.0 (/8)\nBroadcast: 78.255.255.255\nDefault Gateway: 78.0.0.254\nDNS Server: 78.0.0.100",
]

for txt in bob_2_1:
    add_paragraph_justified(doc, txt)

add_paragraph_justified(doc, "2.1-rasm. Tarmoqning umumiy topologiyasi (Cisco Packet Tracer)", first_line_indent=0).alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# 2.2
add_section_title(doc, "2.2. Qurilmalarga IP manzillarni belgilash")

add_paragraph_justified(doc, "Tarmoqdagi har bir qurilmaga 78.0.0.0/8 tarmog'idan IP manzil belgilandi. Maska barcha qurilmalar uchun 255.0.0.0, default gateway esa 78.0.0.254 (Router1) qilib o'rnatildi. Quyidagi jadvalda qurilmalarning IP manzillari ko'rsatilgan:")

# IP jadval
ip_table = doc.add_table(rows=12, cols=4)
ip_table.style = 'Light Grid Accent 1'
ip_headers = ['Qurilma', 'IP manzil', 'Subnet maskasi', 'Default Gateway']
for i, h in enumerate(ip_headers):
    cell = ip_table.rows[0].cells[i]
    cell.text = h
    for run in cell.paragraphs[0].runs:
        run.bold = True
        run.font.size = Pt(12)

ip_data = [
    ('Router1 (Fa0/0)', '78.0.0.254', '255.0.0.0', '—'),
    ('PC0', '78.0.0.1', '255.0.0.0', '78.0.0.254'),
    ('PC1', '78.0.0.2', '255.0.0.0', '78.0.0.254'),
    ('PC2', '78.0.0.3', '255.0.0.0', '78.0.0.254'),
    ('PC3', '78.0.0.4', '255.0.0.0', '78.0.0.254'),
    ('PC4', '78.0.0.5', '255.0.0.0', '78.0.0.254'),
    ('PC5', '78.0.0.6', '255.0.0.0', '78.0.0.254'),
    ('PC6', '78.0.0.7', '255.0.0.0', '78.0.0.254'),
    ('Server0', '78.0.0.100', '255.0.0.0', '78.0.0.254'),
    ('Laptop0', '78.0.0.12', '255.0.0.0', '78.0.0.254'),
    ('Tablet PC0', '78.0.0.14', '255.0.0.0', '78.0.0.254'),
]

for i, row_data in enumerate(ip_data, start=1):
    for j, val in enumerate(row_data):
        cell = ip_table.rows[i].cells[j]
        cell.text = val
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(12)

doc.add_paragraph()

add_paragraph_justified(doc, "Router1 ga IP manzil berish uchun quyidagi CLI buyruqlari ishlatildi:")

router_ip_code = """Router>enable
Router#configure terminal
Router(config)#interface FastEthernet0/0
Router(config-if)#ip address 78.0.0.254 255.0.0.0
Router(config-if)#no shutdown
Router(config-if)#exit
Router(config)#end
Router#write memory"""

add_code_block(doc, router_ip_code)

doc.add_paragraph()

add_paragraph_justified(doc, "PC larga IP manzil berish uchun har bir kompyuterda Desktop > IP Configuration menyusi orqali statik IP manzil, subnet maskasi va default gateway qiymatlari kiritildi.")

doc.add_paragraph()

# 2.3
add_section_title(doc, "2.3. Kommutatorlarda VLAN 1 va VLAN 2 yaratish")

add_paragraph_justified(doc, "Tarmoq xavfsizligini oshirish va segmentatsiya qilish maqsadida kommutatorlarda ikkita VLAN yaratildi: VLAN 1 (MANAGEMENT) va VLAN 2 (DATA). Barcha foydalanuvchi portlari va trunk portlar VLAN 1 ga biriktirildi.")

add_paragraph_justified(doc, "Switch0 ga VLAN sozlamalari:")

switch0_code = """Switch>enable
Switch#configure terminal
Switch(config)#vlan 2
Switch(config-vlan)#name DATA
Switch(config-vlan)#exit
Switch(config)#interface range fa0/1-2
Switch(config-if-range)#switchport mode access
Switch(config-if-range)#switchport access vlan 1
Switch(config-if-range)#exit
Switch(config)#interface fa0/24
Switch(config-if)#switchport mode access
Switch(config-if)#switchport access vlan 1
Switch(config-if)#end
Switch#write memory"""

add_code_block(doc, switch0_code)

doc.add_paragraph()

add_paragraph_justified(doc, "Switch1 ga VLAN sozlamalari (PC0, PC1, PC2 ulangan):")

switch1_code = """Switch>enable
Switch#configure terminal
Switch(config)#vlan 2
Switch(config-vlan)#name DATA
Switch(config-vlan)#exit
Switch(config)#interface range fa0/1-3
Switch(config-if-range)#switchport mode access
Switch(config-if-range)#switchport access vlan 1
Switch(config-if-range)#exit
Switch(config)#interface fa0/24
Switch(config-if)#switchport mode access
Switch(config-if)#switchport access vlan 1
Switch(config-if)#end
Switch#write memory"""

add_code_block(doc, switch1_code)

doc.add_paragraph()

add_paragraph_justified(doc, "Switch2 ga VLAN sozlamalari (PC5, PC6 ulangan):")

switch2_code = """Switch>enable
Switch#configure terminal
Switch(config)#vlan 2
Switch(config-vlan)#name DATA
Switch(config-vlan)#exit
Switch(config)#interface range fa0/1-2
Switch(config-if-range)#switchport mode access
Switch(config-if-range)#switchport access vlan 1
Switch(config-if-range)#exit
Switch(config)#interface fa0/24
Switch(config-if)#switchport mode access
Switch(config-if)#switchport access vlan 1
Switch(config-if)#end
Switch#write memory"""

add_code_block(doc, switch2_code)

doc.add_paragraph()

add_paragraph_justified(doc, "VLAN sozlamalarini tekshirish uchun show vlan brief buyrug'i ishlatildi. Natijada VLAN 1 (default) va VLAN 2 (DATA) ro'yxatda ko'rinadi, hamda har bir port qaysi VLAN ga biriktirilgani aniqlanadi.")

doc.add_paragraph()

# 2.4
add_section_title(doc, "2.4. OSPF protokolini sozlash")

add_paragraph_justified(doc, "Tarmoqda dinamik marshrutlashtirishni amalga oshirish uchun OSPF protokoli ishga tushirildi. OSPF Open Shortest Path First protokoli zamonaviy va kuchli protokol bo'lib, link-state algoritmiga asoslanadi va Dijkstra algoritmi yordamida eng qisqa yo'lni hisoblaydi.")

add_paragraph_justified(doc, "Router1 da OSPF sozlash uchun quyidagi CLI buyruqlari kiritildi:")

ospf_code = """Router>enable
Router#configure terminal
Router(config)#router ospf 1
Router(config-router)#network 78.0.0.0 0.255.255.255 area 0
Router(config-router)#exit
Router(config)#end
Router#write memory"""

add_code_block(doc, ospf_code)

doc.add_paragraph()

add_paragraph_justified(doc, "Yuqoridagi buyruqlarda router ospf 1 — process ID si 1 bo'lgan OSPF protokolini yoqadi. network 78.0.0.0 0.255.255.255 — 78.0.0.0/8 tarmog'ini OSPF ga qo'shadi (wildcard maskasi). area 0 — backbone area (asosiy soha) ga tegishli ekanligini bildiradi.")

doc.add_paragraph()

# 2.5
add_section_title(doc, "2.5. DHCP va DNS serverlarni sozlash")

add_paragraph_justified(doc, "DHCP xizmati Router1 da sozlandi. DHCP pool 78.0.0.0/8 tarmog'i uchun yaratildi va default gateway hamda DNS server parametrlari kiritildi. Birinchi 10 ta IP manzil (78.0.0.1 — 78.0.0.10) statik foydalanish uchun ajratib qo'yildi.")

add_paragraph_justified(doc, "Router1 da DHCP sozlash uchun quyidagi buyruqlar ishlatildi:")

dhcp_code = """Router>enable
Router#configure terminal
Router(config)#ip dhcp pool VLAN1-POOL
Router(dhcp-config)#network 78.0.0.0 255.0.0.0
Router(dhcp-config)#default-router 78.0.0.254
Router(dhcp-config)#dns-server 78.0.0.100
Router(dhcp-config)#exit
Router(config)#ip dhcp excluded-address 78.0.0.1 78.0.0.10
Router(config)#end
Router#write memory"""

add_code_block(doc, dhcp_code)

doc.add_paragraph()

add_paragraph_justified(doc, "DNS xizmati Server0 qurilmasida sozlandi. Buning uchun Server0 ga statik IP 78.0.0.100 berildi va Services > DNS bo'limi orqali DNS xizmati yoqildi. DNS jadvalga server.local domen nomi qo'shildi va u 78.0.0.100 IP manziliga bog'landi.")

add_paragraph_justified(doc, "DNS sozlamalari quyidagicha kiritildi:")

dns_code = """Server0 > Services > DNS:
DNS Service: ON

Name: server.local
Type: A Record
Address: 78.0.0.100"""

add_code_block(doc, dns_code)

doc.add_paragraph()

add_paragraph_justified(doc, "Tarmoq ulanishini tekshirish uchun ping buyrug'idan foydalanildi. Quyidagi sinovlar muvaffaqiyatli o'tdi: PC0 dan Router1 ga (ping 78.0.0.254), PC0 dan PC5 ga (ping 78.0.0.6), PC0 dan Server0 ga (ping 78.0.0.100). Barcha hollarda Reply natijasi olindi va ma'lumot paketlari muvaffaqiyatli yetkazildi.")

doc.add_paragraph()

# 2.6
add_section_title(doc, "2.6. Inkapsulatsiya jarayonini tahlil qilish")

add_paragraph_justified(doc, "Cisco Packet Tracer dasturining Simulation rejimi orqali ma'lumot paketlarining inkapsulatsiya jarayoni o'rganildi. Buning uchun quyidagi qadamlar bajarildi: dastur pastki qismida Simulation tugmasi bosildi; PC0 dan PC5 ga ping 78.0.0.6 buyrug'i yuborildi; Auto Capture / Play tugmasi bosilib paketlar harakati kuzatildi; har bir paket ustiga bosib PDU Details oynasi ochildi.")

add_paragraph_justified(doc, "PDU Details oynasida har bir OSI qatlamida amalga oshirilgan inkapsulatsiya ko'rsatildi:")

add_paragraph_justified(doc, "• Layer 1 (Physical) — paket FastEthernet portlari orqali bitlar oqimi shaklida uzatiladi (masalan, FastEthernet0/3 → FastEthernet0/4);")

add_paragraph_justified(doc, "• Layer 2 (Data Link) — Ethernet II Header qo'shiladi va manba hamda manzil MAC manzillari ko'rsatiladi (masalan, 0001.C717.E6A1 >> 0001.C7A8.BE36);")

add_paragraph_justified(doc, "• Layer 3 (Network) — IP sarlavhasi qo'shiladi va manba IP (78.0.0.1) hamda manzil IP (78.0.0.6) belgilanadi;")

add_paragraph_justified(doc, "• Layer 4 (Transport) — ICMP yoki TCP/UDP sarlavhasi qo'shiladi (ping uchun ICMP);")

add_paragraph_justified(doc, "• Layer 5-7 (Session, Presentation, Application) — bu yuqori qatlamlarda foydalanuvchi ma'lumotlari shakllantiriladi va boshqariladi.")

add_paragraph_justified(doc, "Inkapsulatsiya jarayonida har bir qurilma (PC, Switch, Router) paketni qabul qilib, kerakli qatlamga qadar ochib (dekapsulatsiya), keyin yana inkapsulatsiya qilib uzatadi. Switch faqat Layer 2 ga qadar ochadi (MAC manzil bo'yicha yo'naltiradi), Router esa Layer 3 ga qadar ochadi (IP manzil bo'yicha yo'naltiradi).")

doc.add_paragraph()

# 2.7
add_section_title(doc, "2.7. Marshrutlash jadvalini qurish")

add_paragraph_justified(doc, "Quyida tarmoqdagi barcha qurilmalar uchun to'liq marshrutlash jadvali keltirilgan:")

# Marshrutlash jadvali
mr_table = doc.add_table(rows=12, cols=6)
mr_table.style = 'Light Grid Accent 1'
mr_headers = ['№', 'Qurilma', 'IP manzil', 'Tarmoq IP', 'Maska', 'Broadcast']
for i, h in enumerate(mr_headers):
    cell = mr_table.rows[0].cells[i]
    cell.text = h
    for run in cell.paragraphs[0].runs:
        run.bold = True
        run.font.size = Pt(11)

mr_data = [
    ('1', 'PC0', '78.0.0.1', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('2', 'PC1', '78.0.0.2', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('3', 'PC2', '78.0.0.3', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('4', 'PC3', '78.0.0.4', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('5', 'PC4', '78.0.0.5', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('6', 'PC5', '78.0.0.6', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('7', 'PC6', '78.0.0.7', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('8', 'Laptop0', '78.0.0.12', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('9', 'Tablet PC0', '78.0.0.14', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('10', 'Server0', '78.0.0.100', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
    ('11', 'Router1 (Fa0/0)', '78.0.0.254', '78.0.0.0', '255.0.0.0', '78.255.255.255'),
]

for i, row_data in enumerate(mr_data, start=1):
    for j, val in enumerate(row_data):
        cell = mr_table.rows[i].cells[j]
        cell.text = val
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(11)

doc.add_paragraph()

add_paragraph_justified(doc, "Yuqoridagi jadvaldan ko'rinib turibdiki, barcha qurilmalar bitta tarmoqda — 78.0.0.0/8 da joylashgan. Bu shuni anglatadiki, ular o'zaro to'g'ridan-to'g'ri muloqot qilishi mumkin va ular orasida marshrutlashtirish talab etilmaydi (Router faqat agar tashqi tarmoqlar bo'lsa kerak bo'ladi).")

doc.add_page_break()

# ==================================================================
# XULOSA
# ==================================================================

add_heading_centered(doc, "XULOSA", 14, True)
doc.add_paragraph()

xulosa_text = [
    "Men ushbu kurs ishi davomida Cisco Packet Tracer dasturidan foydalanib 5-variant shartlari asosida 78.0.0.0/8 tarmog'i bo'yicha kompyuter tarmog'ini yaratish va uni to'liq sozlash ishlarini amalga oshirdim. Ish jarayonida bir nechta switch, hub, router, server, kompyuter, laptop va tablet qurilmalaridan iborat gibrid topologiya tuzildi.",

    "Dastlab tarmoq topologiyasi qurildi va har bir qurilmaga 78.0.0.0/8 tarmog'idan IP manzil berildi. Router1 ga 78.0.0.254 IP manzil berildi va u barcha qurilmalar uchun default gateway bo'lib xizmat qildi. Server0 ga 78.0.0.100 IP berildi va u DNS server vazifasini bajardi. Maska sifatida 255.0.0.0 (/8) ishlatildi.",

    "Keyingi bosqichda kommutatorlarda VLAN sozlamalari bajarildi. Switch0, Switch1 va Switch2 da VLAN 1 (MANAGEMENT) va VLAN 2 (DATA) yaratildi. Barcha portlar access mode da sozlanib, mos VLAN ga biriktirildi. Bu tarmoqning xavfsizligini oshirish va segmentatsiya qilish imkonini berdi.",

    "OSPF marshrutlashtirish protokoli Router1 da ishga tushirildi. router ospf 1 va network 78.0.0.0 0.255.255.255 area 0 buyruqlari yordamida tarmoq OSPF ga qo'shildi. Bu protokol link-state algoritmiga asoslangan bo'lib, Dijkstra algoritmi yordamida eng qisqa yo'lni hisoblaydi.",

    "DHCP xizmati Router1 da sozlandi. VLAN1-POOL nomli DHCP pool yaratilib, network 78.0.0.0 255.0.0.0, default-router 78.0.0.254 va dns-server 78.0.0.100 parametrlari kiritildi. Birinchi 10 ta IP manzil (78.0.0.1 — 78.0.0.10) statik foydalanish uchun ajratib qo'yildi.",

    "DNS xizmati Server0 da sozlandi. Server0 ga statik IP 78.0.0.100 berildi, DNS xizmati yoqildi va server.local domen nomi 78.0.0.100 IP manziliga bog'landi. Shu orqali foydalanuvchilar serverga IP manzil orqali emas, balki domen nomi orqali murojaat qilish imkoniga ega bo'ldilar.",

    "Tarmoq ishini sinash uchun ping buyrug'idan foydalanildi. PC0 dan Router1 ga, PC0 dan PC5 ga va PC0 dan Server0 ga muvaffaqiyatli ping yuborildi. Barcha sinovlarda Reply javoblari olindi, bu esa tarmoqning to'g'ri ishlayotganini tasdiqladi.",

    "Inkapsulatsiya jarayoni Cisco Packet Tracer ning Simulation rejimi orqali o'rganildi. PDU Details oynasida har bir OSI qatlamida amalga oshirilgan inkapsulatsiya — Layer 1 da port belgilanishi, Layer 2 da MAC manzil va Ethernet II Header, Layer 3 da IP sarlavhasi va manzil IP — batafsil tahlil qilindi. Bu OSI modelining asosiy qatlamlari va ularning vazifalarini chuqurroq tushunishga yordam berdi.",

    "Ushbu kurs ishini bajarish jarayonida men tarmoq topologiyasini yaratish, router va switchlarni sozlash, OSPF marshrutlash protokolini qo'llash, DHCP, DNS xizmatlarini ishga tushirish, VLAN texnologiyasini amaliyotda qo'llash hamda inkapsulatsiya jarayonini tahlil qilish bo'yicha muhim bilim va ko'nikmalarga ega bo'ldim. Mazkur ish tarmoq texnologiyalarini chuqurroq tushunishga, real muhitga yaqin sharoitda tajriba orttirishga va kelajakdagi kasbiy faoliyatim uchun zarur amaliy bilimlarni mustahkamlashga xizmat qildi.",
]

for txt in xulosa_text:
    add_paragraph_justified(doc, txt)

doc.add_page_break()

# ==================================================================
# FOYDALANILGAN ADABIYOTLAR
# ==================================================================

add_heading_centered(doc, "FOYDALANILGAN ADABIYOTLAR", 14, True)
doc.add_paragraph()

adabiyotlar = [
    "1. E. Tanenbaum, D. Weatherall \"Kompyuter tarmoqlari\" 5-nashr. — Sankt-Peterburg: Piter, 2017. — 960 b.",
    "2. A. Sergeev \"Lokal kompyuter tarmoqlari asoslari\". — M.: Vilyams, 2016. — 384 b.",
    "3. W. Odom \"Rasmiy Cisco qo'llanmasiga tayyorgarlik ko'rish CCNA ICND2 200-101 sertifikatlash imtihonlari\". — Cisco Press, 2013. — 720 b.",
    "4. Astaxova I.F. \"Kompyuter fanlari. Daraxtlar, operatsion tizimlar, tarmoqlar\" / I.F. Astaxova, I.K. Astanin va boshqalar. — M.: Fizmatlit, 2013. — 88 b.",
    "5. Kuzin A.V. \"Kompyuter tarmoqlari\": Darslik / A.V. Kuzin. — M.: Forum, SIC INFRA-M, 2013. — 192 b.",
    "6. Kuzmenko N.G. \"Kompyuter tarmoqlari va tarmoq texnologiyalari\" / N.G. Kuzmenko. — Sankt-Peterburg: Fan va texnologiya, 2013. — 368 b.",
    "7. Cisco Networking Academy. \"Routing and Switching Essentials Companion Guide\". — Cisco Press, 2014. — 720 b.",
    "8. Olifer V.G., Olifer N.A. \"Kompyuter tarmoqlari. Prinsiplar, texnologiyalar, protokollar\": Universitetlar uchun darslik. 5-nashr. — Sankt-Peterburg: Piter, 2016. — 992 b.",
]

for adab in adabiyotlar:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    run = p.add_run(adab)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)

# Saqlash
output_path = '/projects/sandbox/Feruzbek/Kurs_ishi_5_variant.docx'
doc.save(output_path)
print(f"Word fayl yaratildi: {output_path}")
print(f"Sahifa soni: ~{len(doc.paragraphs) // 25}")
