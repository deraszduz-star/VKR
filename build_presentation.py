# -*- coding: utf-8 -*-
"""
Генератор презентации ВКР.
Тема: «Моделирование и анализ бизнес-процессов в образовательной организации».
Формат: 16:9, шрифт PT Sans, фирменная палитра КФУ.
Запуск:  python3 build_presentation.py
Результат: VKR_presentation.pptx
"""

from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from lxml import etree

# ---------- ПАЛИТРА ----------
PRIMARY  = RGBColor(0x17, 0x36, 0x5D)   # тёмно-синий КФУ
ACCENT   = RGBColor(0x4A, 0x7B, 0xB7)
SUCCESS  = RGBColor(0x9B, 0xBB, 0x59)
WARNING  = RGBColor(0xF7, 0x96, 0x46)
DANGER   = RGBColor(0xC0, 0x50, 0x4D)
LIGHT_BG = RGBColor(0xF2, 0xF2, 0xF2)
GREY     = RGBColor(0x59, 0x59, 0x59)
GOLD     = RGBColor(0xF2, 0xC8, 0x11)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK    = RGBColor(0x00, 0x00, 0x00)

FONT = "PT Sans"

# Слайд 16:9: 25.4 х 14.29 см
SLIDE_W = Cm(25.4)
SLIDE_H = Cm(14.29)


# ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------
def add_rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, line_w=None):
    sh = slide.shapes.add_shape(shape, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        if line_w:
            sh.line.width = line_w
    sh.shadow.inherit = False
    return sh


def add_text(slide, x, y, w, h, text, font_size=14, bold=False, italic=False,
             color=BLACK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font_name=FONT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
    return tb


def slide_header(slide, title, subtitle=None, slide_no=None):
    """Стандартная шапка слайда (для слайдов 2-9)."""
    # Левая цветная полоска-акцент
    add_rect(slide, Cm(0), Cm(0), Cm(0.4), SLIDE_H, PRIMARY)
    # Заголовок
    add_text(slide, Cm(0.7), Cm(0.3), Cm(20), Cm(1.0),
             title.upper(), font_size=22, bold=True, color=PRIMARY)
    # Подзаголовок
    if subtitle:
        add_text(slide, Cm(0.7), Cm(1.2), Cm(20), Cm(0.6),
                 subtitle, font_size=12, italic=True, color=GREY)
    # Линия-разделитель
    line = add_rect(slide, Cm(0.7), Cm(1.85), Cm(24.0), Cm(0.05), PRIMARY)
    line.line.fill.background()
    # Номер слайда
    if slide_no:
        add_text(slide, Cm(23.5), Cm(13.5), Cm(1.5), Cm(0.6),
                 str(slide_no), font_size=12, bold=True, color=PRIMARY,
                 align=PP_ALIGN.RIGHT)


def kpi_tile(slide, x, y, w, h, big, label, source=None,
             top_color=PRIMARY, big_color=None, big_size=44, label_size=11):
    """Карточка KPI: цветной верхний бордюр, гигантская цифра, подпись, источник."""
    # Подложка белая с лёгкой тенью (имитируем светлой рамкой)
    add_rect(slide, x, y, w, h, WHITE, line=LIGHT_BG, line_w=Pt(0.75))
    # Цветной верхний бордюр
    add_rect(slide, x, y, w, Cm(0.25), top_color)
    # Большая цифра
    big_color = big_color or top_color
    add_text(slide, x, y + Cm(0.5), w, Cm(2.0),
             big, font_size=big_size, bold=True, color=big_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Подпись
    add_text(slide, x + Cm(0.2), y + Cm(2.5), w - Cm(0.4), h - Cm(3.5),
             label, font_size=label_size, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    # Источник
    if source:
        add_text(slide, x + Cm(0.2), y + h - Cm(1.0), w - Cm(0.4), Cm(0.9),
                 source, font_size=9, italic=True, color=GREY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)


def bottom_band(slide, text, y=None, color=PRIMARY, txt_color=WHITE, font_size=13):
    """Нижняя «итоговая» плашка."""
    if y is None:
        y = Cm(12.6)
    band = add_rect(slide, Cm(0.7), y, Cm(24.0), Cm(0.85), color)
    add_text(slide, Cm(0.8), y, Cm(23.8), Cm(0.85),
             text, font_size=font_size, bold=True, italic=True, color=txt_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- СБОРКА ПРЕЗЕНТАЦИИ ----------
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
blank = prs.slide_layouts[6]


# =================================================================
# СЛАЙД 1. ТИТУЛЬНЫЙ
# =================================================================
slide = prs.slides.add_slide(blank)
# Тёмно-синий фон
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)
# Декоративная золотая полоса слева
add_rect(slide, Cm(0), Cm(0), Cm(0.6), SLIDE_H, GOLD)

# Шапка университета
add_text(slide, Cm(1.5), Cm(0.8), Cm(22.4), Cm(1.5),
         "Министерство науки и высшего образования Российской Федерации\n"
         "ФГАОУ ВО «Казанский (Приволжский) федеральный университет»\n"
         "Институт управления, экономики и финансов",
         font_size=12, color=WHITE, align=PP_ALIGN.CENTER)

# Тип работы
add_text(slide, Cm(1.5), Cm(3.5), Cm(22.4), Cm(0.8),
         "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
         font_size=14, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

# Тема — главный блок
add_text(slide, Cm(1.5), Cm(4.8), Cm(22.4), Cm(2.5),
         "Моделирование и анализ\nбизнес-процессов\nв образовательной организации",
         font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Разделитель
add_rect(slide, Cm(10.7), Cm(8.5), Cm(4.0), Cm(0.06), GOLD)

# Автор и научный руководитель
add_text(slide, Cm(2.0), Cm(9.4), Cm(10.0), Cm(2.5),
         "ВЫПОЛНИЛ:\n"
         "ФИО студента\n"
         "группа __________\n"
         "направление 38.03.05",
         font_size=12, color=WHITE, align=PP_ALIGN.LEFT)

add_text(slide, Cm(13.5), Cm(9.4), Cm(10.0), Cm(2.5),
         "НАУЧНЫЙ РУКОВОДИТЕЛЬ:\n"
         "ФИО, учёная степень,\n"
         "учёное звание, должность",
         font_size=12, color=WHITE, align=PP_ALIGN.LEFT)

# Подвал
add_text(slide, Cm(1.5), Cm(13.2), Cm(22.4), Cm(0.8),
         "Казань — 2026",
         font_size=12, italic=True, color=GOLD, align=PP_ALIGN.CENTER)


# =================================================================
# СЛАЙД 2. АКТУАЛЬНОСТЬ — 3 KPI-плитки
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Актуальность", slide_no=2)

# 3 плитки
tile_w = Cm(7.4)
tile_h = Cm(8.5)
y_tile = Cm(2.5)
gap = Cm(0.4)
x0 = Cm(0.7)

# Плитка 1
kpi_tile(slide, x0, y_tile, tile_w, tile_h,
         "95 %", "управленческих процессов\nвузов в электронном виде\nк 2030 году",
         source="Распоряжение Правительства РФ\n№ 3759-р от 21.12.2021",
         top_color=PRIMARY, big_size=60)

# Плитка 2
kpi_tile(slide, x0 + tile_w + gap, y_tile, tile_w, tile_h,
         "29 %", "российских вузов —\n«цифровые аутсайдеры»\n(48 % — последователи,\n23 % — лидеры)",
         source="НИУ ВШЭ, исследование\nцифровой трансформации ВО",
         top_color=WARNING, big_color=WARNING, big_size=60)

# Плитка 3
kpi_tile(slide, x0 + 2*(tile_w + gap), y_tile, tile_w, tile_h,
         "64 %", "вузов: главный барьер\nцифровизации — дефицит\nИТ-кадров (52 % — финансы,\n47 % — низкая грамотность)",
         source="НИУ ВШЭ, исследование\nцифровой трансформации ВО",
         top_color=DANGER, big_color=DANGER, big_size=60)

bottom_band(slide,
    "Нормативное давление + низкая зрелость + кадровый барьер  =  "
    "необходимость моделирования и анализа БП")


# =================================================================
# СЛАЙД 3. ЦЕЛЬ И ЗАДАЧИ
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Цель и задачи исследования", slide_no=3)

# Плашка ЦЕЛЬ
goal_y = Cm(2.4)
add_rect(slide, Cm(0.7), goal_y, Cm(24.0), Cm(2.5), PRIMARY)
add_text(slide, Cm(1.0), goal_y, Cm(2.5), Cm(2.5),
         "ЦЕЛЬ", font_size=22, bold=True, color=GOLD,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, Cm(4.0), goal_y + Cm(0.2), Cm(20.5), Cm(2.1),
         "Изучение теоретических основ моделирования и анализа бизнес-процессов и их "
         "практическое применение в деятельности структурных подразделений КФУ "
         "(Дирекция музеев, Управление кадров, Управление документооборота).",
         font_size=13, color=WHITE, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

# 5 задач — круги-ступени
tasks = [
    ("1", "ТЕОРИЯ", "IDEF0 · BPMN · EPC\nLean: 5С · 7 потерь\nДЦ/НДБ/НДЦ"),
    ("2", "ОТРАСЛЬ", "Анализ ВО РФ\nКФУ vs 7 ФУ\nКритерии отбора"),
    ("3", "AS IS", "9 моделей\nУзкие места\nДоля НДЦ/НДБ"),
    ("4", "TO BE", "9 моделей\n5С · Битрикс24\nOCR · QR-маршрут"),
    ("5", "ЭКОНОМИКА", "NPV · IRR · DPP\nЧувствительность\nПорядок внедрения"),
]

n = len(tasks)
total_w = Cm(24.0)
circle_d = Cm(3.6)
arrow_w = Cm(0.6)
gap_total = total_w - n * circle_d
gap = Emu((gap_total - arrow_w * (n - 1)) / (n - 1))
y_circles = Cm(6.0)

# Соединительные стрелки (фоном)
for i in range(n - 1):
    cx = Cm(0.7) + circle_d + i * (circle_d + gap + arrow_w)
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, cx, y_circles + Cm(1.5),
                                   gap + arrow_w, Cm(0.6))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = ACCENT
    arrow.line.fill.background()

# Круги задач (с градиентом цвета по индексу)
task_colors = [
    RGBColor(0x4A, 0x7B, 0xB7),
    RGBColor(0x3A, 0x68, 0xA0),
    RGBColor(0x2C, 0x55, 0x88),
    RGBColor(0x22, 0x46, 0x72),
    RGBColor(0x17, 0x36, 0x5D),
]

for i, (num, ttl, body) in enumerate(tasks):
    cx = Cm(0.7) + i * (circle_d + gap + arrow_w)
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, y_circles, circle_d, circle_d)
    c.fill.solid()
    c.fill.fore_color.rgb = task_colors[i]
    c.line.color.rgb = WHITE
    c.line.width = Pt(2)
    # Номер
    add_text(slide, cx, y_circles + Cm(0.2), circle_d, Cm(1.2),
             num, font_size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Подзаголовок задачи
    add_text(slide, cx, y_circles + Cm(1.4), circle_d, Cm(0.8),
             ttl, font_size=11, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Тело — под кругом
    add_text(slide, cx - Cm(0.3), y_circles + circle_d + Cm(0.2), circle_d + Cm(0.6), Cm(2.8),
             body, font_size=10, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)


# =================================================================
# СЛАЙД 4. ТЕОРЕТИКО-МЕТОДОЛОГИЧЕСКАЯ БАЗА (Задача 1)
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Теоретико-методологическая база",
             subtitle="Результаты по задаче 1", slide_no=4)

# Пирамида (3 трапеции) слева
pyr_x = Cm(1.0); pyr_y = Cm(2.5); pyr_h = Cm(8.5)
# Трапеция 1 (IDEF0) — узкая сверху
t1 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(2.5), pyr_y, Cm(3.5), Cm(2.5))
t1.fill.solid(); t1.fill.fore_color.rgb = PRIMARY; t1.line.color.rgb = WHITE; t1.line.width = Pt(2)
add_text(slide, pyr_x + Cm(2.5), pyr_y + Cm(0.3), Cm(3.5), Cm(2.0),
         "IDEF0\nстратегия", font_size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

t2 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(1.5), pyr_y + Cm(2.5), Cm(5.5), Cm(2.7))
t2.fill.solid(); t2.fill.fore_color.rgb = ACCENT; t2.line.color.rgb = WHITE; t2.line.width = Pt(2)
add_text(slide, pyr_x + Cm(1.5), pyr_y + Cm(2.7), Cm(5.5), Cm(2.3),
         "BPMN 2.0\nоперации, исполнители", font_size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

t3 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(0.3), pyr_y + Cm(5.2), Cm(7.9), Cm(3.0))
t3.fill.solid(); t3.fill.fore_color.rgb = WARNING; t3.line.color.rgb = WHITE; t3.line.width = Pt(2)
add_text(slide, pyr_x + Cm(0.3), pyr_y + Cm(5.4), Cm(7.9), Cm(2.6),
         "EPC\nдетализация, узкие места", font_size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 3 карточки справа
card_x = Cm(10.0); card_w = Cm(4.7); card_h = Cm(8.5); card_y = Cm(2.5); card_gap = Cm(0.2)
card_data = [
    ("НОТАЦИИ", "•  IDEF0\n   FIPS PUB 183 (NIST)\n•  BPMN 2.0\n   OMG, 2011\n•  EPC\n   Business Studio", PRIMARY),
    ("LEAN", "•  7 видов потерь\n   (Тайити Оно)\n•  ДЦ / НДБ / НДЦ\n•  Диаграмма Ишикавы\n•  5С\n•  DMAIC", ACCENT),
    ("ЭКОНОМИКА", "•  NPV\n•  IRR\n•  DPP\n•  Анализ\n   чувствительности\n•  SMART", DANGER),
]

for i, (ttl, body, col) in enumerate(card_data):
    cx = card_x + i * (card_w + card_gap)
    add_rect(slide, cx, card_y, card_w, card_h, WHITE, line=col, line_w=Pt(1.5))
    # Шапка
    add_rect(slide, cx, card_y, card_w, Cm(1.2), col)
    add_text(slide, cx, card_y, card_w, Cm(1.2),
             ttl, font_size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx + Cm(0.3), card_y + Cm(1.4), card_w - Cm(0.6), card_h - Cm(1.6),
             body, font_size=12, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "Стек: НОТАЦИИ + LEAN + ИНВЕСТИЦИОННЫЙ АНАЛИЗ — применён к 3 подразделениям КФУ")


# =================================================================
# СЛАЙД 5. ПОЛОЖЕНИЕ КФУ (Задача 2)
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Положение КФУ в отрасли",
             subtitle="Результаты по задаче 2", slide_no=5)

# Радар-чарт слева
chart_data = CategoryChartData()
chart_data.categories = ['Балл ЕГЭ', 'Иностр., %', 'НИОКР, млрд', 'НПР со ст.', 'Внебюдж., %']
chart_data.add_series('КФУ', (74.2, 19.1, 25.3, 4.5, 46.2))      # масштаб НИОКР *10
chart_data.add_series('Медиана РФ', (62.8, 4.9, 3.0, 3.1, 25.0))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.RADAR_FILLED, Cm(0.7), Cm(2.3),
    Cm(11.5), Cm(8.5), chart_data
).chart
chart.has_title = True
chart.chart_title.text_frame.text = "КФУ vs медиана 7 ФУ РФ (2025)"
for run in chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(12); run.font.bold = True; run.font.color.rgb = PRIMARY
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
# Цвета серий
ser1 = chart.series[0]; ser1.format.fill.solid(); ser1.format.fill.fore_color.rgb = PRIMARY
ser1.format.line.color.rgb = PRIMARY
ser2 = chart.series[1]; ser2.format.fill.solid(); ser2.format.fill.fore_color.rgb = WARNING
ser2.format.line.color.rgb = WARNING

# Карточки лидерства справа
lead_x = Cm(13.0); lead_y = Cm(2.3); lead_w = Cm(11.7)
add_rect(slide, lead_x, lead_y, lead_w, Cm(0.8), PRIMARY)
add_text(slide, lead_x, lead_y, lead_w, Cm(0.8),
         "ЛИДЕРСТВО КФУ СРЕДИ 7 ФЕДЕРАЛЬНЫХ ВУЗОВ",
         font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

leaders = [
    ("Балл ЕГЭ", "74,2", "★ 1/7"),
    ("Иностр. студенты", "19,1 %", "★ 1/7"),
    ("Внебюджет", "46,2 %", "★ 2/7"),
    ("НИОКР, млрд ₽", "2,53", "★ 2/7"),
]
for i, (ttl, val, place) in enumerate(leaders):
    row_y = lead_y + Cm(1.0) + i * Cm(1.0)
    add_rect(slide, lead_x, row_y, lead_w, Cm(0.9), LIGHT_BG)
    add_text(slide, lead_x + Cm(0.3), row_y, Cm(5.5), Cm(0.9),
             ttl, font_size=12, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, lead_x + Cm(5.5), row_y, Cm(3.5), Cm(0.9),
             val, font_size=14, bold=True, color=PRIMARY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, lead_x + Cm(8.7), row_y, Cm(2.7), Cm(0.9),
             place, font_size=13, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Объекты исследования (мини-карточки внизу)
obj_y = Cm(11.0)
# (убираем — переносим в подвал, чтобы не перегружать)
bottom_band(slide,
    "🏛 Дирекция музеев (800 тыс. ед.)   ·   👥 ЦРКП (6 615 ППС)   ·   "
    "📄 УДК (170 тыс. ед. архива)")


# =================================================================
# СЛАЙД 6. AS IS (Задача 3)
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Анализ текущего состояния (AS IS)",
             subtitle="Результаты по задаче 3", slide_no=6)

# 3 карточки подразделений с «термометром»
units = [
    ("Дирекция музеев", "45 мин", "на экспонат",
     "•  ожидание ПО 10 мин\n•  9 экспонатов / день\n•  ручная фотосъёмка\n•  отсутствие шаблонов",
     1.0),  # «заполнение» 100%
    ("ЦРКП  (Управление кадров)", "5 дней", "на заявление ППС",
     "•  20 сотрудников\n•  40 ч/мес транспорт\n•  бумажный документооборот\n•  192 мин потерь",
     1.0),
    ("УДК (Документооборот)", "22 мин", "на справку",
     "•  архив оцифрован 12 %\n•  ОРД 7 ч ожидания\n•  176 сек / регистрация\n•  ручной поиск",
     0.7),
]

card_w = Cm(7.7); card_h = Cm(7.0); card_y = Cm(2.4); card_gap = Cm(0.3)
for i, (ttl, big, sub, body, fill_pct) in enumerate(units):
    cx = Cm(0.7) + i * (card_w + card_gap)
    add_rect(slide, cx, card_y, card_w, card_h, WHITE, line=DANGER, line_w=Pt(1.5))
    # Шапка
    add_rect(slide, cx, card_y, card_w, Cm(1.0), DANGER)
    add_text(slide, cx, card_y, card_w, Cm(1.0),
             ttl, font_size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # «Термометр» слева
    therm_x = cx + Cm(0.4); therm_y = card_y + Cm(1.4); therm_h = Cm(4.5)
    add_rect(slide, therm_x, therm_y, Cm(0.5), therm_h, LIGHT_BG)
    fill_h = Emu(int(therm_h * fill_pct))
    add_rect(slide, therm_x, therm_y + therm_h - fill_h, Cm(0.5), fill_h, DANGER)
    # Большая цифра
    add_text(slide, cx + Cm(1.2), card_y + Cm(1.4), card_w - Cm(1.4), Cm(1.6),
             big, font_size=30, bold=True, color=DANGER,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx + Cm(1.2), card_y + Cm(2.9), card_w - Cm(1.4), Cm(0.6),
             sub, font_size=11, italic=True, color=GREY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    # Список причин
    add_text(slide, cx + Cm(1.2), card_y + Cm(3.7), card_w - Cm(1.4), card_h - Cm(3.9),
             body, font_size=11, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

# Stacked bar — структура операций
sb_y = Cm(9.7)
add_text(slide, Cm(0.7), sb_y, Cm(24.0), Cm(0.6),
         "СТРУКТУРА ОПЕРАЦИЙ В ИССЛЕДОВАННЫХ ПРОЦЕССАХ:",
         font_size=11, bold=True, color=PRIMARY)
total_w_bar = Cm(24.0)
bar_h = Cm(1.1)
bar_y = sb_y + Cm(0.7)
# Сегменты
seg_ndc = Emu(int(total_w_bar * 0.64))
seg_ndb = Emu(int(total_w_bar * 0.24))
seg_dc  = Emu(total_w_bar - seg_ndc - seg_ndb)
add_rect(slide, Cm(0.7),                              bar_y, seg_ndc, bar_h, DANGER)
add_rect(slide, Cm(0.7) + seg_ndc,                    bar_y, seg_ndb, bar_h, WARNING)
add_rect(slide, Cm(0.7) + seg_ndc + seg_ndb,          bar_y, seg_dc,  bar_h, SUCCESS)
add_text(slide, Cm(0.7),                              bar_y, seg_ndc, bar_h,
         "НДЦ — устранить  64 %", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, Cm(0.7) + seg_ndc,                    bar_y, seg_ndb, bar_h,
         "НДБ — авто  24 %", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, Cm(0.7) + seg_ndc + seg_ndb,          bar_y, seg_dc,  bar_h,
         "ДЦ  12 %", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

bottom_band(slide,
    "9 моделей  ·  6 из 7 видов потерь Тайити Оно  ·  3 узких места: "
    "ручной ввод · бумага · ожидание ИТ")


# =================================================================
# СЛАЙД 7. TO BE (Задача 4)
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Проектирование целевых процессов (TO BE)",
             subtitle="Результаты по задаче 4", slide_no=7)

# 3 строки: «AS IS → TO BE → решение»
rows = [
    ("Дирекция музеев", "45 мин", "9 экс/день",
     "18 мин", "30 экс/день",
     "5С + шаблоны\nпривлечение лаборанта",
     "65 тыс. ₽", "+233 % произв."),
    ("ЦРКП", "5 дней", "20 чел.",
     "20 мин", "3 чел.",
     "Битрикс24\nэлектронная подпись",
     "450 тыс. ₽", "−17 шт. ед."),
    ("УДК", "12 % архива", "ОРД 7 ч",
     "100 % архива", "ОРД −49,9 %",
     "OCR + студенты\nQR + ИИ-проверка",
     "277,5 тыс. ₽", "регистр. +21,9 %"),
]

row_h = Cm(2.7)
row_y0 = Cm(2.4)
row_gap = Cm(0.2)
for i, r in enumerate(rows):
    (unit, asis_a, asis_b, tobe_a, tobe_b, sol, inv, gain) = r
    ry = row_y0 + i * (row_h + row_gap)
    # Подложка строки
    add_rect(slide, Cm(0.7), ry, Cm(24.0), row_h, LIGHT_BG)
    # Левая шапка-подразделение
    add_rect(slide, Cm(0.7), ry, Cm(3.5), row_h, PRIMARY)
    add_text(slide, Cm(0.7), ry, Cm(3.5), row_h,
             unit, font_size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # AS IS — красная плитка
    asis_x = Cm(4.4); asis_w = Cm(5.0)
    add_rect(slide, asis_x, ry + Cm(0.3), asis_w, row_h - Cm(0.6), WHITE, line=DANGER, line_w=Pt(1.5))
    add_text(slide, asis_x, ry + Cm(0.3), asis_w, Cm(0.6),
             "AS IS", font_size=10, bold=True, color=DANGER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    add_text(slide, asis_x, ry + Cm(0.8), asis_w, Cm(1.0),
             asis_a, font_size=20, bold=True, color=DANGER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, asis_x, ry + Cm(1.8), asis_w, Cm(0.5),
             asis_b, font_size=11, color=GREY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    # Стрелка
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                  asis_x + asis_w + Cm(0.2), ry + Cm(0.9), Cm(1.0), Cm(0.8))
    arr.fill.solid(); arr.fill.fore_color.rgb = PRIMARY; arr.line.fill.background()
    # TO BE — зелёная плитка
    tobe_x = asis_x + asis_w + Cm(1.4); tobe_w = Cm(5.0)
    add_rect(slide, tobe_x, ry + Cm(0.3), tobe_w, row_h - Cm(0.6), WHITE, line=SUCCESS, line_w=Pt(1.5))
    add_text(slide, tobe_x, ry + Cm(0.3), tobe_w, Cm(0.6),
             "TO BE", font_size=10, bold=True, color=SUCCESS,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    add_text(slide, tobe_x, ry + Cm(0.8), tobe_w, Cm(1.0),
             tobe_a, font_size=20, bold=True, color=SUCCESS,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, tobe_x, ry + Cm(1.8), tobe_w, Cm(0.5),
             tobe_b, font_size=11, color=GREY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    # Решение и инвестиции
    sol_x = tobe_x + tobe_w + Cm(0.3); sol_w = Cm(8.5)
    add_text(slide, sol_x, ry + Cm(0.2), sol_w, Cm(1.4),
             sol, font_size=11, bold=True, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    add_text(slide, sol_x, ry + Cm(1.6), sol_w, Cm(0.5),
             f"I₀ {inv}", font_size=11, italic=True, color=PRIMARY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    add_text(slide, sol_x, ry + Cm(2.1), sol_w, Cm(0.5),
             f"▲ {gain}", font_size=11, bold=True, color=SUCCESS,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "9 моделей TO BE  ·  совокупные инвестиции 792,5 тыс. ₽  ·  "
    "3 узких места = 3 класса решений (lean + cloud + OCR)")


# =================================================================
# СЛАЙД 8. ЭКОНОМИКА (Задача 5)
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Экономическая эффективность",
             subtitle="Результаты по задаче 5", slide_no=8)

# 5 KPI плиток сверху
kpis = [
    ("I₀",      "792,5",  "тыс. ₽"),
    ("ΔCF/год", "1 709,7", "тыс. ₽"),
    ("NPV",     "4 000",   "тыс. ₽ (5 лет)"),
    ("IRR",     "215 %",   "ставка 23 %"),
    ("DPP",     "≈ 7",     "месяцев"),
]
kpi_y = Cm(2.3); kpi_h = Cm(2.3); kpi_w = Cm(4.7); kpi_gap = Cm(0.15)
for i, (lbl, val, sub) in enumerate(kpis):
    cx = Cm(0.7) + i * (kpi_w + kpi_gap)
    add_rect(slide, cx, kpi_y, kpi_w, kpi_h, PRIMARY)
    add_text(slide, cx, kpi_y + Cm(0.05), kpi_w, Cm(0.5),
             lbl, font_size=11, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, kpi_y + Cm(0.5), kpi_w, Cm(1.2),
             val, font_size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, kpi_y + Cm(1.65), kpi_w, Cm(0.6),
             sub, font_size=10, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Линейный график — накопительный NPV
line_data = CategoryChartData()
line_data.categories = ['0', '1', '2', '3', '4', '5']
line_data.add_series('Накопительный NPV, млн ₽', (-0.79, 0.80, 1.99, 3.06, 3.55, 4.00))
line_chart = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS, Cm(0.7), Cm(5.0),
    Cm(13.0), Cm(7.5), line_data
).chart
line_chart.has_title = True
line_chart.chart_title.text_frame.text = "Накопительный NPV (горизонт 5 лет, r = 23 %)"
for run in line_chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(12); run.font.bold = True; run.font.color.rgb = PRIMARY
line_chart.has_legend = False
ser = line_chart.series[0]
ser.format.line.color.rgb = PRIMARY
ser.format.line.width = Pt(3)
# Точки-маркеры
for pt in ser.points:
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = PRIMARY
# Подписи данных
ser.data_labels.show_value = True
ser.data_labels.font.size = Pt(10)
ser.data_labels.font.bold = True
ser.data_labels.font.color.rgb = PRIMARY

# Tornado-bar — чувствительность
tornado_data = CategoryChartData()
tornado_data.categories = ['−50 % CF', 'Песс. CF+r', '−30 % CF', '+50 % затр.', '−18 % r=30 %', 'Базовый']
tornado_data.add_series('NPV, тыс. ₽', (1596, 2020, 2543, 3604, 3261, 4000))
tor_chart = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Cm(14.0), Cm(5.0),
    Cm(10.7), Cm(7.5), tornado_data
).chart
tor_chart.has_title = True
tor_chart.chart_title.text_frame.text = "Анализ чувствительности NPV"
for run in tor_chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(12); run.font.bold = True; run.font.color.rgb = PRIMARY
tor_chart.has_legend = False
plot = tor_chart.plots[0]
plot.gap_width = 60
ser2 = tor_chart.series[0]
# Цвета баров: первые — красные, базовый — зелёный, средние — оранжевые
bar_colors = [DANGER, DANGER, WARNING, WARNING, ACCENT, SUCCESS]
for idx, pt in enumerate(ser2.points):
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = bar_colors[idx]
ser2.data_labels.show_value = True
ser2.data_labels.font.size = Pt(10)
ser2.data_labels.font.bold = True

bottom_band(slide,
    "Очерёдность внедрения:   ① ЦРКП  →  ② УДК  →  ③ МУЗЕИ   ·   запас прочности 83 %")


# =================================================================
# СЛАЙД 9. ВЫВОДЫ
# =================================================================
slide = prs.slides.add_slide(blank)
slide_header(slide, "Выводы",
             subtitle="Цель достигнута · 5 задач решены · методика воспроизводима",
             slide_no=9)

# 5 плиток-достижений: 3 в верхнем ряду, 2 в нижнем
achievements = [
    ("МЕТОДИКА",   "3 уровня",      "IDEF0 + BPMN + EPC", PRIMARY),
    ("ОТРАСЛЬ",    "★ 1/7",          "КФУ среди ФУ РФ",      ACCENT),
    ("AS IS",      "9",              "моделей · 64 % НДЦ",   WARNING),
    ("TO BE",      "−17",            "штатных единиц ЦРКП",  SUCCESS),
    ("ЭКОНОМИКА",  "4,0 млн ₽",       "NPV · IRR 215 % · DPP 7 мес.", DANGER),
]

# Верхний ряд: 3 плитки
top_w = Cm(7.7); top_h = Cm(4.7); top_y = Cm(2.5); top_gap = Cm(0.3)
for i in range(3):
    ttl, big, body, col = achievements[i]
    cx = Cm(0.7) + i * (top_w + top_gap)
    add_rect(slide, cx, top_y, top_w, top_h, WHITE, line=col, line_w=Pt(1.5))
    add_rect(slide, cx, top_y, top_w, Cm(0.9), col)
    add_text(slide, cx, top_y, top_w, Cm(0.9),
             ttl, font_size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, top_y + Cm(1.1), top_w, Cm(2.0),
             big, font_size=36, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, top_y + Cm(3.2), top_w, Cm(1.4),
             body, font_size=12, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

# Нижний ряд: 2 плитки центрированно
bot_w = Cm(11.5); bot_h = Cm(4.5); bot_y = Cm(7.5); bot_gap = Cm(0.5)
total_bot = 2 * bot_w + bot_gap
bot_x0 = (SLIDE_W - total_bot) / 2
for i, j in enumerate([3, 4]):
    ttl, big, body, col = achievements[j]
    cx = bot_x0 + i * (bot_w + bot_gap)
    add_rect(slide, cx, bot_y, bot_w, bot_h, WHITE, line=col, line_w=Pt(1.5))
    add_rect(slide, cx, bot_y, bot_w, Cm(0.9), col)
    add_text(slide, cx, bot_y, bot_w, Cm(0.9),
             ttl, font_size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, bot_y + Cm(1.1), bot_w, Cm(1.8),
             big, font_size=36, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, bot_y + Cm(3.0), bot_w, Cm(1.4),
             body, font_size=12, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "Результаты применимы в КФУ и тиражируемы на 7 федеральных университетов РФ")


# =================================================================
# СЛАЙД 10. СПАСИБО
# =================================================================
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)
add_rect(slide, Cm(0), Cm(0), Cm(0.6), SLIDE_H, GOLD)

add_text(slide, Cm(1), Cm(5.0), Cm(23.4), Cm(2.5),
         "СПАСИБО ЗА ВНИМАНИЕ", font_size=44, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(slide, Cm(11.2), Cm(7.6), Cm(3.0), Cm(0.06), GOLD)
add_text(slide, Cm(1), Cm(8.0), Cm(23.4), Cm(1.0),
         "Готов ответить на ваши вопросы", font_size=18, italic=True, color=GOLD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- СОХРАНЕНИЕ ----------
out = "VKR_presentation.pptx"
prs.save(out)
print(f"OK: {out}")
