# -*- coding: utf-8 -*-
"""
Генератор презентации ВКР НА ОСНОВЕ ШАБЛОНА «ВКР ИиИ_2026».
Тема: «Моделирование и анализ бизнес-процессов в образовательной организации».
Формат: 16:9 (25.4 х 14.29 см), фирменный стиль шаблона ИиИ КФУ.
Запуск:  python3 build_presentation.py
Результат: VKR_presentation.pptx

Логика:
1. Открываем файл «Шаблон презентации ВКР ИиИ_2026 (2).pptx».
2. На каждом слайде шаблона удаляем плейсхолдеры с заголовком/контентом
   (но сохраняем все декоративные картинки: вертикальную полосу слева, лого,
   круглый акцент, номер слайда).
3. В освободившуюся зону (x>=2.85 см, w<=22 см) рисуем наш контент:
   плитки KPI, диаграммы, таблицы AS IS/TO BE, выводы.
"""

import os
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# ---------- ПАЛИТРА (совпадает с темой шаблона) ----------
PRIMARY  = RGBColor(0x17, 0x36, 0x5D)   # тёмно-синий КФУ (= шаблон)
ACCENT   = RGBColor(0x4A, 0x7B, 0xB7)
SUCCESS  = RGBColor(0x9B, 0xBB, 0x59)
WARNING  = RGBColor(0xF7, 0x96, 0x46)
DANGER   = RGBColor(0xC0, 0x50, 0x4D)
LIGHT_BG = RGBColor(0xF2, 0xF2, 0xF2)
GREY     = RGBColor(0x59, 0x59, 0x59)
GOLD     = RGBColor(0xF2, 0xC8, 0x11)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK    = RGBColor(0x00, 0x00, 0x00)

FONT = "Arial"  # шаблон использует Arial

# Слайд 16:9: 25.4 х 14.29 см
SLIDE_W = Cm(25.4)
SLIDE_H = Cm(14.29)

# Контент-зона шаблона (декор шаблона занимает левую полосу 0..2.5 см)
LEFT  = Cm(2.85)         # левый край контента
RIGHT = Cm(25.0)         # правый край контента
CW    = RIGHT - LEFT     # ширина контента  (= ~22.15 см)
TOP   = Cm(2.55)         # верхний край контента (под заголовком слайда)
BOT   = Cm(13.0)         # нижний край контента (над футером)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "Шаблон презентации ВКР ИиИ_2026 (2).pptx")
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "VKR_presentation.pptx")


# ---------- ОЧИСТКА ПЛЕЙСХОЛДЕРОВ ШАБЛОНА ----------
def clear_template_placeholders(slide, keep_slide_no=True):
    """Удаляем все плейсхолдеры с текстом (шаблонные заголовки и подвалы),
    оставляя картинки (декор: полоса, лого, круг, фон)."""
    to_remove = []
    for shape in list(slide.shapes):
        # MSO_SHAPE_TYPE: 13 = PICTURE, 14 = PLACEHOLDER, 17 = TEXT_BOX, 9 = LINE
        if shape.shape_type == 14:  # placeholder — почти всегда текстовая шапка
            to_remove.append(shape)
        elif shape.shape_type == 17:  # textbox шаблона ("Казань — 2026", ФИО и пр.)
            # Сохраняем номер слайда (textbox с именем "Номер слайда" + единственная цифра)
            if keep_slide_no and "Номер слайда" in (shape.name or ""):
                continue
            to_remove.append(shape)
        elif shape.shape_type == 9:  # LINE-разделитель на титуле
            to_remove.append(shape)
    for sh in to_remove:
        sh._element.getparent().remove(sh._element)


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


def slide_header(slide, title, subtitle=None):
    """Заголовок и подзаголовок слайда — в стилистике шаблона
    (заголовок капсом, тёмно-синий, выровнен в шапке шаблона: y≈0.5 см)."""
    add_text(slide, LEFT, Cm(0.5), CW, Cm(1.3),
             title.upper(), font_size=22, bold=True, color=PRIMARY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, LEFT, Cm(1.65), CW, Cm(0.6),
                 subtitle, font_size=12, italic=True, color=GREY,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    # Тонкая линия-разделитель в фирменных цветах
    add_rect(slide, LEFT, Cm(2.30), CW, Cm(0.05), PRIMARY)


def kpi_tile(slide, x, y, w, h, big, label, source=None,
             top_color=PRIMARY, big_color=None, big_size=44, label_size=11):
    add_rect(slide, x, y, w, h, WHITE, line=LIGHT_BG, line_w=Pt(0.75))
    add_rect(slide, x, y, w, Cm(0.25), top_color)
    big_color = big_color or top_color
    add_text(slide, x, y + Cm(0.5), w, Cm(2.0),
             big, font_size=big_size, bold=True, color=big_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, x + Cm(0.2), y + Cm(2.5), w - Cm(0.4), h - Cm(3.5),
             label, font_size=label_size, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    if source:
        add_text(slide, x + Cm(0.2), y + h - Cm(1.0), w - Cm(0.4), Cm(0.9),
                 source, font_size=9, italic=True, color=GREY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)


def bottom_band(slide, text, y=None, color=PRIMARY, txt_color=WHITE, font_size=12):
    """Нижняя итоговая плашка (слегка приподнята, чтобы не закрыть номер слайда шаблона)."""
    if y is None:
        y = Cm(12.30)
    band = add_rect(slide, LEFT, y, CW, Cm(0.85), color)
    add_text(slide, LEFT, y, CW, Cm(0.85),
             text, font_size=font_size, bold=True, italic=True, color=txt_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- ОТКРЫТИЕ ШАБЛОНА ----------
prs = Presentation(TEMPLATE_PATH)
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

slides = list(prs.slides)
assert len(slides) >= 10, "Шаблон должен содержать минимум 10 слайдов"
# Очистим шаблонный текст на каждом слайде
for s in slides[:10]:
    clear_template_placeholders(s)


# =================================================================
# СЛАЙД 1. ТИТУЛЬНЫЙ — фон шаблона уже на месте (image1.png).
# Поверх добавляем только текстовые блоки.
# =================================================================
slide = slides[0]
# Фон шаблонной картинки тёмный → текст белым/золотым.
# Шапка университета — сверху по центру
add_text(slide, Cm(1.5), Cm(0.5), Cm(22.4), Cm(1.5),
         "Министерство науки и высшего образования Российской Федерации\n"
         "ФГАОУ ВО «Казанский (Приволжский) федеральный университет»\n"
         "Институт управления, экономики и финансов",
         font_size=11, color=WHITE, align=PP_ALIGN.CENTER)

# Тип работы
add_text(slide, Cm(1.5), Cm(2.5), Cm(22.4), Cm(0.7),
         "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
         font_size=13, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

# Тема
add_text(slide, Cm(1.5), Cm(3.4), Cm(22.4), Cm(3.0),
         "Моделирование и анализ\nбизнес-процессов\nв образовательной организации",
         font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Декоративный разделитель
add_rect(slide, Cm(11.2), Cm(6.8), Cm(3.0), Cm(0.05), GOLD)

# Автор и научный руководитель
add_text(slide, Cm(2.0), Cm(7.5), Cm(10.0), Cm(2.5),
         "ВЫПОЛНИЛ:\n"
         "ФИО студента\n"
         "группа __________\n"
         "направление 38.03.05",
         font_size=11, color=WHITE, align=PP_ALIGN.LEFT)

add_text(slide, Cm(13.5), Cm(7.5), Cm(10.0), Cm(2.5),
         "НАУЧНЫЙ РУКОВОДИТЕЛЬ:\n"
         "ФИО, учёная степень,\n"
         "учёное звание, должность",
         font_size=11, color=WHITE, align=PP_ALIGN.LEFT)

# Подвал
add_text(slide, Cm(1.5), Cm(13.0), Cm(22.4), Cm(0.8),
         "Казань — 2026",
         font_size=12, italic=True, color=GOLD, align=PP_ALIGN.CENTER)


# =================================================================
# СЛАЙД 2. АКТУАЛЬНОСТЬ — 3 KPI-плитки
# =================================================================
slide = slides[1]
slide_header(slide, "Актуальность")

n_tiles = 3
gap = Cm(0.3)
tile_w = (CW - gap * (n_tiles - 1)) / n_tiles
tile_h = Cm(8.5)
y_tile = Cm(2.7)

kpi_tile(slide, LEFT, y_tile, tile_w, tile_h,
         "95 %", "управленческих процессов\nвузов в электронном виде\nк 2030 году",
         source="Распоряжение Правительства РФ\n№ 3759-р от 21.12.2021",
         top_color=PRIMARY, big_size=56)

kpi_tile(slide, LEFT + tile_w + gap, y_tile, tile_w, tile_h,
         "29 %", "российских вузов —\n«цифровые аутсайдеры»\n(48 % — последователи,\n23 % — лидеры)",
         source="НИУ ВШЭ, исследование\nцифровой трансформации ВО",
         top_color=WARNING, big_color=WARNING, big_size=56)

kpi_tile(slide, LEFT + 2 * (tile_w + gap), y_tile, tile_w, tile_h,
         "64 %", "вузов: главный барьер\nцифровизации — дефицит\nИТ-кадров (52 % — финансы,\n47 % — низкая грамотность)",
         source="НИУ ВШЭ, исследование\nцифровой трансформации ВО",
         top_color=DANGER, big_color=DANGER, big_size=56)

bottom_band(slide,
    "Нормативное давление + низкая зрелость + кадровый барьер  =  "
    "необходимость моделирования и анализа БП")


# =================================================================
# СЛАЙД 3. ЦЕЛЬ И ЗАДАЧИ
# =================================================================
slide = slides[2]
slide_header(slide, "Цель и задачи исследования")

# Плашка ЦЕЛЬ
goal_y = Cm(2.6)
add_rect(slide, LEFT, goal_y, CW, Cm(2.3), PRIMARY)
add_text(slide, LEFT + Cm(0.3), goal_y, Cm(2.5), Cm(2.3),
         "ЦЕЛЬ", font_size=20, bold=True, color=GOLD,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, LEFT + Cm(3.3), goal_y + Cm(0.2), CW - Cm(3.5), Cm(2.0),
         "Изучение теоретических основ моделирования и анализа бизнес-процессов и их "
         "практическое применение в деятельности структурных подразделений КФУ "
         "(Дирекция музеев, Управление кадров, Управление документооборота).",
         font_size=12, color=WHITE, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

# 5 задач — круги-ступени
tasks = [
    ("1", "ТЕОРИЯ", "IDEF0 · BPMN · EPC\nLean: 5С · 7 потерь\nДЦ/НДБ/НДЦ"),
    ("2", "ОТРАСЛЬ", "Анализ ВО РФ\nКФУ vs 7 ФУ\nКритерии отбора"),
    ("3", "AS IS", "9 моделей\nУзкие места\nДоля НДЦ/НДБ"),
    ("4", "TO BE", "9 моделей\n5С · Битрикс24\nOCR · QR-маршрут"),
    ("5", "ЭКОНОМИКА", "NPV · IRR · DPP\nЧувствительность\nПорядок внедрения"),
]

n = len(tasks)
circle_d = Cm(3.2)
arrow_w = Cm(0.5)
gap_total = CW - n * circle_d
gap = Emu((gap_total - arrow_w * (n - 1)) / (n - 1))
y_circles = Cm(5.6)

# Соединительные стрелки (фоном)
for i in range(n - 1):
    cx = LEFT + circle_d + i * (circle_d + gap + arrow_w)
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, cx, y_circles + Cm(1.3),
                                   gap + arrow_w, Cm(0.6))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = ACCENT
    arrow.line.fill.background()

task_colors = [
    RGBColor(0x4A, 0x7B, 0xB7),
    RGBColor(0x3A, 0x68, 0xA0),
    RGBColor(0x2C, 0x55, 0x88),
    RGBColor(0x22, 0x46, 0x72),
    RGBColor(0x17, 0x36, 0x5D),
]

for i, (num, ttl, body) in enumerate(tasks):
    cx = LEFT + i * (circle_d + gap + arrow_w)
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, y_circles, circle_d, circle_d)
    c.fill.solid()
    c.fill.fore_color.rgb = task_colors[i]
    c.line.color.rgb = WHITE
    c.line.width = Pt(2)
    add_text(slide, cx, y_circles + Cm(0.2), circle_d, Cm(1.0),
             num, font_size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, y_circles + Cm(1.3), circle_d, Cm(0.7),
             ttl, font_size=10, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Тело — под кругом, не выходим за пределы контент-зоны
    body_x = cx - Cm(0.2)
    body_w = circle_d + Cm(0.4)
    if body_x < LEFT:
        body_x = LEFT
    if body_x + body_w > RIGHT:
        body_w = RIGHT - body_x
    add_text(slide, body_x, y_circles + circle_d + Cm(0.2), body_w, Cm(2.8),
             body, font_size=9, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)


# =================================================================
# СЛАЙД 4. ТЕОРЕТИКО-МЕТОДОЛОГИЧЕСКАЯ БАЗА (Задача 1)
# =================================================================
slide = slides[3]
slide_header(slide, "Теоретико-методологическая база",
             subtitle="Результаты по задаче 1")

# Пирамида (3 трапеции) слева
pyr_x = LEFT + Cm(0.3); pyr_y = Cm(2.7); pyr_h = Cm(8.0)
t1 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(2.5), pyr_y, Cm(3.5), Cm(2.4))
t1.fill.solid(); t1.fill.fore_color.rgb = PRIMARY; t1.line.color.rgb = WHITE; t1.line.width = Pt(2)
add_text(slide, pyr_x + Cm(2.5), pyr_y + Cm(0.3), Cm(3.5), Cm(1.9),
         "IDEF0\nстратегия", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

t2 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(1.5), pyr_y + Cm(2.5), Cm(5.5), Cm(2.5))
t2.fill.solid(); t2.fill.fore_color.rgb = ACCENT; t2.line.color.rgb = WHITE; t2.line.width = Pt(2)
add_text(slide, pyr_x + Cm(1.5), pyr_y + Cm(2.7), Cm(5.5), Cm(2.1),
         "BPMN 2.0\nоперации, исполнители", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

t3 = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, pyr_x + Cm(0.3), pyr_y + Cm(5.1), Cm(7.9), Cm(2.8))
t3.fill.solid(); t3.fill.fore_color.rgb = WARNING; t3.line.color.rgb = WHITE; t3.line.width = Pt(2)
add_text(slide, pyr_x + Cm(0.3), pyr_y + Cm(5.3), Cm(7.9), Cm(2.4),
         "EPC\nдетализация, узкие места", font_size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 3 карточки справа — динамически подгоняются к ширине контента
right_x = LEFT + Cm(8.7)
right_w_total = CW - Cm(8.7)
card_gap = Cm(0.2)
card_w = (right_w_total - 2 * card_gap) / 3
card_h = Cm(8.0); card_y = Cm(2.7)

card_data = [
    ("НОТАЦИИ", "•  IDEF0\n   FIPS PUB 183 (NIST)\n•  BPMN 2.0\n   OMG, 2011\n•  EPC\n   Business Studio", PRIMARY),
    ("LEAN", "•  7 видов потерь\n   (Тайити Оно)\n•  ДЦ / НДБ / НДЦ\n•  Диаграмма Ишикавы\n•  5С\n•  DMAIC", ACCENT),
    ("ЭКОНОМИКА", "•  NPV\n•  IRR\n•  DPP\n•  Анализ\n   чувствительности\n•  SMART", DANGER),
]
for i, (ttl, body, col) in enumerate(card_data):
    cx = right_x + i * (card_w + card_gap)
    add_rect(slide, cx, card_y, card_w, card_h, WHITE, line=col, line_w=Pt(1.5))
    add_rect(slide, cx, card_y, card_w, Cm(1.0), col)
    add_text(slide, cx, card_y, card_w, Cm(1.0),
             ttl, font_size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx + Cm(0.25), card_y + Cm(1.2), card_w - Cm(0.5), card_h - Cm(1.4),
             body, font_size=11, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "Стек: НОТАЦИИ + LEAN + ИНВЕСТИЦИОННЫЙ АНАЛИЗ — применён к 3 подразделениям КФУ")


# =================================================================
# СЛАЙД 5. ПОЛОЖЕНИЕ КФУ (Задача 2)
# =================================================================
slide = slides[4]
slide_header(slide, "Положение КФУ в отрасли",
             subtitle="Результаты по задаче 2")

# Радар-чарт слева
chart_data = CategoryChartData()
chart_data.categories = ['Балл ЕГЭ', 'Иностр., %', 'НИОКР, млрд', 'НПР со ст.', 'Внебюдж., %']
chart_data.add_series('КФУ', (74.2, 19.1, 25.3, 4.5, 46.2))
chart_data.add_series('Медиана РФ', (62.8, 4.9, 3.0, 3.1, 25.0))

chart_w = Cm(10.5); chart_h = Cm(8.3)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.RADAR_FILLED, LEFT, Cm(2.5), chart_w, chart_h, chart_data
).chart
chart.has_title = True
chart.chart_title.text_frame.text = "КФУ vs медиана 7 ФУ РФ (2025)"
for run in chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(11); run.font.bold = True; run.font.color.rgb = PRIMARY
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
ser1 = chart.series[0]; ser1.format.fill.solid(); ser1.format.fill.fore_color.rgb = PRIMARY
ser1.format.line.color.rgb = PRIMARY
ser2 = chart.series[1]; ser2.format.fill.solid(); ser2.format.fill.fore_color.rgb = WARNING
ser2.format.line.color.rgb = WARNING

# Карточки лидерства справа
lead_x = LEFT + chart_w + Cm(0.3)
lead_y = Cm(2.5)
lead_w = RIGHT - lead_x
add_rect(slide, lead_x, lead_y, lead_w, Cm(0.8), PRIMARY)
add_text(slide, lead_x, lead_y, lead_w, Cm(0.8),
         "ЛИДЕРСТВО КФУ СРЕДИ 7 ФЕДЕРАЛЬНЫХ ВУЗОВ",
         font_size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

leaders = [
    ("Балл ЕГЭ", "74,2", "★ 1/7"),
    ("Иностр. студенты", "19,1 %", "★ 1/7"),
    ("Внебюджет", "46,2 %", "★ 2/7"),
    ("НИОКР, млрд ₽", "2,53", "★ 2/7"),
]
for i, (ttl, val, place) in enumerate(leaders):
    row_y = lead_y + Cm(1.0) + i * Cm(1.0)
    add_rect(slide, lead_x, row_y, lead_w, Cm(0.85), LIGHT_BG)
    col1_w = lead_w * 0.50
    col2_w = lead_w * 0.28
    col3_w = lead_w * 0.22
    add_text(slide, lead_x + Cm(0.3), row_y, col1_w - Cm(0.3), Cm(0.85),
             ttl, font_size=11, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, lead_x + col1_w, row_y, col2_w, Cm(0.85),
             val, font_size=13, bold=True, color=PRIMARY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, lead_x + col1_w + col2_w, row_y, col3_w, Cm(0.85),
             place, font_size=12, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

bottom_band(slide,
    "🏛 Дирекция музеев (800 тыс. ед.)   ·   👥 ЦРКП (6 615 ППС)   ·   "
    "📄 УДК (170 тыс. ед. архива)")


# =================================================================
# СЛАЙД 6. AS IS (Задача 3)
# =================================================================
slide = slides[5]
slide_header(slide, "Анализ текущего состояния (AS IS)",
             subtitle="Результаты по задаче 3")

units = [
    ("Дирекция музеев", "45 мин", "на экспонат",
     "•  ожидание ПО 10 мин\n•  9 экспонатов / день\n•  ручная фотосъёмка\n•  отсутствие шаблонов",
     1.0),
    ("ЦРКП  (Управление кадров)", "5 дней", "на заявление ППС",
     "•  20 сотрудников\n•  40 ч/мес транспорт\n•  бумажный документооборот\n•  192 мин потерь",
     1.0),
    ("УДК (Документооборот)", "22 мин", "на справку",
     "•  архив оцифрован 12 %\n•  ОРД 7 ч ожидания\n•  176 сек / регистрация\n•  ручной поиск",
     0.7),
]

card_gap = Cm(0.25)
card_w = (CW - 2 * card_gap) / 3
card_h = Cm(6.4); card_y = Cm(2.55)
for i, (ttl, big, sub, body, fill_pct) in enumerate(units):
    cx = LEFT + i * (card_w + card_gap)
    add_rect(slide, cx, card_y, card_w, card_h, WHITE, line=DANGER, line_w=Pt(1.5))
    add_rect(slide, cx, card_y, card_w, Cm(0.9), DANGER)
    add_text(slide, cx, card_y, card_w, Cm(0.9),
             ttl, font_size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    therm_x = cx + Cm(0.3); therm_y = card_y + Cm(1.2); therm_h = Cm(4.2)
    add_rect(slide, therm_x, therm_y, Cm(0.4), therm_h, LIGHT_BG)
    fill_h = Emu(int(therm_h * fill_pct))
    add_rect(slide, therm_x, therm_y + therm_h - fill_h, Cm(0.4), fill_h, DANGER)
    add_text(slide, cx + Cm(1.0), card_y + Cm(1.2), card_w - Cm(1.2), Cm(1.4),
             big, font_size=26, bold=True, color=DANGER,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx + Cm(1.0), card_y + Cm(2.5), card_w - Cm(1.2), Cm(0.5),
             sub, font_size=10, italic=True, color=GREY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    add_text(slide, cx + Cm(1.0), card_y + Cm(3.2), card_w - Cm(1.2), card_h - Cm(3.4),
             body, font_size=10, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

# Stacked bar — структура операций
sb_y = Cm(9.2)
add_text(slide, LEFT, sb_y, CW, Cm(0.6),
         "СТРУКТУРА ОПЕРАЦИЙ В ИССЛЕДОВАННЫХ ПРОЦЕССАХ:",
         font_size=11, bold=True, color=PRIMARY)
bar_h = Cm(1.0)
bar_y = sb_y + Cm(0.7)
seg_ndc = Emu(int(CW * 0.64))
seg_ndb = Emu(int(CW * 0.24))
seg_dc  = Emu(CW - seg_ndc - seg_ndb)
add_rect(slide, LEFT,                              bar_y, seg_ndc, bar_h, DANGER)
add_rect(slide, LEFT + seg_ndc,                    bar_y, seg_ndb, bar_h, WARNING)
add_rect(slide, LEFT + seg_ndc + seg_ndb,          bar_y, seg_dc,  bar_h, SUCCESS)
add_text(slide, LEFT,                              bar_y, seg_ndc, bar_h,
         "НДЦ — устранить  64 %", font_size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, LEFT + seg_ndc,                    bar_y, seg_ndb, bar_h,
         "НДБ — авто  24 %", font_size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, LEFT + seg_ndc + seg_ndb,          bar_y, seg_dc,  bar_h,
         "ДЦ  12 %", font_size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

bottom_band(slide,
    "9 моделей  ·  6 из 7 видов потерь Тайити Оно  ·  3 узких места: "
    "ручной ввод · бумага · ожидание ИТ")


# =================================================================
# СЛАЙД 7. TO BE (Задача 4)
# =================================================================
slide = slides[6]
slide_header(slide, "Проектирование целевых процессов (TO BE)",
             subtitle="Результаты по задаче 4")

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

row_h = Cm(2.85)
row_y0 = Cm(2.55)
row_gap = Cm(0.2)
unit_w = Cm(3.2)
asis_w = Cm(4.4)
arrow_gap = Cm(0.15)
arrow_total = Cm(0.9)
tobe_w = Cm(4.4)
sol_x_start = LEFT + unit_w + asis_w + arrow_total + tobe_w + Cm(0.4)
sol_w = RIGHT - sol_x_start

for i, r in enumerate(rows):
    (unit, asis_a, asis_b, tobe_a, tobe_b, sol, inv, gain) = r
    ry = row_y0 + i * (row_h + row_gap)
    add_rect(slide, LEFT, ry, CW, row_h, LIGHT_BG)
    add_rect(slide, LEFT, ry, unit_w, row_h, PRIMARY)
    add_text(slide, LEFT, ry, unit_w, row_h,
             unit, font_size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    asis_x = LEFT + unit_w + Cm(0.15)
    add_rect(slide, asis_x, ry + Cm(0.25), asis_w, row_h - Cm(0.5), WHITE, line=DANGER, line_w=Pt(1.5))
    add_text(slide, asis_x, ry + Cm(0.25), asis_w, Cm(0.5),
             "AS IS", font_size=9, bold=True, color=DANGER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    add_text(slide, asis_x, ry + Cm(0.7), asis_w, Cm(1.0),
             asis_a, font_size=18, bold=True, color=DANGER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, asis_x, ry + Cm(1.85), asis_w, Cm(0.5),
             asis_b, font_size=10, color=GREY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                  asis_x + asis_w + arrow_gap, ry + Cm(0.95),
                                  arrow_total - 2 * arrow_gap, Cm(0.7))
    arr.fill.solid(); arr.fill.fore_color.rgb = PRIMARY; arr.line.fill.background()
    tobe_x = asis_x + asis_w + arrow_total
    add_rect(slide, tobe_x, ry + Cm(0.25), tobe_w, row_h - Cm(0.5), WHITE, line=SUCCESS, line_w=Pt(1.5))
    add_text(slide, tobe_x, ry + Cm(0.25), tobe_w, Cm(0.5),
             "TO BE", font_size=9, bold=True, color=SUCCESS,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    add_text(slide, tobe_x, ry + Cm(0.7), tobe_w, Cm(1.0),
             tobe_a, font_size=18, bold=True, color=SUCCESS,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, tobe_x, ry + Cm(1.85), tobe_w, Cm(0.5),
             tobe_b, font_size=10, color=GREY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    add_text(slide, sol_x_start, ry + Cm(0.2), sol_w, Cm(1.4),
             sol, font_size=10, bold=True, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    add_text(slide, sol_x_start, ry + Cm(1.6), sol_w, Cm(0.5),
             f"I₀ {inv}", font_size=10, italic=True, color=PRIMARY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    add_text(slide, sol_x_start, ry + Cm(2.1), sol_w, Cm(0.5),
             f"▲ {gain}", font_size=10, bold=True, color=SUCCESS,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "9 моделей TO BE  ·  совокупные инвестиции 792,5 тыс. ₽  ·  "
    "3 узких места = 3 класса решений (lean + cloud + OCR)")


# =================================================================
# СЛАЙД 8. ЭКОНОМИКА (Задача 5)
# =================================================================
slide = slides[7]
slide_header(slide, "Экономическая эффективность",
             subtitle="Результаты по задаче 5")

kpis = [
    ("I₀",      "792,5",  "тыс. ₽"),
    ("ΔCF/год", "1 709,7", "тыс. ₽"),
    ("NPV",     "4 000",   "тыс. ₽ (5 лет)"),
    ("IRR",     "215 %",   "ставка 23 %"),
    ("DPP",     "≈ 7",     "месяцев"),
]
kpi_y = Cm(2.55); kpi_h = Cm(2.1); kpi_gap = Cm(0.15)
kpi_w = (CW - kpi_gap * 4) / 5
for i, (lbl, val, sub) in enumerate(kpis):
    cx = LEFT + i * (kpi_w + kpi_gap)
    add_rect(slide, cx, kpi_y, kpi_w, kpi_h, PRIMARY)
    add_text(slide, cx, kpi_y + Cm(0.05), kpi_w, Cm(0.45),
             lbl, font_size=10, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, kpi_y + Cm(0.45), kpi_w, Cm(1.1),
             val, font_size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, kpi_y + Cm(1.5), kpi_w, Cm(0.6),
             sub, font_size=9, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Линейный график — накопительный NPV
ch_y = Cm(5.0); ch_h = Cm(7.0)
left_chart_w = Cm(11.5)
right_chart_x = LEFT + left_chart_w + Cm(0.3)
right_chart_w = RIGHT - right_chart_x

line_data = CategoryChartData()
line_data.categories = ['0', '1', '2', '3', '4', '5']
line_data.add_series('Накопительный NPV, млн ₽', (-0.79, 0.80, 1.99, 3.06, 3.55, 4.00))
line_chart = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS, LEFT, ch_y, left_chart_w, ch_h, line_data
).chart
line_chart.has_title = True
line_chart.chart_title.text_frame.text = "Накопительный NPV (горизонт 5 лет, r = 23 %)"
for run in line_chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(11); run.font.bold = True; run.font.color.rgb = PRIMARY
line_chart.has_legend = False
ser = line_chart.series[0]
ser.format.line.color.rgb = PRIMARY
ser.format.line.width = Pt(3)
for pt in ser.points:
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = PRIMARY
ser.data_labels.show_value = True
ser.data_labels.font.size = Pt(9)
ser.data_labels.font.bold = True
ser.data_labels.font.color.rgb = PRIMARY

# Tornado-bar — чувствительность
tornado_data = CategoryChartData()
tornado_data.categories = ['−50 % CF', 'Песс. CF+r', '−30 % CF', '+50 % затр.', '−18 % r=30 %', 'Базовый']
tornado_data.add_series('NPV, тыс. ₽', (1596, 2020, 2543, 3604, 3261, 4000))
tor_chart = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, right_chart_x, ch_y, right_chart_w, ch_h, tornado_data
).chart
tor_chart.has_title = True
tor_chart.chart_title.text_frame.text = "Анализ чувствительности NPV"
for run in tor_chart.chart_title.text_frame.paragraphs[0].runs:
    run.font.size = Pt(11); run.font.bold = True; run.font.color.rgb = PRIMARY
tor_chart.has_legend = False
plot = tor_chart.plots[0]
plot.gap_width = 60
ser2 = tor_chart.series[0]
bar_colors = [DANGER, DANGER, WARNING, WARNING, ACCENT, SUCCESS]
for idx, pt in enumerate(ser2.points):
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = bar_colors[idx]
ser2.data_labels.show_value = True
ser2.data_labels.font.size = Pt(9)
ser2.data_labels.font.bold = True

bottom_band(slide,
    "Очерёдность внедрения:   ① ЦРКП  →  ② УДК  →  ③ МУЗЕИ   ·   запас прочности 83 %")


# =================================================================
# СЛАЙД 9. ВЫВОДЫ
# =================================================================
slide = slides[8]
slide_header(slide, "Выводы",
             subtitle="Цель достигнута · 5 задач решены · методика воспроизводима")

achievements = [
    ("МЕТОДИКА",   "3 уровня",      "IDEF0 + BPMN + EPC", PRIMARY),
    ("ОТРАСЛЬ",    "★ 1/7",          "КФУ среди ФУ РФ",      ACCENT),
    ("AS IS",      "9",              "моделей · 64 % НДЦ",   WARNING),
    ("TO BE",      "−17",            "штатных единиц ЦРКП",  SUCCESS),
    ("ЭКОНОМИКА",  "4,0 млн ₽",       "NPV · IRR 215 % · DPP 7 мес.", DANGER),
]

# Верхний ряд: 3 плитки
top_gap = Cm(0.25)
top_w = (CW - 2 * top_gap) / 3
top_h = Cm(4.4); top_y = Cm(2.6)
for i in range(3):
    ttl, big, body, col = achievements[i]
    cx = LEFT + i * (top_w + top_gap)
    add_rect(slide, cx, top_y, top_w, top_h, WHITE, line=col, line_w=Pt(1.5))
    add_rect(slide, cx, top_y, top_w, Cm(0.85), col)
    add_text(slide, cx, top_y, top_w, Cm(0.85),
             ttl, font_size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, top_y + Cm(1.0), top_w, Cm(1.9),
             big, font_size=32, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, top_y + Cm(2.95), top_w, Cm(1.4),
             body, font_size=11, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

# Нижний ряд: 2 плитки центрированно
bot_gap = Cm(0.4)
bot_w = (CW - bot_gap) / 2
bot_h = Cm(4.2); bot_y = Cm(7.3)
for i, j in enumerate([3, 4]):
    ttl, big, body, col = achievements[j]
    cx = LEFT + i * (bot_w + bot_gap)
    add_rect(slide, cx, bot_y, bot_w, bot_h, WHITE, line=col, line_w=Pt(1.5))
    add_rect(slide, cx, bot_y, bot_w, Cm(0.85), col)
    add_text(slide, cx, bot_y, bot_w, Cm(0.85),
             ttl, font_size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, bot_y + Cm(1.0), bot_w, Cm(1.7),
             big, font_size=32, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, cx, bot_y + Cm(2.8), bot_w, Cm(1.3),
             body, font_size=11, color=BLACK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)

bottom_band(slide,
    "Результаты применимы в КФУ и тиражируемы на 7 федеральных университетов РФ")


# =================================================================
# СЛАЙД 10. СПАСИБО — фон шаблона уже на месте
# =================================================================
slide = slides[9]
# На слайде 10 шаблона — большая декоративная картинка-фон.
# Поверх — наш заголовок и подзаголовок.
add_text(slide, Cm(1), Cm(5.0), Cm(23.4), Cm(2.5),
         "СПАСИБО ЗА ВНИМАНИЕ", font_size=40, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(slide, Cm(11.2), Cm(7.6), Cm(3.0), Cm(0.06), GOLD)
add_text(slide, Cm(1), Cm(8.0), Cm(23.4), Cm(1.0),
         "Готов ответить на ваши вопросы", font_size=16, italic=True, color=GOLD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- СОХРАНЕНИЕ ----------
prs.save(OUT_PATH)
print(f"OK: {OUT_PATH}")
