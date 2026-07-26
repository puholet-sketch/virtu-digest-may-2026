# -*- coding: utf-8 -*-
"""Generate PPTX from structured digest data — май 2026."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor

BG = RGBColor(0x1A, 0x23, 0x32)
WHITE = RGBColor(0xE8, 0xED, 0xF4)
MUTED = RGBColor(0x8B, 0x9C, 0xB3)
ACCENT = RGBColor(0x3B, 0x82, 0xF6)
GREEN = RGBColor(0x10, 0xB9, 0x81)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
PURPLE = RGBColor(0x8B, 0x5C, 0xF6)
BORDER = RGBColor(0x2D, 0x3A, 0x4F)

COLORS = [ACCENT, GREEN, AMBER, PURPLE, RGBColor(0xEF, 0x44, 0x44), RGBColor(0x06, 0xB6, 0xD4)]

PLATFORM_COLORS = {
    "VFOS РГС": ACCENT,
    "ОФР.ИФЛ": GREEN,
    "ОФР.Базовая / Сервисная": AMBER,
    "GuideWire": PURPLE,
    "Импульс": PURPLE,
    "Госуслуги.Дом": ACCENT,
    "Кросс-продажи ПВС МВД": PURPLE,
    "Автовыгрузка": GREEN,
    "НСИС": AMBER,
    "Keycloak": ACCENT,
    "РГС Ипотека": ACCENT,
    "Ренессанс Ипотека": GREEN,
    "Инсайт Ипотека": PURPLE,
}

GRATITUDE_COLORS = {
    "VFOS РГС": ACCENT,
    "ОФР.ИФЛ": GREEN,
    "ОФР.Базовая / Сервисная": GREEN,
    "GuideWire": AMBER,
    "Импульс": AMBER,
    "Согласие Ипотека": PURPLE,
    "Согласие Вита / НПФ": PURPLE,
    "РГС Ипотека": ACCENT,
    "Ренессанс Ипотека": GREEN,
    "ВСК / МСГ": ACCENT,
}

VFOS_DIRECTION_COLORS = {
    "Госуслуги.Дом": AMBER,
    "Кросс-продажи ПВС МВД": AMBER,
    "Автовыгрузка": AMBER,
    "НСИС": AMBER,
    "Keycloak": AMBER,
    "3-й интегр. контур": AMBER,
    "Инфраструктура": ACCENT,
    "Архитектура": PURPLE,
    "Альфа": ACCENT,
    "Райффайзен": PURPLE,
}


def set_slide_bg(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_title(slide, title, subtitle=None, title_size=28, title_height=0.8, subtitle_top=1.05, subtitle_height=0.5):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(title_height))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(title_size)
    p.font.bold = True
    p.font.color.rgb = WHITE
    if subtitle:
        box2 = slide.shapes.add_textbox(Inches(0.5), Inches(subtitle_top), Inches(9), Inches(subtitle_height))
        tf2 = box2.text_frame
        tf2.word_wrap = True
        for i, line in enumerate(subtitle.split("\n")):
            sp = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            sp.text = line
            sp.font.size = Pt(14)
            sp.font.color.rgb = MUTED


def add_section_label(slide, text, left=0.5, top=1.35, color=GREEN, width=9, height=0.3):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = color


def add_bullets(slide, items, left=0.5, top=1.6, width=9, height=5, font_size=13):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = WHITE
        p.space_after = Pt(6)


def add_kpi_row(slide, items, top=1.35):
    """items: list of (value, label)"""
    w = 9.0 / len(items)
    for i, (val, lbl) in enumerate(items):
        b = slide.shapes.add_textbox(Inches(0.4 + i * w), Inches(top), Inches(w - 0.1), Inches(0.9))
        t = b.text_frame.paragraphs[0]
        t.text = val
        t.font.size = Pt(18)
        t.font.bold = True
        t.font.color.rgb = COLORS[i % len(COLORS)]
        t2 = b.text_frame.add_paragraph()
        t2.text = lbl
        t2.font.size = Pt(9)
        t2.font.color.rgb = MUTED


def add_bar_chart(slide, title, categories, values, left=0.5, top=1.5, width=9, height=4.5, horiz=False, label_color=WHITE, title_color=None, show_axis_title=True, value_max=None, show_chart_title=True, labels_outside=False):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("", values)
    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if horiz else XL_CHART_TYPE.COLUMN_CLUSTERED
    chart = slide.shapes.add_chart(chart_type, Inches(left), Inches(top), Inches(width), Inches(height), chart_data).chart
    chart.has_legend = False
    chart.has_title = show_chart_title
    if show_chart_title:
        chart.chart_title.text_frame.text = title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(10 if title_color else 12)
        if title_color:
            chart.chart_title.text_frame.paragraphs[0].font.color.rgb = title_color
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = "#,##0"
    plot.data_labels.font.size = Pt(9)
    plot.data_labels.font.color.rgb = label_color
    if labels_outside and horiz:
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    for i, pt in enumerate(plot.series[0].points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = COLORS[i % len(COLORS)]
    if show_axis_title:
        chart.value_axis.has_title = True
        chart.value_axis.axis_title.text_frame.text = "Человек" if horiz else "Количество"
    else:
        chart.value_axis.has_title = False
    if value_max is not None:
        chart.value_axis.maximum_scale = value_max
    if horiz:
        chart.category_axis.tick_labels.font.size = Pt(10)
        chart.category_axis.tick_labels.font.color.rgb = WHITE
        chart.category_axis.has_major_gridlines = False
    return chart


def add_team_bar_chart(slide, label, label_color, categories, values, top, height=1.55, width=3.5, left=0.4):
    label_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.25))
    p = label_box.text_frame.paragraphs[0]
    p.text = label
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = label_color
    chart_top = top + 0.23
    add_bar_chart(
        slide, "Состав команды (чел.)", categories, values,
        left=left, top=chart_top, width=width, height=height,
        horiz=True, labels_outside=True,
    )
    return chart_top + height


def _set_cell_text(cell, text, font_size=9, color=WHITE, bold=False):
    lines = str(text).split("\n")
    tf = cell.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.bold = bold


def add_table(slide, headers, rows, left=0.4, top=1.5, width=9.2, row_height=0.35, col_colors=None, color_col=0):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(row_height * n_rows))
    table = table_shape.table
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        _set_cell_text(cell, h, font_size=10, color=MUTED, bold=True)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            color_key = None
            color_idx = None
            if col_colors:
                if color_col < len(row) and row[color_col] in col_colors:
                    color_idx, color_key = color_col, row[color_col]
                elif row[0] in col_colors:
                    color_idx, color_key = 0, row[0]
            block_color = col_colors.get(color_key) if color_key else None
            use_color = block_color if block_color and j == color_idx else WHITE
            use_bold = bool(block_color and j == color_idx)
            _set_cell_text(
                cell, val,
                font_size=9 if j > 0 else 10,
                color=use_color,
                bold=use_bold,
            )
    return table


def add_hline(slide, left, top, width, color=BORDER):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.01))
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()


def add_team_block(slide, title, items, top, color, left=4.0, width=5.55, font_size=7.5, row_unit=0.24):
    block_height = row_unit * len(items) + 0.06
    add_section_label(slide, title, left=left, top=top, color=color)
    block_left = left + 0.14
    block_top = top + 0.22
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(block_left - 0.06), Inches(block_top),
        Inches(0.025), Inches(block_height))
    accent.fill.solid()
    accent.fill.fore_color.rgb = color
    accent.line.fill.background()
    add_bullets(slide, items, left=block_left, top=block_top, width=width - 0.14, height=block_height, font_size=font_size)
    row_h = block_height / len(items)
    for i in range(1, len(items)):
        add_hline(slide, block_left, block_top + i * row_h - 0.015, width - 0.14)
    return block_top + block_height + 0.1


def build():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Cover
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    box = s.shapes.add_textbox(Inches(0.8), Inches(2.1), Inches(8.4), Inches(2.4))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Статус проектов компании"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    for line in (
        "с января по май 2026 г",
        "Руководитель Технической дирекции - Столбов А,",
        "Руководитель Проектного офиса - Сорванов О",
    ):
        p2 = tf.add_paragraph()
        p2.text = line
        p2.font.size = Pt(16)
        p2.font.color.rgb = MUTED
        p2.alignment = PP_ALIGN.CENTER

    # 1 Portfolio
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Портфель по страховым компаниям", "с января по май 2026 года")
    add_table(s,
        ["СК", "Проект", "Выполнено", "В работе"],
        [
            ["РГС", "VFOS", "61 круп. задача", "21"],
            ["РГС", "ОФР", "13 блоков функционала", "3 блока"],
            ["РГС", "Аутстафф GuideWire / Импульс", "KASKO/OSAGO, ипотека", "—"],
            ["РГС", "РГС Ипотека", "27", "5"],
            ["Согласие", "Согласие Ипотека", "29", "9"],
            ["Согласие", "Вита / НПФ / Банки", "45", "52"],
            ["Согласие", "VFOS Согласие", "12 задач", "—"],
            ["Альфа", "VFOS Альфа", "2FA, jump-сервер", "—"],
            ["Райффайзен", "VFOS Райффайзен", "Архит. описания", "—"],
            ["Ренессанс", "Ренессанс Ипотека", "9", "6"],
            ["Инсайт", "Инсайт Ипотека", "—", "—"],
            ["ВСК/МСГ", "ВСК Автострахование, ВСК ДМС, МСГ", "21", "8"],
            ["—", "Virtu Drive", "6 доработок в период паузы", "—"],
        ], top=1.35, row_height=0.38)

    # 3 RGS VFOS — team chart + activities
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "VFOS: РГС (ДМС, АВТО, ИФЛ, ВЗР), АЛЬФА Ипотека, Согласие ВЗР, Райффайзен (КСН, МББ, ВЗР)",
              "PM: Сорванов О.Н. · Аккаунт: Черноус С. · 61 выполнено · 21 в работе · команда 7 чел.",
              title_size=17, title_height=0.95, subtitle_top=1.38)
    add_team_bar_chart(s, "КОМАНДА VFOS", ACCENT,
                       ["Аналитики", "Разработчики", "QA"], [2, 4, 1],
                       top=1.62, height=2.2)
    add_section_label(s, "Реализованы ключевые активности:", left=4.0, top=1.62)
    add_table(s,
        ["СК", "Направление", "Реализованный функционал", "План"],
        [
            ["РГС", "Госуслуги.Дом", "КСП ИФЛ, .NET 10, OAuth2/Keycloak, очередь с ретраями, mock-контур", "60 млн ₽"],
            ["РГС", "Кросс-продажи ПВС МВД", "«Добро пожаловать в гости» + «РГС Гость_2016» через API", "100 млн ₽"],
            ["РГС", "Автовыгрузка", "НС, ИФЛ, ДМС в ЕКИС без ручной загрузки XML", "—"],
            ["РГС", "НСИС", "Передача данных личного страхования по требованиям ЦБ", "—"],
            ["РГС", "Keycloak", "Переход авторизации (Волосников А.); ~10 000 польз./день", "—"],
            ["РГС", "3-й интегр. контур", "Развёрнут третий интеграционный контур по запросу РГС", "—"],
            ["Альфа", "Инфраструктура", "Осуществлён переход на 2FA и jump-сервер для работы с ресурсами СК", "—"],
            ["Райффайзен", "Архитектура", "Детальные описания архитектурных решений: интеграция VFOS — банковская система СК", "—"],
        ], left=4.0, top=1.88, width=5.6, row_height=0.42,
        col_colors={**VFOS_DIRECTION_COLORS, "РГС": ACCENT}, color_col=1)
    vfos_table = s.shapes[-1].table
    vfos_table.columns[0].width = Inches(0.55)
    vfos_table.columns[1].width = Inches(1.15)
    vfos_table.columns[2].width = Inches(3.35)
    vfos_table.columns[3].width = Inches(0.55)
    for r in range(1, len(vfos_table.rows)):
        for p in vfos_table.cell(r, 0).text_frame.paragraphs:
            p.font.size = Pt(9)
            if r <= 6:
                p.font.color.rgb = ACCENT
            p.font.bold = True
    for r in range(1, len(vfos_table.rows)):
        for p in vfos_table.cell(r, 2).text_frame.paragraphs:
            p.font.color.rgb = WHITE
            p.font.bold = False

    # 3 OFR — two columns: chart left, all activities right
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "ОФР и Аутстафф РГС", "PM: Сорванов О. · Аккаунт: Черноус С. · ОФР — 14 чел. · Аутстафф — 6 чел.")
    ofr_bottom = add_team_bar_chart(s, "КОМАНДА ОФР", ACCENT,
                                    ["TypeScript", "QA", "БА"], [6, 4, 4],
                                    top=1.32, height=1.95)
    add_team_bar_chart(s, "КОМАНДА АУТСТАФФ", AMBER,
                       ["GoSu", "Аналитик", "TypeScript", "Java"], [2, 1, 1, 2],
                       top=ofr_bottom + 0.35, height=1.95)
    add_section_label(s, "Реализованы ключевые активности:", left=4.0, top=1.35)
    ifl_items = [
        "getSellerProfile — единый сервис для всех команд: данные продавца и подчинённых при недоступности систем-источников",
        "«Умное» получение договоров — снижение нагрузки на GuideWire",
        "Платёжная машина: транзакционность, блокировка одной транзакции в единицу времени, логи, алерты об ошибках оплаты в messenger",
    ]
    y = add_team_block(s, "ОФР. ИФЛ", ifl_items, 1.58, GREEN, row_unit=0.22)
    base_items = [
        "3 продукта на Резолют (2 с пролонгацией) + 1 на GW с пролонгацией",
        "Стабилизация и ускорение продуктов КСП",
        "Оплата на следующий день — улучшение процесса оплаты",
        "Улучшение работы списка договоров",
        "ofr-core-rights — получение и проверка прав, «умное» включение прав",
    ]
    y = add_team_block(s, "ОФР. Базовая и Сервисная", base_items, y + 0.06, GREEN, row_unit=0.19)
    shared_items = [
        "Blitz — перевод авторизации на импортозамещённую систему",
        "Поиск коллег и старых ЛНР по продавцам — отображение и перехват задач",
        "«Флагман» — кэширование и оптимизация запуска и работы BFF",
        "Мониторинг 15 внешних систем (каждые 5 мин, графики для ТП)",
        "Подготовка перехода на GitLab",
    ]
    y = add_team_block(s, "ОФР. ИФЛ и Базовая", shared_items, y + 0.06, GREEN, row_unit=0.19)
    outstaff_items = [
        "GuideWire — KASKO/OSAGO OpenAPI; дефекты и интеграции (NSIS, Dadata); новые продукты; изменение и расторжение",
        "GuideWire (поддержка) — с янв. 89 задач, ~95% баги и инц.; остановка пролонгации ОСАГО — 2 инц. высок. приоритета устранены",
        "Импульс — андеррайтинг ипотеки: логика решений, оформление, продуктовые сценарии",
    ]
    y = add_team_block(s, "Аутстафф", outstaff_items, y + 0.06, AMBER, row_unit=0.22)
    ofr_wip = [
        "Внесение изменений и андеррайтинг",
        "Структура пользователей «Флагман»",
        "Переход на сервисы «Импульс» для данных пользователей",
    ]
    add_team_block(s, "ОФР. ИФЛ — в работе", ofr_wip, y + 0.06, GREEN, row_unit=0.19)

    # 4 Mortgage — team chart + activities
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Ипотечный портфель", "PM: Кузнецова А. · Аккаунт: Черноус С., Клейн Т., Конкин М. · команда 8 чел. · 0 инцидентов за май")
    mortgage_table_top = 1.35
    mortgage_row_height = 0.72
    mortgage_table_rows = 4
    mortgage_chart_height = mortgage_table_top + mortgage_row_height * mortgage_table_rows - 1.55
    add_team_bar_chart(s, "СОСТАВ КОМАНДЫ", ACCENT,
                       ["Аналитики", "Разработчики", "Настройщики"], [2, 4, 2],
                       top=1.32, height=mortgage_chart_height)
    add_section_label(s, "Реализованы ключевые активности:", left=4.0, top=1.35)
    add_table(s,
        ["Площадка", "Реализованный функционал", "Сделано / в работе"],
        [
            ["РГС Ипотека", "Доп. соглашения с доплатой; S3 для аттачей (dev/preprod); prod и −2/3 БД — июнь", "27 / 5"],
            ["Ренессанс Ипотека", "В B2B реализовали возможность менеджеру вернуть договор из онлайн-пролонгации клиента на ручную пролонгацию", "9 / 6"],
            ["Инсайт Ипотека", "Прод, развитие; регресс и нагрузочное тестирование; 0 инцидентов за май", "—"],
        ], left=4.0, top=1.65, width=5.6, row_height=0.72, col_colors=PLATFORM_COLORS)

    # 5 Soglasie — team chart + activities
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Согласие — Ипотека и Вита; Гелиос, Солидарность, ВСК, МСГ",
              "PM: Кочетков Н., Царёва А. · Аккаунт: Клейн Т., Конкин М.\n"
              "Согласие Ипотека — 29 вып, 9 в работе\n"
              "ВИТА\\НПФ\\Банки\\Гелиос\\Солидарность — 45 вып, 52 в работе",
              title_size=22, title_height=0.85, subtitle_top=1.25, subtitle_height=0.72)
    vita_bottom = add_team_bar_chart(s, "КОМАНДА ВИТА / НПФ", PURPLE,
                                     ["Аналитики", "Разработчики", "Настройщики"], [2, 4, 1],
                                     top=1.98, height=1.35)
    ipoteka_bottom = add_team_bar_chart(s, "КОМАНДА ИПОТЕКА", PURPLE,
                                        ["Аналитики", "Разработчики", "Настройщики"], [2, 4, 3],
                                        top=vita_bottom + 0.35, height=1.35)
    add_team_bar_chart(s, "КОМАНДА ВСК / МСГ", ACCENT,
                       ["Аналитики", "Разработчики"], [1, 2],
                       top=ipoteka_bottom + 0.35, height=1.2)
    add_section_label(s, "Реализованы ключевые активности:", left=4.0, top=1.68)
    add_table(s,
        ["Блок", "Детализация", "Результат"],
        [
            ["Интеграция ВТБ ELM", "Интеграция пользователей JX→ELM; сервис обновления данных (сброс пароля, регистрация); интеграция с календарём ELM; автоматическая отправка пула договоров к пролонгации в ELM", "~1 000 ч"],
            ["Спецусловия партнёров", "Настройка спецтарифов для групп партнёров (группы тарификации); сервис присвоения группы по ИКП; автоматическое определение группы и применение тарифов при создании договора", "—"],
            ["Журнал пролонгации", "Крупный блок доработок журнала пролонгации и механизмов пролонгации многолетних договоров", "> 8 000 ч"],
            ["Вита / НПФ / Банки", "Разработка лендинга ПДС; регулярные доработки существующих продуктов различной степени сложности", "45 вып / 52 в работе\n> 8 млн ₽"],
            ["МП Эльбрус", "Финальное согласование макетов МП для ЛК клиента, подготовка окружения, передача заказчику на тестирование.", "Финал работ в июне 2026 г."],
            ["ВСК / МСГ", "Миграция данных Автострахование. Восстановление системы. Добавление программ «Гарантия авто Грузовые» и «Гарантия Li Auto»", "—"],
        ], left=4.0, top=1.78, width=5.6, row_height=0.48,
        col_colors={
            "Интеграция ВТБ ELM": AMBER,
            "Спецусловия партнёров": AMBER,
            "Журнал пролонгации": AMBER,
            "Вита / НПФ / Банки": GREEN,
            "МП Эльбрус": PURPLE,
            "ВСК / МСГ": ACCENT,
        })

    # 6 ALL Tech Support — 3 columns, aligned grid
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Техническая поддержка — все проекты и системы", "с января по май 2026 · Руководитель поддержки: Смирнова А. · Руководитель тех.дирекции: Столбов А.")
    tp_content_top = 1.32
    tp_label_h = 0.36
    tp_chart_h = 1.40
    tp_row_gap = 0.12
    tp_row_slot = tp_label_h + tp_chart_h + tp_row_gap
    tp_cols = {1: (0.4, 2.85), 2: (3.35, 2.85)}
    tp_summary_left = 6.35
    tp_summary_w = 3.25

    def add_tp_widget(row, col, label, categories, values, value_max):
        left, width = tp_cols[col]
        label_top = tp_content_top + row * tp_row_slot
        chart_top = label_top + tp_label_h
        add_section_label(s, label, left=left, top=label_top, color=AMBER, width=width, height=tp_label_h)
        add_bar_chart(s, "", categories, values, left=left, top=chart_top, width=width, height=tp_chart_h,
                      label_color=AMBER, show_axis_title=False, value_max=value_max, show_chart_title=False)

    add_tp_widget(0, 1, "VFOS — 3 897 задач", ["Доступы", "Изм. БД", "Инцид.", "Прод.", "Проч."], [2874, 596, 238, 138, 51], 3400)
    add_tp_widget(0, 2, "Согласие Ипотека — 29 задач", ["Конс.", "Нетип."], [69, 278], 310)
    add_tp_widget(1, 1, "Ипотечный портфель (РГС, Ренессанс, Инсайт) — 864 задачи", ["Нетип.", "Подкл.", "Конс.", "Аварии"], [537, 204, 27, 11], 650)
    add_tp_widget(1, 2, "Альфа VFOS — 8 задач", ["Конс.", "Ошиб.", "Изм-ния", "Сбой"], [3, 1, 4, 0], 5)
    add_tp_widget(2, 1, "Virtu Drive — 80 заявок", ["Конс.", "Нетип."], [7, 73], 85)
    add_tp_widget(2, 2, "Согласие VFOS — 14 задач", ["Конс.", "Ошиб.", "Изм-ния", "Сбой"], [6, 2, 4, 2], 7)

    summary_label_top = tp_content_top
    summary_chart_top = tp_content_top + tp_label_h
    summary_chart_h = tp_content_top + 2 * tp_row_slot + tp_label_h + tp_chart_h - summary_chart_top
    add_section_label(s, "Закрытые задачи ТП по системам (шт.)", left=tp_summary_left, top=summary_label_top,
                      color=AMBER, width=tp_summary_w, height=tp_label_h)
    add_bar_chart(s, "",
                  ["Virtu Drive", "Альфа+Согл.", "Согл. ип.", "Ипотека (РГС, Рен., Инс.)", "VFOS"],
                  [80, 22, 29, 864, 3897], left=tp_summary_left, top=summary_chart_top, width=tp_summary_w,
                  height=summary_chart_h, horiz=True, label_color=AMBER, show_axis_title=False, value_max=4600,
                  show_chart_title=False, labels_outside=True)

    # 7 Pre-sale — team chart height aligned with 4-row table
    presale_table_top = 1.35
    presale_row_height = 0.55
    presale_table_rows = 5
    presale_chart_top = 1.55
    presale_chart_height = presale_table_top + presale_row_height * presale_table_rows - presale_chart_top
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "VirtuDrive, Pre-sale и новые инициативы", "PM: Шейко Н. · Аккаунт: Чесноков Д.")
    add_team_bar_chart(s, "КОМАНДА VIRTU DRIVE", AMBER,
                       ["Аналитик", "Java"], [1, 2],
                       top=1.32, height=presale_chart_height)
    add_table(s,
        ["Проект", "Статус", "Следующий шаг", "Дедлайн"],
        [
            ["Virtu Drive", "Пауза (бессрочно)", "On-hold", "—"],
            ["Тендер Автофинанс", "Ожидание", "Ответ заказчика", "01.06.2026"],
            ["Мобильное приложение ЛК", "Демо", "КП T-Лизинг, встреча ОЛА", "15.06.2026"],
            ["Конструктор комиссионных", "Discovery", "Встречи с ДЦ", "—"],
        ], left=4.0, top=presale_table_top, width=5.6, row_height=presale_row_height)

    # 9 Gratitude table
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Благодарность сотрудникам за реализацию функционала")
    add_table(s,
        ["Площадка", "Сотрудники", "Реализованный функционал"],
        [
            ["VFOS РГС", "Рыжикова Ю., Рудченко П., Сивогривова Н., Нижурин С., Агалетдинов А., Кукарин Д., Волосников А.",
             "Госуслуги.Дом; кросс-продажи; автовыгрузка ЕКИС; НСИС; Keycloak"],
            ["РГС Ипотека", "Павлова И., Пустохин А., Крысина А., Ситник Д., Спирчина Н., Редько Г., Аравопуло Е.",
             "Допсоглашение по техкоррекции с доплатой; онлайн-пролонгация; онлайн-пролонгация однолетних договоров в коробке"],
            ["Ренессанс Ипотека", "Ульянова М., Ситник Д., Белоусов Е.",
             "Дополнения к реализации ПК по НСиБ; не прекращать онлайн-пролонгацию при ручных действиях в UI; скидка «Скидка андеррайтера» для БИ"],
            ["ОФР.ИФЛ", "Грабер Д., Кошкина О., Ташбулатова В., Кудяков С., Лобов О., Тетерин М., Максимова Е.",
             "Продажи, пролонгация, скидки, сверхлимиты, перехват задач, расторжение, техкоррекция"],
            ["ОФР.Базовая / Сервисная", "Попков Д., Искаков Т., Тамаев Р.", "4 продукта ДМС через B2B Резолют"],
            ["GuideWire", "Шеламов А., Щербаков Д.", "KASKO/OSAGO OpenAPI; новые продукты; изменение и расторжение"],
            ["Импульс", "Резниченко Д.", "Андеррайтинг ипотеки: расчёт, оформление, оплата"],
            ["Согласие Ипотека", "Мазанова А., Базуева И., Дмитриев Э., Сергеева Е., Купченков А., Воронов И., Савельева А., Демшина А., Ситник Д.",
             "Интеграция ВТБ ELM; спецусловия партнёров; журнал пролонгации и многолетние договоры"],
            ["Согласие Вита / НПФ", "Патрушева О., Понина В., Купченков А., Яковлев А., Коротков А., Сергеева Е., Дорофеева Ю.",
             "Лендинг ПДС; регулярные доработки продуктов — 35 доработок, > 8 млн ₽"],
            ["ВСК / МСГ", "Бактикова Ф., Костарев А., Воронов И.",
             "Миграция данных Автострахование. Восстановление системы. Добавление программ «Гарантия авто Грузовые» и «Гарантия Li Auto»"],
        ], top=1.35, row_height=0.48, col_colors=GRATITUDE_COLORS)

    # 10 PM matrix
    s = prs.slides.add_slide(blank)
    set_slide_bg(s)
    add_title(s, "Матрица ответственности PM", "с января по май 2026 года")
    add_table(s,
        ["PM", "Проекты", "Разработка\n(вып. / в раб.)", "ТП (задач)"],
        [
            ["Сорванов\u00a0О.Н.", "VFOS РГС, Альфастрахование, Согласие, Райффайзен, ОФР", "61 / 21", "3 897 + 22"],
            ["Кузнецова А.", "Ипотека (3 СК)", "27/5 (РГС) + 9/6 (Рен.)", "864"],
            ["Кочетков Н.,\nЦарёва А.",
             "Согласие Вита, Согласие Банки, Согласие НПФ, Согласие Мигрант, Солидарность, Гелиос",
             "45 / 52",
             "—"],
            ["Кочетков Н.,\nЦарёва А.",
             "Согласие Ипотека",
             "29 / 9",
             "69 + 278"],
            ["Кочетков Н.",
             "ВСК АВТО и ВСК ДМС, МСГ, ВСК Банки",
             "21 / 8",
             "—"],
            ["Царёва А.", "МП Эльбрус", "—", "—"],
            ["Шейко Н.", "Virtu Drive, тендеры, ЛК", "6 доработок", "80 заявок"],
        ], top=1.35, row_height=0.68)
    pm_table = s.shapes[-1].table
    pm_table.columns[0].width = Inches(1.45)
    pm_table.columns[1].width = Inches(4.0)
    for row in pm_table.rows:
        row.cells[0].text_frame.word_wrap = False
        row.cells[1].text_frame.word_wrap = True
        row.cells[2].text_frame.word_wrap = True

    out = r"D:\projects\дайджест\май\output\презентация_статус_проектов_май_2026.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
