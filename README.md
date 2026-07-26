# Дайджест статуса проектов (май 2026)

Корпоративная презентация статуса проектов Virtu Systems за период **январь–май 2026** и шаблон Word для обратного сбора данных.

## Смотреть онлайн

- **Репозиторий:** https://github.com/puholet-sketch/virtu-digest-may-2026
- **Презентация (GitHub Pages):** https://puholet-sketch.github.io/virtu-digest-may-2026/
- **Шаблон Word:** кнопка на обложке презентации или  
  https://puholet-sketch.github.io/virtu-digest-may-2026/%D1%88%D0%B0%D0%B1%D0%BB%D0%BE%D0%BD_%D0%B2%D0%B2%D0%BE%D0%B4%D0%B0_%D0%B4%D0%B0%D0%BD%D0%BD%D1%8B%D1%85_%D0%B4%D0%B0%D0%B9%D0%B4%D0%B6%D0%B5%D1%81%D1%82.docx  
  (файл [`docs/шаблон_ввода_данных_дайджест.docx`](docs/шаблон_ввода_данных_дайджест.docx))

После публикации GitHub Pages сайт доступен по ссылкам выше.

## Что в репозитории

| Путь | Назначение |
|------|------------|
| `docs/index.html` | Презентация (GitHub Pages) |
| `docs/шаблон_ввода_данных_дайджест.docx` | Структурированный шаблон ввода |
| `docs/дайджест_структура.md` | Markdown-экспорт структуры |
| `output/` | Исходники генерации (HTML, PPTX-скрипт, MD) |
| `templates/generate_digest_input_docx.py` | Пересборка Word-шаблона |

## Как обновить дайджест

1. Скачайте **шаблон Word** с обложки презентации.
2. Заполните разделы 1–12 (портфель, команды, активности, ТП, благодарности, матрица PM).
3. Сохраните как `дайджест_<месяц>_<год>_ввод.docx` и передайте в проектный офис / Cursor.
4. По шаблону обновляются HTML → PPTX (`python output/generate_pptx.py`).

Пересобрать сам шаблон:

```bash
python templates/generate_digest_input_docx.py
```

Пересобрать PPTX:

```bash
python output/generate_pptx.py
```

## Локальный просмотр

Откройте `docs/index.html` или `output/презентация_статус_проектов_май_2026.html` в браузере.

## Стиль Word

Красно-чёрная палитра Virtu/JX (`#C00000` / `#1A1A1A`), без синих акцентов.
