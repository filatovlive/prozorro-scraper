# Prozorro Search Scraper

## Призначення

`prozorro_search_scraper.py` збирає тендери через публічний search API Prozorro і зберігає результат у `.xlsx`.

Поточна реалізація працює в режимі `ui_like_text_search`:

- робить `POST /api/search/tenders` з параметром `text`
- за потреби резолвить `buyer` у ЄДРПОУ через `POST /api/search/organizations`
- проходить кілька сторінок search backend
- локально відсікає записи по року

## Де знаходиться скрипт

- Скрипт: [prozorro_search_scraper.py](</tender-scraper/prozorro_search_scraper.py>)
- Output за замовчуванням: [prozorro_tenders_atomenergomash_2026.xlsx](</tender-scraper/output/prozorro_tenders_2026.xlsx>)

## Як працює скрипт

1. бере налаштування з `SCRAPER_CONFIG`
2. формує параметри запуску
3. за потреби резолвить назву організації в ЄДРПОУ
4. формує text-запит до `POST /api/search/tenders`
5. проходить сторінки search backend
6. локально залишає лише записи потрібного року
7. розплющує JSON у табличний вигляд
8. записує результат в Excel

## Основні параметри запуску

- `entity` — зараз підтримується тільки `tenders`
- `timeout` — таймаут HTTP-запитів
- `resolve_org_names` — чи резолвити назви організацій
- `resolve_exact_name_first` — чи шукати точний збіг назви перед першим наближеним
- `fetch_all_pages` — чи проходити кілька сторінок search backend
- `max_pages` — максимальна кількість сторінок
- `per_page` — розмір сторінки search backend
- `output_path` — шлях до `.xlsx`
- `filters` — набір фільтрів

## Підтримані фільтри

У поточній версії реально застосовуються тільки:

- `text`
- `buyer`
- `year`

### `text`

Текстовий запит до search backend.

Можна передавати:

- ЄДРПОУ
- назву організації
- номер закупівлі
- слова з назви закупівлі

Приклад:

```python
"text": "26444970"
```

### `buyer`

Поле для організації, яку треба перетворити в text-запит.

Можна передавати:

- ЄДРПОУ рядком
- назву організації

Як працює:

- якщо значення цифрове, воно використовується як є
- якщо це назва, скрипт викликає `search/organizations`
- результат резолву використовується не як `buyer=...`, а як `text=<buyer_id>`
- після цього скрипт локально залишає лише записи, де `procuringEntity.identifier.id` збігається з резолвленим buyer-id

Це навмисно. У прямих перевірках backend `search/tenders?buyer=...` не працював як надійний фільтр, а `text=<buyer_id>` повертав релевантні записи. Локальний post-filter по `procuringEntity.identifier.id` прибирає сторонні результати, які text search теж підтягує.

Приклад:

```python
"buyer": ['ТОВ "СУПЕРСИМЕТРІЯ"']
```

### `year`

Локальний фільтр після отримання search results.

Як працює:

- скрипт не передає `year` у backend
- він залишає лише записи, де рік дорівнює потрібному
- рік береться з `tenderID` формату `UA-YYYY-...`
- якщо `tenderID` не підходить, скрипт пробує `dateCreated`, `dateModified`, `date`

Приклад:

```python
"year": [2026]
```

## Порядок пріоритету

1. якщо `text` непорожній, використовується він
2. якщо `text` порожній, але є `buyer`, buyer резолвиться в ЄДРПОУ і використовується як `text`
3. якщо заповнений `year`, результати локально відсікаються по року

## Непідтримані поля

У шаблоні конфігу залишилися поля форми пошуку сайту, але зараз вони не застосовуються:

- `tenderer`
- `supplier`
- `producer`
- `procuringEntity`
- `status`
- `proc_type`
- `cpv`
- `cpv_mask`
- `region`
- `funders`
- `awardCriteria`
- `proc_rationale`
- `milestone`
- `local_share`
- `yearCreated`
- `contract_status`
- `product_status`
- `framework_types`
- `agreement`
- `group`
- `sort_by`
- `order`
- `page`
- `per_page`
- `value`
- `contract_value`
- `date`

Якщо будь-яке з цих полів заповнене, скрипт завершується помилкою.

## Як запускати

```bash
python "scripts\base-scraper\prozorro_search_scraper.py"
```

## Що буде в Excel

Листи:

- `data` — результати пошуку
- `meta` — службова інформація

У `meta` пишуться:

- `entity`
- `total_items_returned`
- `params_json`
- `mode`
- `text_queries_json`
- `total_by_query_json`
- `pages_fetched_by_query_json`

## Важливі обмеження

- це не strict schema-level buyer matching
- результат залежить від того, як Prozorro search backend індексує дані
- `year` — це локальна післяобробка, а не backend-фільтр
- повнота результату залежить від `fetch_all_pages`, `max_pages` і `per_page`
- search backend може мати власні обмеження глибини або поведінки пагінації

## Залежності

- стандартна бібліотека Python
- `openpyxl`
