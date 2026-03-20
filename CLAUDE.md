# НефтеУчёт — расширение 1С:БП 3.0

## Суть

Расширение (.cfe) для загрузки сменных отчётов АЗС и поступлений топлива (ТТН) из облака в 1С:Бухгалтерию 3.0.

## Текущая задача

MVP — минимальное рабочее расширение:
1. HTTP-клиент к API STS (pos.autooplata.ru/tms)
2. Создание документов: Поступление, Комплектация, Перемещение, Розн.продажи
3. Простая форма: загрузить → провести

## Правила учёта

- **41.01** — опт (тонны), **41.02** — розница (литры)
- **Комплектация** для пересчёта тонны → литры (не Перемещение!)
- Плотность из ТТН на уровне партии
- Виртуальные склады: Карты, Талоны, Ведомости
- НДС 22%

## Префикс объектов

`НУ_` (НефтеУчёт)

## Разработка

- Исходники: `xml/` (XML-выгрузка) + `src/` (.bsl модули)
- Сборка: `scripts/build.ps1` → `1cv8.exe DESIGNER /LoadConfigFromFiles /DumpCfg`
- Язык: BSL (встроенный язык 1С)
- VS Code + BSL плагин для редактирования

## API

**Base URL:** `https://pos.autooplata.ru/tms`
- `POST /v1/login` → JWT (UserTest / sys5tem6 для тестов)
- `GET /v1/report/shift_report?system=&station=&shift=` → смена
- `GET /v1/report/receipts?system=&station=&shift=` → ТТН
- Подробнее: `ELSYPLUS/TF_1C_Projects/docs/API_AUTOOPLATA.md`

## Связь с экосистемой

- **GIG Ledger** (`OneDrive\GIG Ledger\`) — клиент, стратегия, аудиты
- **ClearLedger** (`OneDrive\Ledger\`) — платформа
- **TF_1C_Projects** (`ELSYPLUS\TF_1C_Projects\`) — шаблоны .bsl, API-доку
