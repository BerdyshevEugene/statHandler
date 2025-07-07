# statHandler

[![Ruff](https://github.com/BerdyshevEugene/statHandler/actions/workflows/ruff.yml/badge.svg?cache=buster)](https://github.com/BerdyshevEugene/statHandler/actions/workflows/ruff.yml)

---

## Описание

**statHandler** — инструмент для автоматизированной обработки статистики, получаемой через RabbitMQ, и сохранения данных в Excel-отчёты. Подходит для интеграции с корпоративными системами и автоматизации отчётности.

---

## Быстрый старт

1. **Клонируйте репозиторий:**
   ```bash
   git clone https://github.com/BerdyshevEugene/statHandler.git
   cd statHandler
   ```
2. **Установите Python 3.11+**
3. **Создайте файл `.env` в корне проекта:**
   ```env
   EXCEL_PATH=/path/to/report.xlsx
   RABBITMQ_URL=amqp://guest:guest@localhost/
   STATSCRAPER=statScraper
   ```
4. **Установите зависимости:**
   ```bash
   uv venv .venv
   uv pip install -r requirements.txt
   ```
5. **Запустите проект:**
   ```bash
   py main.py
   ```

---

## Структура проекта

<details>
<summary>Показать структуру</summary>

```python
statHandler/
│
├── app/
│   ├── config.py            # настройки (пути, переменные окружения)
│   ├── excel_processor.py   # обработка и запись данных в Excel
│   └── mq_consumer.py       # приём сообщений из RabbitMQ
│
├── main.py                  # точка входа
│
├── requirements.txt         # зависимости
├── .env                     # переменные окружения
├── logger/                  # конфиг логгера
│   └── logger.py
├── logs/                    # логи
│   └── debug.log/errors.log
└── Dockerfile               # контейнеризация
```
</details>

---

## Установка и запуск (подробно)

1. **Создайте виртуальное окружение:**
   ```bash
   uv venv .venv  # создаёт виртуальное окружение на python 3.11
   ```
2. **Установите зависимости:**
   ```bash
   uv pip install -r requirements.txt
   ```
3. **Запустите программу:**
   ```bash
   py main.py
   ```

---

## Компиляция в исполняемый файл

Для сборки standalone-скрипта используйте [PyInstaller](https://pyinstaller.org/):

```bash
pyinstaller main.py --onefile --copy-metadata aio-pika --copy-metadata pamqp --copy-metadata yarl
```

---

## Для разработчиков

### Использование UV

<details>
<summary>📦 Установка и команды UV</summary>

**Установка UV:**
- macOS/Linux:
  ```bash
  curl -LsSf https://astral.sh/uv/install.sh | sh
  ```
- Windows (PowerShell):
  ```powershell
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```
- Через PyPI:
  ```bash
  pip install uv
  ```

**Обновление UV:**
```bash
uv self update
```

**Установка Python:**
```bash
uv python install 3.13
```

**Синхронизация зависимостей:**
```bash
uv sync
```

**Запуск команд в окружении:**
```bash
uv run <COMMAND>
```
</details>

### Интеграция с Ruff

<details>
<summary>🔍 Проверка кода с помощью Ruff</summary>

[Ruff](https://github.com/astral-sh/ruff) — быстрый линтер для Python.

**Установка и запуск:**
```bash
uvx ruff
uvx ruff check .
```
</details>

---

## Лицензия и авторы

```
CompanyName: GMG
FileDescription: statHandler
InternalName: statHandler
ProductName: statHandler
Author: Berdyshev E.A.
Development and support: Berdyshev E.A.
LegalCopyright: © GMG. All rights reserved.
```