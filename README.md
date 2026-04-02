# ELEMENTS EML Парсер

Десктоп апликација за парсирање `.eml` имејлова типа **"ELEMENTS Info: Status projekata"** и извоз чистих Excel и PDF извештаја.

---

## Садржај

- [Захтеви](#захтеви)
- [Инсталација](#инсталација)
- [Покретање](#покретање)
- [Коришћење](#коришћење)
- [Извоз извештаја](#извоз-извештаја)
- [Прављење Windows .exe фајла](#прављење-windows-exe-фајла)

---

## Захтеви

- Python 3.11+
- Windows, macOS или Linux
- Tkinter (укључен у стандардни Python)
- Фонт DejaVu Sans (препоручен за исправан приказ ћирилице у PDF-у)

---

## Инсталација

```bash
# 1. Клонирајте репозиторијум
git clone https://github.com/proditserbia/eml-parser.git
cd eml-parser

# 2. (Препоручено) Направите виртуелно окружење
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

# 3. Инсталирајте зависности
pip install -r requirements.txt
```

---

## Покретање

```bash
python main.py
```

Отвориће се главни прозор апликације.

---

## Коришћење

1. Кликните **📂 Учитај EML** и одаберите `.eml` датотеку.
2. Метаподаци, подаци о складишту и листе радних простора биће приказани аутоматски.
3. Користите поље **Претрага** за филтрирање радних простора по називу.
4. Кликните **🗑 Очисти** да бисте уклонили учитане податке.

### Тестирање парсера (самотест)

```bash
python parser.py
```

---

## Извоз извештаја

| Дугме | Формат | Опис |
|---|---|---|
| 📊 Извези у Excel | `.xlsx` | Три листа: Преглед, Активни, Старији од 90 дана |
| 📄 Извези у PDF | `.pdf` | Потпун извештај са метаподацима и листама |
| 💾 Сачувај JSON | `.json` | Сирови JSON за отклањање грешака |

---

## Прављење Windows .exe фајла

```bash
# Инсталирај PyInstaller
pip install pyinstaller

# Направи standalone .exe
pyinstaller --onefile --windowed --name "EML_Parser" main.py
```

Извршна датотека биће у фолдеру `dist/EML_Parser.exe`.

> **Напомена:** Да би ћирилица исправно радила у PDF извозу на Windows-у, уверите се да је `DejaVuSans.ttf` доступан у системском фолдеру фонтова (`C:\Windows\Fonts`).

---

## Структура пројекта

```
eml-parser/
├── main.py          # Улазна тачка апликације
├── gui.py           # Tkinter GUI (MainWindow)
├── parser.py        # EML парсер (EmlParser)
├── exporters.py     # Excel и PDF извозници
├── models.py        # Dataclass модели
├── requirements.txt # Python зависности
└── README.md        # Ова датотека
```

---

## Техничке напомене

- Парсер прво покушава да чита `text/plain` делове имејла.
- Ако `text/plain` није доступан, аутоматски прелази на `text/html` и извлачи текст.
- Подржано кодирање: UTF-8, quoted-printable.
- Апликација је тестирана на Windows 10/11 са Python 3.11.
