# app.py
import asyncio
import os
import re
import warnings
from datetime import datetime
from io import BytesIO
from threading import Thread

import pandas as pd
import pymupdf
from aiogram import Bot, Dispatcher, F, types
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.types import BufferedInputFile
from flask import Flask

# --- Инициализация Flask для Render ---
app = Flask(__name__)

@app.route('/')
def health_check():
    return "Bot is running", 200

# --- Код бота (парсер и хэндлеры) ---
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
if not TELEGRAM_TOKEN:
    raise SystemExit("Не задан TELEGRAM_TOKEN в переменных окружения")

warnings.filterwarnings("ignore", category=FutureWarning)

COLUMNS = [
    ("№", 0, 20), ("Код", 20, 31), ("Номенклатура", 31, 60),
    ("Количество", 60, 71), ("Цена", 71, 82), ("Сумма", 82, 93),
    ("Сумма продажи", 93, 500),
]
SERVICE_RE = re.compile(
    r"(Итого|Итоговая|Итог|Всего|Сумма продаж|руб\.|рубль|копеек|"
    r"Комитент|Комиссионер|Отчет комитенту|Наименован)",
    re.IGNORECASE,
)
REPORT_DATE_RE = re.compile(r"от\s+(\d{2})\.(\d{2})\.(\d{4})")
NUM_RE = r"(\d[\d\s.,]*)"

def col_of(x):
    for name, x0, x1 in COLUMNS:
        if x0 <= x < x1:
            return name
    return None

def parse_num(s):
    if s is None:
        return None
    s = str(s).replace("\xa0", "").replace(" ", "").strip().replace(",", ".")
    if s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(s) if s not in ("", "-") else None
    except ValueError:
        return None

def group_rows(words, y_tol=1.5):
    rows = []
    for w in sorted(words, key=lambda w: (w[1], w[0])):
        x0, y0 = w[0], w[1]
        text = w[4]
        placed = False
        for r in rows:
            if abs(r["y"] - y0) < y_tol:
                r["words"].append((x0, text))
                placed = True
                break
        if not placed:
            rows.append({"y": y0, "words": [(x0, text)]})
    return rows

def extract_report_date(doc):
    for page in doc:
        text = page.get_text()[:500]
        m = REPORT_DATE_RE.search(text)
        if m:
            d, mo, y = m.groups()
            return datetime(int(y), int(mo), int(d)).date()
    return None

def collect_lines_with_pymupdf(doc):
    lines = []
    for page in doc:
        words = page.get_text("words")
        if not words:
            continue
        rows = group_rows(words)
        for r in rows:
            line = " ".join(w[1] for w in sorted(r["words"], key=lambda p: p[0]))
            lines.append(line)
    return lines

def extract_total_from_pdf(doc):
    lines = collect_lines_with_pymupdf(doc)
    joined = " ".join(lines).replace("\xa0", " ")
    joined = re.sub(r"\s+", " ", joined)
    patterns = [
        rf"Итог\w*\s+сумм\w*\s+продаж\w*[^\d]*{NUM_RE}\s*руб",
        rf"Итог\w*[^\d]{{0,30}}?{NUM_RE}\s*руб",
        rf"на\s+сумму\s+{NUM_RE}\s*руб",
    ]
    for pat in patterns:
        m = re.search(pat, joined, re.IGNORECASE)
        if m:
            return parse_num(m.group(1))
    return None

def parse_report(pdf_bytes: bytes):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    records = []
    current = None

    for page in doc:
        words = page.get_text("words")
        rows = group_rows(words)
        header_seen = False
        for r in rows:
            cells = {name: [] for name, _, _ in COLUMNS}
            for x, text in sorted(r["words"], key=lambda p: p[0]):
                c = col_of(x)
                if c:
                    cells[c].append(text)
            row_data = {name: " ".join(v).strip() for name, v in cells.items()}
            if not header_seen and "Номенклатура" in row_data.values():
                header_seen = True
                continue
            if not header_seen:
                continue
            num = row_data.get("№", "")
            if re.match(r"^\d+$", num):
                if current:
                    records.append(current)
                name = SERVICE_RE.split(row_data["Номенклатура"])[0].strip()
                current = {
                    "№": int(num),
                    "Код": row_data["Код"],
                    "Номенклатура": name,
                    "Количество": parse_num(row_data["Количество"]),
                    "Цена": parse_num(row_data["Цена"]),
                    "Сумма": parse_num(row_data["Сумма"]),
                    "Сумма продажи": parse_num(row_data["Сумма продажи"]),
                }
            else:
                cont = row_data.get("Номенклатура", "").strip()
                if not cont or SERVICE_RE.search(cont) or not re.search(r"[A-Za-zА-Яа-я]", cont):
                    continue
                if current:
                    current["Номенклатура"] = (
                        current["Номенклатура"] + " " + cont
                    ).strip()
        if current:
            records.append(current)
            current = None

    report_date = extract_report_date(doc)
    pdf_total = extract_total_from_pdf(doc)

    df = pd.DataFrame(records)
    if not df.empty:
        df["Дата отчёта"] = report_date
    return df, pdf_total

bot = Bot(token=TELEGRAM_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

@dp.message(CommandStart())
async def command_start_handler(message: types.Message) -> None:
    await message.answer(
        f"Привет, {message.from_user.full_name}!\n\n"
        "Я помогу превратить PDF-отчёт от дистрибьютора в удобный Excel-файл.\n"
        "Просто отправь мне PDF-документ."
    )

@dp.message(F.document)
async def document_handler(message: types.Message) -> None:
    document = message.document
    file_name = document.file_name or "report.pdf"

    if not file_name.lower().endswith(".pdf"):
        await message.answer("Пожалуйста, отправьте файл в формате PDF.")
        return

    processing_msg = await message.answer("📄 Обрабатываю ваш отчёт...")

    try:
        file_info = await bot.get_file(document.file_id)
        file_bytes_io = await bot.download_file(file_info.file_path)
        pdf_bytes = file_bytes_io.read()

        df, pdf_total = parse_report(pdf_bytes)

        if df.empty:
            await processing_msg.edit_text("❌ Не удалось найти данные в этом PDF.")
            return

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Продажи", index=False)
            summary = pd.DataFrame({
                "Показатель": ["Файл", "Позиций", "Сумма продажи", "Итог из PDF"],
                "Значение": [
                    file_name, len(df), df["Сумма продажи"].sum(),
                    pdf_total if pdf_total is not None else "Не найден",
                ],
            })
            summary.to_excel(writer, sheet_name="Сводка", index=False)
        output.seek(0)

        out_name = file_name.rsplit(".", 1)[0] + ".xlsx"
        await message.answer_document(
            BufferedInputFile(output.read(), filename=out_name),
            caption=f"✅ Готово! Найдено {len(df)} позиций.\nСумма продаж: {df['Сумма продажи'].sum():,.2f} ₽"
        )
        await processing_msg.delete()

    except Exception as e:
        await processing_msg.edit_text(f"⚠️ Произошла ошибка: {e}")

# --- Запуск бота в отдельном потоке ---
def run_bot():
    asyncio.run(dp.start_polling(bot, handle_signals=False))

def start_bot_thread():
    bot_thread = Thread(target=run_bot)
    bot_thread.daemon = True  # Поток завершится вместе с основным процессом
    bot_thread.start()

# --- Запуск ---
if __name__ == "__main__":
    start_bot_thread()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
