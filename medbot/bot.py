"""
Telegram-бот для анализа медицинских PDF-файлов.
Стек: python-telegram-bot + google-generativeai + openpyxl
"""

import os
import io
import json
import logging
import sqlite3
from datetime import datetime
from pathlib import Path

import google.generativeai as genai
from telegram import Update, Document
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Config (из .env) ──────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

DB_PATH = Path("data/analyses.db")
DB_PATH.parent.mkdir(exist_ok=True)

# ── База данных ───────────────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS analyses (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id   INTEGER NOT NULL,
            date      TEXT    NOT NULL,
            filename  TEXT,
            raw_json  TEXT    NOT NULL,
            created   TEXT    DEFAULT (datetime('now'))
        )
    """)
    con.commit()
    con.close()

def save_analysis(user_id: int, date: str, filename: str, data: list):
    con = sqlite3.connect(DB_PATH)
    con.execute(
        "INSERT INTO analyses (user_id, date, filename, raw_json) VALUES (?,?,?,?)",
        (user_id, date, filename, json.dumps(data, ensure_ascii=False))
    )
    con.commit()
    con.close()

def load_history(user_id: int) -> list[dict]:
    """Возвращает список {'date': ..., 'data': [...]} для пользователя."""
    con = sqlite3.connect(DB_PATH)
    rows = con.execute(
        "SELECT date, raw_json FROM analyses WHERE user_id=? ORDER BY date ASC",
        (user_id,)
    ).fetchall()
    con.close()
    return [{"date": r[0], "data": json.loads(r[1])} for r in rows]

def clear_history(user_id: int):
    con = sqlite3.connect(DB_PATH)
    con.execute("DELETE FROM analyses WHERE user_id=?", (user_id,))
    con.commit()
    con.close()

# ── Извлечение данных через Gemini ────────────────────────────────────────────
EXTRACT_PROMPT = """
Ты — медицинский ассистент. Перед тобой текст лабораторного анализа.
Извлеки ВСЕ показатели и верни ТОЛЬКО валидный JSON-массив без лишнего текста.

Формат каждого элемента:
{
  "category": "ОАК" | "Биохимия" | "Гормоны" | "Липиды" | "Другое",
  "name": "Название показателя",
  "value": <число или null>,
  "unit": "единица измерения",
  "ref_low": <нижняя граница нормы или null>,
  "ref_high": <верхняя граница нормы или null>,
  "date": "ГГГГ-ММ-ДД"  // дата из бланка, если есть
}

Правила:
- value всегда число (float/int), не строка
- Если нормы написаны как "> 60" — ref_low=60, ref_high=null
- Если "<5.0" — ref_low=null, ref_high=5.0
- Дату бери из шапки бланка. Если не найдена — сегодняшняя.
- Верни ТОЛЬКО JSON-массив, никакого текста вокруг.

Текст анализа:
"""

async def extract_markers(pdf_bytes: bytes, filename: str) -> tuple[list, str]:
    """Загружает PDF в Gemini, возвращает (список_маркеров, дата)."""
    # Загружаем файл через Files API
    file_obj = genai.upload_file(
        io.BytesIO(pdf_bytes),
        mime_type="application/pdf",
        display_name=filename
    )

    response = model.generate_content(
        [file_obj, EXTRACT_PROMPT],
        generation_config={"temperature": 0.1}
    )

    raw = response.text.strip()
    # Убираем возможные markdown-обёртки
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    markers = json.loads(raw)

    # Определяем дату
    date = datetime.today().strftime("%Y-%m-%d")
    for m in markers:
        if m.get("date"):
            date = m["date"]
            break

    return markers, date

# ── Генерация Excel ───────────────────────────────────────────────────────────
# Цветовая схема (тёмная спортивная)
DARK_BG   = "161616"
ROW_ALT   = "1E1E1E"
ROW_EVEN  = "242424"
RED_HOT   = "E8001F"
RED_DARK  = "9B0015"
RED_HEAD  = "6B0010"
GRAY_LINE = "3A3A3A"
WHITE_T   = "F2F2F2"
GRAY_T    = "AAAAAA"
GREEN_OK  = "00C851"
RED_BAD   = "FF1744"
AMBER     = "FF9100"

def fl(h): return PatternFill("solid", fgColor=h)
def fn(h, bold=False, sz=10):
    return Font(color=h, bold=bold, size=sz, name="Calibri")
thin = Side(style="thin",   color=GRAY_LINE)
red  = Side(style="medium", color=RED_HOT)
BORD = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT  = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def build_excel(history: list[dict], user_name: str = "") -> bytes:
    """
    history = [{'date': '2026-03-05', 'data': [...]}, ...]
    Возвращает байты xlsx-файла.
    """
    wb = Workbook()

    # ── Sheet 1: Сводная таблица ──────────────────────────────────────────────
    ws = wb.active
    ws.title = "Сводная таблица"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = RED_HOT

    # Фон
    for r in range(1, 300):
        for c in range(1, 15):
            ws.cell(r, c).fill = fl(DARK_BG)

    # Даты (от новых к старым)
    dates = sorted({e["date"] for h in history for e in h["data"] if e.get("date")}, reverse=True)
    if not dates:
        dates = [h["date"] for h in history]
    dates = sorted(set(dates), reverse=True)

    # Заголовок
    total_cols = 3 + len(dates) + 2  # кат + показатель + даты + ед + норма + статус
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    t = ws["A1"]
    t.value = f"  {'ПРЯЖЕНЦЕВ А.В. — ' if not user_name else user_name.upper() + ' — '}МЕДИЦИНСКИЕ АНАЛИЗЫ"
    t.font  = fn(WHITE_T, bold=True, sz=14)
    t.fill  = fl(RED_DARK)
    t.alignment = LFT
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{get_column_letter(total_cols)}2")
    s = ws["A2"]
    s.value = "  " + "  →  ".join(dates) + "   (новые слева → старые справа)"
    s.font  = fn(GRAY_T, sz=9)
    s.fill  = fl("1A0000")
    s.alignment = LFT
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6

    # Заголовки столбцов
    headers = ["Категория", "Показатель"] + [d[5:] + "\n" + d[:4] for d in dates] + ["Ед. изм.", "Норма", "Статус"]
    col_widths = [11, 32] + [12]*len(dates) + [9, 14, 8]
    ws.row_dimensions[4].height = 36
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(4, ci)
        cell.value     = h
        cell.font      = fn(WHITE_T, bold=True, sz=10)
        cell.fill      = fl(RED_HEAD)
        cell.alignment = CTR
        cell.border    = Border(left=red, right=red, top=red, bottom=red)
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "C5"

    # Собираем все уникальные показатели
    all_markers: dict[tuple, dict] = {}  # (category, name) -> {date: value, unit, ref_low, ref_high}
    for h in history:
        for m in h["data"]:
            key = (m.get("category","Другое"), m.get("name",""))
            if key not in all_markers:
                all_markers[key] = {"unit": m.get("unit",""), "ref_low": m.get("ref_low"), "ref_high": m.get("ref_high"), "dates": {}}
            all_markers[key]["dates"][m.get("date", h["date"])] = m.get("value")

    # Сортировка по категории
    CAT_ORDER = {"Гормоны":0, "Биохимия":1, "Липиды":2, "ОАК":3, "Другое":4}
    sorted_keys = sorted(all_markers.keys(), key=lambda k: (CAT_ORDER.get(k[0],5), k[1]))

    row = 5
    prev_cat = None
    for (cat, name), info in sorted_keys:
        # Заголовок категории
        if cat != prev_cat:
            if prev_cat is not None:
                for c in range(1, total_cols+1): ws.cell(row, c).fill = fl(DARK_BG)
                ws.row_dimensions[row].height = 6
                row += 1
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            ch = ws.cell(row, 1)
            ch.value = f"  ◆  {cat.upper()}"
            ch.font  = fn(RED_HOT, bold=True, sz=11)
            ch.fill  = fl("1C0005")
            ch.alignment = LFT
            ch.border = Border(left=red, bottom=Side(style="thin", color=RED_HOT))
            ws.row_dimensions[row].height = 22
            row += 1
            prev_cat = cat

        row_bg = ROW_ALT if row % 2 == 0 else ROW_EVEN
        ws.row_dimensions[row].height = 18

        # Категория
        ws.cell(row,1).value = cat; ws.cell(row,1).font = fn("555555",sz=9)
        ws.cell(row,1).fill = fl(row_bg); ws.cell(row,1).alignment = CTR
        ws.cell(row,1).border = Border(left=red, right=thin, top=thin, bottom=thin)

        # Показатель
        ws.cell(row,2).value = name; ws.cell(row,2).font = fn(WHITE_T)
        ws.cell(row,2).fill = fl(row_bg); ws.cell(row,2).alignment = LFT
        ws.cell(row,2).border = BORD

        # Значения по датам
        lo  = info.get("ref_low")
        hi  = info.get("ref_high")
        any_flag = "норма"

        for di, d in enumerate(dates):
            v    = info["dates"].get(d)
            cell = ws.cell(row, 3 + di)
            cell.alignment = CTR
            if v is None:
                cell.value = "—"; cell.font = fn("444444"); cell.fill = fl(row_bg)
            else:
                disp = round(float(v),3) if isinstance(v,float) and v != int(v) else v
                is_hi = hi is not None and hi < 999 and float(v) > hi
                is_lo = lo is not None and lo > 0  and float(v) < lo
                if is_hi or is_lo:
                    cell.fill = fl("3B0008"); cell.font = fn(RED_BAD, bold=True)
                    cell.value = str(disp) + (" ▲" if is_hi else " ▼")
                    any_flag = "выше" if is_hi else "ниже"
                else:
                    cell.fill = fl(row_bg); cell.font = fn(GREEN_OK)
                    cell.value = disp
            cell.border = BORD

        # Единица
        cu = ws.cell(row, 3 + len(dates))
        cu.value = info.get("unit",""); cu.font = fn(GRAY_T,sz=9)
        cu.fill = fl(row_bg); cu.alignment = CTR; cu.border = BORD

        # Норма (один столбец)
        cr = ws.cell(row, 4 + len(dates))
        if hi is None and lo is None:   ref_txt = "—"
        elif hi is None or hi >= 999:   ref_txt = f"> {lo}"
        elif lo is None or lo == 0:     ref_txt = f"< {hi}"
        else:                           ref_txt = f"{lo} – {hi}"
        cr.value = ref_txt; cr.font = fn(GRAY_T,sz=9)
        cr.fill = fl(row_bg); cr.alignment = CTR; cr.border = BORD

        # Статус
        cs = ws.cell(row, 5 + len(dates))
        if any_flag == "выше":
            cs.value="▲ ВЫШЕ"; cs.font=fn(RED_BAD,bold=True,sz=9); cs.fill=fl("2A0000")
        elif any_flag == "ниже":
            cs.value="▼ НИЖЕ"; cs.font=fn(AMBER,bold=True,sz=9); cs.fill=fl("1A1000")
        else:
            cs.value="✓ НОРМА"; cs.font=fn(GREEN_OK,sz=9); cs.fill=fl("001A05")
        cs.alignment = CTR
        cs.border = Border(left=thin, right=red, top=thin, bottom=thin)

        row += 1

    # ── Sheet 2: Текстовая выжимка ────────────────────────────────────────────
    ws2 = wb.create_sheet("Выжимка")
    ws2.sheet_view.showGridLines = False
    ws2.sheet_properties.tabColor = RED_HOT
    for r in range(1, 60):
        for c in range(1, 5):
            ws2.cell(r, c).fill = fl(DARK_BG)

    ws2.column_dimensions["A"].width = 3
    ws2.column_dimensions["B"].width = 24
    ws2.column_dimensions["C"].width = 68

    ws2.merge_cells("A1:C1")
    ws2["A1"].value = "  ПОКАЗАТЕЛИ ВНЕ НОРМЫ"
    ws2["A1"].font  = fn(WHITE_T, bold=True, sz=13)
    ws2["A1"].fill  = fl(RED_DARK)
    ws2["A1"].alignment = LFT
    ws2.row_dimensions[1].height = 28

    r2 = 3
    for (cat, name), info in sorted_keys:
        lo = info.get("ref_low"); hi = info.get("ref_high")
        flags = []
        for d, v in info["dates"].items():
            if v is None: continue
            is_hi = hi is not None and hi < 999 and float(v) > hi
            is_lo = lo is not None and lo > 0  and float(v) < lo
            if is_hi or is_lo:
                flags.append(f"{d[:10]}: {v} {'▲' if is_hi else '▼'}")
        if not flags: continue

        ws2.cell(r2,2).value = name; ws2.cell(r2,2).font = fn(RED_BAD,bold=True,sz=10)
        ws2.cell(r2,2).fill = fl(ROW_ALT); ws2.cell(r2,2).alignment = LFT
        ws2.cell(r2,2).border = Border(left=red, right=thin, top=thin, bottom=thin)

        ref_str = f"норма {ref_txt}" if (ref_txt := (f"{lo}–{hi}" if lo and hi and hi<999 else (f">{lo}" if hi is None or hi>=999 else f"<{hi}"))) else ""
        ws2.cell(r2,3).value = "  |  ".join(flags) + (f"   [{ref_str}]" if ref_str else "")
        ws2.cell(r2,3).font = fn(WHITE_T,sz=10)
        ws2.cell(r2,3).fill = fl(ROW_ALT); ws2.cell(r2,3).alignment = LFT
        ws2.cell(r2,3).border = Border(left=thin, right=red, top=thin, bottom=thin)
        ws2.row_dimensions[r2].height = 18
        r2 += 1

    # Сохраняем в байты
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Хендлеры Telegram ─────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name or "друг"
    await update.message.reply_text(
        f"Привет, {name}! 👋\n\n"
        "Я анализирую медицинские PDF-файлы и создаю Excel-отчёт с:\n"
        "• Таблицей показателей и подсветкой отклонений\n"
        "• Динамикой по времени (если пришлёте несколько анализов)\n"
        "• Выжимкой того, что вне нормы\n\n"
        "Просто пришлите PDF — и я всё сделаю.\n\n"
        "Команды:\n"
        "/history — посмотреть сколько анализов сохранено\n"
        "/report  — сгенерировать Excel из всех сохранённых анализов\n"
        "/clear   — удалить историю и начать заново"
    )

async def cmd_history(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    h = load_history(uid)
    if not h:
        await update.message.reply_text("Анализов пока нет. Пришлите PDF!")
        return
    lines = [f"  {i+1}. {e['date']}" for i, e in enumerate(h)]
    await update.message.reply_text("Сохранённые анализы:\n" + "\n".join(lines))

async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    name = update.effective_user.full_name or ""
    h    = load_history(uid)
    if not h:
        await update.message.reply_text("Нет сохранённых анализов. Сначала пришлите PDF!")
        return
    await update.message.reply_text(f"Генерирую Excel из {len(h)} анализов...")
    xlsx = build_excel(h, name)
    await update.message.reply_document(
        document=io.BytesIO(xlsx),
        filename=f"Анализы_{datetime.today().strftime('%Y%m%d')}.xlsx",
        caption="Готово! Открывайте в Excel или Google Таблицах."
    )

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    clear_history(update.effective_user.id)
    await update.message.reply_text("История очищена. Можно загружать новые анализы.")

async def handle_pdf(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc: Document = update.message.document
    if not doc.mime_type == "application/pdf":
        await update.message.reply_text("Пришлите файл в формате PDF.")
        return

    uid  = update.effective_user.id
    name = update.effective_user.full_name or ""

    msg = await update.message.reply_text("⏳ Читаю PDF и извлекаю показатели...")

    try:
        file = await ctx.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()

        markers, date = await extract_markers(bytes(pdf_bytes), doc.file_name or "analysis.pdf")

        save_analysis(uid, date, doc.file_name, markers)

        # Быстрая выжимка
        flags = [m for m in markers if m.get("value") is not None and (
            (m.get("ref_high") is not None and m["ref_high"] < 999 and float(m["value"]) > m["ref_high"]) or
            (m.get("ref_low")  is not None and m["ref_low"]  > 0   and float(m["value"]) < m["ref_low"])
        )]
        flag_text = ""
        if flags:
            lines = [f"• {f['name']}: {f['value']} {f.get('unit','')} {'▲' if float(f['value']) > (f.get('ref_high') or 1e9) else '▼'}" for f in flags[:8]]
            flag_text = "\n\n⚠️ Отклонения:\n" + "\n".join(lines)
            if len(flags) > 8:
                flag_text += f"\n...и ещё {len(flags)-8}"

        await msg.edit_text(
            f"✅ Извлечено {len(markers)} показателей (дата: {date}){flag_text}\n\n"
            f"Нажмите /report чтобы скачать Excel со всеми анализами."
        )

        # Если анализов >= 1 — сразу генерируем Excel
        h = load_history(uid)
        xlsx = build_excel(h, name)
        await update.message.reply_document(
            document=io.BytesIO(xlsx),
            filename=f"Анализы_{date.replace('-','')}.xlsx",
            caption=f"Excel готов! Всего анализов в истории: {len(h)}"
        )

    except json.JSONDecodeError:
        await msg.edit_text("Не смог разобрать ответ AI. Попробуйте другой файл.")
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        await msg.edit_text(f"Ошибка при обработке: {str(e)[:200]}")

# ── Запуск ────────────────────────────────────────────────────────────────────
def main():
    init_db()
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("history", cmd_history))
    app.add_handler(CommandHandler("report",  cmd_report))
    app.add_handler(CommandHandler("clear",   cmd_clear))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))

    logger.info("Бот запущен!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
