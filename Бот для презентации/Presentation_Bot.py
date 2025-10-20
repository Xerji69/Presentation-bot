# Presentation_Bot_part1.py
# ЧАСТЬ 1/3 — инициализация, БД, шаблоны, базовые хендлеры
# Сохрани как Presentation_Bot_part1.py и дождись частей 2/3,
# или вставляй в начало итогового Presentation_Bot.py.

import os
import sys
import logging
import sqlite3
from pathlib import Path
from datetime import datetime, timezone, timedelta
from io import BytesIO
import tempfile
import shutil
import asyncio

from dotenv import load_dotenv
from telegram import (
    Update, InputFile, KeyboardButton, ReplyKeyboardMarkup,
    InlineKeyboardButton, InlineKeyboardMarkup
)
from telegram.ext import (
    ApplicationBuilder, ContextTypes, CommandHandler,
    MessageHandler, filters, ConversationHandler, CallbackQueryHandler
)

# PPTX libs (частично будут использоваться в Части 2)
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image, ImageDraw, ImageFont

# Try import openai (может быть не установлено) — в Части 2 будет mock-AI
try:
    import openai
    OPENAI_AVAILABLE = True
except Exception:
    OPENAI_AVAILABLE = False

# -------------------------
# Конфигурация / директории
# -------------------------
load_dotenv()  # загружает .env из рабочей папки, если есть
print("BOT_TOKEN =", os.getenv("BOT_TOKEN"))  # ← проверка

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()  # ← токен берётся из .env
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()  # опционально
BOT_DISPLAY_NAME = os.getenv("BOT_DISPLAY_NAME", "Presentation_Bot.ru")
DATA_DIR = Path(os.getenv("DATA_DIR", "data"))
STATIC_DIR = Path(os.getenv("STATIC_DIR", "user_data"))
TEMPLATES_DIR = Path(os.getenv("TEMPLATES_DIR", "templates"))
LOG_DIR = Path(os.getenv("LOG_DIR", "logs"))
RETENTION_DAYS = int(os.getenv("RETENTION_DAYS", "40"))

# Админ (твой ID подставлен в финале; можно расширить через ENV ADMIN_IDS)
DEFAULT_ADMIN_ID = -1003105251694
ADMIN_IDS = [DEFAULT_ADMIN_ID] + [int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]

# Создаём папки при необходимости
for p in (DATA_DIR, STATIC_DIR, TEMPLATES_DIR, LOG_DIR):
    p.mkdir(parents=True, exist_ok=True)

# -------------------------
# Логирование
# -------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "bot.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Инициализация OpenAI (если указан ключ и библиотека установлена)
if OPENAI_API_KEY and OPENAI_AVAILABLE:
    openai.api_key = OPENAI_API_KEY
elif OPENAI_API_KEY and not OPENAI_AVAILABLE:
    logger.warning("OPENAI_API_KEY задан, но пакет openai не установлен. Установите openai или оставьте ключ пустым.")

# -------------------------
# SQLite: инициализация + обёртки
# -------------------------
DB_PATH = DATA_DIR / "presentation_bot.db"

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            first_name TEXT,
            last_name TEXT,
            language_code TEXT,
            first_seen TEXT,
            last_seen TEXT,
            interactions INTEGER DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS presentations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            filename TEXT,
            created_at TEXT,
            template TEXT,
            title TEXT,
            path TEXT,
            theme TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            text TEXT,
            created_at TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS stats_topics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            topic TEXT,
            count INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()
logger.info("DB initialized at %s", DB_PATH)

def touch_user(user):
    """Добавляет или обновляет запись о пользователе."""
    if user is None:
        return
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now(timezone.utc).isoformat()
    cur.execute("SELECT user_id FROM users WHERE user_id = ?", (user.id,))
    row = cur.fetchone()
    if row:
        cur.execute("""
            UPDATE users SET username=?, first_name=?, last_name=?, language_code=?, last_seen=?, interactions = interactions + 1
            WHERE user_id = ?
        """, (user.username, user.first_name, user.last_name, user.language_code, now, user.id))
    else:
        cur.execute("""
            INSERT INTO users (user_id, username, first_name, last_name, language_code, first_seen, last_seen, interactions)
            VALUES (?, ?, ?, ?, ?, ?, ?, 1)
        """, (user.id, user.username, user.first_name, user.last_name, user.language_code, now, now))
    conn.commit()
    conn.close()

def add_presentation(user_id, filename, template, title, path, theme):
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now(timezone.utc).isoformat()
    cur.execute("""
        INSERT INTO presentations (user_id, filename, created_at, template, title, path, theme)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (user_id, filename, now, template, title, path, theme))
    conn.commit(); conn.close()

def list_user_presentations(user_id, limit=50):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM presentations WHERE user_id = ? ORDER BY created_at DESC LIMIT ?", (user_id, limit))
    rows = cur.fetchall(); conn.close()
    return rows

def stats_topic_increment(topic):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id,count FROM stats_topics WHERE topic = ?", (topic,))
    r = cur.fetchone()
    if r:
        cur.execute("UPDATE stats_topics SET count = count + 1 WHERE id = ?", (r["id"],))
    else:
        cur.execute("INSERT INTO stats_topics (topic, count) VALUES (?, 1)", (topic,))
    conn.commit(); conn.close()

def feedback_add(user_id, text):
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now(timezone.utc).isoformat()
    cur.execute("INSERT INTO feedback (user_id, text, created_at) VALUES (?, ?, ?)", (user_id, text, now))
    conn.commit(); conn.close()

# -------------------------
# Шаблоны / темы (создание минимального шаблона если отсутствует)
# -------------------------
TEMPLATE_NAMES = [
    "modern_dark", "modern_light", "corporate_blue", "corporate_gray",
    "creative_color", "minimal_white", "minimal_black", "photo_background",
    "gradient", "classic"
]

THEMES = [
    "Business", "Education", "Technology", "Healthcare", "Marketing",
    "Finance", "Science", "Art & Design", "Travel", "Startup Pitch"
]

def create_minimal_template(path: Path, name: str):
    prs = Presentation()
    layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(1.0))
    tb.text_frame.paragraphs[0].text = f"{name.replace('_',' ').title()} Template"
    tb.text_frame.paragraphs[0].font.size = Pt(28)
    tb.text_frame.paragraphs[0].font.bold = True
    prs.save(path)
    logger.info("Created minimal template %s", path)

def ensure_templates():
    for name in TEMPLATE_NAMES:
        p = TEMPLATES_DIR / f"{name}.pptx"
        if not p.exists():
            create_minimal_template(p, name)

# -------------------------
# UI helpers (keyboards)
# -------------------------
def main_menu_kb(lang="ru"):
    kb = [
        [KeyboardButton("/new"), KeyboardButton("/ai")],
        [KeyboardButton("/templates"), KeyboardButton("/history")],
        [KeyboardButton("/my_presentations"), KeyboardButton("/privacy")]
    ]
    return ReplyKeyboardMarkup(kb, resize_keyboard=True)

# -------------------------
# Базовые хендлеры
# -------------------------
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    touch_user(user)
    lang = (user.language_code or "ru").lower()
    text = ("Привет! Я бот для создания презентаций.\n"
            "Команды: /new, /ai, /templates, /history, /my_presentations, /privacy, /help")
    if lang.startswith("en"):
        text = ("Hello! I'm a presentation generator bot.\n"
                "Commands: /new, /ai, /templates, /history, /my_presentations, /privacy, /help")
    await update.message.reply_text(text, reply_markup=main_menu_kb(lang))

async def help_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user; touch_user(user)
    text = (
        "/new — создать презентацию вручную\n"
        "/ai — AI-режим (если задан OPENAI_API_KEY)\n"
        "/templates — выбрать шаблон\n"
        "/upload_images — загрузить изображения\n"
        "/my_presentations — скачать архив ваших презентаций\n"
        "/privacy — политика конфиденциальности\n"
        "/delete_my_data — удалить все ваши файлы и данные"
    )
    await update.message.reply_text(text, reply_markup=main_menu_kb())

async def templates_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user; touch_user(user)
    buttons = []
    for t in TEMPLATE_NAMES:
        label = t.replace("_", " ").title()
        buttons.append([InlineKeyboardButton(label, callback_data=f"tpl::{t}")])
    await update.message.reply_text("Выберите шаблон (нажмите кнопку):", reply_markup=InlineKeyboardMarkup(buttons))

async def tpl_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    data = q.data or ""
    if data.startswith("tpl::"):
        tpl = data.split("::", 1)[1]
        context.user_data['template_key'] = tpl
        path = TEMPLATES_DIR / f"{tpl}.pptx"
        context.user_data['template_path'] = str(path) if path.exists() else None
        await q.edit_message_text(f"Шаблон *{tpl}* выбран. Теперь отправьте /new чтобы начать создание.", parse_mode="Markdown")

# -------------------------
# Заготовка build_app (завершится в Части 2/3)
# -------------------------
def build_app():
    # Инициализация БД и шаблонов
    init_db()
    ensure_templates()

    # Если есть OPENAI ключ и установлен пакет — настроим (частично)
    if OPENAI_AVAILABLE and OPENAI_API_KEY:
        openai.api_key = OPENAI_API_KEY

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # Регистрируем базовые обработчики
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("help", help_handler))
    app.add_handler(CommandHandler("templates", templates_handler))
    app.add_handler(CallbackQueryHandler(tpl_callback, pattern="^tpl::"))

    # Заготовки: в Части 2 я добавлю полную логику /new, /ai, загрузку изображений, генерацию pptx/pdf/png
    # Также в Части 3 добавлю admin, cleanup task, feedback, archive и финальный run.
    return app

# Если запустили эту часть отдельно — проверим, что всё инициализируется корректно.
if __name__ == "__main__":
    print("Presentation_Bot — Часть 1. Инициализация и базовые хендлеры.")
    init_db()
    ensure_templates()
    print("DB and templates created. Добавь Части 2 и 3 чтобы иметь полностью рабочий бот.")




# ЧАСТЬ 2/3 — генерация презентаций (ручная, AI и mock-AI), экспорт
# ===============================

import re
import textwrap
from pptx.enum.text import PP_ALIGN

# ---------- Mock AI генерация ----------
def local_mock_generate_markdown(topic: str) -> str:
    """
    Простая локальная генерация презентации без OpenAI.
    Возвращает markdown-текст с разделением слайдов через '---'.
    """
    topic = topic.strip().capitalize()
    slides = [
        f"# {topic}\n### Презентация",
        f"## Введение\n- Почему {topic.lower()} важно\n- Краткое определение\n- Исторический контекст",
        f"## Проблема\n- Основные вызовы\n- Ошибки и трудности\n- Примеры из практики",
        f"## Решения\n- Эффективные методы\n- Современные подходы\n- Применение технологий",
        f"## Преимущества\n- Экономия времени\n- Увеличение эффективности\n- Применимость",
        f"## Примеры\n- Реальные кейсы\n- Демонстрации\n- Краткие результаты",
        f"## Заключение\n- Ключевые выводы\n- Перспективы развития\n- Благодарность за внимание!"
    ]
    return "\n\n---\n\n".join(slides)

async def ai_generate_markdown(topic: str) -> str:
    """Вызывает реальный OpenAI API, если доступен, иначе mock."""
    if OPENAI_API_KEY and OPENAI_AVAILABLE:
        try:
            completion = await asyncio.to_thread(
                openai.chat.completions.create,
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": f"Создай презентацию по теме: {topic}. Формат: Markdown со слайдами, разделёнными ---"}],
                max_tokens=700
            )
            content = completion.choices[0].message.content
            return content
        except Exception as e:
            logger.error("OpenAI API error: %s", e)
            return local_mock_generate_markdown(topic)
    else:
        return local_mock_generate_markdown(topic)

# ---------- Конвертация markdown -> PPTX ----------
def md_to_pptx(md_text: str, template_path: str | None = None) -> BytesIO:
    prs = Presentation(template_path) if template_path and os.path.exists(template_path) else Presentation()
    slides = re.split(r'\n---+\n', md_text)
    for raw_slide in slides:
        lines = [ln.strip() for ln in raw_slide.strip().splitlines() if ln.strip()]
        if not lines:
            continue
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        y = Inches(0.8)
        for line in lines:
            if line.startswith("# "):
                tf = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.5)).text_frame
                tf.text = line[2:]
                tf.paragraphs[0].font.size = Pt(36)
                tf.paragraphs[0].font.bold = True
            elif line.startswith("## "):
                tf = slide.shapes.add_textbox(Inches(0.5), y, Inches(9), Inches(1.0)).text_frame
                tf.text = line[3:]
                tf.paragraphs[0].font.size = Pt(28)
                tf.paragraphs[0].font.bold = True
                y += Inches(0.8)
            elif line.startswith("- "):
                tf = slide.shapes.add_textbox(Inches(1.0), y, Inches(8), Inches(3.0)).text_frame
                for b in [l[2:] for l in lines if l.startswith("- ")]:
                    p = tf.add_paragraph()
                    p.text = b
                    p.font.size = Pt(20)
                break
            else:
                tf = slide.shapes.add_textbox(Inches(1.0), y, Inches(8.5), Inches(3.0)).text_frame
                p = tf.paragraphs[0]
                p.text = line
                p.font.size = Pt(20)
                p.alignment = PP_ALIGN.LEFT
                y += Inches(0.6)
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ---------- Экспорт PPTX -> PDF / PNG ----------
def export_pptx(pptx_buffer: BytesIO, base_filename: str, export_dir: Path):
    export_dir.mkdir(parents=True, exist_ok=True)

    pptx_path = export_dir / f"{base_filename}.pptx"
    with open(pptx_path, "wb") as f:
        f.write(pptx_buffer.getvalue())

    # Простая PNG-обложка
    cover = Image.new("RGB", (800, 450), (40, 40, 60))
    draw = ImageDraw.Draw(cover)
    title_text = textwrap.fill(base_filename, width=25)
    draw.text((40, 180), title_text, fill="white")
    cover.save(export_dir / f"{base_filename}.png", "PNG")

    # Фейковая PDF-заглушка
    pdf_path = export_dir / f"{base_filename}.pdf"
    with open(pdf_path, "w") as f:
        f.write("PDF-файл Mock PDF placeholder\n")

    return {
        "pptx_path": pptx_path,
        "png_path": export_dir / f"{base_filename}.png",
        "pdf_path": pdf_path
    }

# ---------- Команды /new и /ai ----------
NEW_STATE_TOPIC = range(1)

async def new_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user; touch_user(user)
    await update.message.reply_text("Введите тему презентации:")
    return NEW_STATE_TOPIC

async def new_topic_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = update.message.text.strip()
    user = update.effective_user
    touch_user(user)

    md = local_mock_generate_markdown(topic)
    tpl = context.user_data.get("template_path")
    pptx_buffer = md_to_pptx(md, tpl)
    filename = f"{topic.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}"
    user_dir = STATIC_DIR / str(user.id)
    paths = export_pptx(pptx_buffer, filename, user_dir)
    add_presentation(user.id, f"{filename}.pptx", tpl or "default", topic, str(paths["pptx"]), "manual")

    await update.message.reply_text(f"✅ Презентация по теме *{topic}* готова!", parse_mode="Markdown")
    await update.message.reply_document(document=InputFile(paths["pptx"]), caption="Ваш файл .pptx")
    return ConversationHandler.END

async def ai_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    touch_user(user)
    await update.message.reply_text("Отправьте тему для AI-презентации:")
    return NEW_STATE_TOPIC

async def ai_topic_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = update.message.text.strip()
    user = update.effective_user
    touch_user(user)
    await update.message.reply_text("🧠 Генерирую презентацию, подождите...")
    md = await ai_generate_markdown(topic)
    tpl = context.user_data.get("template_path")
    pptx_buffer = md_to_pptx(md, tpl)
    filename = f"AI_{topic.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}"
    user_dir = STATIC_DIR / str(user.id)
    paths = export_pptx(pptx_buffer, filename, user_dir)
    add_presentation(user.id, f"{filename}.pptx", tpl or "default", topic, str(paths["pptx"]), "AI")
    await update.message.reply_text(f"✅ AI-презентация по теме *{topic}* готова!", parse_mode="Markdown")
    await update.message.reply_document(document=InputFile(paths["pptx"]), caption="Ваш файл .pptx")
    stats_topic_increment(topic)
    return ConversationHandler.END

# ---------- История /my_presentations ----------
async def my_presentations_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    touch_user(user)
    presentations = list_user_presentations(user.id)
    if not presentations:
        await update.message.reply_text("У вас пока нет презентаций.")
        return
    lines = []
    for row in presentations:
        dt = row["created_at"].split("T")[0]
        lines.append(f"• {row['title']} ({dt}) — {row['template']}")
    await update.message.reply_text("\n".join(lines), reply_markup=main_menu_kb())

# ---------- Обновление build_app() ----------
def build_app():
    init_db()
    ensure_templates()
    if OPENAI_AVAILABLE and OPENAI_API_KEY:
        openai.api_key = OPENAI_API_KEY

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # Базовые хендлеры из Части 1
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("help", help_handler))
    app.add_handler(CommandHandler("templates", templates_handler))
    app.add_handler(CallbackQueryHandler(tpl_callback, pattern="^tpl:"))

    # Создание вручную
    new_conv = ConversationHandler(
        entry_points=[CommandHandler("new", new_start)],
        states={NEW_STATE_TOPIC: [MessageHandler(filters.TEXT & ~filters.COMMAND, new_topic_received)]},
        fallbacks=[]
    )
    app.add_handler(new_conv)

    # AI генерация
    ai_conv = ConversationHandler(
        entry_points=[CommandHandler("ai", ai_start)],
        states={NEW_STATE_TOPIC: [MessageHandler(filters.TEXT & ~filters.COMMAND, ai_topic_received)]},
        fallbacks=[]
    )
    app.add_handler(ai_conv)

    app.add_handler(CommandHandler("my_presentations", my_presentations_handler))

    return app   # ← теперь внутри функции, ошибка исчезнет



# ЧАСТЬ 3/3 — админка, privacy, удаление, фоновая очистка, запуск
# ===============================

import asyncio

# ---------- Privacy и удаление данных ----------
async def privacy_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    touch_user(user)
    lang = (user.language_code or "ru").lower()
    if lang.startswith("en"):
        await update.message.reply_text(PRIVACY_EN if 'PRIVACY_EN' in globals() else "Privacy policy not available.")
    else:
        await update.message.reply_text(PRIVACY_RU if 'PRIVACY_RU' in globals() else "Политика конфиденциальности не задана.")

async def delete_my_data_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not user:
        await update.message.reply_text("Не удалось получить информацию о пользователе.")
        return
    touch_user(user)
    # удаляем файлы пользователя
    user_dir = STATIC_DIR / str(user.id)
    try:
        if user_dir.exists():
            shutil.rmtree(user_dir)
    except Exception as e:
        logger.exception("Failed to remove user dir: %s", e)
    # удаляем из БД
    conn = get_conn(); cur = conn.cursor()
    cur.execute("DELETE FROM presentations WHERE user_id = ?", (user.id,))
    cur.execute("DELETE FROM users WHERE user_id = ?", (user.id,))
    conn.commit(); conn.close()
    await update.message.reply_text("Ваши локальные файлы и данные были удалены. Если хотите полностью удалить — также удалите резервные копии, если они есть.")

# ---------- Feedback (если не добавляли ранее) ----------
# Если в Части 1/2 уже добавлены feedback handlers — повторы безопасны (функции просто перезапишутся).
async def feedback_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user; touch_user(user)
    await update.message.reply_text("Спасибо, оставьте ваш отзыв — я сохраню его для администрации. Отправьте текст сейчас.")
    return 1

async def feedback_receive(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user; touch_user(user)
    text = update.message.text or ""
    if not text.strip():
        await update.message.reply_text("Пустой отзыв не сохранён.")
        return ConversationHandler.END
    feedback_add(user.id, text.strip())
    await update.message.reply_text("Спасибо! Ваш отзыв сохранён.")
    return ConversationHandler.END

# ---------- Админ-панель ----------
# -------- Админ-панель --------
async def admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pass  # <-- временно пустая функция, чтобы не было ошибки

async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users")
    total_users = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM presentations")
    total_presentations = cur.fetchone()[0]
    cur.execute("SELECT topic, COUNT(*) FROM interactions GROUP BY topic")
    popular = cur.fetchall()
    cur.execute("SELECT user_id, username, first_seen, last_seen, interactions FROM users ORDER BY last_seen DESC LIMIT 15")
    recent = cur.fetchall()
    conn.close()

    text = f"📊 Статистика ({BOT_DISPLAY_NAME})\n\n"
    text += f"👥 Пользователи: {total_users}\n🗂 Презентации: {total_presentations}\n\n"
    text += "🔥 Популярные темы:\n"
    if popular:
        for r in popular:
            text += f" - {r[0]}: {r[1]}\n"
    else:
        text += "Нет данных.\n"
    text += "\n🕒 Недавние пользователи:\n"
    for r in recent:
        uname = r[1] or str(r[0])
        text += f" - {uname} (id: {r[0]}, hits: {r[4]})\n"

    await update.message.reply_text(text)
    
# ---------- Фоновая задача: очистка старых файлов (и DB-очистка) ----------
async def cleanup_task(app):
    await asyncio.sleep(2)  # небольшой стартовый таймаут
    logger.info("Cleanup task started (retention days = %s)", RETENTION_DAYS)
    while True:
        try:
            cutoff = datetime.now(timezone.utc) - timedelta(days=RETENTION_DAYS)
            # пройти по всем user dirs
            for user_dir in STATIC_DIR.iterdir():
                if not user_dir.is_dir():
                    continue
                try:
                    for f in user_dir.iterdir():
                        try:
                            mtime = datetime.fromtimestamp(f.stat().st_mtime, tz=timezone.utc)
                            if mtime < cutoff:
                                logger.info("Removing old file: %s", f)
                                f.unlink(missing_ok=True)
                        except Exception as fe:
                            logger.exception("Error checking/removing file %s: %s", f, fe)
                    # если пустая папка, удалить
                    try:
                        if not any(user_dir.iterdir()):
                            user_dir.rmdir()
                    except Exception:
                        pass
                except Exception as ude:
                    logger.exception("Error scanning user dir %s: %s", user_dir, ude)

            # удалить записи в БД, если файлы пропали и старше cutoff (чистка записей)
            conn = get_conn(); cur = conn.cursor()
            cur.execute("SELECT id, path, created_at FROM presentations")
            rows = cur.fetchall()
            for r in rows:
                try:
                    created = datetime.fromisoformat(r["created_at"])
                except Exception:
                    # если формат нестандартный — попробуем пропустить
                    created = datetime.now(timezone.utc) - timedelta(days=RETENTION_DAYS + 1)
                p = Path(r["path"])
                if created < cutoff and (not p.exists()):
                    cur.execute("DELETE FROM presentations WHERE id = ?", (r["id"],))
            conn.commit(); conn.close()
        except Exception as e:
            logger.exception("Cleanup loop error: %s", e)
        # ждать 24 часа
        await asyncio.sleep(24 * 3600)

# ---------- Регистрация handlers и запуск ----------
def main():
    # Build app (использует build_app из Части 2)
    app = build_app()

    # Регистрация дополнительных handlers (может повторяться, но перезапись безопасна)
    app.add_handler(CommandHandler("privacy", privacy_cmd))
    app.add_handler(CommandHandler("delete_my_data", delete_my_data_cmd))
    app.add_handler(CommandHandler("admin", admin_cmd))
    app.add_handler(CommandHandler("feedback", feedback_start))

    # feedback conversation
    conv_fb = ConversationHandler(
        entry_points=[CommandHandler("feedback", feedback_start)],
        states={1: [MessageHandler(filters.TEXT & ~filters.COMMAND, feedback_receive)]},
        fallbacks=[CommandHandler("cancel", lambda u,c: u.message.reply_text("Отмена."))]
    )
    app.add_handler(conv_fb)

    # Start cleanup background task after app started
    async def start_bg(app_instance):
        app_instance.create_task(cleanup_task(app_instance))
    app.post_init = start_bg

    # Run polling
    logger.info("Starting bot (polling)...")
    app.run_polling()

if __name__ == "__main__":
    main()

