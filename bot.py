import logging
from telegram import Update, WebAppInfo, KeyboardButton, ReplyKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
)
import sqlite3
import json
from datetime import datetime

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ВАЖНО: Замените это на ваш URL от GitHub Pages
WEB_APP_URL = "https://ваш-username.github.io/telegram-mini-app/"

# Инициализация базы данных
def init_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            name TEXT,
            phone TEXT,
            email TEXT,
            registration_date TEXT
        )
    ''')
    conn.commit()
    conn.close()
    logger.info("База данных инициализирована")

# Проверка регистрации пользователя
def is_registered(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('SELECT user_id FROM users WHERE user_id = ?', (user_id,))
    result = c.fetchone()
    conn.close()
    return result is not None

# Получение данных пользователя
def get_user_data(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('SELECT * FROM users WHERE user_id = ?', (user_id,))
    result = c.fetchone()
    conn.close()
    return result

# Сохранение пользователя
def save_user(user_id, username, name, phone, email):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    try:
        c.execute('''
            INSERT OR REPLACE INTO users (user_id, username, name, phone, email, registration_date)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (user_id, username, name, phone, email, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        logger.info(f"Пользователь {user_id} сохранён в базу")
        return True
    except Exception as e:
        logger.error(f"Ошибка сохранения: {e}")
        return False
    finally:
        conn.close()

# Команда /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    if is_registered(user.id):
        # Пользователь уже зарегистрирован - НЕ показываем кнопку регистрации
        user_data = get_user_data(user.id)
        
        # Создаём кнопки для зарегистрированных пользователей
        keyboard = [
            [KeyboardButton(text="👤 Мой профиль")],
            [KeyboardButton(text="❓ Помощь")]
        ]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        
        await update.message.reply_text(
            f"👋 С возвращением, {user_data[2]}!\n\n"
            "✅ Вы уже зарегистрированы в системе.\n\n"
            f"📝 Имя: {user_data[2]}\n"
            f"📱 Телефон: {user_data[3]}\n"
            f"📧 Email: {user_data[4]}\n\n"
            "Доступные команды:\n"
            "/profile - Посмотреть полный профиль\n"
            "/help - Справка",
            reply_markup=reply_markup
        )
        logger.info(f"Зарегистрированный пользователь {user.id} ({user_data[2]}) вернулся")
    else:
        # Новый пользователь - показываем кнопку регистрации
        keyboard = [
            [KeyboardButton(
                text="📝 Зарегистрироваться",
                web_app=WebAppInfo(url=WEB_APP_URL)
            )]
        ]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        
        await update.message.reply_text(
            f"👋 Привет, {user.first_name}!\n\n"
            "Добро пожаловать в наш бот! 🎉\n\n"
            "⚠️ Для использования бота необходимо пройти регистрацию.\n\n"
            "Нажмите кнопку ниже, чтобы заполнить форму 👇",
            reply_markup=reply_markup
        )
        logger.info(f"Новый пользователь {user.id} ({user.first_name}) начал регистрацию")

# Обработка данных из Web App
async def web_app_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    try:
        # Получаем данные из Web App
        data = json.loads(update.effective_message.web_app_data.data)
        
        if data.get('action') == 'registration':
            user_data = data.get('data', {})
            
            # Сохраняем в базу
            success = save_user(
                user.id,
                user.username or "N/A",
                user_data.get('name'),
                user_data.get('phone'),
                user_data.get('email')
            )
            
            if success:
                await update.message.reply_text(
                    "✅ Регистрация успешно завершена!\n\n"
                    f"📝 Ваши данные:\n"
                    f"• Имя: {user_data.get('name')}\n"
                    f"• Телефон: {user_data.get('phone')}\n"
                    f"• Email: {user_data.get('email')}\n\n"
                    "Используйте /profile для просмотра профиля"
                )
                logger.info(f"Регистрация завершена для пользователя {user.id}")
            else:
                await update.message.reply_text(
                    "❌ Произошла ошибка при сохранении данных.\n"
                    "Попробуйте ещё раз."
                )
    except json.JSONDecodeError:
        logger.error("Ошибка декодирования JSON")
        await update.message.reply_text(
            "❌ Ошибка обработки данных. Попробуйте снова."
        )
    except Exception as e:
        logger.error(f"Ошибка обработки Web App данных: {e}")
        await update.message.reply_text(
            "❌ Произошла ошибка. Попробуйте позже."
        )

# Команда /profile
async def profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    
    if not is_registered(user.id):
        keyboard = [
            [KeyboardButton(
                text="📝 Пройти регистрацию",
                web_app=WebAppInfo(url=WEB_APP_URL)
            )]
        ]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        
        await update.message.reply_text(
            "⚠️ Вы ещё не зарегистрированы!\n\n"
            "Нажмите кнопку ниже для регистрации 👇",
            reply_markup=reply_markup
        )
        return
    
    user_data = get_user_data(user.id)
    
    await update.message.reply_text(
        "👤 Ваш профиль:\n\n"
        f"📝 Имя: {user_data[2]}\n"
        f"📱 Телефон: {user_data[3]}\n"
        f"📧 Email: {user_data[4]}\n"
        f"📅 Дата регистрации: {user_data[5]}"
    )

# Команда /help
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 Доступные команды:\n\n"
        "/start - Начать работу с ботом\n"
        "/profile - Посмотреть свой профиль\n"
        "/stats - Статистика (только для админов)\n"
        "/help - Показать это сообщение\n\n"
        "💡 Нажмите кнопку 'Открыть регистрацию' "
        "для доступа к приложению"
    )

# Команда /stats (для администратора)
async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Здесь добавьте ID вашего Telegram аккаунта
    ADMIN_ID = 7774588164  # Замените на ваш ID
    
    user = update.effective_user
    
    if user.id != ADMIN_ID:
        await update.message.reply_text("❌ У вас нет доступа к этой команде.")
        return
    
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('SELECT COUNT(*) FROM users')
    total_users = c.fetchone()[0]
    conn.close()
    
    await update.message.reply_text(
        f"📊 Статистика бота:\n\n"
        f"👥 Всего пользователей: {total_users}"
    )

# Обработка обычных сообщений (включая кнопки)
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user = update.effective_user
    
    # Обработка кнопок
    if text == "👤 Мой профиль":
        await profile(update, context)
    elif text == "❓ Помощь":
        await help_command(update, context)
    else:
        # Проверка регистрации для всех остальных сообщений
        if not is_registered(user.id):
            keyboard = [
                [KeyboardButton(
                    text="📝 Зарегистрироваться",
                    web_app=WebAppInfo(url=WEB_APP_URL)
                )]
            ]
            reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            
            await update.message.reply_text(
                "⚠️ Для использования бота необходимо зарегистрироваться!\n\n"
                "Нажмите кнопку ниже 👇",
                reply_markup=reply_markup
            )
        else:
            await update.message.reply_text(
                "Я вас не понял. Используйте /help для просмотра доступных команд."
            )

def main():
    # Инициализация базы данных
    init_db()
    
    # ВАЖНО: Замените на токен вашего бота от @BotFather
    TOKEN = '8046331797:AAHq48Wbyu3ihFkoM8HiFflbnpbN0-couGU'
    
    if TOKEN == 'YOUR_BOT_TOKEN_HERE':
        print("❌ ОШИБКА: Укажите токен бота!")
        print("Получите токен у @BotFather и замените YOUR_BOT_TOKEN_HERE в коде")
        return
    
    # Создание приложения
    application = Application.builder().token(TOKEN).build()
    
    # Регистрация обработчиков
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("profile", profile))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("stats", stats))
    application.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, web_app_data))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    # Запуск бота
    print("=" * 60)
    print("🤖 Бот успешно запущен!")
    print("=" * 60)
    print(f"📱 Web App URL: {WEB_APP_URL}")
    print("✅ Бот готов к работе...")
    print("=" * 60)
    
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
