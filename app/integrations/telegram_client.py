from app.core.config import settings


def telegram_listener_configured() -> bool:
    return bool(settings.telegram_api_id and settings.telegram_api_hash)


def start_listener() -> str:
    if not telegram_listener_configured():
        return "Telegram listener is not configured. Set TELEGRAM_API_ID and TELEGRAM_API_HASH."

    return "Telegram listener placeholder started. Connect Telethon event handlers next."
