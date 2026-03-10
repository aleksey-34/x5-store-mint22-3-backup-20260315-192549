from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    app_name: str = "X5_Storage API"
    app_env: str = "dev"
    debug: bool = False

    database_url: str = "sqlite:///./x5_storage.db"
    object_root: str = "docflow/objects/x5-ufa-e2_logistics_park"

    telegram_api_id: int | None = None
    telegram_api_hash: str | None = None
    telegram_session_name: str = "x5_storage_session"

    local_llm_enabled: bool = True
    local_llm_base_url: str = "http://127.0.0.1:11434"
    local_llm_model: str = "llama3.2:3b"
    local_llm_timeout_seconds: int = 180
    local_llm_system_prompt: str = (
        "Ты помощник АРМ X5_Storage для стройплощадки. "
        "Отвечай по-русски, кратко и практично. "
        "Сфокусируйся на документообороте объекта, журналах, нарядах-допусках, "
        "ОТ/ПБ, исполнительной документации, сканировании и печати."
    )

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )


settings = Settings()
