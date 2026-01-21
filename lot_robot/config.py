import logging

CONFIG = {
    "BASE_URL": "https://zakupki.gov.ru",
    "RESULTS_PER_PAGE": 20,
    "REQUEST_TIMEOUT": 30,  # Увеличил таймаут
    "REQUEST_DELAY": 1,
    "MAX_RETRIES": 3,
    "USER_AGENT": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

CONFIG["LLM_PROVIDER"] = "vsellm_yandex"
CONFIG["LLM_PROVIDER_KEYS"] = {
    "cloudru": "OWMyYjA4ZjItNGY2Ni00OTNjLWJlMmUtN2Y5YTI1MjYwYWNi.9b5fe2f4b7aa0c1358219ee59f9b2b25",
    "vsellm_yandex": "sk-FazTRmiqKGjpZwRob5lvRw",
    "vsellm_deepseek": "sk-FazTRmiqKGjpZwRob5lvRw",
    "mistral": "",
}
CONFIG["LLM_PROVIDERS"] = {
    "cloudru": {
        "api_url": "https://foundation-models.api.cloud.ru/v1/chat/completions",
        "model": "Qwen/Qwen3-235B-A22B-Instruct-2507",
        "env_key": "API_KEY",
    },
    "vsellm_yandex": {
        "api_url": "https://api.vsellm.ru/v1/chat/completions",
        "model": "yandex/gpt5-lite",
        "env_key": "VSELLM_API_KEY",
    },
    "vsellm_deepseek": {
        "api_url": "https://api.vsellm.ru/v1/chat/completions",
        "model": "deepseek/deepseek-chat-v3-0324",
        "env_key": "VSELLM_API_KEY",
    },
    "mistral": {
        "api_url": "https://api.mistral.ai/v1/chat/completions",
        "model": "mistral-large-latest",
        "env_key": "MISTRAL_API_KEY",
    },
}
CONFIG["LLM_API_KEY"] = ""  # optional override for active provider
CONFIG["LLM_API_URL"] = ""  # optional override for active provider
CONFIG["LLM_MODEL"] = ""  # optional override for active provider
CONFIG["LLM_REQUEST_TIMEOUT"] = 90

# Search filter options
PURCHASE_STAGES = {
    "": "Все этапы",
    "SUBMISSION": "Подача заявок",
    "EVALUATION": "Работа комиссии",
    "COMPLETED": "Закупка завершена",
    "CANCELLED": "Закупка отменена",
}

LAWS = {"": "Все законы", "44FZ": "44-ФЗ", "223FZ": "223-ФЗ", "PP615": "ПП РФ 615"}

# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)
