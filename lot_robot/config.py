import logging

CONFIG = {
    "BASE_URL": "https://zakupki.gov.ru",
    "RESULTS_PER_PAGE": 20,
    "REQUEST_TIMEOUT": 30,  # Увеличил таймаут
    "REQUEST_DELAY": 1,
    "MAX_RETRIES": 3,
    "USER_AGENT": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

CONFIG["LLM_API_KEY"] = "cCT8j22cdaqVFZGceW5iAigM1ULuIz0S"  # или оставь пустым и используй переменную окружения MISTRAL_API_KEY
CONFIG["LLM_API_URL"] = "https://api.mistral.ai/v1/chat/completions"
CONFIG["LLM_MODEL"] = "mistral-large-latest"
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
