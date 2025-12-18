# Конфигурация для mektep_scraper.py

# URL платформы
BASE_URL = "https://mektep.edu.kz"
LOGIN_URL = f"{BASE_URL}/_monitor/index.php"
REPORTS_URL = f"{BASE_URL}/_monitor/pg_reports.php"

# Учетные данные (используйте переменные окружения для безопасности)
import os
from dotenv import load_dotenv

# Загружаем переменные окружения из .env файла (если есть)
load_dotenv()

LOGIN = os.getenv("EDUS_LOGIN", "")  # Логин из переменной окружения EDUS_LOGIN
PASSWORD = os.getenv("EDUS_PASSWORD", "")  # Пароль из переменной окружения EDUS_PASSWORD

# Настройки браузера
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"  # Headless режим из переменной окружения
BROWSER_TIMEOUT = int(os.getenv("BROWSER_TIMEOUT", "60"))  # Таймаут ожидания элементов (секунды)
IMPLICIT_WAIT = int(os.getenv("IMPLICIT_WAIT", "10"))  # Неявное ожидание (секунды)
PAGE_LOAD_TIMEOUT = int(os.getenv("PAGE_LOAD_TIMEOUT", "30"))  # Таймаут загрузки страницы (секунды)

# Настройки для извлечения данных
MIN_TABLE_ROWS = int(os.getenv("MIN_TABLE_ROWS", "60"))  # Минимальное количество строк в таблице школ
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "success_data.xlsx")  # Имя выходного Excel файла

