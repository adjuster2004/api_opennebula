# config.py
# Конфигурация для подключения к OpenNebula

# Настройки подключения
ENDPOINT = "http://FQDN:2633/RPC2"
USERNAME = "your_username"
TOKEN = "your_token"

# Настройки пагинации
BATCH_SIZE = 50  # Количество VM в одном запросе
MAX_DISPLAY = 1000  # Максимальное количество VM для отображения в таблице

# Настройки вывода
VERBOSE = False  # Подробный вывод
EXPORT_CSV = True  # Экспорт в CSV
EXPORT_FILENAME = "vm_list_export.csv""
