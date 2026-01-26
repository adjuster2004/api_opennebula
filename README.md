# Check all vm OpenNebula - We are collecting complete information on the VM


## 🛠️ Порядок действий
- **Склонировать репозиторий** смотри раздел "Клонирование репозитория"
- **Поменять** в файле config.py свои значения:

```
#Настройки подключения
ENDPOINT = "http://FQDN:2633/RPC2"
USERNAME = "your_username"
TOKEN = "your_token"chrome://extensions/
```

- **Запустить** check_vm_all.py:

```bash
python3 check_vm_all.py
``` 

или

  ```bash
python check_vm_all.py
``` 
- **Открыть** файл vm_list_export.csv


### 👥 Клонирование репозитория

```bash
git clone https://github.com/adjuster2004/api_opennebula/
cd /api_opennebula
```

## 📄 Лицензия
Этот проект распространяется под лицензией **MIT**.

Copyright (c) 2025 Sergey S @adjuster2004

Подробности в файле [LICENSE](LICENSE).
