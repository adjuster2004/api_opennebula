#!/usr/bin/env python3
"""
Рабочий скрипт для OpenNebula API с полной информацией о VM
Включая MAC адреса и метки (LABELS) с правильной пагинацией
"""

from pyone import OneServer
import warnings
from collections import OrderedDict
import time
from datetime import datetime
import csv
import os

# Импортируем конфигурацию
try:
    from config import *
except ImportError:
    print("❌ Ошибка: не найден файл config.py")
    print("Создайте файл config.py с настройками подключения")
    exit(1)

# Отключаем warnings о SSL
warnings.filterwarnings('ignore')

# Функции для поиска CPU и vCPU
def get_value_from_template(vm, key, default=None):
    """Получение значения из шаблона VM с поиском в разных местах"""
    if not hasattr(vm, 'TEMPLATE'):
        return default

    template = vm.TEMPLATE

    # Если шаблон - словарь, ищем в нем (регистронезависимо)
    if isinstance(template, (dict, OrderedDict)):
        # Ищем напрямую (точное совпадение)
        if key in template:
            value = template[key]
            if value not in [None, '', ' ']:
                return str(value)

        # Ищем регистронезависимо
        for template_key in template.keys():
            if str(template_key).upper() == key.upper():
                value = template[template_key]
                if value not in [None, '', ' ']:
                    return str(value)

        # Ищем в USER_TEMPLATE внутри TEMPLATE
        if 'USER_TEMPLATE' in template:
            user_template = template['USER_TEMPLATE']
            if isinstance(user_template, (dict, OrderedDict)):
                # Точное совпадение
                if key in user_template:
                    value = user_template[key]
                    if value not in [None, '', ' ']:
                        return str(value)

                # Регистронезависимый поиск
                for user_key in user_template.keys():
                    if str(user_key).upper() == key.upper():
                        value = user_template[user_key]
                        if value not in [None, '', ' ']:
                            return str(value)

    # Проверяем USER_TEMPLATE как атрибут VM
    if hasattr(vm, 'USER_TEMPLATE'):
        user_template = vm.USER_TEMPLATE
        if isinstance(user_template, (dict, OrderedDict)):
            # Точное совпадение
            if key in user_template:
                value = user_template[key]
                if value not in [None, '', ' ']:
                    return str(value)

            # Регистронезависимый поиск
            for user_key in user_template.keys():
                if str(user_key).upper() == key.upper():
                    value = user_template[user_key]
                    if value not in [None, '', ' ']:
                        return str(value)

    return default

def get_cpu_vcpu_from_vm(vm):
    """Получение CPU и vCPU из VM"""
    cpu = 0
    vcpu = 0

    if not hasattr(vm, 'TEMPLATE'):
        return cpu, vcpu

    template = vm.TEMPLATE

    # Ищем CPU (физические процессоры)
    cpu_value = get_value_from_template(vm, 'CPU')
    if cpu_value:
        try:
            cpu = float(cpu_value)
        except (ValueError, TypeError):
            cpu = 0

    # Ищем VCPU (виртуальные процессоры)
    vcpu_value = None

    # 1. Сначала проверяем напрямую в TEMPLATE
    if isinstance(template, (dict, OrderedDict)):
        # Проверяем разные варианты написания
        vcpu_keys = ['VCPU', 'vcpu', 'Vcpu']
        for key in vcpu_keys:
            if key in template:
                vcpu_value = template[key]
                break

        # Если не нашли стандартными ключами, проверяем все ключи
        if not vcpu_value:
            for key in template.keys():
                if str(key).strip().upper() == 'VCPU':
                    vcpu_value = template[key]
                    break

    # 2. Если не нашли, используем функцию get_value_from_template
    if not vcpu_value:
        vcpu_value = get_value_from_template(vm, 'VCPU')

    # 3. Если все еще не нашли, проверяем 'vcpu' (маленькими буквами)
    if not vcpu_value:
        vcpu_value = get_value_from_template(vm, 'vcpu')

    if vcpu_value:
        try:
            vcpu = float(vcpu_value)
        except (ValueError, TypeError):
            vcpu = 0

    # Если VCPU не найден, но есть CPU, используем CPU как VCPU
    if vcpu == 0 and cpu > 0:
        vcpu = cpu

    return cpu, vcpu

def test_with_pyone():
    """Тест через библиотеку pyone"""

    print("="*60)
    print("ТЕСТ ЧЕРЕЗ БИБЛИОТЕКУ PYONE")
    print("="*60)

    try:
        # Ключевая строка: session="username:token"
        one = OneServer(
            ENDPOINT,
            session=f"{USERNAME}:{TOKEN}",
            https_verify=False
        )

        # Тест 1: Получить версию
        version = one.system.version()
        print(f"✅ OpenNebula версия: {version}")

        # Тест 2: Получить информацию о пользователе
        user_info = one.user.info(-1)  # -1 = текущий пользователь
        print(f"✅ Текущий пользователь: {user_info.NAME}")
        print(f"✅ ID пользователя: {user_info.ID}")

        return one

    except Exception as e:
        print(f"❌ Ошибка pyone: {e}")
        return None

def get_labels_from_vm(vm):
    """Получение меток (LABELS) из объекта VM"""
    labels = {}

    if not vm:
        return labels

    # Метод 1: Проверяем атрибут LABELS у VM
    if hasattr(vm, 'LABELS'):
        labels_value = vm.LABELS
        if labels_value and labels_value not in [None, '', ' ']:
            labels.update(parse_labels_string(labels_value))

    # Метод 2: Проверяем USER_TEMPLATE
    if hasattr(vm, 'USER_TEMPLATE'):
        user_template = vm.USER_TEMPLATE
        if isinstance(user_template, (dict, OrderedDict)) and 'LABELS' in user_template:
            labels_value = user_template['LABELS']
            if labels_value and labels_value not in [None, '', ' ']:
                labels.update(parse_labels_string(labels_value))

    # Метод 3: Проверяем TEMPLATE
    if hasattr(vm, 'TEMPLATE'):
        template = vm.TEMPLATE
        if isinstance(template, (dict, OrderedDict)):
            # Проверяем LABELS в основном шаблоне
            if 'LABELS' in template:
                labels_value = template['LABELS']
                if labels_value and labels_value not in [None, '', ' ']:
                    labels.update(parse_labels_string(labels_value))

            # Проверяем USER_TEMPLATE внутри TEMPLATE
            if 'USER_TEMPLATE' in template:
                user_template = template['USER_TEMPLATE']
                if isinstance(user_template, (dict, OrderedDict)) and 'LABELS' in user_template:
                    labels_value = user_template['LABELS']
                    if labels_value and labels_value not in [None, '', ' ']:
                        labels.update(parse_labels_string(labels_value))

    return labels

def parse_labels_string(labels_value):
    """Парсинг строки с метками"""
    labels = {}

    if not labels_value:
        return labels

    # Если это строка
    if isinstance(labels_value, str):
        labels_str = labels_value.strip()

        # Формат 1: Простая строка "vtnrf"
        if '=' not in labels_str and ',' not in labels_str:
            labels[labels_str] = "true"

        # Формат 2: Строка с разделителями "ключ=значение,ключ2=значение2"
        elif '=' in labels_str:
            pairs = labels_str.split(',')
            for pair in pairs:
                pair = pair.strip()
                if '=' in pair:
                    key, value = pair.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    # Игнорируем ключи типа LABEL_1, LABEL_2 и т.д.
                    if not key.startswith('LABEL_'):
                        labels[key] = value
                elif pair and not pair.startswith('LABEL_'):  # Просто значение без ключа
                    labels[pair] = "true"

        # Формат 3: Просто значения через запятую "val1,val2,val3"
        else:
            values = labels_str.split(',')
            for value in values:
                value = value.strip()
                if value and not value.startswith('LABEL_'):
                    labels[value] = "true"

    # Если это словарь
    elif isinstance(labels_value, (dict, OrderedDict)):
        for key, value in labels_value.items():
            if value and value not in [None, '', ' ']:
                # Игнорируем ключи типа LABEL_1, LABEL_2 и т.д.
                if not str(key).startswith('LABEL_'):
                    labels[key] = str(value)

    # Если это список
    elif isinstance(labels_value, list):
        for value in labels_value:
            if value and value not in [None, '', ' ']:
                value_str = str(value)
                if not value_str.startswith('LABEL_'):
                    labels[value_str] = "true"

    return labels

def get_disk_info(disk):
    """Получение информации о диске"""
    disk_info = {}

    if isinstance(disk, (dict, OrderedDict)):
        # Размер диска
        if 'SIZE' in disk:
            try:
                size_mb = int(disk['SIZE'])
                disk_info['size_mb'] = size_mb
                disk_info['size_gb'] = size_mb / 1024
            except (ValueError, TypeError):
                disk_info['size_mb'] = 0
                disk_info['size_gb'] = 0

        # Образ
        if 'IMAGE' in disk:
            disk_info['image'] = disk['IMAGE']

        # Тип диска
        if 'TYPE' in disk:
            disk_info['type'] = disk['TYPE']

        # Формат
        if 'FORMAT' in disk:
            disk_info['format'] = disk['FORMAT']

        # Целевое устройство
        if 'TARGET' in disk:
            disk_info['target'] = disk['TARGET']

    return disk_info

def get_nic_info(nic):
    """Получение информации о сетевом интерфейсе"""
    nic_info = {}

    if isinstance(nic, (dict, OrderedDict)):
        # IP адрес
        if 'IP' in nic:
            nic_info['ip'] = nic['IP']

        # MAC адрес
        if 'MAC' in nic:
            nic_info['mac'] = nic['MAC']

        # Сеть
        if 'NETWORK' in nic:
            nic_info['network'] = nic['NETWORK']

        # VNET ID
        if 'NETWORK_ID' in nic:
            nic_info['network_id'] = nic['NETWORK_ID']

        # Имя интерфейса
        if 'NIC_ID' in nic:
            nic_info['nic_id'] = nic['NIC_ID']

        # VLAN ID
        if 'VLAN_ID' in nic:
            nic_info['vlan_id'] = nic['VLAN_ID']

        # Security Groups
        if 'SECURITY_GROUPS' in nic:
            nic_info['security_groups'] = nic['SECURITY_GROUPS']

    return nic_info

def get_all_vms_simple(one_connection):
    """Простой метод получения всех VM (теперь с полным доступом)"""
    print(f"\n📋 Получение ВСЕХ виртуальных машин...")

    all_vms = []
    start_time = time.time()

    try:
        # Используем метод, который работал в диагностике: state=-2, limit=-1
        print("  Используем метод: vmpool.info(-2, -1, -1, -1)")
        vmpool = one_connection.vmpool.info(-2, -1, -1, -1)

        if hasattr(vmpool, 'VM'):
            vms = vmpool.VM
            if not isinstance(vms, list):
                vms = [vms]

            all_vms.extend(vms)

            elapsed_time = time.time() - start_time
            print(f"  ✅ Получено {len(all_vms)} VM за {elapsed_time:.2f} секунд")

            # Сортируем по ID для удобства
            all_vms.sort(key=lambda vm: int(vm.ID))

            # Выводим информацию о диапазоне
            if all_vms:
                min_id = int(all_vms[0].ID)
                max_id = int(all_vms[-1].ID)
                print(f"  📊 Диапазон ID VM: {min_id} - {max_id}")

                # Проверяем пропуски
                ids = [int(vm.ID) for vm in all_vms]
                expected_ids = set(range(min_id, max_id + 1))
                missing_ids = expected_ids - set(ids)

                if missing_ids:
                    print(f"  ⚠️  Пропущенные ID: {len(missing_ids)} штук")
                    if len(missing_ids) <= 10:
                        print(f"     Пропущенные ID: {sorted(list(missing_ids))[:10]}")

            return all_vms
        else:
            print("  ❌ Не удалось получить VM")
            return []

    except Exception as e:
        print(f"  ❌ Ошибка при получении VM: {e}")
        return []

def get_vm_resources(vm, one_connection=None):
    """Получение ресурсов VM из шаблона"""
    cpu = 0
    vcpu = 0

    # ВАЖНО: Всегда получаем свежие данные VM через one_connection
    if one_connection:
        try:
            vm_fresh = one_connection.vm.info(vm.ID)
            vm = vm_fresh
        except Exception as e:
            print(f"❌ Ошибка получения полной информации о VM {vm.ID}: {e}")
            # Продолжаем с имеющимися данными

    if not hasattr(vm, 'TEMPLATE'):
        return {
            'cpu': cpu,
            'vcpu': vcpu,
            'memory': 0,
            'memory_gb': 0,
            'total_disk_gb': 0,
            'nics': [],
            'labels': {},
            'labels_count': 0
        }

    template = vm.TEMPLATE

    # Ищем CPU и VCPU напрямую из TEMPLATE
    if isinstance(template, (dict, OrderedDict)):
        # CPU
        if 'CPU' in template:
            try:
                cpu = float(template['CPU'])
            except (ValueError, TypeError):
                cpu = 0

        # VCPU
        if 'VCPU' in template:
            try:
                vcpu = float(template['VCPU'])
            except (ValueError, TypeError):
                vcpu = 0
        else:
            # Пробуем найти ключ с VCPU в разных регистрах
            for key in template.keys():
                if str(key).strip().upper() == 'VCPU':
                    try:
                        vcpu = float(template[key])
                        break
                    except (ValueError, TypeError):
                        pass

    # Если VCPU не найден, но есть CPU, используем CPU как VCPU
    if vcpu == 0 and cpu > 0:
        vcpu = cpu

    # Если не нашли в TEMPLATE, пробуем другие методы
    if vcpu == 0:
        cpu, vcpu = get_cpu_vcpu_from_vm(vm)

    # Остальной код остается без изменений...
    memory = 0
    memory_gb = 0
    total_disk_gb = 0
    nics = []
    labels = {}

    # Память
    memory_value = get_value_from_template(vm, 'MEMORY', '0')
    if memory_value not in [None, '', ' ']:
        try:
            memory = int(memory_value)
            memory_gb = memory / 1024
        except (ValueError, TypeError):
            memory = 0
            memory_gb = 0

    # Диски
    if isinstance(template, (dict, OrderedDict)) and 'DISK' in template:
        disks = template['DISK']
        if not isinstance(disks, list):
            disks = [disks]

        for disk in disks:
            disk_info = get_disk_info(disk)
            if 'size_gb' in disk_info:
                total_disk_gb += disk_info['size_gb']

    # Сетевые интерфейсы
    if isinstance(template, (dict, OrderedDict)) and 'NIC' in template:
        nic_list = template['NIC']
        if not isinstance(nic_list, list):
            nic_list = [nic_list]

        for nic in nic_list:
            nic_info = get_nic_info(nic)
            nics.append(nic_info)

    # Метки
    labels = get_labels_from_vm(vm)

    return {
        'cpu': cpu,
        'vcpu': vcpu,
        'memory': memory,
        'memory_gb': memory_gb,
        'total_disk_gb': total_disk_gb,
        'nics': nics,
        'labels': labels,
        'labels_count': len(labels)
    }

def collect_vm_data_for_display_and_export(vms, one_connection, display_limit=None):
    """Сбор данных VM для отображения в консоли и экспорта в CSV"""

    print(f"\n📊 Сбор данных для {len(vms)} VM...")
    print(f"⚠️  ВНИМАНИЕ: Запрашивается полная информация для каждой VM...")
    print(f"   Это может занять некоторое время.")

    start_time = time.time()

    # Данные для отображения в консоли
    display_data = []
    # Данные для экспорта в CSV
    export_data = []
    # Все метки для CSV
    all_labels = set()
    # Максимальное количество сетевых интерфейсов
    max_nics = 0
    # Статистика для быстрого расчета
    total_cpu = 0
    total_vcpu = 0
    total_memory_gb = 0
    total_disk_gb = 0
    total_ips = 0
    total_macs = 0
    total_labels = 0
    vms_with_labels = 0
    state_counts = {}
    owner_counts = {}

    # Прогресс-бар для больших наборов данных
    if len(vms) > 50:
        print("   Прогресс: ", end='', flush=True)

    for i, vm in enumerate(vms):
        # Получаем полные ресурсы с использованием one_connection
        resources = get_vm_resources(vm, one_connection)

        # Собираем статистику
        total_cpu += resources['cpu']
        total_vcpu += resources['vcpu']
        total_memory_gb += resources['memory_gb']
        total_disk_gb += resources['total_disk_gb']

        # Считаем IP и MAC адреса
        for nic in resources['nics']:
            if 'ip' in nic and nic['ip']:
                total_ips += 1
            if 'mac' in nic and nic['mac']:
                total_macs += 1

        total_labels += resources['labels_count']

        if resources['labels_count'] > 0:
            vms_with_labels += 1

        # Статистика по состояниям
        state_code = int(vm.STATE)
        states = {
            0: 'INIT', 1: 'PENDING', 2: 'HOLD',
            3: 'ACTIVE', 4: 'STOPPED', 5: 'SUSPENDED',
            6: 'DONE', 8: 'POWEROFF', 9: 'UNDEPLOYED'
        }
        state = states.get(state_code, f'STATE_{state_code}')
        state_counts[state] = state_counts.get(state, 0) + 1

        # Статистика по владельцам
        owner = vm.UNAME
        owner_counts[owner] = owner_counts.get(owner, 0) + 1

        # Данные для отображения в консоли
        display_item = {
            'vm': vm,
            'resources': resources,
            'state': state
        }
        display_data.append(display_item)

        # Данные для экспорта
        creation_date = ""
        if hasattr(vm, 'STIME'):
            try:
                creation_date = datetime.fromtimestamp(int(vm.STIME)).strftime('%Y-%m-%d %H:%M:%S')
            except:
                creation_date = str(vm.STIME)

        modification_date = ""
        if hasattr(vm, 'ETIME') and vm.ETIME:
            try:
                modification_date = datetime.fromtimestamp(int(vm.ETIME)).strftime('%Y-%m-%d %H:%M:%S')
            except:
                modification_date = str(vm.ETIME)

        export_item = {
            'vm': vm,
            'resources': resources,
            'state': state,
            'creation_date': creation_date,
            'modification_date': modification_date
        }
        export_data.append(export_item)

        # Собираем метки для CSV
        for label_key in resources['labels'].keys():
            all_labels.add(label_key)

        # Определяем максимальное количество сетевых интерфейсов
        if len(resources['nics']) > max_nics:
            max_nics = len(resources['nics'])

        # Вывод прогресса для больших наборов
        if len(vms) > 50 and i % (len(vms) // 20) == 0:
            print(f"▮", end='', flush=True)

        # Небольшая задержка чтобы не перегружать API
        if i < len(vms) - 1:
            time.sleep(0.01)  # Уменьшена задержка для скорости

    if len(vms) > 50:
        print()

    elapsed_time = time.time() - start_time
    print(f"⏱️  Время сбора данных: {elapsed_time:.2f} секунд")

    # Сортируем метки для стабильного порядка столбцов
    sorted_labels = sorted(list(all_labels))

    # Пакетная статистика
    statistics = {
        'total_vms': len(vms),
        'total_cpu': total_cpu,
        'total_vcpu': total_vcpu,
        'total_memory_gb': total_memory_gb,
        'total_disk_gb': total_disk_gb,
        'total_ips': total_ips,
        'total_macs': total_macs,
        'total_labels': total_labels,
        'vms_with_labels': vms_with_labels,
        'state_counts': state_counts,
        'owner_counts': owner_counts,
        'max_nics': max_nics,
        'sorted_labels': sorted_labels
    }

    return display_data, export_data, statistics

def display_vm_table(display_data, display_limit=None):
    """Отображение таблицы VM в консоли"""

    if not display_data:
        print("❌ Нет данных для отображения")
        return

    if display_limit is None:
        display_limit = min(MAX_DISPLAY, len(display_data))

    print(f"\n📋 ОТОБРАЖЕНИЕ ПЕРВЫХ {display_limit} VM (из {len(display_data)}):")
    print(f"{'ID':<8} {'Имя':<20} {'Состояние':<10} {'Владелец':<12} {'CPU':<5} {'vCPU':<5} {'Память':<8} {'Диск':<8} {'IP':<15} {'MAC':<17} {'Метки'}")
    print("-"*150)

    for i in range(min(display_limit, len(display_data))):
        item = display_data[i]
        vm = item['vm']
        resources = item['resources']
        state = item['state']

        # Формируем строки с IP и MAC адресами
        if resources['nics']:
            ip_str = resources['nics'][0].get('ip', '-')
            mac_str = resources['nics'][0].get('mac', '-')

            if len(resources['nics']) > 1:
                ip_str += f" (+{len(resources['nics']) - 1})"
                mac_str += f" (+{len(resources['nics']) - 1})"
        else:
            ip_str = "-"
            mac_str = "-"

        # Формируем строку с метки
        if resources['labels']:
            # Берем первые 2 метки
            label_keys = list(resources['labels'].keys())[:2]
            label_str = ", ".join([f"{k}:{resources['labels'][k]}" for k in label_keys])
            if len(resources['labels']) > 2:
                label_str += f" (+{len(resources['labels']) - 2})"
        else:
            label_str = "-"

        print(f"{vm.ID:<8} {vm.NAME[:19]:<20} {state:<10} {vm.UNAME[:11]:<12} "
              f"{resources['cpu']:<5.1f} {resources['vcpu']:<5} {resources['memory_gb']:<6.1f}GB "
              f"{resources['total_disk_gb']:<7.1f}GB {ip_str:<15} {mac_str:<17} {label_str}")

    if len(display_data) > display_limit:
        print(f"\n... и еще {len(display_data) - display_limit} VM (всего {len(display_data)})")

def display_statistics(statistics):
    """Отображение статистики"""

    total_vms = statistics['total_vms']

    if total_vms == 0:
        print("❌ Нет данных для статистики")
        return

    print(f"\n📈 ОБЩАЯ СТАТИСТИКА:")
    print(f"   • Всего VM: {total_vms}")
    print(f"   • Всего физических CPU: {statistics['total_cpu']:.1f}")
    print(f"   • Всего vCPU: {statistics['total_vcpu']:.1f}")
    print(f"   • Всего памяти: {statistics['total_memory_gb']:.1f} GB")
    print(f"   • Всего дискового пространства: {statistics['total_disk_gb']:.1f} GB")
    print(f"   • Всего IP адресов: {statistics['total_ips']}")
    print(f"   • Всего MAC адресов: {statistics['total_macs']}")
    print(f"   • Всего меток: {statistics['total_labels']}")
    print(f"   • VM с метками: {statistics['vms_with_labels']} ({statistics['vms_with_labels']/total_vms*100:.1f}%)")

    print(f"\n📊 РАСПРЕДЕЛЕНИЕ ПО СОСТОЯНИЯМ:")
    for state, count in sorted(statistics['state_counts'].items()):
        percentage = (count / total_vms) * 100
        print(f"   • {state:<12}: {count:>4} VM ({percentage:.1f}%)")

    print(f"\n👥 РАСПРЕДЕЛЕНИЕ ПО ВЛАДЕЛЬЦАМ (ТОП-10):")
    sorted_owners = sorted(statistics['owner_counts'].items(), key=lambda x: x[1], reverse=True)[:10]
    for owner, count in sorted_owners:
        percentage = (count / total_vms) * 100
        print(f"   • {owner:<20}: {count:>4} VM ({percentage:.1f}%)")

    print(f"\n💾 СРЕДНИЕ ЗНАЧЕНИЯ НА ОДНУ VM:")
    if total_vms > 0:
        print(f"   • Среднее физических CPU: {statistics['total_cpu']/total_vms:.1f}")
        print(f"   • Среднее vCPU: {statistics['total_vcpu']/total_vms:.1f}")
        print(f"   • Средняя память: {statistics['total_memory_gb']/total_vms:.1f} GB")
        print(f"   • Средний диск: {statistics['total_disk_gb']/total_vms:.1f} GB")
        print(f"   • Среднее IP адресов: {statistics['total_ips']/total_vms:.1f}")
        print(f"   • Среднее MAC адресов: {statistics['total_macs']/total_vms:.1f}")
        print(f"   • Среднее меток: {statistics['total_labels']/total_vms:.1f}")

def export_to_csv(export_data, statistics, filename=None):
    """Экспорт данных VM в CSV файл"""

    # Если имя файла не указано, используем значение из конфига или генерируем автоматически
    if filename is None:
        if 'EXPORT_FILENAME' in globals():
            filename = EXPORT_FILENAME
        else:
            # Генерируем имя файла с датой
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"vm_list_{timestamp}.csv"

    print(f"\n💾 Экспорт данных в CSV файл: {filename}")
    print(f"   Записей: {len(export_data)}")

    start_time = time.time()

    try:
        max_nics = statistics['max_nics']
        sorted_labels = statistics['sorted_labels']

        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            # Формируем заголовки столбцов
            fieldnames = [
                'ID', 'Имя', 'Состояние', 'Владелец',
                'Физические_CPU', 'vCPU', 'Память_MB', 'Память_GB', 'Диск_GB'
            ]

            # Добавляем столбцы для сетевых интерфейсов
            for i in range(max_nics):
                fieldnames.extend([
                    f'IP_{i+1}',
                    f'MAC_{i+1}',
                    f'Сеть_{i+1}',
                    f'VLAN_{i+1}'
                ])

            # Добавляем столбцы для меток
            fieldnames.extend(sorted_labels)

            # Добавляем даты
            fieldnames.extend([
                'Дата_создания', 'Дата_изменения'
            ])

            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            print("   💾 Запись данных в файл...")

            for i, item in enumerate(export_data):
                vm = item['vm']
                resources = item['resources']
                state = item['state']
                creation_date = item['creation_date']
                modification_date = item['modification_date']

                # Создаем базовую строку данных
                row_data = {
                    'ID': vm.ID,
                    'Имя': vm.NAME,
                    'Состояние': state,
                    'Владелец': vm.UNAME,
                    'Физические_CPU': resources['cpu'],
                    'vCPU': resources['vcpu'],
                    'Память_MB': resources['memory'],
                    'Память_GB': f"{resources['memory_gb']:.1f}",
                    'Диск_GB': f"{resources['total_disk_gb']:.1f}"
                }

                # Добавляем данные сетевых интерфейсов
                for j, nic in enumerate(resources['nics']):
                    row_data[f'IP_{j+1}'] = nic.get('ip', '')
                    row_data[f'MAC_{j+1}'] = nic.get('mac', '')
                    row_data[f'Сеть_{j+1}'] = nic.get('network', '')
                    row_data[f'VLAN_{j+1}'] = nic.get('vlan_id', '')

                # Заполняем оставшиеся сетевые интерфейсы пустыми значениями
                for j in range(len(resources['nics']), max_nics):
                    row_data[f'IP_{j+1}'] = ''
                    row_data[f'MAC_{j+1}'] = ''
                    row_data[f'Сеть_{j+1}'] = ''
                    row_data[f'VLAN_{j+1}'] = ''

                # Добавляем метки
                for label_key in sorted_labels:
                    row_data[label_key] = resources['labels'].get(label_key, '')

                # Добавляем даты
                row_data['Дата_создания'] = creation_date
                row_data['Дата_изменения'] = modification_date

                writer.writerow(row_data)

                if (i + 1) % 100 == 0 or (i + 1) == len(export_data):
                    print(f"      Записано: {i + 1}/{len(export_data)} записей")

        elapsed_time = time.time() - start_time
        print(f"✅ Данные успешно экспортированы в {filename}")
        print(f"⏱️  Время экспорта: {elapsed_time:.2f} секунд")

        # Показываем первые 5 строк файла для проверки правильности VCPU
        print(f"\n🔍 ПРОВЕРКА VCPU В ЭКСПОРТИРОВАННОМ ФАЙЛЕ:")
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                lines = f.readlines()[:6]  # Заголовок + 5 строк данных
                print(f"   Заголовок: {lines[0].strip()}")
                for j in range(1, min(6, len(lines))):
                    parts = lines[j].strip().split(',')
                    if len(parts) > 5:
                        vm_id = parts[0]
                        vm_name = parts[1]
                        cpu = parts[4]
                        vcpu = parts[5]
                        print(f"   VM {vm_id} ({vm_name}): CPU={cpu}, vCPU={vcpu}")
        except Exception as e:
            print(f"   Не удалось прочитать файл для проверки: {e}")

        return filename

    except Exception as e:
        print(f"❌ Ошибка при экспорте в CSV: {e}")
        import traceback
        traceback.print_exc()
        return None

def main():
    print("="*80)
    print("OPENNEBULA VM MANAGER - ПОЛНАЯ ИНФОРМАЦИЯ С MAC И МЕТКАМИ (LABELS)")
    print("ВКЛЮЧАЯ ЭКСПОРТ В CSV")
    print("="*80)

    # Отображаем информацию о конфигурации
    print(f"\n⚙️  КОНФИГУРАЦИЯ:")
    print(f"   • Endpoint: {ENDPOINT}")
    print(f"   • Пользователь: {USERNAME}")
    print(f"   • Токен: {'установлен' if TOKEN else 'отсутствует'}")
    print(f"   • Размер пакета: {BATCH_SIZE} VM")
    print(f"   • Отображать в таблице: {MAX_DISPLAY} VM")
    print(f"   • Подробный вывод: {'ВКЛ' if VERBOSE else 'ВЫКЛ'}")

    # Проверяем настройку экспорта
    export_enabled = 'EXPORT_CSV' in globals() and EXPORT_CSV
    print(f"   • Экспорт в CSV: {'ВКЛ' if export_enabled else 'ВЫКЛ (будет предложен)'}")

    # 1. Подключаемся к OpenNebula
    print("\n1. 🔌 Подключение к OpenNebula...")
    one = test_with_pyone()

    if not one:
        print("\n❌ Не удалось подключиться к OpenNebula")
        return

    # 2. Получаем список ВСЕХ VM (теперь с полным доступом)
    print("\n2. 📋 Получение списка ВСЕХ виртуальных машин...")
    print(f"   Используем токен с полным доступом")
    vms = get_all_vms_simple(one)

    if not vms:
        print("\n❌ Не удалось получить VM")
        return

    print(f"\n✅ Всего VM найдено: {len(vms)}")

    if len(vms) > 1000:
        print(f"🎉 Отлично! Получен доступ ко всем VM в системе")
    elif len(vms) > 500:
        print(f"⚠️  Получено {len(vms)} VM. Возможно, есть еще ограничения")

    # 3. Собираем данные ОДИН РАЗ для отображения и экспорта
    print("\n3. 📊 Сбор полных данных...")
    display_data, export_data, statistics = collect_vm_data_for_display_and_export(vms, one)

    if not display_data:
        print("\n❌ Не удалось собрать данные VM")
        return

    # 4. Отображаем таблицу в консоли
    print("\n4. 📋 Отображение таблицы VM...")
    display_vm_table(display_data)

    # 5. Отображаем статистику
    print("\n5. 📊 Статистика...")
    display_statistics(statistics)

    # 6. Экспорт в CSV
    print("\n6. 📁 Экспорт данных...")

    # Определяем имя файла для экспорта
    filename = EXPORT_FILENAME  # Используем имя из конфига

    # АВТОМАТИЧЕСКОЕ переименование файла, если он уже существует
    if os.path.exists(filename):
        print(f"⚠️  Файл {filename} уже существует")

        # Генерируем новое имя файла с временной меткой
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name_without_ext, ext = os.path.splitext(filename)
        new_filename = f"{name_without_ext}_{timestamp}{ext}"

        # Проверяем, не существует ли уже файл с таким именем
        counter = 1
        while os.path.exists(new_filename):
            new_filename = f"{name_without_ext}_{timestamp}_{counter}{ext}"
            counter += 1
            if counter > 100:  # На всякий случай ограничим попытки
                print(f"⚠️  Слишком много попыток, перезаписываю {filename}")
                new_filename = filename
                break

        filename = new_filename
        print(f"📝 Создаем новый файл: {filename}")

    # Выполняем экспорт
    exported_file = export_to_csv(export_data, statistics, filename)

    if exported_file:
        print(f"\n🎉 ЭКСПОРТ УСПЕШНО ЗАВЕРШЕН!")
        print(f"   Файл: {exported_file}")
        print(f"   Записей: {len(export_data)}")
    else:
        print(f"❌ Ошибка экспорта данных")

    # 7. Итоговая информация
    print("\n" + "="*80)
    print("📋 ИТОГОВАЯ ИНФОРМАЦИЯ")
    print("="*80)

    print(f"\n📊 ВСЕГО VM: {len(vms)}")
    print(f"🏷️  VM с метками: {statistics['vms_with_labels']} ({statistics['vms_with_labels']/len(vms)*100:.1f}%)")

    if statistics['sorted_labels']:
        print(f"🔑 Уникальных ключей меток: {len(statistics['sorted_labels'])}")

        # Собираем статистику по меткам
        label_counts = {}
        for item in export_data:
            labels = item['resources']['labels']
            for label_key in labels.keys():
                label_counts[label_key] = label_counts.get(label_key, 0) + 1

        # Показываем топ-5 ключей
        if label_counts:
            print(f"\n📈 ТОП-5 КЛЮЧЕЙ МЕТОК:")
            sorted_labels_by_count = sorted(label_counts.items(), key=lambda x: x[1], reverse=True)[:5]
            for label_key, count in sorted_labels_by_count:
                print(f"   • {label_key}: {count} VM")

    print(f"\n✅ СКРИПТ УСПЕШНО ВЫПОЛНЕН!")
    if exported_file:
        print(f"   Данные сохранены в: {exported_file}")
        file_size = os.path.getsize(exported_file) / 1024 / 1024
        print(f"   Размер файла: {file_size:.2f} MB")
    print(f"   Для повторного запуска: python check_vm_count.py")

if __name__ == "__main__":
    main()
