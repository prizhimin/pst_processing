import os
import argparse
from datetime import datetime, timezone, timedelta
import pypff
import re
import unicodedata
from bs4 import BeautifulSoup
from striprtf.striprtf import rtf_to_text

# Константа для временной зоны GMT+3
GMT3 = timezone(timedelta(hours=3))


def print_header():
    """Выводит заголовок программы"""
    print("\n" + "=" * 80)
    print("PST File Search Tool".center(80))
    print("=" * 80 + "\n")


def ensure_output_dir(output_dir):
    """Создает каталог для сохранения, если он не существует"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"[+] Создан каталог для сохранения: {output_dir}")


def sanitize_filename(filename):
    """Очищает строку для использования в имени файла"""
    filename = unicodedata.normalize('NFKD', filename)
    filename = re.sub(r'[\\/*?:"<>|]', "", filename)
    filename = filename.replace('\n', ' ').replace('\r', ' ')
    filename = filename.strip('. ')
    return filename[:250]


def get_recipients(message):
    """Получает список получателей сообщения"""
    recipients = []
    try:
        for recipient in message.recipients:
            name = getattr(recipient, 'name', '')
            if name:
                recipients.append(name)
    except:
        pass
    return recipients


def convert_to_gmt3(dt):
    """Конвертирует datetime в GMT+3"""
    if dt is None:
        return None
    if dt.tzinfo is None:
        # Если время без временной зоны, считаем что это UTC
        return dt.replace(tzinfo=timezone.utc).astimezone(GMT3)
    return dt.astimezone(GMT3)


def format_datetime_gmt3(dt):
    """Форматирует datetime в строку с указанием GMT+3"""
    if dt is None:
        return "Неизвестно"
    dt_gmt3 = convert_to_gmt3(dt)
    return dt_gmt3.strftime('%Y-%m-%d %H:%M:%S (GMT+3)')


def get_message_body(message):
    """Улучшенное извлечение тела письма с обработкой RTF"""
    try:
        body = getattr(message, 'plain_text_body', None)
        if body:
            if isinstance(body, bytes):
                body = body.decode('utf-8', errors='replace')
            return str(body).strip()

        rtf_body = getattr(message, 'rtf_body', None)
        if isinstance(rtf_body, bytes):
            rtf_body = rtf_body.decode('utf-8', errors='replace')
        if rtf_body:
            return rtf_to_text(rtf_body.strip())

        html_body = getattr(message, 'html_body', None)
        if html_body:
            return BeautifulSoup(html_body, 'html.parser').get_text().strip()

        return "Тело письма отсутствует"
    except Exception as e:
        print(f"[!] Ошибка извлечения тела письма: {e}")
        return "Не удалось извлечь текст"


def get_folder_path(message):
    """Возвращает путь к папке, содержащей сообщение"""
    try:
        folder = message.parent_folder
        path = []
        while folder:
            path.append(getattr(folder, 'name', 'Unknown Folder'))
            folder = getattr(folder, 'parent_folder', None)
        return " > ".join(reversed(path))
    except:
        return "Неизвестная папка"


def check_time_in_range(dt, time_range):
    """Проверяет, попадает ли время в указанный диапазон часов"""
    if not dt:
        return False

    dt_gmt3 = convert_to_gmt3(dt)
    t = dt_gmt3.time()
    start_hour, end_hour = time_range

    if start_hour <= end_hour:
        return start_hour <= t.hour < end_hour
    else:
        return t.hour >= start_hour or t.hour < end_hour


def matches_criteria(sender, recipients, subject, body,
                     received_time, sent_time, criteria):
    """Проверяет соответствие сообщения критериям поиска"""
    if criteria.get('sender') and criteria['sender'].lower() not in sender.lower():
        return False

    if criteria.get('recipient'):
        recipient_match = any(
            criteria['recipient'].lower() in r.lower()
            for r in recipients
        )
        if not recipient_match:
            return False

    if criteria.get('subject') and criteria['subject'].lower() not in subject.lower():
        return False

    if criteria.get('body') and criteria['body'].lower() not in body.lower():
        return False

    # Конвертируем временные метки в GMT+3 перед сравнением
    received_time_gmt3 = convert_to_gmt3(received_time) if received_time else None
    sent_time_gmt3 = convert_to_gmt3(sent_time) if sent_time else None

    if criteria.get('received_after') and received_time_gmt3:
        if received_time_gmt3 < convert_to_gmt3(criteria['received_after']):
            return False

    if criteria.get('received_before') and received_time_gmt3:
        if received_time_gmt3 > convert_to_gmt3(criteria['received_before']):
            return False

    if criteria.get('sent_after') and sent_time_gmt3:
        if sent_time_gmt3 < convert_to_gmt3(criteria['sent_after']):
            return False

    if criteria.get('sent_before') and sent_time_gmt3:
        if sent_time_gmt3 > convert_to_gmt3(criteria['sent_before']):
            return False

    if criteria.get('received_time_range') and received_time_gmt3:
        if not check_time_in_range(received_time_gmt3, criteria['received_time_range']):
            return False

    if criteria.get('sent_time_range') and sent_time_gmt3:
        if not check_time_in_range(sent_time_gmt3, criteria['sent_time_range']):
            return False

    return True


def save_message_as_txt(message, output_dir, msg_num):
    """Безопасное сохранение письма с временем в GMT+3"""
    try:
        sender = str(getattr(message, 'sender_name', None)) or "Неизвестный_отправитель"
        subject = str(getattr(message, 'subject', None)) or "Без_темы"

        # Конвертируем время в GMT+3
        received_time = convert_to_gmt3(getattr(message, 'delivery_time', None))
        sent_time = convert_to_gmt3(getattr(message, 'client_submit_time', None))

        date_part = (received_time or sent_time or datetime.now(GMT3)).strftime('%Y%m%d_%H%M')
        filename = f"{date_part}_{sanitize_filename(sender)}_{sanitize_filename(subject)}.txt"
        filepath = os.path.join(output_dir, filename)

        body = get_message_body(message)

        content = [
            f"ПАПКА: {get_folder_path(message)}",
            f"НОМЕР: {msg_num}",
            f"ОТПРАВИТЕЛЬ: {sender}",
            f"ПОЛУЧАТЕЛИ: {', '.join(get_recipients(message)) or 'Не указаны'}",
            f"ТЕМА: {subject}",
            f"ОТПРАВЛЕНО: {format_datetime_gmt3(sent_time)}",
            f"ПОЛУЧЕНО: {format_datetime_gmt3(received_time)}",
            "\nТЕКСТ ПИСЬМА:",
            "=" * 80,
            body,
            "=" * 80
        ]

        with open(filepath, 'w', encoding='utf-8', errors='replace') as f:
            f.write('\n'.join(content))

        print(f"[+] Сохранено письмо #{msg_num}: {filename}")
        return filepath
    except Exception as e:
        print(f"[!] Критическая ошибка при сохранении письма #{msg_num}: {str(e)}")
        return None


def parse_datetime(dt_str):
    """Преобразует строку в datetime с учетом GMT+3"""
    try:
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d.%m.%Y', '%d.%m.%Y %H:%M:%S'):
            try:
                dt = datetime.strptime(dt_str, fmt)
                return dt.replace(tzinfo=GMT3)
            except ValueError:
                continue
        return None
    except ValueError:
        return None


def parse_time_range(time_str):
    """Парсит строку с диапазоном времени"""
    try:
        start, end = map(int, time_str.split('-'))
        return (start, end)
    except:
        return None


def search_pst(pst_path, search_criteria, output_dir=None):
    """Основная функция поиска в PST-файле"""
    try:
        print(f"[+] Открываю PST-файл: {pst_path}")
        pst = pypff.file()
        pst.open(pst_path)

        if output_dir:
            ensure_output_dir(output_dir)
            print(f"[+] Найденные письма будут сохранены в: {os.path.abspath(output_dir)}")

        root = pst.get_root_folder()
        print(f"[+] Найдено корневых папок: {root.number_of_sub_folders}")

        total_messages = process_folder(root, search_criteria, 0, output_dir)

        print(f"\n[+] Поиск завершен. Обработано сообщений: {total_messages}")
        if output_dir and os.path.exists(output_dir):
            print(f"[+] Сохранено писем: {len(os.listdir(output_dir))}")
        pst.close()
    except IOError as e:
        print(f"[!] Ошибка при открытии файла: {e}")
    except Exception as e:
        print(f"[!] Критическая ошибка: {e}")


def process_folder(folder, search_criteria, counter, output_dir=None):
    """Рекурсивно обрабатывает папки PST"""
    try:
        for message in folder.sub_messages:
            counter += 1
            process_message(message, search_criteria, counter, output_dir)

        for subfolder in folder.sub_folders:
            counter = process_folder(subfolder, search_criteria, counter, output_dir)
    except AttributeError as e:
        print(f"[!] Ошибка доступа к папке: {e}")
    except Exception as e:
        print(f"[!] Ошибка при обработке папки: {e}")
    return counter


def process_message(message, search_criteria, msg_num, output_dir=None):
    """Обрабатывает отдельное сообщение"""
    try:
        sender = getattr(message, 'sender_name', 'Не указан')
        recipients = get_recipients(message)
        subject = getattr(message, 'subject', 'Без темы')
        body = get_message_body(message)

        # Конвертируем время в GMT+3
        received_time = convert_to_gmt3(getattr(message, 'delivery_time', None))
        sent_time = convert_to_gmt3(getattr(message, 'client_submit_time', None))

        if not matches_criteria(sender, recipients, subject, body,
                                received_time, sent_time, search_criteria):
            return

        print(f"\n[+] Найдено письмо #{msg_num}:")
        print(f"    Отправитель: {sender}")
        print(f"    Тема: {subject}")
        if sent_time:
            print(f"    Отправлено: {format_datetime_gmt3(sent_time)}")

        if output_dir:
            save_message_as_txt(message, output_dir, msg_num)
    except Exception as e:
        print(f"[!] Ошибка при обработке сообщения #{msg_num}: {e}")


def main():
    print_header()
    parser = argparse.ArgumentParser(
        description='Поиск в PST-файле с сохранением результатов',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('pst_file', help='Путь к PST-файлу')
    parser.add_argument('--output-dir', required=True,
                        help='Каталог для сохранения найденных писем')
    parser.add_argument('--sender', help='Фильтр по отправителю')
    parser.add_argument('--recipient', help='Фильтр по получателю')
    parser.add_argument('--subject', help='Фильтр по теме письма')
    parser.add_argument('--body', help='Фильтр по тексту письма')
    parser.add_argument('--sent-after', help='Письма, отправленные после указанной даты (YYYY-MM-DD HH:MM:SS)')
    parser.add_argument('--sent-before', help='Письма, отправленные до указанной даты (YYYY-MM-DD HH:MM:SS)')
    parser.add_argument('--received-after', help='Письма, полученные после указанной даты (YYYY-MM-DD HH:MM:SS)')
    parser.add_argument('--received-before', help='Письма, полученные до указанной даты (YYYY-MM-DD HH:MM:SS)')
    parser.add_argument('--sent-time', help='Диапазон часов отправки (формат: HH-HH, например 8-17 или 22-6)')
    parser.add_argument('--received-time', help='Диапазон часов получения (формат: HH-HH, например 8-17 или 22-6)')

    args = parser.parse_args()
    criteria = {}
    if args.sender: criteria['sender'] = args.sender
    if args.recipient: criteria['recipient'] = args.recipient
    if args.subject: criteria['subject'] = args.subject
    if args.body: criteria['body'] = args.body

    if args.sent_after:
        criteria['sent_after'] = parse_datetime(args.sent_after)
    if args.sent_before:
        criteria['sent_before'] = parse_datetime(args.sent_before)
    if args.received_after:
        criteria['received_after'] = parse_datetime(args.received_after)
    if args.received_before:
        criteria['received_before'] = parse_datetime(args.received_before)

    if args.sent_time:
        time_range = parse_time_range(args.sent_time)
        if time_range:
            criteria['sent_time_range'] = time_range
        else:
            print("[!] Неверный формат диапазона времени для --sent-time")

    if args.received_time:
        time_range = parse_time_range(args.received_time)
        if time_range:
            criteria['received_time_range'] = time_range
        else:
            print("[!] Неверный формат диапазона времени для --received-time")

    search_pst(args.pst_file, criteria, args.output_dir)


if __name__ == '__main__':
    main()
