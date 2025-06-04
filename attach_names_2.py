import os
import pypff

def detect_file_type(data):
    if data.startswith(b'%PDF'):
        return 'pdf'
    elif data.startswith(b'PK\x03\x04'):
        return 'zip_or_office'
    elif data.startswith(b'Rar!\x1A\x07\x00'):
        return 'rar'
    else:
        return 'bin'

def extract_attachments(pst_file_path, output_dir="extracted_attachments"):
    pst = pypff.file()
    pst.open(pst_file_path)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    root_folder = pst.get_root_folder()

    def process_folder(folder):
        print(f"\nПапка: {folder.name}")

        for message in folder.sub_messages:
            print(f"\n  Письмо: {message.subject}")

            if message.number_of_attachments > 0:
                print("  Вложения:")
                for i, attachment in enumerate(message.attachments, 1):
                    try:
                        # Получаем размер вложения
                        size = attachment.get_size()  # или attachment.size
                        if size <= 0:
                            print(f"    [Пропущено: вложение с нулевым размером]")
                            continue

                        # Пробуем получить данные вложения
                        try:
                            data = attachment.read_buffer()
                        except:
                            data = attachment.get_data()  # Попробуем альтернативный метод

                        # Если данные не получены, пропускаем вложение
                        if not data:
                            print(f"    [Ошибка получения данных вложения]")
                            continue

                    except Exception as e:
                        print(f"    [Ошибка чтения вложения: {e}]")
                        continue

                    # Определяем тип файла
                    file_type = detect_file_type(data)
                    if file_type == 'pdf':
                        ext = 'pdf'
                    elif file_type == 'rar':
                        ext = 'rar'
                    elif file_type == 'zip_or_office':
                        ext = 'zip'
                    else:
                        ext = 'bin'

                    base_name = message.subject or "без_темы"
                    safe_name = "".join(c if c.isalnum() else "_" for c in base_name)[:50]
                    filename = f"{safe_name}_attachment_{i}.{ext}"

                    file_path = os.path.join(output_dir, filename)

                    with open(file_path, "wb") as f:
                        f.write(data)

                    print(f"    - сохранено как: {filename} (тип: {ext})")

        for subfolder in folder.sub_folders:
            process_folder(subfolder)

    process_folder(root_folder)
    pst.close()

# Пример использования
extract_attachments(r"R:\PST\TestPST.pst")
