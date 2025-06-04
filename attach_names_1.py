import pypff


def extract_attachments(pst_file_path):
    pst = pypff.file()
    pst.open(pst_file_path)

    root_folder = pst.get_root_folder()

    def process_folder(folder):
        print(f"\nПапка: {folder.name}")

        for message in folder.sub_messages:
            print(f"\n  Письмо: {message.subject}")

            if message.number_of_attachments > 0:
                print("  Вложения:")
                for i, attachment in enumerate(message.attachments, 1):
                    # Пытаемся получить имя вложения разными способами
                    attachment_name = getattr(attachment, "name", None)  # Стандартный способ

                    # Если имя не найдено, пробуем альтернативные методы
                    if not attachment_name:
                        try:
                            # Иногда имя хранится в другом атрибуте или как свойство
                            attachment_name = attachment.get_name()  # Альтернативный метод (если есть)
                        except:
                            attachment_name = f"Вложение_{i}"  # Дефолтное имя, если ничего не найдено

                    print(f"    - {attachment_name}")

        for subfolder in folder.sub_folders:
            process_folder(subfolder)

    process_folder(root_folder)
    pst.close()


# Пример использования
extract_attachments(r"C:\Users\prizh\Documents\PST\Alexandr.Smirnov.pst")
