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
                    # Попытка получить имя вложения через стандартные поля
                    print(f'{dir(attachment)}')
                    print(attachment.get_number_of_entries())
                    attachment_name = None
                    for attr in ['long_filename', 'short_filename', 'filename']:
                        if hasattr(attachment, attr):
                            attachment_name = getattr(attachment, attr)
                            if attachment_name:
                                break

                    if not attachment_name:
                        attachment_name = f"Вложение_{i}"

                    print(f"    - {attachment_name}")

        for subfolder in folder.sub_folders:
            process_folder(subfolder)

    process_folder(root_folder)
    pst.close()

# Пример использования
extract_attachments(r"R:\PST\TestPST.pst")
