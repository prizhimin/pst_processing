import pypff
from email import policy
from email.parser import BytesParser


def get_attachment_names(pst_file_path):
    pst_file = pypff.file()
    pst_file.open(pst_file_path)

    root_folder = pst_file.get_root_folder()
    attachment_names = []

    def process_folder(folder):
        for sub_folder in folder.sub_folders:
            process_folder(sub_folder)

        for message in folder.sub_messages:
            # Получаем тело сообщения в виде байтов
            raw_email = message.get_plain_text_body()  # Получаем тело сообщения в виде байтов (или get_rtf_body() для RTF)
            print(raw_email)

            if raw_email:  # Если тело сообщения не пустое
                # Парсим его с использованием BytesParser
                msg = BytesParser(policy=policy.default).parsebytes(raw_email)
                print(msg)

                # Проходим по всем частям сообщения (включая вложения)
                for part in msg.walk():
                    # Проверяем, является ли часть вложением через Content-Type и наличие имени файла
                    content_type = part.get_content_type()
                    filename = part.get_filename()

                    # Если тип контента - это обычное вложение, изображение или любой другой файл
                    if filename or (content_type not in ["text/plain", "text/html", "multipart/alternative"]):
                        attachment_names.append(filename if filename else f"Unnamed_{content_type}")

    process_folder(root_folder)
    pst_file.close()
    return attachment_names


# Пример использования функции
if __name__ == "__main__":
    pst_file_path = r"r:\TestPST.pst"  # Путь к вашему PST файлу
    attachment_names = get_attachment_names(pst_file_path)
    print("Названия вложений:", attachment_names)
