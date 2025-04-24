import zipfile
import os
import filetype

def extract_embedded_files_from_office(file_path, output_folder):
    # Определяем, что ищем: для Word - word/embeddings, для Excel - xl/embeddings
    search_paths = ['word/embeddings/', 'xl/embeddings/']

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        file_list = zip_ref.namelist()

        embedded_files = [
            f for f in file_list 
            if any(f.startswith(path) for path in search_paths)
        ]

        if not embedded_files:
            print('Вложенных файлов не найдено.')
            return

        os.makedirs(output_folder, exist_ok=True)

        for embedded_file in embedded_files:
            filename = os.path.basename(embedded_file)
            output_path = os.path.join(output_folder, filename)

            with open(output_path, 'wb') as out_file:
                out_file.write(zip_ref.read(embedded_file))
            print(f'Извлечен: {output_path}')

            # Пытаемся угадать тип и переименовать
            kind = filetype.guess(output_path)
            if kind:
                new_output_path = os.path.splitext(output_path)[0] + '.' + kind.extension
                os.rename(output_path, new_output_path)
                print(f'Определен тип: {kind.mime}, файл переименован в {new_output_path}')
            else:
                print('Не удалось определить тип файла, оставлен исходный .bin')
# Пример использования
extract_embedded_files_from_office('./docs/Копия Интерактивная схема Переводы_для АКЦ.xlsx', 'output_folder')
