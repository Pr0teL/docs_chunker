# 🧩 Document Chunker

Document Chunker — это Python-инструмент для иерархического разбиения документов на чанки по заголовкам и подзаголовкам с дополнительной обработкой таблиц, изображений и вложений.

## 🚀 Возможности

- 📄 Поддержка форматов: **`.docx`**, **`.xlsx`**
- 🧱 Разделение текста по заголовкам и подзаголовкам с сохранением структуры
- 🖼️ Извлечение изображений и преобразование в **Base64**
- 🧾 Преобразование таблиц в **HTML**
- 📎 Извлечение embedded-документов и сохранение в указанную директорию
- 🗃️ Возврат структуры документа в формате файла txt с разделением или JSON

---

## 📦 Установка

```bash
git clone https://github.com/Pr0teL/docs_chunker.git
cd document-chunker
python -m venv .venv
source .venv/bin/activate  # Для Windows: .venv\Scripts\activate
pip install -r requirements.txt
