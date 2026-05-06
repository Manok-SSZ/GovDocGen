# GovDocGen — Система подготовки правительственных документов

**GovDocGen** — модульная программа на Python с графическим интерфейсом (tkinter) для генерации документов Microsoft Word (DOCX) и их конвертации в многостраничные JPG. Предназначена для автоматизации подготовки приказов, уставов, кадровых документов и других официальных бумаг в государственных структурах. Сделано для проекта [RMRP](https://rmrp.ru/).

<img width="713" height="692" alt="изображение" src="https://github.com/user-attachments/assets/e5679990-25c9-44a0-b6a4-003437f1bc2d" />

## Возможности

- Модульная архитектура — каждый тип документа оформляется в виде отдельного Python-модуля.
- Поддержка **Microsoft Word** и **LibreOffice** для конвертации DOCX → PDF → JPG.
- Создание **многостраничных JPG** (если документ не помещается на одной странице).
- Контекстное меню «Вставить» для полей ввода (работает в любой раскладке клавиатуры).
- Сохранение последних введённых данных и настроек.
- Единые общие поля: дата, номер документа, подписывающий, электронная подпись (CANVAS).

## Системные требования
- Windows 7 / 8 / 10 / 11
- Установленный **Microsoft Word** *или* **LibreOffice** (для конвертации в JPG)
- **Poppler** (для конвертации PDF → JPG)

## Зависимости Python
docxtpl pdf2image pillow comtypes

## При первом запуске, программа создаст следующую структуру (если таковая отсутствует)
- TEMPLATES/                 # Папка с шаблонами .docx
- MODULES/                   # Папка с модулями .py
- COMPLETE/                  # Готовые DOCX
- IMAGES/                    # Готовые JPG
- LOGS/                      # Логи ошибок
- CANVAS.jpg                 # Изображение с вашей подписью (замените на свою в формате JPG)

## Модули и шаблоны
Модуль пишется на python и помещается в папку MODULES. Шаблон в формате docx помещается в папку TEMPLATES

### Имя модуля, например: MY_MODULE.py
```python 
MODULE_INFO = {
    "name": "Название документа (будет в имени файла)",
    "tab_name": "Название вкладки в интерфейсе",
    "template": "имя_шаблона.docx"   # файл в папке TEMPLATES
}

def create_ui(parent, app_instance, module_name):
    """Создание интерфейса модуля"""
    frame = ttk.Frame(parent)
    # ... создание полей ввода ...
    return frame
```

### Пример модуля
```python 
MODULE_INFO = {
    "name": "Пример",
    "tab_name": "Пример",
    "template": "EXAMPLE_template.docx"
}

def create_ui(parent, app_instance, module_name):
    frame = ttk.Frame(parent)
    entry = ttk.Entry(frame)
    entry.pack()
    
    def get_context(common_data):
        return {'MY_VAR': entry.get()}
    
    def validate_inputs():
        if not entry.get():
            return False, "Заполните поле"
        return True, ""
    
    def save_data():
        return {'value': entry.get()}
    
    def load_data(data):
        if data and 'value' in data:
            entry.insert(0, data['value'])
    
    def get_template_path():
        from pathlib import Path
        return Path.cwd() / "TEMPLATES" / "EXAMPLE_template.docx"
    
    frame.get_context = get_context
    frame.validate_inputs = validate_inputs
    frame.save_data = save_data
    frame.load_data = load_data
    frame.get_template_path = get_template_path
    
    return frame
```

## Обязательные методы, которые должны быть у возвращаемого frame
- get_context(common_data) -	Возвращает словарь переменных для шаблона
- get_template_path() -	Возвращает путь к файлу шаблона
- validate_inputs() -	Проверяет заполнение полей, возвращает (True, "") или (False, "сообщение об ошибке")
- save_data() -	Возвращает словарь для сохранения состояния модуля
- load_data(data) -	Загружает сохранённое состояние

## Общие переменные для шаблонов
- {{dd}}	День
- {{mm}}	Месяц прописью
- {{yyyy}}	Год
- {{num}}	Номер документа
- {{signer}}	Подписывающий

