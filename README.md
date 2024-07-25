### README (English)

## Excel to Word Automation Tool

This is a Tkinter-based GUI application for automating the process of generating Word documents from Excel data.

### Features
- Select an Excel file to extract data.
- Select a Word template to populate with extracted data.
- Generate a Word document with customized data and save it to a specified output folder.
- Dynamic user interface for selecting names and inputting custom values.

### Requirements
- Python 3.x
- Tkinter library
- openpyxl library
- python-docx library

### Installation

1. **Clone the Repository:**
   ```sh
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **Install Required Libraries:**
   ```sh
   pip install tkinter openpyxl python-docx
   ```

### Usage

1. **Run the Application:**
   ```sh
   python <script-name>.py
   ```

2. **Select Files and Output Folder:**
   - Click "Browse" to select an Excel file.
   - Click "Browse" to select a Word template.
   - Click "Browse" to select an output folder.

3. **Select Names and Input Custom Values:**
   - Check the boxes next to the names as required.
   - Enter custom values in the corresponding entry fields.

4. **Generate Document:**
   - Click "Generate Document" to create a Word document based on the provided data and save it to the selected output folder.

### Customization

- **Adding/Removing Names:**
  - Modify the `names` list in the script to add or remove names.
  
- **Changing Key Mapping:**
  - Modify the `presence_keys` and `time_keys` dictionaries to change the key mappings for name presence and time values.

### Troubleshooting

- **Error Reading Excel File:**
  - Ensure the Excel file is in `.xlsx` or `.xls` format.
  - Check the cell reference used in the script (`L43`) to ensure it is correct.

- **Error Loading Word Template:**
  - Ensure the Word template is in `.docx` format.

- **Error Saving Document:**
  - Ensure the output folder path is correct and accessible.

### License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

### README (Українська)

## Інструмент автоматизації Excel до Word

Це графічний додаток на базі Tkinter для автоматизації процесу створення документів Word з даних Excel.

### Можливості
- Вибір файлу Excel для вилучення даних.
- Вибір шаблону Word для заповнення витягнутими даними.
- Створення документа Word з налаштованими даними та збереження його у вказану вихідну папку.
- Динамічний інтерфейс для вибору імен та введення користувацьких значень.

### Вимоги
- Python 3.x
- Бібліотека Tkinter
- Бібліотека openpyxl
- Бібліотека python-docx

### Встановлення

1. **Клонуйте репозиторій:**
   ```sh
   git clone <url-репозиторію>
   cd <каталог-репозиторію>
   ```

2. **Встановіть необхідні бібліотеки:**
   ```sh
   pip install tkinter openpyxl python-docx
   ```

### Використання

1. **Запустіть додаток:**
   ```sh
   python <ім'я-скрипта>.py
   ```

2. **Виберіть файли та вихідну папку:**
   - Натисніть "Огляд", щоб вибрати файл Excel.
   - Натисніть "Огляд", щоб вибрати шаблон Word.
   - Натисніть "Огляд", щоб вибрати вихідну папку.

3. **Виберіть імена та введіть користувацькі значення:**
   - Відмітьте необхідні імена.
   - Введіть користувацькі значення у відповідні поля.

4. **Створіть документ:**
   - Натисніть "Generate Document", щоб створити документ Word на основі наданих даних та зберегти його у вибрану вихідну папку.

### Налаштування

- **Додавання/видалення імен:**
  - Змініть список `names` у скрипті, щоб додати або видалити імена.
  
- **Зміна ключових відповідностей:**
  - Змініть словники `presence_keys` та `time_keys`, щоб змінити ключові відповідності для наявності імен та значень часу.

### Усунення несправностей

- **Помилка читання файлу Excel:**
  - Переконайтеся, що файл Excel має формат `.xlsx` або `.xls`.
  - Перевірте посилання на комірку, що використовується у скрипті (`L43`), щоб переконатися, що воно правильне.

- **Помилка завантаження шаблону Word:**
  - Переконайтеся, що шаблон Word має формат `.docx`.

- **Помилка збереження документа:**
  - Переконайтеся, що шлях до вихідної папки правильний і доступний.

### Ліцензія

Цей проект ліцензований відповідно до ліцензії MIT. Див. файл `LICENSE` для отримання деталей.
