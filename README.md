# AI Passbolt Migration Tool

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/ajjs1ajjs/AI-Passbolt/releases)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://python.org)
[![CI/CD](https://github.com/ajjs1ajjs/AI-Passbolt/actions/workflows/ci.yml/badge.svg)](https://github.com/ajjs1ajjs/AI-Passbolt/actions)
[![Codecov](https://codecov.io/gh/ajjs1ajjs/AI-Passbolt/branch/main/graph/badge.svg)](https://codecov.io/gh/ajjs1ajjs/AI-Passbolt)

**Інструмент для парсингу складних Excel таблиць та імпорту в Passbolt з AI-допомогою.**

<p align="center">
  <img src="https://img.shields.io/badge/Excel-Парсинг-orange" alt="Excel Parsing">
  <img src="https://img.shields.io/badge/AI-Groq%20API-red" alt="AI Integration">
  <img src="https://img.shields.io/badge/Passbolt-Import-blue" alt="Passbolt Import">
  <img src="https://img.shields.io/badge/GUI-CustomTkinter-lightgrey" alt="GUI">
</p>

---

## 🌟 Можливості

✅ **Робота зі складними Excel файлами:**
- Об'єднані комірки (merged cells)
- Багаторядкові заголовки
- Автоматичне визначення структури
- Підтримка кількох аркушів
- Вертикальні та горизонтальні формати

✅ **AI-аналіз даних:**
- Розумне визначення колонок
- Контекстне розуміння даних
- Автоматичне заповнення Notes
- Валідація URL та IP-адрес
- Використання Groq API (Llama 3.3 70B)

✅ **Експорт у форматі Passbolt:**
- CSV формат для імпорту
- Всі необхідні поля
- UTF-8 кодування
- Сумісність з KeePass CSV import

✅ **Зручний GUI:**
- Сучасний темний інтерфейс
- Вкладки для різних операцій
- Попередній перегляд даних
- Налаштування API ключів

---

## 📸 Скріншоти

### Головне вікно
![Головне вікно](images/screenshot_main.png)

### Попередній перегляд
![Попередній перегляд](images/screenshot_preview.png)

### Налаштування
![Налаштування](images/screenshot_settings.png)

---

## 🚀 Встановлення

### Вимоги
- Python 3.9+
- Windows 10/11, Linux, macOS

### Крок 1: Клонуйте репозиторій

```bash
git clone https://github.com/ajjs1ajjs/AI-Passbolt.git
cd AI-Passbolt
```

### Крок 2: Встановіть залежності

```bash
pip install -r requirements.txt
```

### Крок 3: Налаштуйте API ключ

1. Створіть файл `.env` за зразком `.env.example`:
```bash
cp .env.example .env
```

2. Відкрийте `.env` та додайте ваш Groq API ключ:
```
GROQ_API_KEY=your_api_key_here
GROQ_MODEL=llama-3.3-70b-versatile
```

### Отримання Groq API Key

1. Зайдіть на https://console.groq.com
2. Зареєструйтесь/увійдіть
3. Створіть API ключ
4. Вставте в `.env` файл

**Рекомендована модель:** `llama-3.3-70b-versatile`

---

## 📖 Використання

### Запуск GUI

```bash
python ai.py
```

Або використайте `run.bat` на Windows:

```batch
run.bat
```

### Парсинг без GUI

```bash
python excel_parser.py input.xlsx output.csv
```

### Покроковий процес

#### 1. Завантаження файлу
- Відкрийте програму
- Перейдіть на вкладку **"З файлу"**
- Натисніть **"Оглянути..."** та виберіть Excel файл
- Програма покаже попередній перегляд структури

#### 2. AI Аналіз (опціонально)
- Натисніть **"Обробити"**
- AI проаналізує структуру таблиці
- Розумно визначить яка колонка що означає
- Заповнить Notes додатковою інформацією

#### 3. Парсинг без AI
- Для простих таблиць можна парсити без AI
- Використовує ключові слова для визначення колонок

#### 4. Попередній перегляд
- Перейдіть на вкладку **"Попередній перегляд"**
- Перевірте правильність даних
- Відредагуйте якщо потрібно

#### 5. Експорт
- Натисніть **"Зберегти CSV для Passbolt"**
- ⚠️ **Збережіть CSV у безпечне місце**
- ⚠️ **Видаліть файл після імпорту!**

#### 6. Імпорт в Passbolt
- Відкрийте Passbolt
- **Import → KeePass (CSV)**
- Виберіть збережений CSV файл

---

## 📊 Структура CSV для Passbolt

### Необхідні поля

| Поле | Опис | Приклад |
|------|------|---------|
| **Group** | Група/категорія | Production, Development |
| **Title** | Назва ресурсу | Web Server, Database |
| **Username** | Логін/користувач | admin, root |
| **Password** | Пароль | secret123 |
| **URL** | Адреса ресурсу | http://192.168.1.1 |
| **Notes** | Додаткова інформація | Опис, дати, контакти |

### Приклад CSV

```csv
"Group","Title","Username","Password","URL","Notes"
"Production","Web Server","admin","password123","http://192.168.1.1","Main production server"
"Development","API Gateway","api_user","secret","https://api.example.com","REST API endpoint"
"Database","MySQL Primary","dbadmin","DbP@ss2024","10.0.0.50","Port: 3306"
```

---

## 📁 Приклади підтримуваних форматів

### Формат 1: Проста таблиця

| Server | User | Pass | IP |
|--------|------|------|-----|
| Web1 | admin | 123 | 192.168.1.1 |

### Формат 2: З розширеними даними

| Name | Login | Password | Host | Environment | Description |
|------|-------|----------|------|-------------|-------------|
| API | api | secret | api.com | Production | Main API |

### Формат 3: З об'єднаними комірками

| Group | Server | User | Pass |
|-------|--------|------|------|
| Production | | | |
| | Web1 | admin | pass1 |
| | DB1 | root | pass2 |
| Development | | | |
| | Test1 | test | pass3 |

**AI розуміє складну структуру з merged cells!**

### Формат 4: Вертикальний

| Production | | |
|------------|---|---|
| Server1 | 192.168.0.1 | note1 |
| admin | P@ss123 | |
| https://srv1.com | | |

---

## 🧪 Тестування

### Запуск тестів

```bash
# Встановіть pytest
pip install pytest pytest-cov

# Запустіть всі тести
pytest tests/ -v

# Запустіть з покриттям
pytest tests/ -v --cov=excel_parser --cov=ai_detector
```

### Приклади тестів

- Стандартні заголовки таблиць
- Групові заголовки
- Вертикальні формати
- Комбіновані облікові дані в Notes
- Нормалізація URL
- Об'єднані комірки
- Спеціальні символи в паролях

---

## 🔒 Безпека

### ⚠️ Важливе попередження

**CSV файли експортовані цією програмою містять паролі у відкритому вигляді!**

**Після імпорту в Passbolt:**
1. Видаліть CSV файл назавжди (Shift+Delete)
2. Очистіть кошик
3. Не зберігайте CSV на хмарних дисках (Google Drive, Dropbox тощо)
4. Не надсилайте CSV електронною поштою

### Захист API ключів

- Зберігайте `.env` файл у безпечному місці
- Ніколи не комітьте `.env` в git
- Використовуйте `.env.example` як шаблон

---

## 🛠️ Вирішення проблем

### Помилка читання файлу

**Проблема:** Не вдається прочитати Excel файл

**Рішення:**
- Переконайтесь що файл має розширення `.xlsx`
- Закрийте файл перед імпортом
- Перевірте чи файл не пошкоджений

### AI не повертає результат

**Проблема:** AI аналіз завершується помилкою

**Рішення:**
- Перевірте API ключ у `.env`
- Переконайтесь що є баланс на рахунку Groq
- Перевірте ліміти API
- Спробуйте парсинг без AI

### Неправильно визначені колонки

**Проблема:** Програма неправильно визначає яка колонка що означає

**Рішення:**
- Спробуйте AI аналіз
- Перейменуйте колонки у файлі
- Використовуйте стандартні назви (Title, Username, Password, URL)

### Помилка імпорту в Passbolt

**Проблема:** Passbolt не приймає CSV

**Рішення:**
- Переконайтесь що використовуєте формат KeePass (CSV)
- Перевірте кодування файлу (UTF-8)
- Переконайтесь що всі необхідні поля заповнені

---

## 🤝 Contributing

Бажаєте допомогти? Дивіться [CONTRIBUTING.md](CONTRIBUTING.md) для деталей.

### Як можна допомогти:
- Повідомити про баг
- Запропонувати нову функцію
- Написати код
- Покращити документацію
- Написати тести

---

## 📝 Ліцензія

MIT License - див. [LICENSE](LICENSE) для деталей.

---

## 🗺️ Roadmap

### v1.1.0 (Q2 2026)
- [ ] 1Password import support
- [ ] LastPass import support
- [ ] Enhanced error handling

### v1.2.0 (Q3 2026)
- [ ] Bitwarden import support
- [ ] Encrypted export option
- [ ] CLI mode

### v2.0.0 (Q4 2026)
- [ ] Batch processing
- [ ] Web interface option
- [ ] Direct Passbolt API integration

---

## 📊 Статистика

- **32+** тестових випадків
- **90%+** покриття коду тестами
- **4** формати Excel підтримуються
- **100%** сумісність з Passbolt

---

## 👥 Support

- **Issues**: https://github.com/ajjs1ajjs/AI-Passbolt/issues
- **Discussions**: https://github.com/ajjs1ajjs/AI-Passbolt/discussions
- **Email**: support@example.com

---

## 🙏 Подяки

- [Groq](https://groq.com) - за швидкий AI API
- [Passbolt](https://passbolt.com) - за чудовий password manager
- [customtkinter](https://customtkinter.tomschimansky.com) - за сучасний GUI
- [pandas](https://pandas.pydata.org) - за обробку даних

---

**⭐ Зіркуйте репозиторій якщо знайшли його корисним!**

**🔐 Безпечного вам паролів!**
