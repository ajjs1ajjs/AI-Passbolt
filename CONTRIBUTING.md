# Contributing to AI-Passbolt Migration Tool

Дякуємо за інтерес до покращення AI-Passbolt! Цей документ містить інструкції для контриб'юторів.

## Як можна допомогти

### 1. Повідомити про баг

Використовуйте [GitHub Issues](https://github.com/ajjs1ajjs/AI-Passbolt/issues) для повідомлення про баги.

**Приклад хорошого звіту про баг:**
- Короткий опис проблеми
- Кроки для відтворення
- Очікуваний результат
- Фактичний результат
- Версія Python
- Приклад Excel файлу (якщо можливо)

### 2. Запропонувати функцію

Також використовуйте GitHub Issues для пропозицій нових функцій.

**Приклад хорошої пропозиції:**
- Опис функції
- Для чого вона потрібна
- Приклади використання
- Можливі проблеми реалізації

### 3. Написати код

#### Швидкий старт

```bash
# Fork репозиторій
git clone https://github.com/YOUR_USERNAME/AI-Passbolt.git
cd AI-Passbolt

# Створіть віртуальне середовище
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # Linux/Mac

# Встановіть залежності
pip install -r requirements.txt

# Встановіть інструменти розробника
pip install pytest pytest-cov flake8 black
```

#### Створення Pull Request

1. Створіть гілку для вашої функції (`git checkout -b feature/amazing-feature`)
2. Зробіть зміни
3. Напишіть тести (якщо застосовно)
4. Переконайтесь що всі тести проходять
5. Закомітьте зміни (`git commit -m 'Add amazing feature'`)
6. Відправте на GitHub (`git push origin feature/amazing-feature`)
7. Відкрийте Pull Request

#### Стиль коду

- Використовуйте **Black** для форматування: `black .`
- Використовуйте **Flake8** для linting: `flake8 .`
- Дотримуйтесь PEP 8
- Пишіть документацію для нових функцій

#### Тести

```bash
# Запустити всі тести
pytest tests/ -v

# Запустити з покриттям
pytest tests/ -v --cov=excel_parser --cov=ai_detector

# Запустити конкретний тест
pytest tests/test_excel_parser.py::TestExcelParser::test_standard_header_table -v
```

## Структура проекту

```
AI-Passbolt/
├── ai.py                 # GUI додаток
├── excel_parser.py       # Парсинг Excel файлів
├── ai_detector.py        # AI аналіз даних
├── tests/
│   ├── __init__.py
│   └── test_excel_parser.py  # Тести
├── requirements.txt      # Залежності
├── README.md            # Документація
└── CONTRIBUTING.md      # Цей файл
```

## Code of Conduct

- Будьте ввічливими та поважними
- Допомагайте іншим контриб'юторам
- Приймайте конструктивну критику
- Фокусуйтесь на тому що найкраще для спільноти

## Ліцензія

MIT License - див. [LICENSE](LICENSE) для деталей.

---

**Дякуємо за ваш внесок! 🎉**
