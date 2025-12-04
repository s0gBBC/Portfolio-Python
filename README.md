Sales-Analytics-Automation/
│
├── README.md                    # Главное описание проекта
├── requirements.txt             # Зависимости
├── config.yaml                  # Конфигурация (если нужно)
│
├── src/                         # Исходный код
│   ├── __init__.py
│   ├── email_monitor.py         # Основной код мониторинга
│   ├── data_processor.py        # Логика обработки данных
│   ├── excel_updater.py         # Работа с Excel
│   └── utils.py                 # Вспомогательные функции
│
├── docs/                        # Документация
│   ├── business_context.md      # Бизнес-контекст задачи
│   ├── architecture.md          # Архитектура решения
│   └── results_metrics.md       # Результаты и метрики
│
├── examples/                    # Примеры
│   ├── input_example.xlsx       # Пример входного файла (без реальных данных)
│   ├── output_example.xlsx      # Пример выходного файла
│   └── dashboard_screenshot.png # Скриншот дашборда (если есть)
│
├── tests/                       # Тесты
│   └── test_data_processor.py
│
└── deployment/                  # Деплоймент
    ├── dockerfile               # Docker-контейнер
    ├── airflow_dag.py           # DAG для Apache Airflow
    └── windows_service.ps1      # Скрипт для запуска как службы Windows
