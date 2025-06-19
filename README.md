# CAN DBC Converter Tool

Version: 1.1 (Beta)

Last Updated: 2025-06-05

Working link: [Converter2dbc.app](https://convert2dbc-beta.streamlit.app/)

## Описание

Этот инструмент преобразует CAN-матрицу из Excel-файла в DBC-формат, используемый в автомобильных проектах. Поддерживает:

 - Парсинг сигналов, сообщений и их атрибутов из Excel.

 - Генерацию DBC-файла с поддержкой CAN FD.

 - Автоматическую валидацию имен и параметров по правилам CES.


## Поддерживаемые функции

 - Конвертация:

    - Message ID, Name, Length, Cycle Time.

    - Signal Name, Start Bit, Length, Byte Order (Intel/Motorola).
    
    - Физические значения (Factor, Offset, Min/Max).
    
    - Описания сигналов и их значений (Value Descriptions).

 - Валидация:

    - Проверка формата имен (CES_*, NM_*).
    
    - Контроль длины сообщений (8/64 байта для CAN FD).
    
    - Проверка порядка байт (Motorola MSB).


## Требования
- Python 3.9+

- Библиотеки:

  ```bash
  pip install cantools pandas openpyxl PyYAML
  ```

## Как использовать
### Подготовка Excel-файла:

   Данные должны быть в листе Matrix.

### Обязательные колонки:

```
Message ID, Message Name, Signal Name, Start Byte, Start Bit, Bit Length, Factor, Offset, Initinal Value(Hex), Invalid Value(Hex), Min Value, Max Value, Unit, Receiver, Byte Order, Data Type, Msg Cycle Time, Msg Send Type, Description, Msg Length, Signal Value Description, Senders
```
### Запуск конвертации:

```bash
python xlsx2dbc.py --input <file_name>.xlsx --output <file_name>.dbc
```

### Аргументы командной строки:

| Аргумент | Описание |
|---|---|
| --input | Путь к Excel-файлу (обязательно). |
| --output | Имя выходного DBC-файла. |
| --validate | Включить валидацию (опционально). |

## Ограничения
`Beta-версия:`

Не все типы данных (например, Float) полностью протестированы.

Нет поддержки мультиплексных сигналов.

`Требования к Excel:`

Колонки должны строго соответствовать шаблону.

## Планы по развитию
 - Добавить поддержку ARXML.

 - Интеграция с CI/CD (автопроверка при коммитах).

 - Генерация отчетов в HTML/PDF.



 ![asdsd](https://placehold.co/600x400?text=CAN+LIN+Tools)