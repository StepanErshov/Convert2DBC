# CAN DBC Converter Tool

Version: 1.0 (Beta)

Last Updated: 2025-05-14

---

## Описание

Этот инструмент преобразует CAN-матрицу из Excel-файла в DBC-формат, используемый в автомобильных проектах. Поддерживает:

 - Парсинг сигналов, сообщений и их атрибутов из Excel.

 - Генерацию DBC-файла с поддержкой CAN FD.

 - Автоматическую валидацию имен и параметров по правилам CES.

---

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

---

## Требования
- Python 3.9+

- Библиотеки:

  ```bash
  pip install cantools pandas openpyxl PyYAML
  ```
