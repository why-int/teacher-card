# Teacher Card Builder

Десктоп-приложение для формирования карточек преподавателей из РУП.

## Что умеет

- загружает РУП (`.xlsx/.xlsm/.xls`);
- даёт выбрать преподавателей;
- позволяет править таблицу по каждому преподавателю;
- формирует выходной Excel по шаблону.

## Запуск без зависимостей (для пользователя)

Пользователю ничего ставить не нужно: запускается готовый файл двойным кликом.

- Windows: `dist\ExcelDataViewer.exe`
- macOS: `dist/ExcelDataViewer.app`

Важно: `.exe` собирается на Windows, `.app` собирается на macOS.

## Готовые файлы без локальной сборки

Если не хочешь собирать локально:

1. Загрузи проект в GitHub.
2. Открой вкладку `Actions`.
3. Запусти workflow `Build Desktop Apps` (или просто сделай push в `main`).
4. После завершения скачай артефакты:
   - `ExcelDataViewer-windows` (внутри `.exe`)
   - `ExcelDataViewer-macos` (внутри `.dmg`)

## Сборка

### Windows

```bat
run\build_windows.bat
```

### macOS

```bash
chmod +x run/build_macos.sh
./run/build_macos.sh
```

Скрипты автоматически вшивают в сборку:

- `assets/` (иконка и ресурсы),
- `Пример.xlsx` и `Шаблон и пример.xlsm` (если файлы есть в корне проекта).

## Разработка

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```
