# Excel DOOM

Мини-версия DOOM-подобной игры на VBA, которая запускается прямо в Excel как `.xlsm`.

## Что внутри

- быстрый ASCII-рендер через raycasting;
- повышенное разрешение экрана `120x40`;
- управление с клавиатуры через `Application.OnKey`;
- пауза, быстрый рестарт и подсказки по цели;
- демоны, которые преследуют игрока и стреляют в ответ;
- стрельба, здоровье и патроны;
- миникарта и HUD на листе Excel;
- воспроизводимая сборка через PowerShell и Excel COM.

## Сборка

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_excel_doom.ps1
```

Готовый файл появится в `output/spreadsheet/ExcelDoom.xlsm`.

## Проверка

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\smoke_test_excel_doom.ps1
```

## Запуск

1. Открой `output/spreadsheet/ExcelDoom.xlsm` в Excel.
2. Разреши макросы.
3. Нажми `START` на листе `DOOM` или, если нужно, открой список макросов (`Alt + F8`) и запусти `ExcelDoom_StartGame`.

## Управление

- `W` / `↑`: идти вперёд
- `S` / `↓`: идти назад
- `A`, `D` / `←`, `→`: поворот
- `Shift + ←`, `Shift + →`: шаг вбок
- `Space`: выстрел
- `F8` или `P`: пауза / продолжить
- `F5` или `R`: быстрый рестарт

## Ограничения

Это не полноценный порт оригинального DOOM, а компактная Excel-реализация в том же духе: быстрый first-person ASCII-рендер, коридоры, демоны и бой в рамках ограничений VBA и листа Excel.
