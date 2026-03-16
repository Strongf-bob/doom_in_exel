# Excel DOOM

Мини-версия DOOM-подобной игры на VBA, которая запускается прямо в Excel как `.xlsm`.

## Что внутри

- псевдо-3D рендер через raycasting;
- управление с клавиатуры через `Application.OnKey`;
- враги-спрайты, стрельба, здоровье и патроны;
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
3. Если игра не стартовала автоматически с подготовленного экрана, открой список макросов (`Alt + F8`) и запусти `ExcelDoom_StartGame`.

## Управление

- `W` / `↑`: идти вперёд
- `S` / `↓`: идти назад
- `A`, `D` / `←`, `→`: поворот
- `Q`, `E`: шаг вбок
- `Space`: выстрел

## Ограничения

Это не полноценный порт оригинального DOOM, а компактная Excel-реализация в том же духе: псевдо-3D коридоры, враги и стрельба в рамках ограничений VBA и листа Excel.
