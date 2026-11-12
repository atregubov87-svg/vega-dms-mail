# Сортировка писем ДМС

Сканер ящика `vegam2020@rambler.ru`: разбирает вложения-списки ДМС, ставит звёздочку письмам
с адресами Сергиево-Посадского округа, убирает из INBOX списки без местных пациентов,
раскладывает письма по отправителям (ИНВИТРО / Корзина).

## Что где

| Что | Где |
|---|---|
| Запуск (задача «Vega DMS Mail Scan», каждые 15 мин) | `C:\ProgramData\VegaDMS\scan.cmd` → `scan_flag.js` |
| Основной скрипт | `scan_flag.js` |
| Настройки IMAP | `config.json` (пароль; не копировать наружу) |
| Состояние (UID-ы писем) | `scan_state.json` |
| Отчёт последнего прогона | `last_scan.json` |
| Логи | `scan.log`, `scheduler.log` |
| Справочник посёлков | `sergiev_posad_settlements.json` |
| Разовые диагностические скрипты | `tools\` (читают `config.json` из родительской папки) |
| Старые бэкапы состояния | `archive\` |
| Описания проекта | `docs\` |

## Ручной запуск

```powershell
cd C:\Users\Администратор\vegasrv\dms-mail
$env:DRY_RUN=1; node scan_flag.js   # проверка: без звёздочек, перемещений и записи состояния
node scan_flag.js                    # боевой прогон
$env:RESCAN=1; node scan_flag.js     # перепроверить всё заново
Remove-Item Env:DRY_RUN
```

Задачу можно запустить вручную без ожидания:

```powershell
Start-ScheduledTask -TaskName 'Vega DMS Mail Scan'
```
