#!/bin/bash
# Переходим в папку проекта
cd /home/asokolova/tg-reporter

# Создаем папку для бэкапов, если её нет
mkdir -p ./backups

cp report_main.xlsx ./backups/backup_main_$(date +%F).xlsx


# Удаляем бэкапы старше 7 дней (чтобы не забить диск)
find ./backups -type f -name "backup_main_report_*.xlsx" -mtime +7 -delete
