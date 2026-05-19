#!/bin/bash
echo "--- Запуск задачи: $(date) ---"
cd /home/asokolova/tg-reporter
echo "Запуск отчета..."
/usr/bin/docker-compose run --rm reporter
echo "Создаем бекап"
# Создаем папку для бэкапов, если её нет
mkdir -p ./backups

cp report_main.xlsx ./backups/backup_main_$(date +%F).xlsx


# Удаляем бэкапы старше 7 дней (чтобы не забить диск)
find ./backups -type f -name "backup_main_*.xlsx" -mtime +7 -delete

echo "--- Задача завершена: $(date) ---"

