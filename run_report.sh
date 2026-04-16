#!/bin/bash
cd /home/asokolova/tg-reporter
/usr/bin/docker-compose run --rm reporter
find /home/asokolova/tg-reporter -name "*.xlsx" -mtime +7 -delete
