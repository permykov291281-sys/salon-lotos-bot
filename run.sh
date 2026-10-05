#!/bin/bash
# Перезапуск Настены в фоне: ./run.sh
cd "$(dirname "$0")"
pkill -f "^python3 -u bot.py --lotos" && sleep 1
setsid nohup python3 -u bot.py --lotos >> data/bot.log 2>&1 < /dev/null &
