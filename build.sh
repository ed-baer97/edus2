#!/usr/bin/env bash
# build.sh - Скрипт для установки зависимостей на Render

set -o errexit  # Выход при ошибке

echo "Установка системных зависимостей..."

# Обновляем список пакетов
apt-get update -qq

# Устанавливаем Chrome/Chromium и необходимые зависимости
apt-get install -y -qq \
    wget \
    gnupg \
    unzip \
    curl \
    ca-certificates \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libc6 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libexpat1 \
    libfontconfig1 \
    libgbm1 \
    libgcc1 \
    libglib2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libstdc++6 \
    libx11-6 \
    libx11-xcb1 \
    libxcb1 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxi6 \
    libxrandr2 \
    libxrender1 \
    libxss1 \
    libxtst6 \
    lsb-release \
    xdg-utils

# Устанавливаем Google Chrome
echo "Установка Google Chrome..."
wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add - || true
echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list
apt-get update -qq
apt-get install -y -qq google-chrome-stable || apt-get install -y -qq google-chrome-beta || apt-get install -y -qq google-chrome-unstable

# Проверяем установку Chrome
if command -v google-chrome &> /dev/null; then
    echo "✓ Google Chrome установлен: $(google-chrome --version)"
else
    echo "⚠ Google Chrome не установлен, используем Chromium..."
    apt-get install -y -qq chromium-browser || apt-get install -y -qq chromium
fi

# Устанавливаем Python зависимости
echo "Установка Python зависимостей..."
pip install --upgrade pip
pip install -r requirements.txt

echo "✓ Сборка завершена успешно!"
