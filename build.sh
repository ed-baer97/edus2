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
    fonts-liberation2 \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libatspi2.0-0 \
    libc6 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
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
    xdg-utils \
    xvfb

# Устанавливаем Google Chrome
echo "Установка Google Chrome..."
if [ ! -f /etc/apt/sources.list.d/google-chrome.list ]; then
    wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/google-chrome-keyring.gpg
    echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome-keyring.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list
    apt-get update -qq
fi

# Пробуем установить Chrome
if ! apt-get install -y -qq google-chrome-stable 2>/dev/null; then
    echo "Пробуем установить google-chrome-beta..."
    if ! apt-get install -y -qq google-chrome-beta 2>/dev/null; then
        echo "Пробуем установить google-chrome-unstable..."
        apt-get install -y -qq google-chrome-unstable || echo "Не удалось установить Chrome, используем Chromium"
    fi
fi

# Проверяем установку Chrome
if command -v google-chrome &> /dev/null; then
    CHROME_VERSION=$(google-chrome --version 2>/dev/null || google-chrome-stable --version 2>/dev/null || echo "unknown")
    echo "✓ Google Chrome установлен: $CHROME_VERSION"
    
    # Создаем символическую ссылку, если нужно
    if [ ! -f /usr/bin/google-chrome ]; then
        CHROME_PATH=$(which google-chrome-stable || which google-chrome-beta || which google-chrome-unstable)
        if [ -n "$CHROME_PATH" ]; then
            ln -s "$CHROME_PATH" /usr/bin/google-chrome 2>/dev/null || true
        fi
    fi
else
    echo "⚠ Google Chrome не установлен, используем Chromium..."
    apt-get install -y -qq chromium-browser || apt-get install -y -qq chromium || echo "⚠ Не удалось установить ни Chrome, ни Chromium"
fi

# Устанавливаем Python зависимости
echo "Установка Python зависимостей..."
pip install --upgrade pip
pip install -r requirements.txt

# Устанавливаем ChromeDriver (опционально, webdriver-manager тоже может это сделать)
echo "Проверка ChromeDriver..."
if ! command -v chromedriver &> /dev/null; then
    echo "ChromeDriver не найден в PATH, webdriver-manager установит его автоматически при первом запуске"
else
    CHROMEDRIVER_VERSION=$(chromedriver --version 2>/dev/null || echo "unknown")
    echo "✓ ChromeDriver найден: $CHROMEDRIVER_VERSION"
fi

echo "✓ Сборка завершена успешно!"

