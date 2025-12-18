// Глобальные переменные
let currentStep = 0;  // Начинаем с шага 0 (авторизация)
let statusInterval = null;
let logsInterval = null;
let logsExpanded = false;
let authTimerInterval = null;  // Интервал для таймера авторизации
let credentialsSaved = false;  // Флаг сохранения учетных данных

// Инициализация
document.addEventListener('DOMContentLoaded', function() {
    console.log('Инициализация приложения...');
    
    // Проверка API
    fetch('/api/test')
        .then(response => response.json())
        .then(data => {
            console.log('API подключен:', data);
            checkCredentialsStatus();
            updateStatus();
            loadLogs();
            
            // Устанавливаем интервалы
            statusInterval = setInterval(updateStatus, 2000);
            logsInterval = setInterval(loadLogs, 3000);
        })
        .catch(error => {
            console.error('Ошибка подключения к API:', error);
        });
});

// Проверка статуса учетных данных
async function checkCredentialsStatus() {
    try {
        const response = await fetch('/api/credentials');
        const data = await response.json();
        
        if (data.has_credentials) {
            credentialsSaved = true;
            showStep(1);  // Переходим к шагу запуска
            const statusBox = document.getElementById('credentials-status');
            if (statusBox) {
                statusBox.style.display = 'flex';
            }
        } else {
            showStep(0);  // Показываем форму авторизации
        }
    } catch (error) {
        console.error('Ошибка при проверке учетных данных:', error);
        showStep(0);  // В случае ошибки показываем форму
    }
}

// Сохранение учетных данных
async function saveCredentials(event) {
    event.preventDefault();
    
    const login = document.getElementById('login-input').value.trim();
    const password = document.getElementById('password-input').value.trim();
    
    if (!login || !password) {
        alert('Пожалуйста, заполните все поля');
        return;
    }
    
    const saveBtn = document.getElementById('save-credentials-btn');
    if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Сохранение...';
    }
    
    try {
        const response = await fetch('/api/credentials', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ login, password })
        });
        
        const result = await response.json();
        
        if (response.ok) {
            credentialsSaved = true;
            addLog('SYSTEM', 'Учетные данные сохранены', 'success');
            
            const statusBox = document.getElementById('credentials-status');
            if (statusBox) {
                statusBox.style.display = 'flex';
                statusBox.className = 'status-box success';
            }
            
            // Переходим к следующему шагу через небольшую задержку
            setTimeout(() => {
                showStep(1);
            }, 500);
        } else {
            alert('Ошибка: ' + (result.error || 'Не удалось сохранить учетные данные'));
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="fas fa-save"></i> Сохранить и продолжить';
            }
        }
    } catch (error) {
        console.error('Ошибка при сохранении:', error);
        alert('Ошибка при сохранении учетных данных');
        if (saveBtn) {
            saveBtn.disabled = false;
            saveBtn.innerHTML = '<i class="fas fa-save"></i> Сохранить и продолжить';
        }
    }
}

// Обновление статуса
async function updateStatus() {
    try {
        const response = await fetch('/api/status/scraper');
        if (!response.ok) return;
        
        const status = await response.json();
        
        // Обновляем прогресс
        const progressContainer = document.getElementById('progress-container');
        const progressFill = document.getElementById('progress-fill');
        const progressText = document.getElementById('progress-text');
        const statusBox = document.getElementById('status-box');
        
        if (status.running && progressContainer) {
            progressContainer.style.display = 'block';
            if (progressFill) {
                progressFill.style.width = (status.progress || 0) + '%';
            }
            if (progressText) {
                progressText.textContent = (status.progress || 0) + '%';
            }
        }
        
        // Обновляем статус бокс и элементы управления
        const startBtn = document.getElementById('start-btn');
        const restartBtn = document.getElementById('restart-btn');
        const authWaiting = document.getElementById('auth-waiting');
        const authTimer = document.getElementById('auth-timer');
        
        if (statusBox) {
            if (status.error) {
                statusBox.className = 'status-box error';
                statusBox.innerHTML = `<i class="fas fa-exclamation-circle"></i><span>Ошибка: ${status.error}</span>`;
                if (startBtn) startBtn.disabled = false;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'none';
            } else if (status.current_step === 'Авторизация' && status.running) {
                // Показываем инструкции по авторизации
                statusBox.className = 'status-box warning';
                statusBox.innerHTML = `<i class="fas fa-key"></i><span>Ожидание авторизации в браузере...</span>`;
                if (startBtn) startBtn.disabled = true;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'block';
                
                // Обновляем таймер ожидания
                if (authTimer) {
                    // Используем время с сервера, если доступно
                    if (status.auth_wait_time !== null && status.auth_wait_time !== undefined) {
                        authTimer.textContent = status.auth_wait_time;
                    } else {
                        // Если сервер не отправляет время, используем локальный таймер
                        if (!authTimerInterval) {
                            let seconds = 0;
                            authTimerInterval = setInterval(() => {
                                seconds++;
                                if (authTimer) {
                                    authTimer.textContent = seconds;
                                }
                            }, 1000);
                        }
                    }
                }
            } else if (status.waiting_for_school) {
                statusBox.className = 'status-box warning';
                statusBox.innerHTML = `<i class="fas fa-hand-pointer"></i><span>Ожидание выбора школы...</span>`;
                if (startBtn) startBtn.disabled = true;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
                if (status.schools) {
                    showStep(2);
                    displaySchools(status.schools);
                }
            } else if (status.waiting_for_class) {
                statusBox.className = 'status-box warning';
                statusBox.innerHTML = `<i class="fas fa-hand-pointer"></i><span>Ожидание выбора класса...</span>`;
                if (startBtn) startBtn.disabled = true;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
                if (status.classes) {
                    showStep(3);
                    displayClasses(status.classes);
                }
            } else if (status.running) {
                statusBox.className = 'status-box';
                statusBox.innerHTML = `<i class="fas fa-spinner fa-spin"></i><span>${status.message || 'Выполняется...'}</span>`;
                if (startBtn) startBtn.disabled = true;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
            } else if (status.progress === 100 && !status.running) {
                statusBox.className = 'status-box success';
                statusBox.innerHTML = `<i class="fas fa-check-circle"></i><span>Процесс завершен!</span>`;
                if (startBtn) startBtn.disabled = true;
                if (restartBtn) restartBtn.style.display = 'inline-flex';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
                showStep(4);
                loadFiles();
            } else if (!status.running && status.progress === 0) {
                // Если процесс не запущен и прогресс 0 - показываем готовность
                statusBox.className = 'status-box';
                statusBox.innerHTML = `<i class="fas fa-info-circle"></i><span>${status.message || 'Готов к запуску'}</span>`;
                if (startBtn) startBtn.disabled = false;
                if (restartBtn) restartBtn.style.display = 'none';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
                // Сбрасываем таймер на 0
                const authTimer = document.getElementById('auth-timer');
                if (authTimer) {
                    authTimer.textContent = '0';
                }
                if (currentStep !== 1 && credentialsSaved) {
                    showStep(1);
                } else if (!credentialsSaved) {
                    showStep(0);
                }
            } else {
                statusBox.className = 'status-box';
                statusBox.innerHTML = `<i class="fas fa-info-circle"></i><span>${status.message || 'Готов к запуску'}</span>`;
                if (startBtn) startBtn.disabled = false;
                if (restartBtn) restartBtn.style.display = 'none';
                if (authWaiting) authWaiting.style.display = 'none';
                // Останавливаем таймер авторизации
                if (authTimerInterval) {
                    clearInterval(authTimerInterval);
                    authTimerInterval = null;
                }
                // Сбрасываем таймер на 0
                const authTimer = document.getElementById('auth-timer');
                if (authTimer) {
                    authTimer.textContent = '0';
                }
            }
        }
        
    } catch (error) {
        console.error('Ошибка при обновлении статуса:', error);
    }
}

// Запуск процесса
async function startProcess() {
    try {
        // Проверяем наличие учетных данных перед запуском
        if (!credentialsSaved) {
            const response = await fetch('/api/credentials');
            const data = await response.json();
            if (!data.has_credentials) {
                alert('Пожалуйста, сначала введите учетные данные');
                showStep(0);
                return;
            }
            credentialsSaved = true;
        }
        
        const startBtn = document.getElementById('start-btn');
        if (startBtn) {
            startBtn.disabled = true;
        }
        
        const response = await fetch('/api/start/scraper', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        });
        
        const result = await response.json();
        
        if (response.ok) {
            addLog('SYSTEM', 'Процесс запущен', 'info');
            showStep(1);
        } else {
            alert('Ошибка: ' + (result.error || 'Не удалось запустить процесс'));
            if (startBtn) {
                startBtn.disabled = false;
            }
        }
    } catch (error) {
        console.error('Ошибка при запуске:', error);
        alert('Ошибка при запуске процесса');
        const startBtn = document.getElementById('start-btn');
        if (startBtn) {
            startBtn.disabled = false;
        }
    }
}

// Отображение шагов
function showStep(stepNumber) {
    // Скрываем все шаги
    for (let i = 0; i <= 4; i++) {
        const step = document.getElementById(`step-${i}`);
        if (step) {
            step.classList.remove('active');
            if (i < stepNumber) {
                step.classList.add('completed');
            } else {
                step.classList.remove('completed');
            }
        }
    }
    
    // Показываем нужный шаг
    const currentStepEl = document.getElementById(`step-${stepNumber}`);
    if (currentStepEl) {
        currentStepEl.classList.add('active');
        currentStep = stepNumber;
    }
}

// Отображение школ
function displaySchools(schools) {
    const loading = document.getElementById('schools-loading');
    const list = document.getElementById('schools-list');
    
    if (loading) loading.style.display = 'none';
    if (!list) return;
    
    list.innerHTML = schools.map(school => `
        <div class="selection-item" onclick="selectSchool(${school.number})">
            <span class="selection-item-number">№ ${school.number}</span>
            <span class="selection-item-name">${escapeHtml(school.name)}</span>
        </div>
    `).join('');
}

// Выбор школы
async function selectSchool(schoolNumber) {
    try {
        const response = await fetch('/api/select/school', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ school_number: schoolNumber })
        });
        
        const result = await response.json();
        
        if (response.ok) {
            addLog('SYSTEM', `Выбрана школа: ${result.school.name}`, 'success');
            
            // Подсвечиваем выбранный элемент
            const items = document.querySelectorAll('#schools-list .selection-item');
            items.forEach((item, index) => {
                item.classList.toggle('selected', index === schoolNumber - 1);
            });
        } else {
            alert('Ошибка: ' + result.error);
        }
    } catch (error) {
        console.error('Ошибка при выборе школы:', error);
        alert('Ошибка при выборе школы');
    }
}

// Отображение классов
function displayClasses(classes) {
    const loading = document.getElementById('classes-loading');
    const list = document.getElementById('classes-list');
    
    if (loading) loading.style.display = 'none';
    if (!list) return;
    
    list.innerHTML = classes.map(cls => `
        <div class="selection-item" onclick="selectClass('${escapeHtml(cls.name)}')">
            <span class="selection-item-number">№ ${cls.number}</span>
            <span class="selection-item-name">${escapeHtml(cls.name)}</span>
        </div>
    `).join('');
}

// Выбор класса
async function selectClass(className) {
    try {
        const response = await fetch('/api/select/class', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ class_name: className })
        });
        
        const result = await response.json();
        
        if (response.ok) {
            addLog('SYSTEM', `Выбран класс: ${result.class_name}`, 'success');
            
            // Подсвечиваем выбранный элемент
            const items = document.querySelectorAll('#classes-list .selection-item');
            items.forEach(item => {
                if (item.textContent.includes(className)) {
                    item.classList.add('selected');
                }
            });
        } else {
            alert('Ошибка: ' + result.error);
        }
    } catch (error) {
        console.error('Ошибка при выборе класса:', error);
        alert('Ошибка при выборе класса');
    }
}

// Загрузка файлов
async function loadFiles() {
    try {
        const response = await fetch('/api/files');
        const data = await response.json();
        
        const container = document.getElementById('files-container');
        if (!container) return;
        
        if (data.files && data.files.length > 0) {
            // Разделяем файлы на промежуточные (raw) и обработанные
            // Raw файлы содержат "_raw" в имени или "raw" перед расширением
            const rawFiles = data.files.filter(file => {
                const name = file.name.toLowerCase();
                return name.includes('_raw') || (name.includes('raw') && name.endsWith('.xlsx'));
            });
            const processedFiles = data.files.filter(file => {
                const name = file.name.toLowerCase();
                return !name.includes('_raw') && !(name.includes('raw') && name.endsWith('.xlsx'));
            });
            
            let html = '';
            
            // Секция промежуточных файлов (для проверки корректности)
            if (rawFiles.length > 0) {
                html += `
                    <div class="files-section">
                        <h3 class="files-section-title">
                            <i class="fas fa-file-code"></i> Промежуточные файлы (для проверки корректности)
                        </h3>
                        <p class="files-section-description">
                            Файлы с оригинальной двухуровневой структурой таблицы. Используйте для проверки корректности извлечения данных.
                        </p>
                        ${rawFiles.map(file => `
                            <div class="file-item file-item-raw">
                                <div class="file-info">
                                    <div class="file-name">
                                        <i class="fas fa-file-excel"></i> ${file.name}
                                    </div>
                                    <div class="file-meta">
                                        Размер: ${formatFileSize(file.size)} | Изменен: ${file.modified}
                                    </div>
                                </div>
                                <a href="/api/download/${file.name}" class="btn btn-primary file-download" download>
                                    <i class="fas fa-download"></i> Скачать
                                </a>
                            </div>
                        `).join('')}
                    </div>
                `;
            }
            
            // Секция обработанных файлов
            if (processedFiles.length > 0) {
                html += `
                    <div class="files-section">
                        <h3 class="files-section-title">
                            <i class="fas fa-file-check"></i> Обработанные файлы
                        </h3>
                        <p class="files-section-description">
                            Финальные файлы с обработанными данными, готовые для использования.
                        </p>
                        ${processedFiles.map(file => `
                            <div class="file-item file-item-processed">
                                <div class="file-info">
                                    <div class="file-name">
                                        <i class="fas fa-file-excel"></i> ${file.name}
                                    </div>
                                    <div class="file-meta">
                                        Размер: ${formatFileSize(file.size)} | Изменен: ${file.modified}
                                    </div>
                                </div>
                                <a href="/api/download/${file.name}" class="btn btn-primary file-download" download>
                                    <i class="fas fa-download"></i> Скачать
                                </a>
                            </div>
                        `).join('')}
                    </div>
                `;
            }
            
            // Если файлов нет ни в одной категории
            if (rawFiles.length === 0 && processedFiles.length === 0) {
                html = '<div class="loading-box"><span>Файлы не найдены</span></div>';
            }
            
            container.innerHTML = html;
        } else {
            container.innerHTML = '<div class="loading-box"><span>Файлы не найдены</span></div>';
        }
    } catch (error) {
        console.error('Ошибка при загрузке файлов:', error);
    }
}

// Перезапуск процесса
async function restartProcess() {
    // Если процесс запущен, сначала останавливаем его
    const response = await fetch('/api/status/scraper');
    const status = await response.json();
    
    if (status.running) {
        if (!confirm('Процесс выполняется. Остановить и начать заново? Все несохраненные данные будут потеряны.')) {
            return;
        }
        // Останавливаем процесс
        try {
            await fetch('/api/stop/scraper', { method: 'POST' });
            // Ждем немного, чтобы процесс остановился
            await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (e) {
            console.error('Ошибка при остановке:', e);
        }
    } else {
        if (!confirm('Вы уверены, что хотите начать заново?')) {
            return;
        }
    }
    
    try {
        // Сбрасываем состояние на сервере
        const response = await fetch('/api/reset', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ error: `HTTP ${response.status}` }));
            throw new Error(errorData.error || `Ошибка сброса состояния: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('Состояние сброшено:', result);
        
        // Сбрасываем локальное состояние
        // НЕ сбрасываем учетные данные, только переходим к шагу запуска
        if (credentialsSaved) {
            currentStep = 1;
            showStep(1);
        } else {
            currentStep = 0;
            showStep(0);
        }
        
        // Сбрасываем UI элементы
        const startBtn = document.getElementById('start-btn');
        if (startBtn) {
            startBtn.disabled = false;
        }
        
        const restartBtn = document.getElementById('restart-btn');
        if (restartBtn) {
            restartBtn.style.display = 'none';
        }
        
        const progressContainer = document.getElementById('progress-container');
        if (progressContainer) {
            progressContainer.style.display = 'none';
        }
        
        const progressFill = document.getElementById('progress-fill');
        if (progressFill) {
            progressFill.style.width = '0%';
        }
        
        const progressText = document.getElementById('progress-text');
        if (progressText) {
            progressText.textContent = '0%';
        }
        
        const statusBox = document.getElementById('status-box');
        if (statusBox) {
            statusBox.className = 'status-box';
            statusBox.innerHTML = '<i class="fas fa-info-circle"></i><span>Готов к запуску</span>';
        }
        
        const authWaiting = document.getElementById('auth-waiting');
        if (authWaiting) {
            authWaiting.style.display = 'none';
        }
        
        // Останавливаем таймер авторизации
        if (authTimerInterval) {
            clearInterval(authTimerInterval);
            authTimerInterval = null;
        }
        
        // Сбрасываем таймер на 0
        const authTimer = document.getElementById('auth-timer');
        if (authTimer) {
            authTimer.textContent = '0';
        }
        
        // Очищаем списки выбора
        const schoolsList = document.getElementById('schools-list');
        if (schoolsList) {
            schoolsList.innerHTML = '';
        }
        
        const classesList = document.getElementById('classes-list');
        if (classesList) {
            classesList.innerHTML = '';
        }
        
        const schoolsLoading = document.getElementById('schools-loading');
        if (schoolsLoading) {
            schoolsLoading.style.display = 'none';
        }
        
        const classesLoading = document.getElementById('classes-loading');
        if (classesLoading) {
            classesLoading.style.display = 'none';
        }
        
        // Очищаем контейнер файлов
        const filesContainer = document.getElementById('files-container');
        if (filesContainer) {
            filesContainer.innerHTML = '<div class="loading-box"><i class="fas fa-spinner fa-spin"></i><span>Загрузка списка файлов...</span></div>';
        }
        
        addLog('SYSTEM', 'Состояние сброшено. Готов к новому запуску.', 'info');
        
        // Обновляем статус
        updateStatus();
        
    } catch (error) {
        console.error('Ошибка при сбросе:', error);
        
        // Если endpoint не найден (404), все равно сбрасываем локальное состояние
        if (error.message.includes('404') || error.message.includes('NOT FOUND')) {
            console.warn('Endpoint /api/reset не найден, выполняем локальный сброс');
            
            // Выполняем локальный сброс
            currentStep = 1;
            showStep(1);
            
            const startBtn = document.getElementById('start-btn');
            if (startBtn) {
                startBtn.disabled = false;
            }
            
            const progressContainer = document.getElementById('progress-container');
            if (progressContainer) {
                progressContainer.style.display = 'none';
            }
            
            const progressFill = document.getElementById('progress-fill');
            if (progressFill) {
                progressFill.style.width = '0%';
            }
            
            const progressText = document.getElementById('progress-text');
            if (progressText) {
                progressText.textContent = '0%';
            }
            
        const statusBox = document.getElementById('status-box');
        if (statusBox) {
            statusBox.className = 'status-box';
            statusBox.innerHTML = '<i class="fas fa-info-circle"></i><span>Готов к запуску</span>';
        }
        
        const restartBtn = document.getElementById('restart-btn');
        if (restartBtn) {
            restartBtn.style.display = 'none';
        }
        
        const authWaiting = document.getElementById('auth-waiting');
        if (authWaiting) {
            authWaiting.style.display = 'none';
        }
        
        // Останавливаем таймер авторизации
        if (authTimerInterval) {
            clearInterval(authTimerInterval);
            authTimerInterval = null;
        }
        
        // Сбрасываем таймер на 0
        const authTimer = document.getElementById('auth-timer');
        if (authTimer) {
            authTimer.textContent = '0';
        }
        
        // Очищаем списки
        const schoolsList = document.getElementById('schools-list');
        if (schoolsList) {
            schoolsList.innerHTML = '';
        }
        
        const classesList = document.getElementById('classes-list');
        if (classesList) {
            classesList.innerHTML = '';
        }
        
        addLog('SYSTEM', 'Локальное состояние сброшено. Перезапустите сервер для полного сброса.', 'warning');
        } else {
            alert('Ошибка при сбросе состояния: ' + error.message + '\n\nПопробуйте перезагрузить страницу (F5)');
        }
    }
}

// Логи
async function loadLogs() {
    try {
        const response = await fetch('/api/logs');
        const data = await response.json();
        
        const container = document.getElementById('logs-container');
        if (!container) return;
        
        if (data.logs && data.logs.length > 0) {
            const recentLogs = data.logs.slice(-30);
            container.innerHTML = recentLogs.map(log => {
                const levelClass = log.level || 'info';
                const sourceClass = log.source.toLowerCase();
                return `
                    <div class="log-entry ${levelClass}">
                        <span class="log-time">${log.timestamp || '--:--:--'}</span>
                        <span class="log-source ${sourceClass}">${log.source || 'SYSTEM'}</span>
                        <span class="log-message">${escapeHtml(log.message || '')}</span>
                    </div>
                `;
            }).join('');
            
            if (logsExpanded) {
                container.scrollTop = container.scrollHeight;
            }
        }
    } catch (error) {
        console.error('Ошибка при загрузке логов:', error);
    }
}

function toggleLogs() {
    logsExpanded = !logsExpanded;
    const content = document.getElementById('logs-content');
    const chevron = document.getElementById('logs-chevron');
    
    if (content) {
        content.classList.toggle('expanded', logsExpanded);
    }
    if (chevron) {
        chevron.style.transform = logsExpanded ? 'rotate(180deg)' : 'rotate(0deg)';
    }
}

function addLog(source, message, level = 'info') {
    const container = document.getElementById('logs-container');
    if (!container) return;
    
    const timestamp = new Date().toLocaleTimeString('ru-RU');
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${level}`;
    logEntry.innerHTML = `
        <span class="log-time">${timestamp}</span>
        <span class="log-source ${source.toLowerCase()}">${source}</span>
        <span class="log-message">${escapeHtml(message)}</span>
    `;
    
    container.appendChild(logEntry);
    if (logsExpanded) {
        container.scrollTop = container.scrollHeight;
    }
}

// Вспомогательные функции
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}
