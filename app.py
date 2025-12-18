# -*- coding: utf-8 -*-
"""
Flask приложение для веб-интерфейса мониторинга успеваемости
"""
from flask import Flask, render_template, jsonify, request, send_file
import threading
import time
import os
import json
import re
from datetime import datetime
from pathlib import Path
import sys

# Импортируем наши модули
from mektep_scraper import MektepScraper
from process_quarters_final import process_success_data
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'

# Глобальное состояние
scraper_state = {
    'running': False,
    'progress': 0,
    'current_step': None,
    'message': 'Готов к запуску',
    'error': None,
    'waiting_for_school': False,
    'waiting_for_class': False,
    'schools': [],
    'classes': [],
    'selected_school': None,
    'selected_class': None,
    'scraper': None,
    'thread': None,
    'logs': [],
    'auth_start_time': None  # Время начала ожидания авторизации
}

# Папка для файлов
OUTPUT_DIR = Path(__file__).parent
UPLOADS_DIR = OUTPUT_DIR / 'uploads'
FILES_DIR = UPLOADS_DIR

# Создаем папку uploads, если её нет
UPLOADS_DIR.mkdir(exist_ok=True)


def add_log(source, message, level='info'):
    """Добавление лога"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    log_entry = {
        'timestamp': timestamp,
        'source': source,
        'message': message,
        'level': level
    }
    scraper_state['logs'].append(log_entry)
    # Ограничиваем количество логов
    if len(scraper_state['logs']) > 1000:
        scraper_state['logs'] = scraper_state['logs'][-1000:]


def cleanup_session_files():
    """Очистка файлов сессии (промежуточный и конечный Excel)"""
    try:
        deleted_files = []
        
        # Удаляем промежуточный файл
        intermediate_file = UPLOADS_DIR / 'success_data.xlsx'
        if intermediate_file.exists():
            intermediate_file.unlink()
            deleted_files.append(intermediate_file.name)
            add_log('SYSTEM', f'Удален промежуточный файл: {intermediate_file.name}', 'info')
        
        # Удаляем все конечные файлы (обработанные)
        for file_path in UPLOADS_DIR.glob('*.xlsx'):
            if file_path.is_file() and file_path.name != 'success_data.xlsx':
                file_path.unlink()
                deleted_files.append(file_path.name)
                add_log('SYSTEM', f'Удален файл: {file_path.name}', 'info')
        
        if deleted_files:
            add_log('SYSTEM', f'Очищено файлов сессии: {len(deleted_files)}', 'success')
        else:
            add_log('SYSTEM', 'Файлы для очистки не найдены', 'info')
        
        return len(deleted_files)
    except Exception as e:
        add_log('SYSTEM', f'Ошибка при очистке файлов: {str(e)}', 'error')
        return 0


def run_scraper():
    """Запуск скрапера в отдельном потоке"""
    try:
        scraper_state['running'] = True
        scraper_state['progress'] = 0
        scraper_state['error'] = None
        scraper_state['current_step'] = 'Инициализация'
        scraper_state['message'] = 'Запуск скрапера...'
        
        add_log('SCRAPER', 'Запуск скрапера', 'info')
        
        # Создаем экземпляр скрапера
        scraper = MektepScraper()
        scraper.setup_driver()
        scraper_state['scraper'] = scraper
        
        # Запускаем основной процесс
        scraper_state['current_step'] = 'Авторизация'
        scraper_state['message'] = 'Ожидание авторизации в браузере...'
        scraper_state['progress'] = 10
        scraper_state['auth_start_time'] = time.time()  # Запоминаем время начала ожидания
        
        # Открываем страницу и ждем авторизации
        if not scraper.login():
            scraper_state['error'] = 'Не удалось авторизоваться'
            scraper_state['running'] = False
            scraper_state['auth_start_time'] = None
            return
        
        scraper_state['auth_start_time'] = None  # Сбрасываем после успешной авторизации
        
        scraper_state['progress'] = 20
        scraper_state['current_step'] = 'Навигация'
        scraper_state['message'] = 'Переход на страницу отчетов...'
        
        # Переходим на страницу отчетов
        if not scraper.navigate_to_reports():
            scraper_state['error'] = 'Не удалось перейти на страницу отчетов'
            scraper_state['running'] = False
            return
        
        scraper_state['progress'] = 30
        scraper_state['current_step'] = 'Загрузка школ'
        scraper_state['message'] = 'Загрузка списка школ...'
        
        # Получаем список школ
        schools = scraper.get_schools_list()
        if not schools:
            scraper_state['error'] = 'Не удалось загрузить список школ'
            scraper_state['running'] = False
            return
        
        # Форматируем школы для фронтенда
        formatted_schools = []
        for idx, school in enumerate(schools, 1):
            formatted_schools.append({
                'number': idx,
                'name': school.get('name', 'Неизвестная школа')
            })
        
        scraper_state['schools'] = formatted_schools
        scraper_state['waiting_for_school'] = True
        scraper_state['message'] = 'Выберите школу из списка'
        scraper_state['progress'] = 40
        
        add_log('SCRAPER', f'Найдено школ: {len(formatted_schools)}', 'success')
        
        # Ждем выбора школы
        timeout = 300  # 5 минут
        start_time = time.time()
        while scraper_state['waiting_for_school'] and scraper_state['running']:
            if time.time() - start_time > timeout:
                scraper_state['error'] = 'Превышено время ожидания выбора школы'
                scraper_state['running'] = False
                return
            time.sleep(0.5)
        
        if not scraper_state['running']:
            return
        
        selected_school = scraper_state['selected_school']
        if not selected_school:
            scraper_state['error'] = 'Школа не выбрана'
            scraper_state['running'] = False
            return
        
        scraper_state['progress'] = 50
        scraper_state['current_step'] = 'Переход к школе'
        scraper_state['message'] = f'Переход к школе: {selected_school["name"]}'
        
        # Переходим к выбранной школе (используем номер из списка школ, 1-based)
        school_index = selected_school['number']  # Номер школы (1-based)
        if not scraper.select_school(school_index):
            scraper_state['error'] = 'Не удалось перейти к выбранной школе'
            scraper_state['running'] = False
            return
        
        scraper_state['progress'] = 60
        scraper_state['current_step'] = 'Загрузка классов'
        scraper_state['message'] = 'Загрузка списка классов...'
        
        # Получаем список классов (вкладок)
        class_tabs = scraper.get_classes_list()
        if not class_tabs:
            scraper_state['error'] = 'Не удалось загрузить список классов'
            scraper_state['running'] = False
            return
        
        # Форматируем вкладки классов для фронтенда
        formatted_classes = []
        for idx, tab in enumerate(class_tabs):
            formatted_classes.append({
                'number': idx + 1,
                'name': f"{tab['number']} класс",  # Например, "11 класс"
                'grade': tab['number'],  # Номер класса для выбора вкладки
                'text': tab.get('text', f"{tab['number']} класс")
            })
        
        scraper_state['classes'] = formatted_classes
        scraper_state['waiting_for_class'] = True
        scraper_state['message'] = 'Выберите класс из списка'
        scraper_state['progress'] = 70
        
        add_log('SCRAPER', f'Найдено классов: {len(formatted_classes)}', 'success')
        
        # Ждем выбора класса (вкладки)
        timeout = 300  # 5 минут
        start_time = time.time()
        while scraper_state['waiting_for_class'] and scraper_state['running']:
            if time.time() - start_time > timeout:
                scraper_state['error'] = 'Превышено время ожидания выбора класса'
                scraper_state['running'] = False
                return
            time.sleep(0.5)
        
        if not scraper_state['running']:
            return
        
        selected_class = scraper_state['selected_class']
        if not selected_class:
            scraper_state['error'] = 'Класс не выбран'
            scraper_state['running'] = False
            return
        
        scraper_state['progress'] = 80
        scraper_state['current_step'] = 'Обработка данных'
        scraper_state['message'] = f'Обработка класса: {selected_class["name"]}'
        
        # Выбираем вкладку класса
        class_grade = selected_class.get('grade')
        if not class_grade:
            scraper_state['error'] = 'Не указан номер класса'
            scraper_state['running'] = False
            return
        
        if not scraper.select_class_tab(class_grade):
            scraper_state['error'] = f'Не удалось выбрать вкладку класса {class_grade}'
            scraper_state['running'] = False
            return
        
        # Получаем список групп классов для выбранной вкладки
        class_groups = scraper.get_class_groups_from_table()
        if not class_groups:
            scraper_state['error'] = 'Не удалось загрузить список групп классов'
            scraper_state['running'] = False
            return
        
        add_log('SCRAPER', f'Найдено групп в классе {class_grade}: {len(class_groups)}', 'success')
        
        # Выбираем класс и обрабатываем данные
        output_file = str(UPLOADS_DIR / 'success_data.xlsx')
        
        # Получаем список групп классов для выбранной вкладки
        class_groups = scraper.get_class_groups_from_table()
        if not class_groups:
            scraper_state['error'] = 'Не удалось загрузить список групп классов'
            scraper_state['running'] = False
            return
        
        # Обрабатываем все классы параллели
        total_groups = len(class_groups)
        for group_idx, group in enumerate(class_groups):
            if not scraper_state['running']:
                break
            
            scraper_state['message'] = f'Обработка класса: {group["name"]} ({group_idx + 1}/{total_groups})'
            scraper_state['progress'] = 80 + int((group_idx + 1) / total_groups * 10)
            
            # Кликаем по кнопке "Успеваемость" для текущего класса
            if not group.get('button'):
                add_log('SCRAPER', f'Кнопка не найдена для {group["name"]}, пропускаем', 'warning')
                continue
            
            try:
                # ШАГ 1: Убеждаемся, что предыдущее модальное окно закрыто (если есть)
                if group_idx > 0:
                    max_wait = 10
                    wait_count = 0
                    while scraper.is_modal_open() and wait_count < max_wait:
                        add_log('SCRAPER', f'Ожидание закрытия предыдущего модального окна... ({wait_count + 1}/{max_wait})', 'info')
                        time.sleep(0.5)
                        wait_count += 1
                        # Пробуем закрыть принудительно
                        if wait_count > 5:
                            scraper.close_modal()
                    
                    if scraper.is_modal_open():
                        add_log('SCRAPER', f'Предыдущее модальное окно не закрылось, пропускаем {group["name"]}', 'warning')
                        continue
                
                # ШАГ 2: Получаем свежую ссылку на кнопку (может стать устаревшей после закрытия модального окна)
                button = None
                try:
                    # Пробуем использовать сохраненную кнопку
                    button = group['button']
                    # Проверяем, что кнопка еще валидна
                    _ = button.is_displayed()
                except Exception:
                    # Кнопка стала устаревшей, переполучаем её
                    add_log('SCRAPER', f'Переполучение кнопки для {group["name"]}...', 'info')
                    try:
                        # Ищем кнопку заново по имени класса
                        # Ищем строку таблицы с нужным классом
                        tables = scraper.driver.find_elements(By.CSS_SELECTOR, "table.table-striped, table.table-bordered")
                        for table in tables:
                            rows = table.find_elements(By.TAG_NAME, "tr")
                            for row in rows:
                                cells = row.find_elements(By.TAG_NAME, "td")
                                if cells:
                                    # Проверяем, содержит ли первая ячейка название класса
                                    if group['name'] in cells[0].text:
                                        # Ищем кнопку в последней ячейке (столбец "Действия")
                                        try:
                                            button = cells[-1].find_element(By.TAG_NAME, "button")
                                            break
                                        except:
                                            pass
                                if button:
                                    break
                            if button:
                                break
                    except Exception as e:
                        add_log('SCRAPER', f'Не удалось найти кнопку для {group["name"]}: {str(e)}', 'warning')
                        continue
                
                if not button:
                    add_log('SCRAPER', f'Кнопка не найдена для {group["name"]}, пропускаем', 'warning')
                    continue
                
                # ШАГ 3: Прокручиваем к кнопке и убеждаемся, что она видна
                try:
                    scraper.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", button)
                    time.sleep(0.5)
                    
                    # Проверяем, что кнопка видна после прокрутки
                    if not button.is_displayed():
                        add_log('SCRAPER', f'Кнопка для {group["name"]} все еще не видна, пропускаем', 'warning')
                        continue
                except Exception as e:
                    add_log('SCRAPER', f'Ошибка при работе с кнопкой для {group["name"]}: {str(e)}', 'warning')
                    continue
                
                # ШАГ 4: Открываем модальное окно
                scraper.driver.execute_script("arguments[0].click();", button)
                add_log('SCRAPER', f'Клик по кнопке для {group["name"]}', 'info')
                time.sleep(0.5)  # Уменьшена пауза
                
                # ШАГ 5: Убеждаемся, что модальное окно открылось
                try:
                    # Ждем появления модального окна
                    scraper.wait.until(
                        EC.any_of(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "#classSapa.modal.show")),
                            EC.presence_of_element_located((By.CSS_SELECTOR, "#classSapa.modal:not(.fade)")),
                            EC.presence_of_element_located((By.ID, "classSapa"))
                        )
                    )
                    time.sleep(0.5)  # Уменьшена пауза
                    
                    # Проверяем, что модальное окно действительно открыто
                    if not scraper.is_modal_open():
                        add_log('SCRAPER', f'Модальное окно не открылось для {group["name"]}', 'warning')
                        continue
                    
                    # Ждем появления таблицы
                    try:
                        scraper.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "#classSapa table"))
                        )
                        # Ждем, пока таблица полностью загрузится (проверяем наличие строк)
                        scraper.wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#classSapa table tbody tr")) > 0)
                        time.sleep(0.3)  # Минимальная пауза для стабильности
                        add_log('SCRAPER', f'Модальное окно открыто и таблица загружена для {group["name"]}', 'success')
                    except TimeoutException:
                        add_log('SCRAPER', f'Таблица не появилась для {group["name"]}, но продолжаем...', 'warning')
                except TimeoutException:
                    add_log('SCRAPER', f'Модальное окно не найдено для {group["name"]}', 'warning')
                    continue
                
                # ШАГ 6: Извлекаем данные из модального окна
                table_data = scraper.extract_modal_table_data()
                if not table_data:
                    add_log('SCRAPER', f'Не удалось извлечь данные для {group["name"]}', 'warning')
                    scraper.close_modal()
                    # Убеждаемся, что закрылось
                    time.sleep(1)
                    if scraper.is_modal_open():
                        add_log('SCRAPER', f'Не удалось закрыть модальное окно для {group["name"]}', 'error')
                    continue
                
                # ШАГ 7: Сохраняем данные в Excel
                class_name = group['name']
                if scraper.save_to_excel(table_data, class_name, output_file):
                    add_log('SCRAPER', f'Данные для {class_name} сохранены', 'success')
                else:
                    add_log('SCRAPER', f'Не удалось сохранить данные для {class_name}', 'error')
                
                # ШАГ 8: Закрываем модальное окно
                if scraper.close_modal():
                    add_log('SCRAPER', f'Модальное окно закрыто для {group["name"]}', 'info')
                else:
                    add_log('SCRAPER', f'Не удалось закрыть модальное окно для {group["name"]}', 'warning')
                
                # ШАГ 9: Убеждаемся, что модальное окно закрыто перед следующим
                max_wait = 10
                wait_count = 0
                while scraper.is_modal_open() and wait_count < max_wait:
                    time.sleep(0.3)
                    wait_count += 1
                
                if scraper.is_modal_open():
                    add_log('SCRAPER', f'Модальное окно все еще открыто для {group["name"]}, принудительно закрываем', 'warning')
                    # Принудительное закрытие
                    try:
                        scraper.driver.execute_script("""
                            var modal = document.getElementById('classSapa');
                            if (modal) {
                                modal.classList.remove('show');
                                modal.style.display = 'none';
                            }
                            var backdrop = document.querySelector('.modal-backdrop');
                            if (backdrop) {
                                backdrop.remove();
                            }
                            document.body.classList.remove('modal-open');
                        """)
                    except:
                        pass
                
                # Пауза перед следующим классом (уменьшена для ускорения)
                time.sleep(0.5)
                
            except Exception as e:
                add_log('SCRAPER', f'Ошибка при обработке {group["name"]}: {str(e)}', 'error')
                import traceback
                add_log('SCRAPER', f'Трассировка: {traceback.format_exc()}', 'error')
                # Пытаемся закрыть модальное окно при ошибке
                try:
                    scraper.close_modal()
                    time.sleep(1)
                except:
                    pass
                continue
        
        scraper_state['progress'] = 90
        scraper_state['current_step'] = 'Обработка файлов'
        scraper_state['message'] = 'Обработка данных по четвертям...'
        
        # Обрабатываем файл через process_quarters_final
        # Определяем имя класса из выбранного класса
        # Извлекаем номер класса из grade (например, "11" -> "11 класс")
        class_name = f"{class_grade} класс"
        
        success, processed_file = process_success_data(
            input_file=output_file,
            output_file=None,  # Будет определено внутри функции
            class_name=class_name,
            output_dir=str(UPLOADS_DIR)  # Передаем папку для сохранения
        )
        
        if success and processed_file:
            scraper_state['progress'] = 100
            scraper_state['current_step'] = 'Завершено'
            scraper_state['message'] = 'Данные успешно обработаны!'
            add_log('SCRAPER', f'Файл сохранен: {processed_file}', 'success')
            
            # Очищаем промежуточный файл после успешной обработки
            intermediate_file = UPLOADS_DIR / 'success_data.xlsx'
            if intermediate_file.exists():
                try:
                    intermediate_file.unlink()
                    add_log('SYSTEM', 'Промежуточный файл удален', 'info')
                except Exception as e:
                    add_log('SYSTEM', f'Не удалось удалить промежуточный файл: {str(e)}', 'warning')
        else:
            scraper_state['error'] = 'Ошибка при обработке данных'
            add_log('SCRAPER', 'Ошибка при обработке данных', 'error')
        
        scraper_state['running'] = False
        
        # Закрываем браузер
        try:
            if scraper and scraper.driver:
                scraper.driver.quit()
        except Exception as e:
            add_log('SCRAPER', f'Ошибка при закрытии браузера: {str(e)}', 'warning')
        
    except Exception as e:
        scraper_state['error'] = str(e)
        scraper_state['running'] = False
        add_log('SCRAPER', f'Ошибка: {str(e)}', 'error')
        import traceback
        traceback.print_exc()
        
        # Закрываем браузер при ошибке
        try:
            scraper = scraper_state.get('scraper')
            if scraper and scraper.driver:
                scraper.driver.quit()
        except Exception as e:
            add_log('SCRAPER', f'Ошибка при закрытии браузера: {str(e)}', 'warning')


@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')


@app.route('/api/test')
def api_test():
    """Тестовый endpoint"""
    return jsonify({'status': 'ok', 'message': 'API работает'})


@app.route('/api/status/scraper')
def api_status_scraper():
    """Статус скрапера"""
    # Вычисляем время ожидания авторизации, если идет процесс авторизации
    auth_wait_time = None
    if scraper_state['auth_start_time'] is not None:
        elapsed = int(time.time() - scraper_state['auth_start_time'])
        auth_wait_time = elapsed
    
    return jsonify({
        'running': scraper_state['running'],
        'progress': scraper_state['progress'],
        'current_step': scraper_state['current_step'],
        'message': scraper_state['message'],
        'error': scraper_state['error'],
        'waiting_for_school': scraper_state['waiting_for_school'],
        'waiting_for_class': scraper_state['waiting_for_class'],
        'schools': scraper_state['schools'],
        'classes': scraper_state['classes'],
        'auth_wait_time': auth_wait_time  # Время ожидания авторизации в секундах
    })


@app.route('/api/start/scraper', methods=['POST'])
def api_start_scraper():
    """Запуск скрапера"""
    if scraper_state['running']:
        return jsonify({'error': 'Скрапер уже запущен'}), 400
    
    # Запускаем в отдельном потоке
    thread = threading.Thread(target=run_scraper, daemon=True)
    thread.start()
    scraper_state['thread'] = thread
    
    return jsonify({'status': 'started'})


@app.route('/api/stop/scraper', methods=['POST'])
def api_stop_scraper():
    """Остановка скрапера"""
    scraper_state['running'] = False
    scraper_state['waiting_for_school'] = False
    scraper_state['waiting_for_class'] = False
    
    # Закрываем браузер, если открыт
    if scraper_state['scraper']:
        try:
            scraper_state['scraper'].driver.quit()
        except:
            pass
    
    # Очищаем файлы сессии
    cleanup_session_files()
    
    return jsonify({'status': 'stopped'})


@app.route('/api/select/school', methods=['POST'])
def api_select_school():
    """Выбор школы"""
    data = request.get_json()
    school_number = data.get('school_number')
    
    if not school_number:
        return jsonify({'error': 'Не указан номер школы'}), 400
    
    schools = scraper_state['schools']
    selected_school = None
    for school in schools:
        if school['number'] == school_number:
            selected_school = school
            break
    
    if not selected_school:
        return jsonify({'error': 'Школа не найдена'}), 404
    
    scraper_state['selected_school'] = selected_school
    scraper_state['waiting_for_school'] = False
    
    add_log('SYSTEM', f'Выбрана школа: {selected_school["name"]}', 'success')
    
    return jsonify({'status': 'ok', 'school': selected_school})


@app.route('/api/select/class', methods=['POST'])
def api_select_class():
    """Выбор класса"""
    data = request.get_json()
    class_name = data.get('class_name')
    
    if not class_name:
        return jsonify({'error': 'Не указано имя класса'}), 400
    
    classes = scraper_state['classes']
    selected_class = None
    for idx, cls in enumerate(classes):
        if cls['name'] == class_name:
            selected_class = {**cls, 'index': idx}
            break
    
    if not selected_class:
        return jsonify({'error': 'Класс не найден'}), 404
    
    scraper_state['selected_class'] = selected_class
    scraper_state['waiting_for_class'] = False
    
    add_log('SYSTEM', f'Выбран класс: {selected_class["name"]}', 'success')
    
    return jsonify({'status': 'ok', 'class_name': selected_class['name']})


@app.route('/api/files')
def api_files():
    """Список файлов"""
    files = []
    
    # Ищем все .xlsx файлы в директории
    for file_path in FILES_DIR.glob('*.xlsx'):
        if file_path.is_file():
            stat = file_path.stat()
            files.append({
                'name': file_path.name,
                'size': stat.st_size,
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            })
    
    # Сортируем по дате изменения (новые первыми)
    files.sort(key=lambda x: x['modified'], reverse=True)
    
    return jsonify({'files': files})


@app.route('/api/download/<filename>')
def api_download(filename):
    """Скачивание файла"""
    file_path = FILES_DIR / filename
    
    if not file_path.exists() or not file_path.is_file():
        return jsonify({'error': 'Файл не найден'}), 404
    
    return send_file(str(file_path), as_attachment=True)


@app.route('/api/logs')
def api_logs():
    """Получение логов"""
    return jsonify({'logs': scraper_state['logs']})


@app.route('/api/reset', methods=['POST'])
def api_reset():
    """Сброс состояния"""
    scraper_state['running'] = False
    scraper_state['progress'] = 0
    scraper_state['current_step'] = None
    scraper_state['message'] = 'Готов к запуску'
    scraper_state['error'] = None
    scraper_state['waiting_for_school'] = False
    scraper_state['waiting_for_class'] = False
    scraper_state['schools'] = []
    scraper_state['classes'] = []
    scraper_state['selected_school'] = None
    scraper_state['selected_class'] = None
    scraper_state['logs'] = []
    scraper_state['auth_start_time'] = None
    
    # Закрываем браузер, если открыт
    if scraper_state['scraper']:
        try:
            scraper_state['scraper'].driver.quit()
        except:
            pass
    scraper_state['scraper'] = None
    
    # Очищаем файлы сессии
    cleanup_session_files()
    
    add_log('SYSTEM', 'Состояние сброшено', 'info')
    
    return jsonify({'status': 'reset'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("="*70)
    print("Запуск Flask приложения")
    print("="*70)
    print(f"Откройте в браузере: http://localhost:{port}")
    print("="*70)
    
    app.run(debug=False, host='0.0.0.0', port=port)
