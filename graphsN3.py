import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import numpy as np
import warnings
from itertools import cycle
import textwrap
warnings.filterwarnings('ignore')

# Настройки визуализации
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")

def read_excel_sheets(file_path):
    """
    Чтение всех листов Excel файла
    """
    print(f"Чтение файла: {file_path}")
    
    try:
        # Чтение всех листов
        xls = pd.ExcelFile(file_path)
        sheets_data = {}
        
        for sheet_name in xls.sheet_names:
            print(f"  Обработка листа: {sheet_name}")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            sheets_data[sheet_name] = df
            
        return sheets_data, xls.sheet_names
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return None, None

def get_available_municipalities(sheets_data):
    """
    Получение списка всех доступных муниципальных районов
    """
    municipalities_set = set()
    
    for sheet_name, df in sheets_data.items():
        if 'МР' in df.columns:
            municipalities_set.update(df['МР'].dropna().unique())
    
    return sorted(list(municipalities_set))

def display_municipality_list(municipalities):
    """
    Отображение списка муниципальных районов с нумерацией
    """
    print("\n" + "="*60)
    print("ДОСТУПНЫЕ МУНИЦИПАЛЬНЫЕ РАЙОНЫ:")
    print("="*60)
    
    for i, mun in enumerate(municipalities, 1):
        print(f"{i:3}. {mun}")
    
    print("="*60)

def display_indicators_list(sheet_names):
    """
    Отображение списка показателей с нумерацией
    """
    print("\n" + "="*60)
    print("ДОСТУПНЫЕ ПОКАЗАТЕЛИ:")
    print("="*60)
    
    for i, indicator in enumerate(sheet_names, 1):
        print(f"{i:3}. {indicator}")
    
    print("="*60)

def select_multiple_indicators(sheet_names):
    """
    Выбор нескольких показателей пользователем
    """
    selected_indicators = []
    
    while True:
        try:
            print("\n" + "="*60)
            print("РЕЖИМ ВЫБОРА ПОКАЗАТЕЛЕЙ:")
            print("="*60)
            print("1. Добавить один показатель")
            print("2. Добавить ВСЕ показатели")
            print("3. Добавить несколько показателей по диапазону номеров")
            print("4. Добавить показатели по ключевым словам")
            print("5. Просмотреть выбранные")
            print("6. Удалить показатель")
            print("7. Начать построение графиков")
            print("8. Очистить все")
            print("9. Выход")
            print("="*60)
            
            choice = input("\nВаш выбор: ").strip()
            
            if choice == '1':  # Добавить один показатель
                display_indicators_list(sheet_names)
                indicator_choice = input("\nВведите номер показателя: ").strip()
                
                if indicator_choice.isdigit():
                    idx = int(indicator_choice) - 1
                    if 0 <= idx < len(sheet_names):
                        selected = sheet_names[idx]
                        if selected not in selected_indicators:
                            selected_indicators.append(selected)
                            print(f"✓ Добавлен: {selected}")
                        else:
                            print(f"⚠ {selected} уже в списке")
                    else:
                        print(f"⚠ Ошибка: номер должен быть от 1 до {len(sheet_names)}")
                else:
                    print("⚠ Ошибка: введите число")
            
            elif choice == '2':  # Добавить ВСЕ показатели
                add_all = input("\nВы уверены, что хотите добавить ВСЕ показатели? (да/нет): ").strip().lower()
                if add_all in ['да', 'д', 'yes', 'y']:
                    selected_indicators = sheet_names.copy()
                    print(f"✓ Добавлено {len(selected_indicators)} показателей")
            
            elif choice == '3':  # Добавить по диапазону
                display_indicators_list(sheet_names)
                print("\nВведите диапазон номеров (например: 1-5 или 2,4,6): ")
                range_input = input("Диапазон: ").strip()
                
                added_count = 0
                
                # Обработка диапазона 1-5
                if '-' in range_input:
                    try:
                        start, end = map(int, range_input.split('-'))
                        for idx in range(start-1, end):
                            if 0 <= idx < len(sheet_names):
                                indicator = sheet_names[idx]
                                if indicator not in selected_indicators:
                                    selected_indicators.append(indicator)
                                    added_count += 1
                        print(f"✓ Добавлено {added_count} показателей")
                    except ValueError:
                        print("⚠ Ошибка: неверный формат диапазона")
                
                # Обработка списка номеров 1,3,5
                elif ',' in range_input:
                    try:
                        numbers = map(int, range_input.split(','))
                        for num in numbers:
                            idx = num - 1
                            if 0 <= idx < len(sheet_names):
                                indicator = sheet_names[idx]
                                if indicator not in selected_indicators:
                                    selected_indicators.append(indicator)
                                    added_count += 1
                        print(f"✓ Добавлено {added_count} показателей")
                    except ValueError:
                        print("⚠ Ошибка: неверный формат списка")
            
            elif choice == '4':  # Добавить по ключевым словам
                keyword = input("\nВведите ключевое слово для поиска показателей: ").strip().lower()
                
                matches = [ind for ind in sheet_names if keyword in ind.lower()]
                
                if matches:
                    print(f"\nНайдено {len(matches)} показателей:")
                    for i, ind in enumerate(matches, 1):
                        print(f"  {i}. {ind}")
                    
                    add_all_matches = input("\nДобавить все найденные показатели? (да/нет): ").strip().lower()
                    
                    if add_all_matches in ['да', 'д', 'yes', 'y']:
                        added_count = 0
                        for ind in matches:
                            if ind not in selected_indicators:
                                selected_indicators.append(ind)
                                added_count += 1
                        print(f"✓ Добавлено {added_count} показателей")
                    else:
                        print("Введите номера показателей для добавления (через запятую): ")
                        nums_input = input("Номера: ").strip()
                        
                        if nums_input:
                            try:
                                numbers = map(int, nums_input.split(','))
                                added_count = 0
                                for num in numbers:
                                    idx = num - 1
                                    if 0 <= idx < len(matches):
                                        ind = matches[idx]
                                        if ind not in selected_indicators:
                                            selected_indicators.append(ind)
                                            added_count += 1
                                print(f"✓ Добавлено {added_count} показателей")
                            except ValueError:
                                print("⚠ Ошибка: неверный формат")
                else:
                    print(f"⚠ Показатели по ключевому слову '{keyword}' не найдены")
            
            elif choice == '5':  # Просмотреть выбранные
                if selected_indicators:
                    print("\n" + "="*40)
                    print(f"ВЫБРАННЫЕ ПОКАЗАТЕЛИ ({len(selected_indicators)}):")
                    print("="*40)
                    for i, ind in enumerate(selected_indicators, 1):
                        print(f"{i:3}. {ind}")
                    print("="*40)
                else:
                    print("\n⚠ Список выбранных показателей пуст")
            
            elif choice == '6':  # Удалить показатель
                if selected_indicators:
                    print("\nТекущий список выбранных показателей:")
                    for i, ind in enumerate(selected_indicators, 1):
                        print(f"{i:3}. {ind}")
                    
                    indicator_choice = input("\nВведите номер для удаления (или 0 для отмены): ").strip()
                    
                    if indicator_choice.isdigit():
                        idx = int(indicator_choice) - 1
                        if 0 <= idx < len(selected_indicators):
                            removed = selected_indicators.pop(idx)
                            print(f"✓ Удален: {removed}")
                        elif idx == -1:
                            print("Отмена удаления")
                        else:
                            print(f"⚠ Ошибка: номер должен быть от 1 до {len(selected_indicators)}")
                else:
                    print("\n⚠ Список выбранных показателей пуст")
            
            elif choice == '7':  # Начать построение графиков
                if selected_indicators:
                    return selected_indicators
                else:
                    print("\n⚠ Не выбрано ни одного показателя!")
            
            elif choice == '8':  # Очистить все
                if selected_indicators:
                    confirm = input("\nОчистить ВЕСЬ список выбранных показателей? (да/нет): ").strip().lower()
                    if confirm in ['да', 'д', 'yes', 'y']:
                        selected_indicators.clear()
                        print("\n✓ Список показателей очищен")
                else:
                    print("\n⚠ Список выбранных показателей уже пуст")
            
            elif choice == '9':  # Выход
                return None
            
            else:
                print("\n⚠ Неверный выбор. Попробуйте снова.")
        
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем.")
            return None
        except Exception as e:
            print(f"\n⚠ Ошибка: {e}. Попробуйте снова.")

def select_multiple_municipalities(municipalities):
    """
    Улучшенный выбор нескольких муниципальных районов с массовыми операциями
    """
    selected_municipalities = []
    
    while True:
        try:
            print("\n" + "="*60)
            print("РЕЖИМ ВЫБОРА МУНИЦИПАЛЬНЫХ РАЙОНОВ:")
            print("="*60)
            print("1. Добавить один муниципальный район")
            print("2. Добавить ВСЕ муниципальные районы")
            print("3. Добавить несколько районов по диапазону номеров")
            print("4. Добавить районы по ключевым словам")
            print("5. Просмотреть выбранные")
            print("6. Удалить муниципальный район")
            print("7. Начать построение графиков")
            print("8. Очистить все")
            print("9. Выход")
            print("="*60)
            
            choice = input("\nВаш выбор: ").strip()
            
            if choice == '1':  # Добавить один МР
                print("\nКак добавить муниципальный район:")
                print("1. Ввести номер из списка")
                print("2. Ввести полное название")
                print("3. Вернуться назад")
                
                sub_choice = input("\nВыбор: ").strip()
                
                if sub_choice == '1':
                    display_municipality_list(municipalities)
                    mun_choice = input("\nВведите номер муниципального района: ").strip()
                    
                    if mun_choice.isdigit():
                        idx = int(mun_choice) - 1
                        if 0 <= idx < len(municipalities):
                            selected = municipalities[idx]
                            if selected not in selected_municipalities:
                                selected_municipalities.append(selected)
                                print(f"✓ Добавлен: {selected}")
                            else:
                                print(f"⚠ {selected} уже в списке")
                        else:
                            print(f"⚠ Ошибка: номер должен быть от 1 до {len(municipalities)}")
                    else:
                        print("⚠ Ошибка: введите число")
                
                elif sub_choice == '2':
                    mun_name = input("\nВведите название муниципального района: ").strip()
                    
                    if mun_name in municipalities:
                        if mun_name not in selected_municipalities:
                            selected_municipalities.append(mun_name)
                            print(f"✓ Добавлен: {mun_name}")
                        else:
                            print(f"⚠ {mun_name} уже в списке")
                    else:
                        # Поиск частичного совпадения
                        matches = [m for m in municipalities if mun_name.lower() in m.lower()]
                        
                        if len(matches) == 1:
                            if matches[0] not in selected_municipalities:
                                selected_municipalities.append(matches[0])
                                print(f"✓ Добавлен: {matches[0]}")
                            else:
                                print(f"⚠ {matches[0]} уже в списке")
                        elif len(matches) > 1:
                            print(f"\n⚠ Найдено несколько совпадений:")
                            for i, m in enumerate(matches, 1):
                                print(f"   {i}. {m}")
                            
                            sub_sub_choice = input("Введите номер или 0 для отмены: ").strip()
                            if sub_sub_choice.isdigit():
                                idx = int(sub_sub_choice) - 1
                                if 0 <= idx < len(matches):
                                    if matches[idx] not in selected_municipalities:
                                        selected_municipalities.append(matches[idx])
                                        print(f"✓ Добавлен: {matches[idx]}")
                                    else:
                                        print(f"⚠ {matches[idx]} уже в списке")
                        else:
                            print(f"⚠ Муниципальный район '{mun_name}' не найден")
            
            elif choice == '2':  # Добавить ВСЕ МР
                add_all = input("\nВы уверены, что хотите добавить ВСЕ муниципальные районы? (да/нет): ").strip().lower()
                if add_all in ['да', 'д', 'yes', 'y']:
                    added_count = 0
                    for mun in municipalities:
                        if mun not in selected_municipalities:
                            selected_municipalities.append(mun)
                            added_count += 1
                    print(f"✓ Добавлено {added_count} муниципальных районов")
                    print(f"✓ Всего выбрано: {len(selected_municipalities)}")
            
            elif choice == '3':  # Добавить по диапазону
                display_municipality_list(municipalities)
                print("\nВведите диапазон номеров (например: 1-10 или 5,7,9): ")
                range_input = input("Диапазон: ").strip()
                
                added_count = 0
                
                # Обработка диапазона 1-10
                if '-' in range_input:
                    try:
                        start, end = map(int, range_input.split('-'))
                        for idx in range(start-1, end):
                            if 0 <= idx < len(municipalities):
                                mun = municipalities[idx]
                                if mun not in selected_municipalities:
                                    selected_municipalities.append(mun)
                                    added_count += 1
                        print(f"✓ Добавлено {added_count} муниципальных районов")
                    except ValueError:
                        print("⚠ Ошибка: неверный формат диапазона")
                
                # Обработка списка номеров 1,3,5
                elif ',' in range_input:
                    try:
                        numbers = map(int, range_input.split(','))
                        for num in numbers:
                            idx = num - 1
                            if 0 <= idx < len(municipalities):
                                mun = municipalities[idx]
                                if mun not in selected_municipalities:
                                    selected_municipalities.append(mun)
                                    added_count += 1
                        print(f"✓ Добавлено {added_count} муниципальных районов")
                    except ValueError:
                        print("⚠ Ошибка: неверный формат списка")
            
            elif choice == '4':  # Добавить по ключевым словам
                keyword = input("\nВведите ключевое слово для поиска районов: ").strip().lower()
                
                matches = [m for m in municipalities if keyword in m.lower()]
                
                if matches:
                    print(f"\nНайдено {len(matches)} районов:")
                    for i, m in enumerate(matches, 1):
                        print(f"  {i}. {m}")
                    
                    add_all_matches = input("\nДобавить все найденные районы? (да/нет): ").strip().lower()
                    
                    if add_all_matches in ['да', 'д', 'yes', 'y']:
                        added_count = 0
                        for mun in matches:
                            if mun not in selected_municipalities:
                                selected_municipalities.append(mun)
                                added_count += 1
                        print(f"✓ Добавлено {added_count} районов")
                    else:
                        print("Введите номера районов для добавления (через запятую): ")
                        nums_input = input("Номера: ").strip()
                        
                        if nums_input:
                            try:
                                numbers = map(int, nums_input.split(','))
                                added_count = 0
                                for num in numbers:
                                    idx = num - 1
                                    if 0 <= idx < len(matches):
                                        mun = matches[idx]
                                        if mun not in selected_municipalities:
                                            selected_municipalities.append(mun)
                                            added_count += 1
                                print(f"✓ Добавлено {added_count} районов")
                            except ValueError:
                                print("⚠ Ошибка: неверный формат")
                else:
                    print(f"⚠ Районы по ключевому слову '{keyword}' не найдены")
            
            elif choice == '5':  # Просмотреть выбранные
                if selected_municipalities:
                    print("\n" + "="*40)
                    print(f"ВЫБРАННЫЕ МУНИЦИПАЛЬНЫЕ РАЙОНЫ ({len(selected_municipalities)}):")
                    print("="*40)
                    for i, mun in enumerate(selected_municipalities, 1):
                        print(f"{i:3}. {mun}")
                    print("="*40)
                    
                    # Статистика
                    print(f"\nСтатистика:")
                    print(f"• Выбрано: {len(selected_municipalities)} из {len(municipalities)}")
                    print(f"• Осталось: {len(municipalities) - len(selected_municipalities)}")
                else:
                    print("\n⚠ Список выбранных муниципальных районов пуст")
            
            elif choice == '6':  # Удалить МР
                if selected_municipalities:
                    print("\nТекущий список выбранных районов:")
                    for i, mun in enumerate(selected_municipalities, 1):
                        print(f"{i:3}. {mun}")
                    
                    print("\nВарианты удаления:")
                    print("1. Удалить по номеру")
                    print("2. Удалить по названию")
                    print("3. Удалить все выбранные")
                    print("4. Отмена")
                    
                    del_choice = input("\nВыбор: ").strip()
                    
                    if del_choice == '1':
                        mun_choice = input("Введите номер для удаления: ").strip()
                        
                        if mun_choice.isdigit():
                            idx = int(mun_choice) - 1
                            if 0 <= idx < len(selected_municipalities):
                                removed = selected_municipalities.pop(idx)
                                print(f"✓ Удален: {removed}")
                            else:
                                print(f"⚠ Ошибка: номер должен быть от 1 до {len(selected_municipalities)}")
                    
                    elif del_choice == '2':
                        mun_name = input("Введите название для удаления: ").strip()
                        
                        if mun_name in selected_municipalities:
                            selected_municipalities.remove(mun_name)
                            print(f"✓ Удален: {mun_name}")
                        else:
                            print(f"⚠ Район '{mun_name}' не найден в выбранных")
                    
                    elif del_choice == '3':
                        confirm = input("Удалить ВСЕ выбранные районы? (да/нет): ").strip().lower()
                        if confirm in ['да', 'д', 'yes', 'y']:
                            selected_municipalities.clear()
                            print("✓ Все районы удалены")
                else:
                    print("\n⚠ Список выбранных муниципальных районов пуст")
            
            elif choice == '7':  # Начать построение графиков
                if selected_municipalities:
                    if len(selected_municipalities) > 10:
                        print(f"\n⚠ Внимание: выбрано {len(selected_municipalities)} районов!")
                        print("Это может привести к перегруженным графикам.")
                        confirm = input("Продолжить? (да/нет): ").strip().lower()
                        if confirm not in ['да', 'д', 'yes', 'y']:
                            continue
                    
                    return selected_municipalities
                else:
                    print("\n⚠ Не выбрано ни одного муниципального района!")
            
            elif choice == '8':  # Очистить все
                if selected_municipalities:
                    confirm = input("\nОчистить ВЕСЬ список выбранных районов? (да/нет): ").strip().lower()
                    if confirm in ['да', 'д', 'yes', 'y']:
                        selected_municipalities.clear()
                        print("\n✓ Список выбранных муниципальных районов очищен")
                else:
                    print("\n⚠ Список выбранных муниципальных районов уже пуст")
            
            elif choice == '9':  # Выход
                return None
            
            else:
                print("\n⚠ Неверный выбор. Попробуйте снова.")
        
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем.")
            return None
        except Exception as e:
            print(f"\n⚠ Ошибка: {e}. Попробуйте снова.")

def clean_data(df):
    """
    Очистка данных: переименование столбцов, обработка пропусков
    """
    # Переименование столбцов для удобства
    if len(df.columns) >= 3:
        df.columns = ['МР', 'Год', 'Значение'] + list(df.columns[3:])
    
    # Преобразование года в целое число (если он в формате float)
    if 'Год' in df.columns:
        df['Год'] = df['Год'].astype(str).str.replace('.0', '').astype(int)
    
    # Удаление строк с пропущенными значениями
    df = df.dropna(subset=['Значение'])
    
    return df

def format_number(value):
    """
    Форматирование чисел для отображения на графиках
    """
    if abs(value) >= 1_000_000_000:
        return f'{value/1_000_000_000:.1f}B'
    elif abs(value) >= 1_000_000:
        return f'{value/1_000_000:.1f}M'
    elif abs(value) >= 1000:
        return f'{value/1000:.0f}K'
    else:
        return f'{value:.0f}'

def plot_single_indicator(sheets_data, indicator_name, municipalities_list):
    """
    Построение графика для одного показателя со всеми районами
    С улучшенной легендой
    """
    df = sheets_data[indicator_name].copy()
    df = clean_data(df)
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # Цветовая палитра для районов
    if len(municipalities_list) <= 10:
        colors = plt.cm.Set2(np.linspace(0, 1, len(municipalities_list)))
    elif len(municipalities_list) <= 20:
        colors = plt.cm.tab20(np.linspace(0, 1, len(municipalities_list)))
    else:
        colors = plt.cm.gist_ncar(np.linspace(0, 1, len(municipalities_list)))
    
    # Создаем легенду
    legend_handles = []
    legend_labels = []
    
    # Строим линии для каждого муниципального района
    for mun_idx, municipality in enumerate(municipalities_list):
        municipality_data = df[df['МР'] == municipality]
        
        if not municipality_data.empty:
            municipality_data = municipality_data.sort_values('Год')
            
            # Используем полное название для легенды при одном показателе
            short_name = (municipality.replace(' муниципальный район', '')
                                   .replace(' национальный', '')
                                   .replace(' эвенкийский', '')
                                   .replace(' (долгано-эвенкийский)', '')
                                   .replace(' городской округ', '')
                                   .replace(' республики', '')
                                   .replace(' Республики Саха (Якутия)', ''))
            
            # Рисуем линию
            line, = ax.plot(municipality_data['Год'], municipality_data['Значение'], 
                           marker='o', linewidth=2, markersize=5, 
                           color=colors[mun_idx % len(colors)], alpha=0.8)
            
            legend_handles.append(line)
            legend_labels.append(f"{short_name}")
    
    # Настройка графика
    ax.set_xlabel('Год', fontsize=12)
    ax.set_ylabel('Значение', fontsize=12)
    
    # Обрезаем длинное название листа для заголовка
    if len(indicator_name) > 80:
        title = indicator_name[:77] + '...'
    else:
        title = indicator_name
    title = '\n'.join(textwrap.wrap(title, 70))
    
    # Добавляем информацию о количестве районов на графике
    num_plotted = len([m for m in municipalities_list if not df[df['МР'] == m].empty])
    ax.set_title(f'{title}\n(районов на графике: {num_plotted}/{len(municipalities_list)})', 
                fontsize=14, fontweight='bold', pad=15)
    
    ax.grid(True, alpha=0.3, linestyle='--')
    
    # Настройка формата оси Y для больших чисел
    y_max = 0
    for municipality in municipalities_list:
        municipality_data = df[df['МР'] == municipality]
        if not municipality_data.empty:
            y_max = max(y_max, municipality_data['Значение'].max())
    
    if y_max > 1000000:
        ax.ticklabel_format(axis='y', style='sci', scilimits=(6, 6))
    elif y_max > 1000:
        ax.ticklabel_format(axis='y', style='sci', scilimits=(3, 3))
    
    # При одном показателе всегда показываем полную легенду
    # Размещаем легенду в зависимости от количества районов
    if len(municipalities_list) <= 15:
        # Для небольшого количества районов - сбоку
        ax.legend(legend_handles, legend_labels, fontsize=9, 
                 loc='upper left', bbox_to_anchor=(1.02, 1),
                 borderaxespad=0., framealpha=0.9, ncol=1)
        plt.subplots_adjust(right=0.78)
    elif len(municipalities_list) <= 30:
        # Для среднего количества - снизу в несколько колонок
        ax.legend(legend_handles, legend_labels, fontsize=8, 
                 loc='upper center', bbox_to_anchor=(0.5, -0.15),
                 borderaxespad=0., framealpha=0.9, ncol=3)
        plt.subplots_adjust(bottom=0.25)
    else:
        # Для многих районов - справа в несколько колонок
        ncols_legend = 2 if len(municipalities_list) <= 40 else 3
        ax.legend(legend_handles, legend_labels, fontsize=7, 
                 loc='upper left', bbox_to_anchor=(1.02, 1),
                 borderaxespad=0., framealpha=0.9, ncol=ncols_legend)
        plt.subplots_adjust(right=0.75 if ncols_legend == 2 else 0.65)
    
    # Сохранение графика
    filename = f"показатель_{indicator_name[:30]}_{len(municipalities_list)}районов.png"
    filename = filename.replace('/', '_').replace('\\', '_').replace(':', '_')
    plt.savefig(filename, dpi=150, bbox_inches='tight')
    print(f"✓ График сохранен как: {filename}")
    
    plt.tight_layout()
    plt.show()

def plot_selected_indicators(sheets_data, selected_indicators, municipalities_list):
    """
    Построение графиков для выбранных показателей
    Каждый показатель на своем графике, все районы на одном графике
    """
    n_indicators = len(selected_indicators)
    
    # Если выбран только один показатель, используем специальную функцию
    if n_indicators == 1:
        print(f"\nПостроение графика для одного показателя: {selected_indicators[0]}")
        plot_single_indicator(sheets_data, selected_indicators[0], municipalities_list)
        return
    
    # Создаем сетку графиков (максимум 6 графиков на странице)
    graphs_per_page = 6
    num_pages = (n_indicators + graphs_per_page - 1) // graphs_per_page
    
    for page in range(num_pages):
        start_idx = page * graphs_per_page
        end_idx = min((page + 1) * graphs_per_page, n_indicators)
        page_indicators = selected_indicators[start_idx:end_idx]
        
        # Определяем размер сетки
        n_cols = 2
        n_rows = (len(page_indicators) + n_cols - 1) // n_cols
        
        fig, axes = plt.subplots(n_rows, n_cols, figsize=(18, 5 * n_rows))
        
        # Преобразуем axes в плоский список
        if n_rows > 1 and n_cols > 1:
            axes = axes.flatten()
        elif n_rows == 1 and n_cols > 1:
            axes = axes
        else:
            axes = [axes]
        
        # Общий заголовок
        if num_pages > 1:
            fig.suptitle(f'Выбранные показатели для {len(municipalities_list)} муниципальных районов\n'
                        f'Страница {page + 1} из {num_pages}', 
                        fontsize=16, fontweight='bold', y=1.02)
        else:
            fig.suptitle(f'Выбранные показатели для {len(municipalities_list)} муниципальных районов', 
                        fontsize=16, fontweight='bold', y=1.02)
        
        # Цветовая палитра для районов
        if len(municipalities_list) <= 10:
            colors = plt.cm.Set2(np.linspace(0, 1, len(municipalities_list)))
        elif len(municipalities_list) <= 20:
            colors = plt.cm.tab20(np.linspace(0, 1, len(municipalities_list)))
        else:
            colors = plt.cm.gist_ncar(np.linspace(0, 1, len(municipalities_list)))
        
        for idx, sheet_name in enumerate(page_indicators):
            ax = axes[idx]
            df = sheets_data[sheet_name].copy()
            df = clean_data(df)
            
            # Создаем легенду
            legend_handles = []
            legend_labels = []
            
            # Строим линии для каждого муниципального района
            for mun_idx, municipality in enumerate(municipalities_list):
                municipality_data = df[df['МР'] == municipality]
                
                if not municipality_data.empty:
                    municipality_data = municipality_data.sort_values('Год')
                    
                    # Сокращаем название для легенды
                    short_name = (municipality.replace(' муниципальный район', '')
                                           .replace(' национальный', '')
                                           .replace(' эвенкийский', '')
                                           .replace(' (долгано-эвенкийский)', '')
                                           .replace(' городской округ', '')
                                           .replace(' республики', '')
                                           .replace(' Республики Саха (Якутия)', '')[:15])
                    
                    # Рисуем линию
                    line, = ax.plot(municipality_data['Год'], municipality_data['Значение'], 
                                   marker='o', linewidth=1.5, markersize=3, 
                                   color=colors[mun_idx % len(colors)], alpha=0.7)
                    
                    legend_handles.append(line)
                    legend_labels.append(f"{short_name}")
            
            # Настройка графика
            ax.set_xlabel('Год', fontsize=10)
            ax.set_ylabel('Значение', fontsize=10)
            
            # Обрезаем длинное название листа для заголовка
            if len(sheet_name) > 50:
                title = sheet_name[:47] + '...'
            else:
                title = sheet_name
            title = '\n'.join(textwrap.wrap(title, 50))
            
            # Добавляем информацию о количестве районов на графике
            num_plotted = len([m for m in municipalities_list if not df[df['МР'] == m].empty])
            ax.set_title(f'{title}\n(районов на графике: {num_plotted}/{len(municipalities_list)})', 
                        fontsize=11, fontweight='bold', pad=10)
            
            ax.grid(True, alpha=0.3, linestyle='--')
            
            # Настройка формата оси Y для больших чисел
            y_max = 0
            for municipality in municipalities_list:
                municipality_data = df[df['МР'] == municipality]
                if not municipality_data.empty:
                    y_max = max(y_max, municipality_data['Значение'].max())
            
            if y_max > 1000000:
                ax.ticklabel_format(axis='y', style='sci', scilimits=(6, 6))
            elif y_max > 1000:
                ax.ticklabel_format(axis='y', style='sci', scilimits=(3, 3))
            
            # Добавляем легенду снаружи графика (только если районов немного)
            if len(municipalities_list) <= 12:
                ax.legend(legend_handles, legend_labels, fontsize=7, 
                         loc='upper left', bbox_to_anchor=(1.02, 1),
                         borderaxespad=0., framealpha=0.8, ncol=1)
            elif len(municipalities_list) <= 20:
                # Для среднего количества районов - компактная легенда
                ax.legend(legend_handles, legend_labels, fontsize=6, 
                         loc='upper left', bbox_to_anchor=(1.02, 1),
                         borderaxespad=0., framealpha=0.8, ncol=2)
            else:
                # Для многих районов показываем упрощенную информацию
                ax.text(0.02, 0.98, f'{len(municipalities_list)} районов', 
                       transform=ax.transAxes, fontsize=8,
                       verticalalignment='top',
                       bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        
        # Удаляем пустые оси
        for idx in range(len(page_indicators), len(axes)):
            fig.delaxes(axes[idx])
        
        plt.tight_layout()
        plt.subplots_adjust(top=0.92, right=0.85 if len(municipalities_list) <= 20 else 0.95)
        
        # Сохранение графика
        if num_pages > 1:
            filename = f"выбранные_показатели_страница_{page+1}_{len(municipalities_list)}районов.png"
        else:
            filename = f"выбранные_показатели_{len(municipalities_list)}районов.png"
        
        plt.savefig(filename, dpi=150, bbox_inches='tight')
        print(f"✓ Страница {page+1} сохранена как: {filename}")
        
        plt.show()
        
        # Статистика по странице
        print(f"\nСтраница {page + 1}/{num_pages}:")
        for indicator in page_indicators:
            print(f"  - {indicator}")
        
        if page < num_pages - 1:
            continue_choice = input("\nНажмите Enter для следующей страницы или 'q' для выхода: ").strip()
            if continue_choice.lower() in ['q', 'quit', 'выход']:
                break

def plot_all_indicators_separate(sheets_data, sheet_names, municipalities_list):
    """
    Построение отдельных графиков для каждого показателя со всеми выбранными районами вместе
    Каждый показатель на своем графике, все районы на одном графике
    """
    plot_selected_indicators(sheets_data, sheet_names, municipalities_list)

def plot_combined_timeseries(sheets_data, sheet_names, municipalities_list, mode='all_separate'):
    """
    Построение графиков временных рядов для нескольких муниципальных районов
    
    Параметры:
    - mode: 'separate' - отдельные графики для каждого листа (старая функция)
            'combined' - все показатели на одном графике (только для 1-3 МР)
            'all_separate' - каждый показатель отдельно, все районы вместе (новая функция)
    """
    
    if mode == 'all_separate':
        # Новый режим: каждый показатель на отдельном графике, все районы вместе
        plot_all_indicators_separate(sheets_data, sheet_names, municipalities_list)
        return
    
    if mode == 'combined' and len(municipalities_list) > 3:
        print("\n⚠ Для режима сравнения на одном графике рекомендуется не более 3 районов.")
        print("  Автоматически переключаюсь на режим отдельных графиков.")
        mode = 'separate'
    
    if mode == 'separate':
        # Отдельные графики для каждого показателя
        n_sheets = min(len(sheet_names), 6)  # Максимум 6 графиков на одной фигуре
        
        fig, axes = plt.subplots(3, 2, figsize=(18, 14))
        axes = axes.flatten()
        
        fig.suptitle(f'Сравнение показателей для {len(municipalities_list)} муниципальных районов', 
                    fontsize=18, fontweight='bold', y=1.02)
        
        colors = plt.cm.Set2(np.linspace(0, 1, len(municipalities_list)))
        
        for idx, sheet_name in enumerate(sheet_names):
            if idx >= n_sheets:
                break
                
            ax = axes[idx]
            df = sheets_data[sheet_name].copy()
            df = clean_data(df)
            
            # Строим линии для каждого муниципального района
            for mun_idx, municipality in enumerate(municipalities_list):
                municipality_data = df[df['МР'] == municipality]
                
                if not municipality_data.empty:
                    municipality_data = municipality_data.sort_values('Год')
                    
                    # Сокращаем название для легенды
                    short_name = municipality.replace(' муниципальный район', '').replace(' национальный', '').replace(' эвенкийский', '').replace(' (долгано-эвенкийский)', '')[:20]
                    
                    ax.plot(municipality_data['Год'], municipality_data['Значение'], 
                           marker='o', linewidth=2, markersize=5, 
                           color=colors[mun_idx], label=short_name)
            
            # Настройка графика
            ax.set_xlabel('Год', fontsize=10)
            ax.set_ylabel('Значение', fontsize=10)
            ax.set_title(f'{sheet_name[:50]}{"..." if len(sheet_name) > 50 else ""}', 
                        fontsize=12, fontweight='bold')
            ax.grid(True, alpha=0.3)
            ax.legend(fontsize=9, loc='best')
            
            # Форматирование оси Y для больших чисел
            y_max = 0
            for municipality in municipalities_list:
                municipality_data = df[df['МР'] == municipality]
                if not municipality_data.empty:
                    y_max = max(y_max, municipality_data['Значение'].max())
            
            if y_max > 10000:
                ax.ticklabel_format(axis='y', style='sci', scilimits=(3, 3))
        
        # Удаляем пустые подграфики
        for idx in range(len(sheet_names), len(axes)):
            fig.delaxes(axes[idx])
        
        plt.tight_layout()
        plt.show()
    
    else:  # mode == 'combined'
        # Все показатели на одном графике (для 1-3 МР)
        fig, axes = plt.subplots(len(municipalities_list), 1, 
                                figsize=(16, 5 * len(municipalities_list)))
        
        if len(municipalities_list) == 1:
            axes = [axes]
        
        fig.suptitle('Сравнение всех показателей по муниципальным районам', 
                    fontsize=18, fontweight='bold', y=1.02)
        
        # Цвета для разных показателей
        colors = cycle(plt.cm.tab10(np.linspace(0, 1, len(sheet_names))))
        
        for mun_idx, municipality in enumerate(municipalities_list):
            ax = axes[mun_idx]
            
            # Собираем все данные для этого МР
            all_data = {}
            
            for sheet_name in sheet_names:
                df = sheets_data[sheet_name].copy()
                df = clean_data(df)
                municipality_data = df[df['МР'] == municipality]
                
                if not municipality_data.empty:
                    municipality_data = municipality_data.sort_values('Год')
                    all_data[sheet_name] = {
                        'years': municipality_data['Год'].tolist(),
                        'values': municipality_data['Значение'].tolist()
                    }
            
            # Строим линии для каждого показателя
            for sheet_name, color in zip(sheet_names, colors):
                if sheet_name in all_data:
                    short_sheet_name = sheet_name[:30] + ('...' if len(sheet_name) > 30 else '')
                    ax.plot(all_data[sheet_name]['years'], all_data[sheet_name]['values'],
                           marker='o', linewidth=2, markersize=6,
                           color=color, label=short_sheet_name)
            
            # Настройка графика
            ax.set_xlabel('Год', fontsize=11)
            ax.set_ylabel('Значение', fontsize=11)
            ax.set_title(municipality, fontsize=14, fontweight='bold')
            ax.grid(True, alpha=0.3)
            ax.legend(fontsize=9, loc='upper left', bbox_to_anchor=(1.02, 1))
        
        plt.tight_layout()
        plt.show()
    
    # Вывод сводной таблицы
    print("\n" + "="*80)
    print("СВОДНАЯ ТАБЛИЦА ДАННЫХ:")
    print("="*80)
    
    for municipality in municipalities_list:
        print(f"\n{municipality}:")
        print("-" * 80)
        
        for sheet_name in sheet_names:
            df = sheets_data[sheet_name].copy()
            df = clean_data(df)
            municipality_data = df[df['МР'] == municipality]
            
            if not municipality_data.empty:
                # Вычисляем статистику
                years = municipality_data['Год'].unique()
                values = municipality_data['Значение']
                
                print(f"{sheet_name[:50]:50} | Годы: {len(years):2} | "
                      f"Min: {format_number(values.min()):>8} | "
                      f"Max: {format_number(values.max()):>8} | "
                      f"Avg: {format_number(values.mean()):>8}")

def main():
    """
    Основная функция программы
    """
    print("="*60)
    print("АНАЛИЗ И СРАВНЕНИЕ ПОКАЗАТЕЛЕЙ МУНИЦИПАЛЬНЫХ РАЙОНОВ")
    print("РЕСПУБЛИКИ САХА (ЯКУТИЯ)")
    print("="*60)
    
    # Укажите путь к вашему файлу
    file_path = "Реакция экономики на шоки на примере муниципальных образований Республики Саха (3).xlsx"
    
    # Проверяем существование файла
    if not Path(file_path).exists():
        print(f"⚠ Файл '{file_path}' не найден!")
        print("Пожалуйста, укажите правильный путь к файлу.")
        file_path = input("Введите путь к файлу Excel: ").strip()
    
    # Читаем данные из Excel
    sheets_data, sheet_names = read_excel_sheets(file_path)
    
    if sheets_data is None:
        print("Не удалось прочитать данные из файла. Программа завершена.")
        return
    
    # Получаем список всех муниципальных районов
    municipalities = get_available_municipalities(sheets_data)
    
    if not municipalities:
        print("В файле не найдены данные о муниципальных районах.")
        return
    
    print(f"\nЗагружено данных: {len(sheet_names)} листов")
    print(f"Найдено муниципальных районов: {len(municipalities)}")
    
    # Основной цикл программы
    while True:
        try:
            print("\n" + "="*60)
            print("ГЛАВНОЕ МЕНЮ:")
            print("="*60)
            print("1. Просмотреть список всех муниципальных районов")
            print("2. Просмотреть список всех показателей")
            print("3. Выбрать муниципальные районы для сравнения")
            print("4. КАЖДЫЙ ПОКАЗАТЕЛЬ ОТДЕЛЬНО, ВСЕ РАЙОНЫ ВМЕСТЕ (быстрый режим)")
            print("5. ВЫБРАТЬ ОТДЕЛЬНЫЕ ПОКАЗАТЕЛИ (новый режим)")
            print("6. Выход")
            print("="*60)
            
            choice = input("\nВаш выбор: ").strip()
            
            if choice == '1':
                display_municipality_list(municipalities)
                input("\nНажмите Enter для продолжения...")
            
            elif choice == '2':
                display_indicators_list(sheet_names)
                input("\nНажмите Enter для продолжения...")
            
            elif choice == '3':
                # Выбор муниципальных районов с улучшенным интерфейсом
                selected_municipalities = select_multiple_municipalities(municipalities)
                
                if selected_municipalities is None:
                    continue
                
                if not selected_municipalities:
                    print("\nНе выбрано ни одного муниципального района.")
                    continue
                
                # Выбор режима отображения
                print("\n" + "="*60)
                print("РЕЖИМ ОТОБРАЖЕНИЯ ГРАФИКОВ:")
                print("="*60)
                print("1. Отдельные графики для каждого показателя (старый режим)")
                print("2. Все показатели на одном графике (только для 1-3 районов)")
                print("3. КАЖДЫЙ ПОКАЗАТЕЛЬ ОТДЕЛЬНО, ВСЕ РАЙОНЫ ВМЕСТЕ (новый режим)")
                print("="*60)
                
                display_mode = input("\nВаш выбор: ").strip()
                
                if display_mode == '1':
                    mode = 'separate'
                elif display_mode == '2':
                    mode = 'combined'
                elif display_mode == '3':
                    mode = 'all_separate'
                else:
                    print("\n⚠ Неверный выбор. Использую новый режим.")
                    mode = 'all_separate'
                
                # Построение графиков
                print(f"\nПостроение графиков для {len(selected_municipalities)} муниципальных районов...")
                print(f"Районы: {', '.join(selected_municipalities[:5])}{'...' if len(selected_municipalities) > 5 else ''}")
                
                plot_combined_timeseries(sheets_data, sheet_names, selected_municipalities, mode)
            
            elif choice == '4':
                # Быстрый режим: все районы сразу в новом формате
                print("\n" + "="*60)
                print("БЫСТРЫЙ РЕЖИМ: КАЖДЫЙ ПОКАЗАТЕЛЬ ОТДЕЛЬНО, ВСЕ РАЙОНЫ ВМЕСТЕ")
                print("="*60)
                
                # Опции для быстрого режима
                print("\nВыберите опцию:")
                print("1. Использовать ВСЕ муниципальные районы")
                print("2. Выбрать определенные районы")
                print("3. Вернуться в главное меню")
                
                quick_choice = input("\nВаш выбор: ").strip()
                
                if quick_choice == '1':
                    # Использовать все районы
                    selected_municipalities = municipalities.copy()
                    print(f"\n✓ Выбраны ВСЕ {len(selected_municipalities)} муниципальных районов")
                    
                    # Предупреждение если районов много
                    if len(selected_municipalities) > 15:
                        print("\n⚠ Внимание: выбрано много районов!")
                        print("Графики могут быть перегружены.")
                        confirm = input("Продолжить? (да/нет): ").strip().lower()
                        if confirm not in ['да', 'д', 'yes', 'y']:
                            continue
                
                elif quick_choice == '2':
                    # Выбрать определенные районы
                    selected_municipalities = select_multiple_municipalities(municipalities)
                    if not selected_municipalities or selected_municipalities is None:
                        continue
                
                elif quick_choice == '3':
                    continue
                else:
                    print("\n⚠ Неверный выбор. Возврат в главное меню.")
                    continue
                
                # Сразу строим графики в новом режиме
                print(f"\nПостроение графиков для {len(selected_municipalities)} муниципальных районов...")
                plot_combined_timeseries(sheets_data, sheet_names, selected_municipalities, 'all_separate')
            
            elif choice == '5':
                # Новый режим: выбор отдельных показателей
                print("\n" + "="*60)
                print("РЕЖИМ: ВЫБОР ОТДЕЛЬНЫХ ПОКАЗАТЕЛЕЙ")
                print("="*60)
                
                # Сначала выбираем районы
                print("\n1. Сначала выберите муниципальные районы:")
                selected_municipalities = select_multiple_municipalities(municipalities)
                if not selected_municipalities or selected_municipalities is None:
                    continue
                
                # Затем выбираем показатели
                print("\n2. Теперь выберите показатели для анализа:")
                selected_indicators = select_multiple_indicators(sheet_names)
                if not selected_indicators or selected_indicators is None:
                    continue
                
                # Построение графиков для выбранных показателей
                print(f"\nПостроение графиков для {len(selected_municipalities)} районов")
                print(f"по {len(selected_indicators)} выбранным показателям...")
                
                plot_selected_indicators(sheets_data, selected_indicators, selected_municipalities)
            
            elif choice == '6':
                print("\nПрограмма завершена. До свидания!")
                break
            
            else:
                print("\n⚠ Неверный выбор. Попробуйте снова.")
        
        except KeyboardInterrupt:
            print("\n\nПрограмма прервана пользователем.")
            break
        except Exception as e:
            print(f"\n⚠ Ошибка: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    main()