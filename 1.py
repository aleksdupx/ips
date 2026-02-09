import pandas as pd
import numpy as np

# Укажите путь к вашему файлу Excel
file_path = "Реакция экономики на шоки на примере муниципальных образований Республики Саха (3).xlsx"

# Загружаем все листы Excel
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

print("Анализ пропусков данных по датам и муниципальным районам")
print("=" * 80)

for sheet_name in sheet_names:
    print(f"\nЛист: {sheet_name}")
    print("-" * 40)
    
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Проверяем структуру данных
    print(f"Структура данных:")
    print(f"  Столбцы: {list(df.columns)}")
    print(f"  Всего записей: {len(df)}")
    
    # Определяем названия столбцов
    # Ищем столбцы с годами и МР (предполагаем, что есть столбцы 'МР', 'Год', 'Значение')
    if 'Год' in df.columns and 'МР' in df.columns:
        # Удаляем '.0' из годов если они в формате float
        df['Год'] = df['Год'].astype(str).str.replace('.0', '', regex=False)
        
        # Получаем уникальные годы и муниципальные районы
        all_years = sorted(df['Год'].dropna().unique())
        all_municipalities = sorted(df['МР'].dropna().unique())
        
        print(f"  Уникальных лет: {len(all_years)} (от {min(all_years)} до {max(all_years)})")
        print(f"  Уникальных МР: {len(all_municipalities)}")
        
        # Создаем полную сетку всех возможных комбинаций год-МР
        full_grid = pd.MultiIndex.from_product([all_years, all_municipalities], names=['Год', 'МР'])
        
        # Создаем индекс из существующих данных
        existing_index = pd.MultiIndex.from_frame(df[['Год', 'МР']])
        
        # Находим пропущенные комбинации
        missing_combinations = full_grid.difference(existing_index)
        
        if len(missing_combinations) > 0:
            print(f"\nПропущенные данные (год-МР): {len(missing_combinations)} комбинаций")
            
            # Группируем пропуски по годам
            print("\nПропуски по годам:")
            missing_by_year = {}
            for year, mr in missing_combinations:
                if year not in missing_by_year:
                    missing_by_year[year] = []
                missing_by_year[year].append(mr)
            
            for year in sorted(missing_by_year.keys()):
                print(f"  {year}: {len(missing_by_year[year])} МР без данных")
                # Можно вывести список МР для каждого года:
                # print(f"    МР без данных: {', '.join(missing_by_year[year][:5])}" + 
                #       ("..." if len(missing_by_year[year]) > 5 else ""))
            
            # Группируем пропуски по МР
            print("\nПропуски по муниципальным районам:")
            missing_by_mr = {}
            for year, mr in missing_combinations:
                if mr not in missing_by_mr:
                    missing_by_mr[mr] = []
                missing_by_mr[mr].append(year)
            
            # Сортируем по количеству пропусков
            sorted_mr = sorted(missing_by_mr.items(), key=lambda x: len(x[1]), reverse=True)
            
            for mr, years in sorted_mr[:10]:  # Показываем топ-10
                print(f"  {mr}: {len(years)} лет без данных")
                print(f"    Годы: {', '.join(sorted(years)[:10])}" + 
                      ("..." if len(years) > 10 else ""))
            
            if len(sorted_mr) > 10:
                print(f"  ... и еще {len(sorted_mr) - 10} МР с пропусками")
            
            # Сохраняем детальную информацию о пропусках в файл
            missing_df = pd.DataFrame(list(missing_combinations), columns=['Год', 'МР'])
            missing_df.to_excel(f"пропуски_{sheet_name[:30]}.xlsx", index=False)
            print(f"\nДетальная информация о пропусках сохранена в файл: пропуски_{sheet_name[:30]}.xlsx")
            
        else:
            print("\n✓ Все возможные комбинации год-МР присутствуют в данных")
        
        # Анализ пропусков значений в столбце 'Значение' (если он есть)
        if 'Значение' in df.columns:
            missing_values = df[df['Значение'].isna()]
            if len(missing_values) > 0:
                print(f"\nПропущенные значения в столбце 'Значение': {len(missing_values)} записей")
                print("Первые 5 записей с пропущенными значениями:")
                print(missing_values[['Год', 'МР', 'Значение']].head().to_string(index=False))
            else:
                print(f"\n✓ В столбце 'Значение' нет пропусков")
    
    else:
        print("  Внимание: Структура данных не соответствует ожидаемой (нет столбцов 'Год' и/или 'МР')")
        print(f"  Фактические столбцы: {list(df.columns)}")

print("\n" + "=" * 80)
print("Анализ завершен")
# Сводный отчет по всем листам
print("\n" + "=" * 80)
print("СВОДНЫЙ ОТЧЕТ ПО ВСЕМ ЛИСТАМ")
print("=" * 80)

all_missing_info = []

for sheet_name in sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    if 'Год' in df.columns and 'МР' in df.columns:
        df['Год'] = df['Год'].astype(str).str.replace('.0', '', regex=False)
        
        all_years = sorted(df['Год'].dropna().unique())
        all_municipalities = sorted(df['МР'].dropna().unique())
        
        full_grid = pd.MultiIndex.from_product([all_years, all_municipalities], names=['Год', 'МР'])
        existing_index = pd.MultiIndex.from_frame(df[['Год', 'МР']])
        missing_combinations = full_grid.difference(existing_index)
        
        # Значения
        missing_values_count = df['Значение'].isna().sum() if 'Значение' in df.columns else 0
        
        all_missing_info.append({
            'Лист': sheet_name,
            'Всего записей': len(df),
            'Уникальных лет': len(all_years),
            'Уникальных МР': len(all_municipalities),
            'Пропущено комбинаций год-МР': len(missing_combinations),
            'Пропущено значений': missing_values_count,
            'Покрытие данных': f"{(1 - len(missing_combinations)/len(full_grid))*100:.1f}%" if len(full_grid) > 0 else "0%"
        })

# Создаем сводную таблицу
summary_df = pd.DataFrame(all_missing_info)
print("\nСводная статистика по пропускам:")
print(summary_df.to_string(index=False))