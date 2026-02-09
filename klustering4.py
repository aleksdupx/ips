import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import numpy as np
import warnings
from itertools import cycle
import textwrap
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans, AgglomerativeClustering, DBSCAN, SpectralClustering, Birch
from sklearn.decomposition import PCA
from sklearn.manifold import TSNE
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score, silhouette_samples
from sklearn.mixture import GaussianMixture
from scipy.cluster.hierarchy import dendrogram, linkage
from scipy.spatial.distance import pdist
import warnings
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
        # Находим столбец с муниципальными районами
        for col in df.columns:
            if 'МР' in str(col) or 'район' in str(col).lower() or 'муниципальн' in str(col).lower():
                municipalities_set.update(df[col].dropna().astype(str).unique())
                break
    
    return sorted(list(municipalities_set))

def get_available_years(sheets_data):
    """
    Получение списка всех доступных годов
    """
    years_set = set()
    
    for sheet_name, df in sheets_data.items():
        # Находим столбец с годом
        for col in df.columns:
            if 'Год' in str(col) or 'год' in str(col).lower() or 'year' in str(col).lower():
                # Пытаемся преобразовать к числовому типу
                try:
                    years = pd.to_numeric(df[col], errors='coerce').dropna()
                    years_set.update(years.astype(int).unique())
                except:
                    pass
                break
    
    return sorted(list(years_set))

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

def display_years_list(years):
    """
    Отображение списка годов с нумерацией
    """
    print("\n" + "="*60)
    print("ДОСТУПНЫЕ ГОДЫ ДЛЯ АНАЛИЗА:")
    print("="*60)
    
    for i, year in enumerate(years, 1):
        print(f"{i:3}. {year}")
    
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

def select_year_for_clustering(years):
    """
    Выбор года для кластеризации
    """
    print("\n" + "="*60)
    print("ВЫБОР ГОДА ДЛЯ КЛАСТЕРИЗАЦИИ:")
    print("="*60)
    print("1. Выбрать конкретный год")
    print("2. Использовать средние значения за все годы")
    print("3. Использовать последний доступный год")
    print("4. Использовать динамику за все годы")
    print("="*60)
    
    choice = input("\nВаш выбор: ").strip()
    
    if choice == '1':
        display_years_list(years)
        year_choice = input("\nВведите номер года: ").strip()
        
        if year_choice.isdigit():
            idx = int(year_choice) - 1
            if 0 <= idx < len(years):
                selected_year = years[idx]
                print(f"✓ Выбран год: {selected_year}")
                return selected_year, 'single_year'
            else:
                print(f"⚠ Ошибка: номер должен быть от 1 до {len(years)}")
                return None, None
        else:
            print("⚠ Ошибка: введите число")
            return None, None
    
    elif choice == '2':
        print("✓ Используются средние значения за все годы")
        return 'mean', 'average'
    
    elif choice == '3':
        last_year = max(years)
        print(f"✓ Используется последний доступный год: {last_year}")
        return last_year, 'single_year'
    
    elif choice == '4':
        print("✓ Используется динамика за все годы")
        return 'all_years', 'dynamic'
    
    else:
        print("⚠ Неверный выбор. Использую средние значения.")
        return 'mean', 'average'

def clean_sheet_data(df):
    """
    Очистка и стандартизация данных из листа
    """
    df = df.copy()
    
    # Находим столбцы
    mr_col = None
    year_col = None
    value_col = None
    
    for col in df.columns:
        col_str = str(col).lower()
        if 'мр' in col_str or 'район' in col_str or 'муниципальн' in col_str:
            mr_col = col
        elif 'год' in col_str or 'year' in col_str:
            year_col = col
        elif 'значен' in col_str or 'value' in col_str or 'показатель' in col_str:
            value_col = col
    
    # Если не нашли стандартные названия, используем первые три столбца
    if mr_col is None and len(df.columns) >= 1:
        mr_col = df.columns[0]
    if year_col is None and len(df.columns) >= 2:
        year_col = df.columns[1]
    if value_col is None and len(df.columns) >= 3:
        value_col = df.columns[2]
    
    # Переименовываем столбцы
    if mr_col is not None and year_col is not None and value_col is not None:
        df = df.rename(columns={mr_col: 'МР', year_col: 'Год', value_col: 'Значение'})
        
        # Преобразуем год в целое число
        df['Год'] = pd.to_numeric(df['Год'], errors='coerce')
        df = df.dropna(subset=['Год'])
        df['Год'] = df['Год'].astype(int)
        
        # Преобразуем значение в числовой тип
        df['Значение'] = pd.to_numeric(df['Значение'], errors='coerce')
        df = df.dropna(subset=['Значение'])
        
        # Преобразуем МР в строковый тип
        df['МР'] = df['МР'].astype(str).str.strip()
    
    return df

def prepare_clustering_data(sheets_data, selected_indicators, municipalities_list, year_selection, mode='average'):
    """
    Подготовка данных для кластеризации
    """
    data_dict = {}
    
    for municipality in municipalities_list:
        municipality_data = {}
        
        for indicator in selected_indicators:
            if indicator not in sheets_data:
                continue
                
            df = clean_sheet_data(sheets_data[indicator])
            
            # Фильтрация по муниципальному району
            if 'МР' not in df.columns:
                continue
                
            mun_data = df[df['МР'] == municipality]
            
            if not mun_data.empty:
                if mode == 'single_year' and year_selection != 'all_years':
                    # Данные за конкретный год
                    year_data = mun_data[mun_data['Год'] == year_selection]
                    if not year_data.empty:
                        municipality_data[indicator] = year_data['Значение'].iloc[0]
                    else:
                        # Если данных за выбранный год нет, используем ближайший доступный
                        available_years = mun_data['Год'].unique()
                        if len(available_years) > 0:
                            nearest_year = min(available_years, key=lambda x: abs(x - year_selection))
                            year_data = mun_data[mun_data['Год'] == nearest_year]
                            if not year_data.empty:
                                municipality_data[indicator] = year_data['Значение'].iloc[0]
                        else:
                            municipality_data[indicator] = np.nan
                
                elif mode == 'average':
                    # Среднее значение за все годы
                    municipality_data[indicator] = mun_data['Значение'].mean()
                
                elif mode == 'dynamic':
                    # Все годы как отдельные признаки
                    for year in sorted(mun_data['Год'].unique()):
                        year_data = mun_data[mun_data['Год'] == year]
                        if not year_data.empty:
                            municipality_data[f"{indicator}_год{year}"] = year_data['Значение'].iloc[0]
        
        # Добавляем данные только если есть хотя бы одно значение
        if municipality_data:
            data_dict[municipality] = municipality_data
    
    if not data_dict:
        print("⚠ Ошибка: не удалось собрать данные для кластеризации")
        return None
    
    # Преобразование в DataFrame
    clustering_df = pd.DataFrame.from_dict(data_dict, orient='index')
    
    # Заполнение пропущенных значений средними по столбцу
    clustering_df = clustering_df.fillna(clustering_df.mean())
    
    # Удаление столбцов с слишком большим количеством пропусков
    threshold = 0.7  # Максимум 70% пропусков
    clustering_df = clustering_df.loc[:, clustering_df.isnull().mean() < threshold]
    
    # Удаление строк (районов) с пропущенными значениями
    clustering_df = clustering_df.dropna(axis=0, how='any')
    
    if clustering_df.empty:
        print("⚠ Ошибка: после очистки данных не осталось районов для кластеризации")
        return None
    
    return clustering_df

def perform_kmeans_clustering(data, n_clusters=3):
    """
    Выполнение кластеризации методом K-means
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=20)
    clusters = kmeans.fit_predict(scaled_data)
    
    return clusters, kmeans, scaled_data

def perform_hierarchical_clustering(data, n_clusters=3):
    """
    Выполнение иерархической кластеризации
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    hierarchical = AgglomerativeClustering(n_clusters=n_clusters, linkage='ward')
    clusters = hierarchical.fit_predict(scaled_data)
    
    return clusters, hierarchical, scaled_data

def perform_dbscan_clustering(data, eps=0.5, min_samples=5):
    """
    Выполнение кластеризации DBSCAN
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    dbscan = DBSCAN(eps=eps, min_samples=min_samples)
    clusters = dbscan.fit_predict(scaled_data)
    
    return clusters, dbscan, scaled_data

def perform_gmm_clustering(data, n_components=3):
    """
    Выполнение кластеризации Gaussian Mixture Models
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    gmm = GaussianMixture(n_components=n_components, random_state=42)
    clusters = gmm.fit_predict(scaled_data)
    
    return clusters, gmm, scaled_data

def perform_spectral_clustering(data, n_clusters=3):
    """
    Выполнение спектральной кластеризации
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    spectral = SpectralClustering(n_clusters=n_clusters, random_state=42, affinity='nearest_neighbors')
    clusters = spectral.fit_predict(scaled_data)
    
    return clusters, spectral, scaled_data

def perform_birch_clustering(data, n_clusters=3):
    """
    Выполнение кластеризации BIRCH
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    birch = Birch(n_clusters=n_clusters)
    clusters = birch.fit_predict(scaled_data)
    
    return clusters, birch, scaled_data

def determine_optimal_clusters(data, max_clusters=10):
    """
    Определение оптимального количества кластеров
    """
    if len(data) < 3:
        return 2, {'silhouette': [], 'davies_bouldin': [], 'calinski_harabasz': []}
    
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    metrics = {
        'silhouette': [],
        'davies_bouldin': [],
        'calinski_harabasz': []
    }
    
    max_clusters = min(max_clusters, len(data) - 1)
    
    for n in range(2, max_clusters + 1):
        try:
            kmeans = KMeans(n_clusters=n, random_state=42, n_init=10)
            clusters = kmeans.fit_predict(scaled_data)
            
            if len(np.unique(clusters)) > 1:
                metrics['silhouette'].append(silhouette_score(scaled_data, clusters))
                metrics['davies_bouldin'].append(davies_bouldin_score(scaled_data, clusters))
                metrics['calinski_harabasz'].append(calinski_harabasz_score(scaled_data, clusters))
            else:
                metrics['silhouette'].append(0)
                metrics['davies_bouldin'].append(np.inf)
                metrics['calinski_harabasz'].append(0)
        except:
            metrics['silhouette'].append(0)
            metrics['davies_bouldin'].append(np.inf)
            metrics['calinski_harabasz'].append(0)
    
    # Находим оптимальное количество кластеров
    if metrics['silhouette']:
        silhouette_scores = np.array(metrics['silhouette'])
        davies_scores = np.array(metrics['davies_bouldin'])
        
        # Оптимальное по силуэту (максимум)
        optimal_by_silhouette = np.argmax(silhouette_scores) + 2
        
        # Оптимальное по Davies-Bouldin (минимум)
        optimal_by_davies = np.argmin(davies_scores) + 2
        
        # Комбинированная оценка
        if len(silhouette_scores) > 1:
            normalized_silhouette = (silhouette_scores - silhouette_scores.min()) / (silhouette_scores.max() - silhouette_scores.min() + 1e-10)
            normalized_davies = 1 - (davies_scores - davies_scores.min()) / (davies_scores.max() - davies_scores.min() + 1e-10)
            
            combined_scores = normalized_silhouette + normalized_davies
            optimal_combined = np.argmax(combined_scores) + 2
        else:
            optimal_combined = optimal_by_silhouette
        
        return optimal_combined, metrics
    else:
        return 2, metrics

def plot_silhouette_analysis(data, cluster_range=(2, 10)):
    """
    Анализ силуэтных коэффициентов для разных количеств кластеров
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    max_clusters = min(cluster_range[1], len(data) - 1)
    
    silhouette_scores = []
    davies_scores = []
    calinski_scores = []
    
    for n_clusters in range(cluster_range[0], max_clusters + 1):
        try:
            kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
            clusters = kmeans.fit_predict(scaled_data)
            
            if len(np.unique(clusters)) > 1:
                silhouette_scores.append(silhouette_score(scaled_data, clusters))
                davies_scores.append(davies_bouldin_score(scaled_data, clusters))
                calinski_scores.append(calinski_harabasz_score(scaled_data, clusters))
            else:
                silhouette_scores.append(0)
                davies_scores.append(np.inf)
                calinski_scores.append(0)
        except:
            silhouette_scores.append(0)
            davies_scores.append(np.inf)
            calinski_scores.append(0)
    
    # График силуэтных коэффициентов
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # 1. Силуэтные коэффициенты
    x_values = list(range(cluster_range[0], max_clusters + 1))
    axes[0, 0].plot(x_values, silhouette_scores, 'bo-', linewidth=2, markersize=8)
    axes[0, 0].set_xlabel('Количество кластеров')
    axes[0, 0].set_ylabel('Силуэтный коэффициент')
    axes[0, 0].set_title('Силуэтный анализ', fontsize=14, fontweight='bold')
    axes[0, 0].grid(True, alpha=0.3)
    
    # Находим максимум
    if silhouette_scores:
        best_k = x_values[np.argmax(silhouette_scores)]
        best_score = max(silhouette_scores)
        axes[0, 0].axvline(x=best_k, color='r', linestyle='--', alpha=0.7)
        axes[0, 0].text(best_k + 0.1, best_score, f'k={best_k}\nscore={best_score:.3f}', 
                       fontsize=10, verticalalignment='bottom')
    
    # 2. Индекс Davies-Bouldin
    axes[0, 1].plot(x_values, davies_scores, 'ro-', linewidth=2, markersize=8)
    axes[0, 1].set_xlabel('Количество кластеров')
    axes[0, 1].set_ylabel('Индекс Davies-Bouldin')
    axes[0, 1].set_title('Индекс Davies-Bouldin', fontsize=14, fontweight='bold')
    axes[0, 1].grid(True, alpha=0.3)
    
    # Находим минимум
    if davies_scores and np.isfinite(davies_scores).any():
        finite_scores = [s if np.isfinite(s) else np.nan for s in davies_scores]
        if not all(np.isnan(s) for s in finite_scores):
            best_k_db = x_values[np.nanargmin(finite_scores)]
            best_score_db = np.nanmin(finite_scores)
            axes[0, 1].axvline(x=best_k_db, color='b', linestyle='--', alpha=0.7)
            axes[0, 1].text(best_k_db + 0.1, best_score_db, f'k={best_k_db}\nscore={best_score_db:.3f}', 
                          fontsize=10, verticalalignment='bottom')
    
    # 3. Индекс Calinski-Harabasz
    axes[1, 0].plot(x_values, calinski_scores, 'go-', linewidth=2, markersize=8)
    axes[1, 0].set_xlabel('Количество кластеров')
    axes[1, 0].set_ylabel('Индекс Calinski-Harabasz')
    axes[1, 0].set_title('Индекс Calinski-Harabasz', fontsize=14, fontweight='bold')
    axes[1, 0].grid(True, alpha=0.3)
    
    # Находим максимум
    if calinski_scores:
        best_k_ch = x_values[np.argmax(calinski_scores)]
        best_score_ch = max(calinski_scores)
        axes[1, 0].axvline(x=best_k_ch, color='r', linestyle='--', alpha=0.7)
        axes[1, 0].text(best_k_ch + 0.1, best_score_ch, f'k={best_k_ch}\nscore={best_score_ch:.1f}', 
                       fontsize=10, verticalalignment='bottom')
    
    # 4. Сравнение всех метрик
    axes[1, 1].plot(x_values, silhouette_scores, 'bo-', label='Silhouette', linewidth=2, markersize=6)
    
    # Нормализуем Davies-Bouldin для сравнения
    if davies_scores and np.isfinite(davies_scores).any():
        finite_scores = np.array([s if np.isfinite(s) else np.nan for s in davies_scores])
        if not all(np.isnan(finite_scores)):
            min_val = np.nanmin(finite_scores)
            max_val = np.nanmax(finite_scores)
            if max_val > min_val:
                normalized_davies = 1 - (finite_scores - min_val) / (max_val - min_val)
                axes[1, 1].plot(x_values, normalized_davies, 'ro-', label='Davies-Bouldin (norm)', linewidth=2, markersize=6)
    
    # Нормализуем Calinski-Harabasz
    if calinski_scores:
        calinski_arr = np.array(calinski_scores)
        min_val = calinski_arr.min()
        max_val = calinski_arr.max()
        if max_val > min_val:
            normalized_calinski = (calinski_arr - min_val) / (max_val - min_val)
            axes[1, 1].plot(x_values, normalized_calinski, 'go-', label='Calinski-Harabasz (norm)', linewidth=2, markersize=6)
    
    axes[1, 1].set_xlabel('Количество кластеров')
    axes[1, 1].set_ylabel('Нормализованные метрики')
    axes[1, 1].set_title('Сравнение метрик (нормализованные)', fontsize=14, fontweight='bold')
    axes[1, 1].legend()
    axes[1, 1].grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.show()
    
    # Детальный силуэтный анализ для лучшего k
    if silhouette_scores:
        best_k = x_values[np.argmax(silhouette_scores)]
        plot_detailed_silhouette_analysis(data, best_k)
    
    return silhouette_scores, davies_scores, calinski_scores

def plot_detailed_silhouette_analysis(data, n_clusters):
    """
    Детальный анализ силуэтов для конкретного количества кластеров
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
    clusters = kmeans.fit_predict(scaled_data)
    
    # Вычисляем силуэтные коэффициенты для каждого образца
    silhouette_vals = silhouette_samples(scaled_data, clusters)
    silhouette_avg = silhouette_score(scaled_data, clusters)
    
    # Создаем график
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
    # 1. График силуэтов
    y_lower = 10
    for i in range(n_clusters):
        # Собираем силуэтные значения для текущего кластера
        ith_cluster_silhouette_vals = silhouette_vals[clusters == i]
        ith_cluster_silhouette_vals.sort()
        
        size_cluster_i = ith_cluster_silhouette_vals.shape[0]
        y_upper = y_lower + size_cluster_i
        
        color = plt.cm.tab10(i / n_clusters)
        ax1.fill_betweenx(np.arange(y_lower, y_upper),
                         0, ith_cluster_silhouette_vals,
                         facecolor=color, edgecolor=color, alpha=0.7)
        
        # Подписываем кластеры
        ax1.text(-0.05, y_lower + 0.5 * size_cluster_i, str(i))
        
        y_lower = y_upper + 10
    
    ax1.set_xlabel('Значение силуэтного коэффициента')
    ax1.set_ylabel('Кластер')
    ax1.set_title(f'Силуэтный анализ для k={n_clusters}\nСредний силуэт: {silhouette_avg:.3f}', 
                 fontsize=14, fontweight='bold')
    
    # Вертикальная линия для среднего значения
    ax1.axvline(x=silhouette_avg, color="red", linestyle="--")
    ax1.set_yticks([])
    ax1.set_xlim([-1, 1])
    ax1.grid(True, alpha=0.3)
    
    # 2. Визуализация кластеров в пространстве PCA
    pca = PCA(n_components=2, random_state=42)
    pca_result = pca.fit_transform(scaled_data)
    
    unique_clusters = np.unique(clusters)
    colors = plt.cm.tab10(np.linspace(0, 1, len(unique_clusters)))
    
    for i, cluster in enumerate(unique_clusters):
        mask = clusters == cluster
        ax2.scatter(pca_result[mask, 0], pca_result[mask, 1], 
                   c=[colors[i]], label=f'Кластер {cluster}', 
                   s=100, alpha=0.7, edgecolors='black')
    
    ax2.set_xlabel(f'Главная компонента 1 ({pca.explained_variance_ratio_[0]:.1%})')
    ax2.set_ylabel(f'Главная компонента 2 ({pca.explained_variance_ratio_[1]:.1%})')
    ax2.set_title('Визуализация кластеров (PCA)', fontsize=14, fontweight='bold')
    ax2.legend()
    ax2.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.show()

def compare_clustering_methods(data, n_clusters=3):
    """
    Сравнение разных методов кластеризации
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    methods = {
        'K-means': perform_kmeans_clustering,
        'Иерархическая': perform_hierarchical_clustering,
        'GMM': perform_gmm_clustering,
        'Спектральная': perform_spectral_clustering,
        'BIRCH': perform_birch_clustering
    }
    
    results = {}
    
    for method_name, method_func in methods.items():
        try:
            print(f"Выполнение {method_name} кластеризации...")
            clusters, model, scaled_data = method_func(data, n_clusters)
            
            # Вычисляем метрики
            if len(np.unique(clusters)) > 1:
                silhouette = silhouette_score(scaled_data, clusters)
                davies = davies_bouldin_score(scaled_data, clusters)
                calinski = calinski_harabasz_score(scaled_data, clusters)
            else:
                silhouette = 0
                davies = np.inf
                calinski = 0
            
            results[method_name] = {
                'clusters': clusters,
                'silhouette': silhouette,
                'davies_bouldin': davies,
                'calinski_harabasz': calinski,
                'n_clusters': len(np.unique(clusters))
            }
            
            print(f"  ✓ {method_name}: Silhouette={silhouette:.3f}, Davies={davies:.3f}")
            
        except Exception as e:
            print(f"  ⚠ {method_name}: ошибка - {str(e)[:50]}")
            results[method_name] = None
    
    # Визуализация сравнения методов
    plot_methods_comparison(results, scaled_data)
    
    return results

def plot_methods_comparison(results, scaled_data):
    """
    Визуализация сравнения методов кластеризации
    """
    valid_methods = {k: v for k, v in results.items() if v is not None}
    
    if not valid_methods:
        print("⚠ Нет данных для сравнения методов")
        return
    
    # График сравнения метрик
    fig, axes = plt.subplots(2, 3, figsize=(16, 10))
    axes = axes.flatten()
    
    methods_list = list(valid_methods.keys())
    
    # 1. Силуэтные коэффициенты
    silhouette_scores = [valid_methods[m]['silhouette'] for m in methods_list]
    axes[0].bar(range(len(methods_list)), silhouette_scores, color='skyblue')
    axes[0].set_xticks(range(len(methods_list)))
    axes[0].set_xticklabels(methods_list, rotation=45, ha='right')
    axes[0].set_ylabel('Силуэтный коэффициент')
    axes[0].set_title('Сравнение силуэтных коэффициентов', fontsize=12, fontweight='bold')
    axes[0].grid(True, alpha=0.3, axis='y')
    
    # Добавляем значения на столбцы
    for i, v in enumerate(silhouette_scores):
        axes[0].text(i, v + 0.01, f'{v:.3f}', ha='center', va='bottom')
    
    # 2. Индекс Davies-Bouldin
    davies_scores = [valid_methods[m]['davies_bouldin'] if np.isfinite(valid_methods[m]['davies_bouldin']) else 10 
                     for m in methods_list]
    axes[1].bar(range(len(methods_list)), davies_scores, color='lightcoral')
    axes[1].set_xticks(range(len(methods_list)))
    axes[1].set_xticklabels(methods_list, rotation=45, ha='right')
    axes[1].set_ylabel('Индекс Davies-Bouldin')
    axes[1].set_title('Сравнение индекса Davies-Bouldin', fontsize=12, fontweight='bold')
    axes[1].grid(True, alpha=0.3, axis='y')
    
    # 3. Индекс Calinski-Harabasz
    calinski_scores = [valid_methods[m]['calinski_harabasz'] for m in methods_list]
    axes[2].bar(range(len(methods_list)), calinski_scores, color='lightgreen')
    axes[2].set_xticks(range(len(methods_list)))
    axes[2].set_xticklabels(methods_list, rotation=45, ha='right')
    axes[2].set_ylabel('Индекс Calinski-Harabasz')
    axes[2].set_title('Сравнение индекса Calinski-Harabasz', fontsize=12, fontweight='bold')
    axes[2].grid(True, alpha=0.3, axis='y')
    
    # 4. Количество найденных кластеров
    n_clusters_list = [valid_methods[m]['n_clusters'] for m in methods_list]
    axes[3].bar(range(len(methods_list)), n_clusters_list, color='gold')
    axes[3].set_xticks(range(len(methods_list)))
    axes[3].set_xticklabels(methods_list, rotation=45, ha='right')
    axes[3].set_ylabel('Количество кластеров')
    axes[3].set_title('Количество найденных кластеров', fontsize=12, fontweight='bold')
    axes[3].grid(True, alpha=0.3, axis='y')
    
    # Визуализация кластеров для каждого метода (первые 3 метода)
    for idx, (method_name, method_data) in enumerate(list(valid_methods.items())[:3]):
        ax_idx = 4 + idx
        if ax_idx < len(axes) and method_data is not None:
            clusters = method_data['clusters']
            
            # PCA для визуализации
            pca = PCA(n_components=2, random_state=42)
            pca_result = pca.fit_transform(scaled_data)
            
            unique_clusters = np.unique(clusters)
            colors = plt.cm.tab10(np.linspace(0, 1, len(unique_clusters)))
            
            for i, cluster in enumerate(unique_clusters):
                mask = clusters == cluster
                axes[ax_idx].scatter(pca_result[mask, 0], pca_result[mask, 1], 
                                   c=[colors[i]], label=f'Кластер {cluster}', 
                                   s=50, alpha=0.6)
            
            axes[ax_idx].set_xlabel('PC1')
            axes[ax_idx].set_ylabel('PC2')
            axes[ax_idx].set_title(f'{method_name}\nSilhouette: {method_data["silhouette"]:.3f}', 
                                 fontsize=10)
            axes[ax_idx].grid(True, alpha=0.3)
    
    # Удаляем пустые оси
    for idx in range(4 + min(3, len(valid_methods)), len(axes)):
        fig.delaxes(axes[idx])
    
    plt.tight_layout()
    plt.show()
    
    # Таблица с результатами
    print("\n" + "="*80)
    print("СРАВНЕНИЕ МЕТОДОВ КЛАСТЕРИЗАЦИИ:")
    print("="*80)
    print(f"{'Метод':<15} {'Кластеры':<10} {'Silhouette':<12} {'Davies-Bouldin':<15} {'Calinski-Harabasz':<15}")
    print("-" * 80)
    
    for method_name, method_data in valid_methods.items():
        if method_data:
            print(f"{method_name:<15} {method_data['n_clusters']:<10} "
                  f"{method_data['silhouette']:<12.3f} "
                  f"{method_data['davies_bouldin']:<15.3f} "
                  f"{method_data['calinski_harabasz']:<15.1f}")

def select_clustering_method():
    """
    Выбор метода кластеризации
    """
    print("\n" + "="*60)
    print("ВЫБОР МЕТОДА КЛАСТЕРИЗАЦИИ:")
    print("="*60)
    print("1. K-means (стандартный метод)")
    print("2. Иерархическая кластеризация")
    print("3. Gaussian Mixture Models (GMM)")
    print("4. Спектральная кластеризация")
    print("5. BIRCH (для больших данных)")
    print("6. DBSCAN (определяет количество кластеров автоматически)")
    print("7. Сравнить все методы")
    print("="*60)
    
    choice = input("\nВаш выбор: ").strip()
    
    method_map = {
        '1': 'kmeans',
        '2': 'hierarchical',
        '3': 'gmm',
        '4': 'spectral',
        '5': 'birch',
        '6': 'dbscan',
        '7': 'compare'
    }
    
    return method_map.get(choice, 'kmeans')

def select_number_of_clusters(data, max_clusters=10):
    """
    Выбор количества кластеров с анализом силуэтов
    """
    print("\n" + "="*60)
    print("АНАЛИЗ ОПТИМАЛЬНОГО КОЛИЧЕСТВА КЛАСТЕРОВ")
    print("="*60)
    
    # Сначала проводим силуэтный анализ
    print("Выполняю силуэтный анализ для определения оптимального k...")
    silhouette_scores, davies_scores, calinski_scores = plot_silhouette_analysis(data, cluster_range=(2, min(10, len(data)-1)))
    
    # Находим рекомендации
    x_values = list(range(2, min(10, len(data)-1) + 1))
    
    if silhouette_scores:
        best_by_silhouette = x_values[np.argmax(silhouette_scores)]
        best_silhouette = max(silhouette_scores)
        
        # Для Davies-Bouldin ищем минимум среди конечных значений
        finite_davies = [s if np.isfinite(s) else np.nan for s in davies_scores]
        if not all(np.isnan(s) for s in finite_davies):
            best_by_davies = x_values[np.nanargmin(finite_davies)]
            best_davies = np.nanmin(finite_davies)
        else:
            best_by_davies = best_by_silhouette
            best_davies = np.nan
        
        print("\n" + "="*60)
        print("РЕКОМЕНДАЦИИ:")
        print("="*60)
        print(f"• По силуэтному коэффициенту: k={best_by_silhouette} (score={best_silhouette:.3f})")
        if not np.isnan(best_davies):
            print(f"• По индексу Davies-Bouldin: k={best_by_davies} (score={best_davies:.3f})")
        print("="*60)
    
    # Предлагаем выбрать
    print("\nВыберите количество кластеров:")
    print("1. Использовать рекомендуемое по силуэту")
    print("2. Ввести свое количество")
    print("3. Выбрать другое рекомендуемое значение")
    
    choice = input("\nВаш выбор: ").strip()
    
    if choice == '1' and silhouette_scores:
        return best_by_silhouette
    
    elif choice == '2':
        while True:
            try:
                max_k = min(max_clusters, len(data)-1)
                n_clusters = int(input(f"Введите количество кластеров (2-{max_k}): ").strip())
                if 2 <= n_clusters <= max_k:
                    return n_clusters
                else:
                    print(f"Ошибка: должно быть от 2 до {max_k}")
            except ValueError:
                print("Ошибка: введите целое число")
    
    elif choice == '3':
        # Показываем силуэтные коэффициенты
        print("\nСилуэтные коэффициенты для разных k:")
        for k, score in zip(x_values, silhouette_scores):
            print(f"  k={k}: {score:.3f}")
        
        while True:
            try:
                n_clusters = int(input(f"\nВведите выбранное количество кластеров (2-{max_k}): ").strip())
                if 2 <= n_clusters <= max_k:
                    return n_clusters
                else:
                    print(f"Ошибка: должно быть от 2 до {max_k}")
            except ValueError:
                print("Ошибка: введите целое число")
    
    else:
        return 3  # Значение по умолчанию

def perform_cluster_analysis(sheets_data, selected_indicators, municipalities_list, year_selection, mode):
    """
    Основная функция кластерного анализа
    """
    print("\n" + "="*60)
    print("НАЧАЛО КЛАСТЕРНОГО АНАЛИЗА")
    print("="*60)
    print(f"Количество районов: {len(municipalities_list)}")
    print(f"Количество показателей: {len(selected_indicators)}")
    
    # Определяем информацию о годе для отчета
    if mode == 'single_year':
        year_info = f"Год: {year_selection}"
    elif mode == 'average':
        year_info = "Средние значения за все годы"
    elif mode == 'dynamic':
        year_info = "Динамика за все годы"
    else:
        year_info = "Неизвестный режим"
    
    print(f"Режим анализа: {year_info}")
    print("="*60)
    
    # Подготовка данных
    print("\nПодготовка данных для кластеризации...")
    clustering_data = prepare_clustering_data(sheets_data, selected_indicators, 
                                            municipalities_list, year_selection, mode)
    
    if clustering_data is None or clustering_data.empty:
        print("⚠ Ошибка: не удалось подготовить данные для кластеризации")
        return None, None, None
    
    # Получаем список районов, которые остались после очистки
    available_municipalities = clustering_data.index.tolist()
    
    print(f"✓ Данные успешно подготовлены")
    print(f"✓ Районов для анализа: {len(available_municipalities)} из {len(municipalities_list)}")
    print(f"✓ Размерность данных: {clustering_data.shape}")
    
    if len(available_municipalities) < 3:
        print("⚠ Ошибка: слишком мало данных для кластеризации (нужно минимум 3 района)")
        return None, None, None
    
    # Выбор метода кластеризации
    method_choice = select_clustering_method()
    
    if method_choice == 'compare':
        # Сравнение методов
        print("\n" + "="*60)
        print("СРАВНЕНИЕ МЕТОДОВ КЛАСТЕРИЗАЦИИ")
        print("="*60)
        
        # Сначала определяем оптимальное количество кластеров
        n_clusters = select_number_of_clusters(clustering_data)
        
        # Выполняем сравнение методов
        results = compare_clustering_methods(clustering_data, n_clusters)
        
        # Пользователь выбирает лучший метод
        print("\n" + "="*60)
        print("ВЫБЕРИТЕ МЕТОД ДЛЯ ДЕТАЛЬНОГО АНАЛИЗА:")
        valid_methods = [m for m, r in results.items() if r is not None]
        
        for i, method in enumerate(valid_methods, 1):
            print(f"{i}. {method} (Silhouette: {results[method]['silhouette']:.3f})")
        
        try:
            method_idx = int(input("\nВаш выбор (номер): ").strip()) - 1
            if 0 <= method_idx < len(valid_methods):
                selected_method = valid_methods[method_idx]
                method_data = results[selected_method]
                
                # Выполняем детальный анализ выбранного метода
                clusters = method_data['clusters']
                silhouette_avg = method_data['silhouette']
                davies_avg = method_data['davies_bouldin']
                calinski_avg = method_data['calinski_harabasz']
                
                print(f"\n✓ Выбран метод: {selected_method}")
                print(f"✓ Силуэтный коэффициент: {silhouette_avg:.3f}")
                
                # Визуализация результатов выбранного метода
                scaler = StandardScaler()
                scaled_data = scaler.fit_transform(clustering_data)
                
                # Визуализация
                visualize_clusters_2d(clustering_data, clusters, available_municipalities, selected_method)
                
                # Детальный силуэтный анализ
                plot_detailed_silhouette_analysis(clustering_data, len(np.unique(clusters)))
                
                # Другие визуализации
                if len(available_municipalities) <= 50:
                    visualize_dendrogram(clustering_data, available_municipalities)
                
                visualize_cluster_characteristics(clustering_data, clusters, available_municipalities)
                visualize_cluster_size_distribution(clusters, available_municipalities)
                
                # Создание отчета
                cluster_info = create_cluster_report(clusters, available_municipalities, 
                                                   clustering_data, year_info)
                
                # Сохранение результатов
                save_clustering_results(clustering_data, clusters, available_municipalities, 
                                      selected_method, silhouette_avg, davies_avg, calinski_avg)
                
                return cluster_info, None, None
                
            else:
                print("⚠ Неверный выбор. Использую K-means.")
                method_choice = 'kmeans'
        except:
            print("⚠ Неверный выбор. Использую K-means.")
            method_choice = 'kmeans'
    
    # Если выбран конкретный метод
    if method_choice != 'compare':
        # Выбор количества кластеров (кроме DBSCAN)
        if method_choice != 'dbscan':
            n_clusters = select_number_of_clusters(clustering_data)
        else:
            n_clusters = None
        
        # Выполнение выбранного метода
        print(f"\nВыполнение {method_choice} кластеризации...")
        
        if method_choice == 'kmeans':
            clusters, model, scaled_data = perform_kmeans_clustering(clustering_data, n_clusters)
            method_name = "K-means"
        elif method_choice == 'hierarchical':
            clusters, model, scaled_data = perform_hierarchical_clustering(clustering_data, n_clusters)
            method_name = "Иерархическая кластеризация"
        elif method_choice == 'gmm':
            clusters, model, scaled_data = perform_gmm_clustering(clustering_data, n_clusters)
            method_name = "Gaussian Mixture Models"
        elif method_choice == 'spectral':
            clusters, model, scaled_data = perform_spectral_clustering(clustering_data, n_clusters)
            method_name = "Спектральная кластеризация"
        elif method_choice == 'birch':
            clusters, model, scaled_data = perform_birch_clustering(clustering_data, n_clusters)
            method_name = "BIRCH"
        elif method_choice == 'dbscan':
            # Для DBSCAN нужно выбрать параметры
            print("\nНастройка параметров DBSCAN:")
            eps = float(input("Введите параметр eps (рекомендуется 0.5): ") or "0.5")
            min_samples = int(input("Введите min_samples (рекомендуется 5): ") or "5")
            clusters, model, scaled_data = perform_dbscan_clustering(clustering_data, eps, min_samples)
            method_name = "DBSCAN"
            n_clusters = len(np.unique(clusters))
        
        print(f"✓ {method_name} кластеризация выполнена")
        print(f"✓ Найдено кластеров: {len(np.unique(clusters))}")
        
        # Вычисляем метрики
        if len(np.unique(clusters)) > 1:
            silhouette_avg = silhouette_score(scaled_data, clusters)
            davies_avg = davies_bouldin_score(scaled_data, clusters)
            calinski_avg = calinski_harabasz_score(scaled_data, clusters)
        else:
            silhouette_avg = 0
            davies_avg = np.inf
            calinski_avg = 0
        
        print(f"✓ Метрики качества:")
        print(f"  • Силуэтный коэффициент: {silhouette_avg:.3f}")
        print(f"  • Индекс Davies-Bouldin: {davies_avg:.3f}")
        print(f"  • Индекс Calinski-Harabasz: {calinski_avg:.1f}")
        
        # Визуализация результатов
        print("\nВизуализация результатов кластеризации...")
        
        # 1. 2D визуализация
        visualize_clusters_2d(clustering_data, clusters, available_municipalities, method_name)
        
        # 2. Детальный силуэтный анализ
        if method_choice in ['kmeans', 'hierarchical', 'gmm']:
            plot_detailed_silhouette_analysis(clustering_data, len(np.unique(clusters)))
        
        # 3. Дендрограмма (если районов не слишком много)
        if len(available_municipalities) <= 50 and method_choice in ['hierarchical', 'kmeans']:
            visualize_dendrogram(clustering_data, available_municipalities)
        
        # 4. Характеристики кластеров
        cluster_stats = visualize_cluster_characteristics(
            clustering_data, clusters, available_municipalities
        )
        
        # 5. Распределение по кластерам
        visualize_cluster_size_distribution(clusters, available_municipalities)
        
        # 6. Создание отчета
        cluster_info = create_cluster_report(clusters, available_municipalities, 
                                           clustering_data, year_info)
        
        # Сохранение результатов
        save_clustering_results(clustering_data, clusters, available_municipalities, 
                              method_name, silhouette_avg, davies_avg, calinski_avg)
        
        return cluster_info, cluster_stats, model
    
    return None, None, None

def visualize_clusters_2d(data, clusters, municipalities, method_name):
    """
    Визуализация кластеров в 2D пространстве с помощью PCA и t-SNE
    """
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    # PCA для уменьшения размерности до 2D
    pca = PCA(n_components=2, random_state=42)
    pca_result = pca.fit_transform(scaled_data)
    
    # t-SNE для нелинейного уменьшения размерности
    tsne_perplexity = min(30, len(data) - 1)
    if tsne_perplexity > 0:
        tsne = TSNE(n_components=2, random_state=42, perplexity=tsne_perplexity)
        tsne_result = tsne.fit_transform(scaled_data)
    else:
        tsne_result = pca_result
    
    # Создаем фигуру с двумя подграфиками
    fig, axes = plt.subplots(1, 2, figsize=(16, 7))
    
    # PCA визуализация
    unique_clusters = np.unique(clusters)
    colors = plt.cm.tab10(np.linspace(0, 1, len(unique_clusters)))
    
    for i, cluster in enumerate(unique_clusters):
        mask = clusters == cluster
        axes[0].scatter(pca_result[mask, 0], pca_result[mask, 1], 
                       c=[colors[i]], label=f'Кластер {cluster}', 
                       s=100, alpha=0.7, edgecolors='black')
    
    axes[0].set_title(f'PCA проекция - {method_name}', fontsize=14, fontweight='bold')
    axes[0].set_xlabel(f'Главная компонента 1 ({pca.explained_variance_ratio_[0]:.1%})')
    axes[0].set_ylabel(f'Главная компонента 2 ({pca.explained_variance_ratio_[1]:.1%})')
    axes[0].legend()
    axes[0].grid(True, alpha=0.3)
    
    # t-SNE визуализация
    for i, cluster in enumerate(unique_clusters):
        mask = clusters == cluster
        axes[1].scatter(tsne_result[mask, 0], tsne_result[mask, 1], 
                       c=[colors[i]], label=f'Кластер {cluster}', 
                       s=100, alpha=0.7, edgecolors='black')
    
    axes[1].set_title(f't-SNE проекция - {method_name}', fontsize=14, fontweight='bold')
    axes[1].set_xlabel('t-SNE компонента 1')
    axes[1].set_ylabel('t-SNE компонента 2')
    axes[1].legend()
    axes[1].grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.show()
    
    return fig, pca_result, tsne_result

def visualize_dendrogram(data, municipalities):
    """
    Визуализация дендрограммы для иерархической кластеризации
    """
    if len(municipalities) > 50:
        return
    
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(data)
    
    distance_matrix = pdist(scaled_data, metric='euclidean')
    linkage_matrix = linkage(distance_matrix, method='ward')
    
    plt.figure(figsize=(12, max(6, len(municipalities) * 0.3)))
    dendrogram(linkage_matrix,
               labels=municipalities,
               orientation='right',
               leaf_font_size=9,
               color_threshold=0.7 * max(linkage_matrix[:, 2]))
    
    plt.title('Дендрограмма иерархической кластеризации', fontsize=14, fontweight='bold')
    plt.xlabel('Расстояние')
    plt.tight_layout()
    plt.show()

def visualize_cluster_characteristics(data, clusters, municipalities):
    """
    Визуализация характеристик кластеров
    """
    cluster_df = pd.DataFrame(data.copy())
    cluster_df['Кластер'] = clusters
    cluster_df.index = municipalities
    
    cluster_stats = cluster_df.groupby('Кластер').mean()
    
    if cluster_stats.empty:
        return None
    
    normalized_stats = (cluster_stats - cluster_stats.mean()) / cluster_stats.std()
    
    plt.figure(figsize=(12, 6))
    
    max_features = 15
    if normalized_stats.shape[1] > max_features:
        feature_variance = normalized_stats.var()
        important_features = feature_variance.nlargest(max_features).index
        normalized_stats = normalized_stats[important_features]
    
    sns.heatmap(normalized_stats.T, 
                cmap='RdBu_r', 
                center=0,
                annot=True, 
                fmt='.2f',
                linewidths=0.5,
                cbar_kws={'label': 'Z-скор (отклонение от среднего)'})
    
    plt.title('Характеристики кластеров (нормализованные значения)', fontsize=14, fontweight='bold')
    plt.xlabel('Кластер')
    plt.ylabel('Показатели')
    plt.tight_layout()
    plt.show()
    
    return cluster_stats

def visualize_cluster_size_distribution(clusters, municipalities):
    """
    Визуализация распределения районов по кластерам
    """
    unique_clusters, cluster_counts = np.unique(clusters, return_counts=True)
    
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    colors = plt.cm.Set3(np.linspace(0, 1, len(unique_clusters)))
    cluster_labels = [f'Кластер {c}' for c in unique_clusters]
    bars = axes[0].bar(cluster_labels, cluster_counts, color=colors, edgecolor='black')
    
    axes[0].set_title('Количество районов в кластерам', fontsize=14, fontweight='bold')
    axes[0].set_ylabel('Количество районов')
    
    for bar, count in zip(bars, cluster_counts):
        height = bar.get_height()
        axes[0].text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{count}', ha='center', va='bottom', fontweight='bold')
    
    if len(cluster_counts) > 1:
        wedges, texts, autotexts = axes[1].pie(cluster_counts, 
                                              labels=cluster_labels,
                                              autopct='%1.1f%%',
                                              startangle=90,
                                              colors=colors)
        axes[1].set_title('Распределение районов по кластерам', fontsize=14, fontweight='bold')
    else:
        axes[1].text(0.5, 0.5, 'Все районы в одном кластере', 
                    ha='center', va='center', fontsize=12)
        axes[1].set_title('Распределение районов по кластерам', fontsize=14, fontweight='bold')
    
    plt.tight_layout()
    plt.show()

def create_cluster_report(clusters, municipalities, data, year_info):
    """
    Создание отчета по кластерам
    """
    cluster_info = pd.DataFrame({
        'Муниципальный район': municipalities,
        'Кластер': clusters
    })
    
    for col in data.columns:
        if col != 'Кластер':
            cluster_info[col] = data[col].values
    
    cluster_info = cluster_info.sort_values(['Кластер', 'Муниципальный район'])
    
    print("\n" + "="*80)
    print("ОТЧЕТ ПО КЛАСТЕРИЗАЦИИ МУНИЦИПАЛЬНЫХ РАЙОНОВ")
    print("="*80)
    print(f"Общее количество районов: {len(municipalities)}")
    print(f"Количество кластеров: {len(np.unique(clusters))}")
    print(f"{year_info}")
    print("="*80)
    
    for cluster in sorted(cluster_info['Кластер'].unique()):
        cluster_data = cluster_info[cluster_info['Кластер'] == cluster]
        print(f"\nКЛАСТЕР {cluster} ({len(cluster_data)} районов):")
        print("-" * 60)
        
        for i, (_, row) in enumerate(cluster_data.iterrows(), 1):
            mun_name = row['Муниципальный район']
            if len(mun_name) > 40:
                mun_name = mun_name[:37] + "..."
            print(f"{i:2}. {mun_name}")
    
    return cluster_info

def save_clustering_results(data, clusters, municipalities, method_name, 
                           silhouette_score, davies_score, calinski_score):
    """
    Сохранение результатов кластеризации
    """
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Основные результаты
    results_df = pd.DataFrame({
        'Муниципальный район': municipalities,
        'Кластер': clusters
    })
    
    for col in data.columns:
        results_df[col] = data[col].values
    
    filename = f"результаты_{method_name}_{timestamp}.csv"
    results_df.to_csv(filename, index=False, encoding='utf-8-sig')
    print(f"✓ Результаты сохранены в файл: {filename}")
    
    # Метрики качества
    metrics_df = pd.DataFrame({
        'Метод': [method_name],
        'Количество кластеров': [len(np.unique(clusters))],
        'Силуэтный коэффициент': [silhouette_score],
        'Индекс Davies-Bouldin': [davies_score],
        'Индекс Calinski-Harabasz': [calinski_score],
        'Дата анализа': [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    })
    
    metrics_filename = f"метрики_{method_name}_{timestamp}.csv"
    metrics_df.to_csv(metrics_filename, index=False, encoding='utf-8-sig')
    print(f"✓ Метрики сохранены в файл: {metrics_filename}")

def main():
    """
    Основная функция программы
    """
    print("="*60)
    print("АНАЛИЗ И СРАВНЕНИЕ ПОКАЗАТЕЛЕЙ МУНИЦИПАЛЬНЫХ РАЙОНОВ")
    print("РЕСПУБЛИКИ САХА (ЯКУТИЯ)")
    print("="*60)
    
    file_path = "Реакция экономики на шоки на примере муниципальных образований Республики Саха (2).xlsx"
    
    if not Path(file_path).exists():
        print(f"⚠ Файл '{file_path}' не найден!")
        print("Пожалуйста, укажите правильный путь к файлу.")
        file_path = input("Введите путь к файлу Excel: ").strip()
    
    sheets_data, sheet_names = read_excel_sheets(file_path)
    
    if sheets_data is None:
        print("Не удалось прочитать данные из файла. Программа завершена.")
        return
    
    municipalities = get_available_municipalities(sheets_data)
    years = get_available_years(sheets_data)
    
    if not municipalities:
        print("В файле не найдены данные о муниципальных районах.")
        return
    
    print(f"\nЗагружено данных: {len(sheet_names)} листов")
    print(f"Найдено муниципальных районов: {len(municipalities)}")
    
    if years:
        print(f"Доступные годы: {min(years)} - {max(years)}")
    
    while True:
        try:
            print("\n" + "="*60)
            print("ГЛАВНОЕ МЕНЮ:")
            print("="*60)
            print("1. Просмотреть список всех муниципальных районов")
            print("2. Просмотреть список всех показателей")
            print("3. Выполнить кластерный анализ")
            print("4. Выход")
            print("="*60)
            
            choice = input("\nВаш выбор: ").strip()
            
            if choice == '1':
                display_municipality_list(municipalities)
                input("\nНажмите Enter для продолжения...")
            
            elif choice == '2':
                display_indicators_list(sheet_names)
                input("\nНажмите Enter для продолжения...")
            
            elif choice == '3':
                print("\n" + "="*60)
                print("КЛАСТЕРНЫЙ АНАЛИЗ МУНИЦИПАЛЬНЫХ РАЙОНОВ")
                print("="*60)
                
                # Выбор муниципальных районов
                print("Выберите муниципальные районы для анализа:")
                print("1. Использовать ВСЕ районы")
                print("2. Выбрать определенные районы")
                
                mun_choice = input("\nВаш выбор (1/2): ").strip()
                
                if mun_choice == '1':
                    selected_municipalities = municipalities.copy()
                    print(f"✓ Выбраны ВСЕ {len(selected_municipalities)} муниципальных районов")
                else:
                    selected_municipalities = municipalities.copy()
                    print("⚠ Использую все районы")
                
                # Выбор показателей
                print("\nВыберите показатели для анализа:")
                print("1. Использовать ВСЕ показатели")
                print("2. Выбрать определенные показатели")
                
                ind_choice = input("\nВаш выбор (1/2): ").strip()
                
                if ind_choice == '1':
                    selected_indicators = sheet_names.copy()
                    print(f"✓ Выбраны ВСЕ {len(selected_indicators)} показателей")
                else:
                    selected_indicators = sheet_names.copy()
                    print("⚠ Использую все показатели")
                
                # Выбор года для анализа
                print("\nВыберите временной период для анализа:")
                if years:
                    year_selection, mode = select_year_for_clustering(years)
                    
                    if year_selection is None:
                        print("⚠ Ошибка выбора года. Отмена анализа.")
                        continue
                else:
                    print("⚠ Годы не найдены. Использую средние значения.")
                    year_selection = 'mean'
                    mode = 'average'
                
                # Проверка данных
                if len(selected_municipalities) < 3:
                    print("⚠ Для кластеризации нужно минимум 3 района.")
                    continue
                
                if len(selected_indicators) < 2:
                    print("⚠ Для кластеризации нужно минимум 2 показателя.")
                    continue
                
                # Запуск кластерного анализа
                print("\n" + "="*60)
                print("ЗАПУСК КЛАСТЕРНОГО АНАЛИЗА")
                print("="*60)
                
                try:
                    results_df, cluster_stats, model = perform_cluster_analysis(
                        sheets_data, selected_indicators, selected_municipalities, 
                        year_selection, mode
                    )
                    
                    if results_df is not None:
                        print("\n✓ Кластерный анализ завершен успешно!")
                    else:
                        print("\n⚠ Кластерный анализ завершился с ошибками.")
                except Exception as e:
                    print(f"\n⚠ Ошибка при выполнении кластерного анализа: {e}")
                    import traceback
                    traceback.print_exc()
                
                input("\nНажмите Enter для продолжения...")
            
            elif choice == '4':
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