import streamlit as st
import pandas as pd
import numpy as np
import datetime
import gc  # Сборщик мусора для очистки оперативной памяти

# Настройка страницы Streamlit
st.set_page_config(page_title="Комбайн-П: Предотказ", layout="wide")

st.title("🚂 Модуль «Комбайн-П: Предотказ»")
st.markdown("### Оптимизированный предиктивный анализ (Версия с каскадными фильтрами)")

# --- ФУНКЦИИ ГЕНЕРАЦИИ ДЕМО-ДАННЫХ ---
@st.cache_data
def get_demo_data():
    """Генерация легких тестовых данных для экономии памяти"""
    np.random.seed(42)
    km_rows = []
    directions = [24602, 91022700, 91031600]
    
    for yr in [2024, 2025, 2026]:
        for mn in [4, 5, 6]:
            for day in [8, 22]:
                for d_kod in directions:
                    for km in range(2320, 2335):
                        ball = np.random.randint(95, 150) if (km == 2328 and d_kod == 24602 and mn == 5) else np.random.randint(10, 35)
                        km_rows.append({
                            'ГОД': yr, 'МЕСЯЦ': mn, 'ДЕНЬ': day, 'КОДНАПР': d_kod, 
                            'ПУТЬ': 1, 'KM': km, 'БАЛЛ': ball, 'ПД': 7
                        })
    df_km = pd.DataFrame(km_rows)
    
    otst_rows = []
    types = ['Просадка', 'Уширение', 'Рихтовка', 'Перекос']
    for idx, row in df_km.iterrows():
        if row['БАЛЛ'] > 40:
            for _ in range(int(row['БАЛЛ'] / 25)):
                otst_rows.append({
                    'ГОД': row['ГОД'], 'МЕСЯЦ': row['МЕСЯЦ'], 'ДЕНЬ': row['ДЕНЬ'],
                    'КОДНАПР': row['КОДНАПР'], 'ПУТЬ': row['ПУТЬ'], 'KM': row['KM'],
                    'М': np.random.randint(100, 900), 'ОТСТУПЛЕНИЕ': np.random.choice(types), 'БАЛЛ': 20
                })
    return df_km, pd.DataFrame(otst_rows)

# --- БОКОВАЯ ПАНЕЛЬ С МУЛЬТИЗАГРУЗКОЙ ---
st.sidebar.header("📁 Архив путеизмерителя")
uploaded_files = st.sidebar.file_uploader(
    "Выберите один или несколько файлов Excel (.xlsx):", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Списки необходимых колонок (в верхнем регистре) для точечной загрузки
REQUIRED_KM_COLS = ['ГОД', 'МЕСЯЦ', 'КОДНАПР', 'ПУТЬ', 'KM', 'БАЛЛ', 'ПД']
REQUIRED_OTST_COLS = ['ГОД', 'МЕСЯЦ', 'КОДНАПР', 'ПУТЬ', 'KM', 'М', 'ОТСТУПЛЕНИЕ', 'БАЛЛ']

km_list = []
otst_list = []

if uploaded_files:
    success_count = 0
    for f in uploaded_files:
        try:
            excel_file = pd.ExcelFile(f)
            sheet_names = excel_file.sheet_names
            
            target_km_sheet = next((s for s in sheet_names if s.strip().lower() == "оценка км"), None)
            target_otst_sheet = next((s for s in sheet_names if s.strip().lower() == "отступления"), None)
            
            if target_km_sheet and target_otst_sheet:
                df_km_cols = pd.read_excel(f, sheet_name=target_km_sheet, nrows=1)
                df_otst_cols = pd.read_excel(f, sheet_name=target_otst_sheet, nrows=1)
                
                km_cols_dict = {col.strip().upper(): col for col in df_km_cols.columns}
                otst_cols_dict = {col.strip().upper(): col for col in df_otst_cols.columns}
                
                if 'КОДНАПРВ' in otst_cols_dict:
                    otst_cols_dict['КОДНАПР'] = otst_cols_dict.pop('КОДНАПРВ')
                
                actual_km_cols = [km_cols_dict[c] for c in REQUIRED_KM_COLS if c in km_cols_dict]
                actual_otst_cols = [otst_cols_dict[c] for c in REQUIRED_OTST_COLS if c in otst_cols_dict]
                
                df_km_single = pd.read_excel(f, sheet_name=target_km_sheet, usecols=actual_km_cols)
                df_otst_single = pd.read_excel(f, sheet_name=target_otst_sheet, usecols=actual_otst_cols)
                
                df_km_single.columns = df_km_single.columns.str.strip().str.upper()
                df_otst_single.columns = df_otst_single.columns.str.strip().str.upper()
                if 'КОДНАПРВ' in df_otst_single.columns:
                    df_otst_single = df_otst_single.rename(columns={'КОДНАПРВ': 'КОДНАПР'})
                
                for col in ['ГОД', 'МЕСЯЦ', 'ПУТЬ', 'KM', 'БАЛЛ', 'ПД']:
                    if col in df_km_single.columns:
                        df_km_single[col] = pd.to_numeric(df_km_single[col], errors='coerce')
                
                for col in ['ГОД', 'МЕСЯЦ', 'ПУТЬ', 'KM', 'М', 'БАЛЛ']:
                    if col in df_otst_single.columns:
                        df_otst_single[col] = pd.to_numeric(df_otst_single[col], errors='coerce')
                
                km_list.append(df_km_single)
                otst_list.append(df_otst_single)
                success_count += 1
                
            del excel_file
            gc.collect()
            
        except Exception as e:
            st.sidebar.error(f"Ошибка в файле {f.name}: {e}")
            
    if success_count > 0:
        st.sidebar.success(f"📊 Объединено файлов: {success_count} шт.")
        df_km_raw = pd.concat(km_list, ignore_index=True)
        df_otst_raw = pd.concat(otst_list, ignore_index=True)
        del km_list, otst_list
        gc.collect()
    else:
        st.sidebar.error("❌ Нужные листы не найдены. Загружены демо-данные.")
        df_km_raw, df_otst_raw = get_demo_data()
else:
    st.sidebar.info("Используются демонстрационные данные для ПЧ-22.")
    df_km_raw, df_otst_raw = get_demo_data()

# Настройки фильтрации
st.sidebar.header("⚙️ Настройки анализа")
current_month = st.sidebar.slider("Прогноз на месяц:", 1, 12, int(datetime.datetime.now().month))
threshold_ball = st.sidebar.number_input("Сигнальный порог баллов для КМ:", value=50, step=5)

# --- ОБРАБОТКА ДАННЫХ ---

hist_km = df_km_raw[df_km_raw['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM', 'БАЛЛ']).copy()
hist_otst = df_otst_raw[df_otst_raw['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM']).copy()

del df_km_raw, df_otst_raw
gc.collect()

if hist_km.empty:
    st.warning(f"⚠️ В загруженном архиве нет исторических записей для месяца № {current_month}.")
else:
    # Строим профиль километров с добавлением ПД (берем моду (наиболее частое значение) или первое попавшееся)
    km_profile = hist_km.groupby(['КОДНАПР', 'ПУТЬ', 'KM']).agg(
        Ср_Балл=('БАЛЛ', 'mean'),
        Макс_Балл=('БАЛЛ', 'max'),
        Кол_Проверок=('БАЛЛ', 'count'),
        Превышений_Порога=('БАЛЛ', lambda x: (x >= threshold_ball).sum()),
        НОМЕР_ПД=('ПД', lambda x: x.dropna().iloc[0] if not x.dropna().empty else 0)
    ).reset_index()

    km_profile['Повторяемость_%'] = (km_profile['Превышений_Порога'] / km_profile['Кол_Проверок'] * 100).round(1)

    dangerous_kms = km_profile[
        (km_profile['Ср_Балл'] >= threshold_ball) | (km_profile['Превышений_Порога'] >= 2)
    ].sort_values(by='Ср_Балл', ascending=False)

    # --- ИНТЕРФЕЙС ---
    c1, c2 = st.columns(2)
    with c1:
        st.metric(label="Найдено км в зоне риска предотказа", value=len(dangerous_kms))
    with c2:
        st.success(f"Анализ проведен по архиву из {int(km_profile['Кол_Проверок'].sum())} проверок км.")

    if dangerous_kms.empty:
        st.success("🎉 Исторических аномалий для этого периода не обнаружено. Все участки стабильны!")
    else:
        st.subheader("📋 Сводная ведомость километров с высоким риском неисправностей")
        dangerous_kms['Ср_Балл'] = dangerous_kms['Ср_Балл'].round(1)
        
        display_df = dangerous_kms[['КОДНАПР', 'ПУТЬ', 'KM', 'Ср_Балл', 'Макс_Балл', 'Повторяемость_%']].copy()
        for col in ['КОДНАПР', 'ПУТЬ', 'KM']:
            display_df[col] = display_df[col].fillna(0).astype(int)

        st.dataframe(display_df.rename(columns={
            'КОДНАПР': 'Код направления', 'ПУТЬ': 'Путь', 'KM': 'Километр',
            'Ср_Балл': 'Исторический ср. балл', 'Макс_Балл': 'Пиковый балл',
            'Повторяемость_%': 'Вероятность повторения %'
        }), use_container_width=True, hide_index=True)

        # --- 1. КАСКАДНЫЕ ФИЛЬТРЫ ДЛЯ ДЕТАЛИЗАЦИИ УЧАСТКА ---
        st.subheader("🔍 Детальный анализ геометрии выбранного участка")
        
        # Создаем три колонки под выпадающие списки фильтров
        f1, f2, f3 = st.columns(3)
        
        with f1:
            available_directions = sorted(dangerous_kms['КОДНАПР'].dropna().unique().astype(int))
            selected_dir = st.selectbox("1. Выберите направление:", available_directions)
            
        # Фильтруем данные под выбранное направление
        df_filtered_dir = dangerous_kms[dangerous_kms['КОДНАПР'] == selected_dir]
        
        with f2:
            available_paths = sorted(df_filtered_dir['ПУТЬ'].dropna().unique().astype(int))
            selected_path = st.selectbox("2. Выберите путь:", available_paths)
            
        # Фильтруем данные под выбранный путь
        df_filtered_path = df_filtered_dir[df_filtered_dir['ПУТЬ'] == selected_path]
        
        with f3:
            available_kms = sorted(df_filtered_path['KM'].dropna().unique().astype(int))
            selected_km = st.selectbox("3. Выберите километр:", available_kms)

        # Вытаскиваем строку с результатами для окончательно выбранной точки
        sel_row = df_filtered_path[df_filtered_path['KM'] == selected_km].iloc[0]
        pd_number = int(sel_row['НОМЕР_ПД']) if pd.notna(sel_row['НОМЕР_ПД']) else 0
        
        # Выгружаем точечные отступления для выбранного КМ из общего массива
        km_defects = hist_otst[
            (hist_otst['КОДНАПР'] == selected_dir) & 
            (hist_otst['ПУТЬ'] == selected_path) & 
            (hist_otst['KM'] == selected_km)
        ].copy()
        
        if km_defects.empty:
            st.info(f"💡 Для {selected_km} км нет детализированных записей по отдельным отступлениям в листах 'Отступления'. Используйте общую статистику баллов.")
        else:
            km_defects['ОТСТУПЛЕНИЕ'] = km_defects['ОТСТУПЛЕНИЕ'].astype(str).str.strip()
            
            dl, dr = st.columns(2)
            with dl:
                st.markdown("**Характерные виды неисправностей (по сумме баллов):**")
                structure = km_defects.groupby('ОТСТУПЛЕНИЕ', as_index=False).agg(
                    Суммарный_Балл=('БАЛЛ', 'sum'),
                    Количество_Случаев=('ОТСТУПЛЕНИЕ', 'count')
                ).sort_values(by='Суммарный_Балл', ascending=False)
                
                st.dataframe(structure.rename(columns={
                    'ОТСТУПЛЕНИЕ': 'Вид неисправности', 
                    'Суммарный_Балл': 'Набрано баллов', 
                    'Количество_Случаев': 'Кол-во случаев'
                }), use_container_width=True, hide_index=True)
                
            with dr:
                st.markdown("**Критические участки километра (Привязка к метрам):**")
                km_defects['Участок_Метры'] = (km_defects['М'] // 100) * 100
                
                # Группируем данные по метрам
                density = km_defects.groupby('Участок_Метры', as_index=False).agg(
                    Баллы=('БАЛЛ', 'sum'),
                    Количество=('БАЛЛ', 'count')
                )
                
                # --- 2. СОРТИРОВКА ПО ВОЗРАСТАНИЮ МЕТРОВ ---
                density = density.sort_values(by='Участок_Метры', ascending=True)
                
                # Создаем красивый текстовый интервал
                density['Интервал (м)'] = density['Участок_Метры'].astype(int).astype(str) + " - " + (density['Участок_Метры'] + 100).astype(int).astype(str)
                
                # Перестраиваем порядок колонок: Интервал -> Количество -> Сумма баллов
                display_density = density[['Интервал (м)', 'Количество', 'Баллы']].copy()
                
                st.dataframe(display_density.rename(columns={
                    'Количество': 'Количество отступлений', 
                    'Баллы': 'Сумма баллов'
                }), use_container_width=True, hide_index=True)

            # Находим самый опасный участок для вывода в рекомендацию (по макс. баллам)
            worst_density = density.sort_values(by='Баллы', ascending=False)

            if not structure.empty and not worst_density.empty:
                main_threat = structure.iloc[0]['ОТСТУПЛЕНИЕ']
                crit_meters = worst_density.iloc[0]['Интервал (м)']
                
                # --- 3. РЕКОМЕНДАЦИЯ С УЧЕТОМ НОМЕРА ПД ---
                pd_prefix = f" ПД-{pd_number}" if pd_number > 0 else ""
                st.warning(f"📋 **Рекомендация для дорожного мастера{pd_prefix}:** На {int(selected_km)} км в этом месяце ожидается ухудшение геометрии пути. "
                           f"Выдайте целевое задание бригадиру (ПДБ) проверить в первую очередь интервал **{crit_meters} метров**. "
                           f"Основное внимание обратить на поиск и превентивное устранение дефектов типа **{main_threat}**.")
