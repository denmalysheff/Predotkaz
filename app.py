import streamlit as st
import pandas as pd
import numpy as np
import datetime
import io
import gc

# Настройка страницы Streamlit
st.set_page_config(page_title="Комбайн-П: Экспорт", layout="wide")

st.title("🚂 Модуль «Комбайн-П: Аналитический Конвейер»")
st.markdown("### Сверхбыстрая обработка больших архивов с выгрузкой предиктивных отчетов в Excel")

# --- БОКОВАЯ ПАНЕЛЬ С МУЛЬТИЗАГРУЗКОЙ ---
st.sidebar.header("📁 Архив путеизмерителя")
uploaded_files = st.sidebar.file_uploader(
    "Выберите файлы Excel (.xlsx) для глубокого анализа:", 
    type=["xlsx"], 
    accept_multiple_files=True
)

REQUIRED_KM_COLS = ['ГОД', 'МЕСЯЦ', 'КОДНАПР', 'ПУТЬ', 'KM', 'БАЛЛ', 'ПД']
REQUIRED_OTST_COLS = ['ГОД', 'МЕСЯЦ', 'КОДНАПР', 'ПУТЬ', 'KM', 'М', 'ОТСТУПЛЕНИЕ', 'БАЛЛ']

# Инициализация переменных в сессии, чтобы данные не сбрасывались при переключении кнопок
if 'df_km_all' not in st.session_state:
    st.session_state.df_km_all = None
if 'df_otst_all' not in st.session_state:
    st.session_state.df_otst_all = None

if uploaded_files:
    if st.sidebar.button("🚀 Запустить экспресс-анализ архива"):
        km_list = []
        otst_list = []
        success_count = 0
        
        progress_bar = st.progress(0)
        for idx, f in enumerate(uploaded_files):
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
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        if success_count > 0:
            st.session_state.df_km_all = pd.concat(km_list, ignore_index=True)
            st.session_state.df_otst_all = pd.concat(otst_list, ignore_index=True)
            st.sidebar.success(f"📥 База данных собрана! Файлов: {success_count} шт.")
        else:
            st.sidebar.error("Нужные листы не найдены.")

# Параметры фильтрации
st.sidebar.header("⚙️ Настройки отчета")
current_month = st.sidebar.slider("Прогнозный месяц (сезон):", 1, 12, int(datetime.datetime.now().month))
threshold_ball = st.sidebar.number_input("Критический порог баллов:", value=50, step=5)

# --- ОСНОВНОЙ БЛОК ОБРАБОТКИ И ВЫГРУЗКИ ---
if st.session_state.df_km_all is not None:
    # Делаем быстрый срез по выбранному месяцу
    hist_km = st.session_state.df_km_all[st.session_state.df_km_all['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM', 'БАЛЛ']).copy()
    hist_otst = st.session_state.df_otst_all[st.session_state.df_otst_all['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM']).copy()
    
    if hist_km.empty:
        st.warning(f"В загруженном архиве нет данных для месяца № {current_month}.")
    else:
        # 1. СТРАТЕГИЧЕСКИЙ СЛОЙ (Оценка КМ)
        km_profile = hist_km.groupby(['КОДНАПР', 'ПУТЬ', 'KM']).agg(
            Ср_Балл=('БАЛЛ', 'mean'),
            Макс_Балл=('БАЛЛ', 'max'),
            Кол_Проверок=('БАЛЛ', 'count'),
            Превышений=('БАЛЛ', lambda x: (x >= threshold_ball).sum()),
            ПД=('ПД', lambda x: int(x.dropna().iloc[0]) if not x.dropna().empty else 0)
        ).reset_index()
        
        km_profile['Повторяемость_%'] = (km_profile['Превышений'] / km_profile['Кол_Проверок'] * 100).round(1)
        dangerous_kms = km_profile[(km_profile['Ср_Балл'] >= threshold_ball) | (km_profile['Превышений'] >= 2)].sort_values(by='Ср_Балл', ascending=False).copy()
        
        if dangerous_kms.empty:
            st.success("🎉 Все километры стабильны! Рисков предотказов на этот месяц не выявлено.")
        else:
            st.info(f"📊 В зоне риска обнаружено {len(dangerous_kms)} км. Нажмите кнопку ниже для генерации комплексного Excel-отчета.")
            
            # --- КНОПКА ГЕНЕРАЦИИ EXCEL ---
            if st.button("📊 Сгенерировать комплексный отчет Excel"):
                
                # Подготовка ТАКТИЧЕСКОГО СЛОЯ (Причины и адресность)
                summary_rows = []
                assignment_rows = []
                
                # Чтобы не делать тяжелых циклов в Streamlit, обрабатываем только проблемные КМ быстрым вектором
                for _, row in dangerous_kms.iterrows():
                    dkod, put, km = row['КОДНАПР'], row['ПУТЬ'], row['KM']
                    
                    # Фильтруем точечные отступления
                    defects = hist_otst[(hist_otst['КОДНАПР'] == dkod) & (hist_otst['ПУТЬ'] == put) & (hist_otst['KM'] == km)].copy()
                    
                    reason_text = "Общая усталость балластной призмы"
                    main_threat = "Не определено"
                    crit_meters_text = "0 - 1000"
                    
                    if not defects.empty:
                        defects['ОТСТУПЛЕНИЕ'] = defects['ОТСТУПЛЕНИЕ'].astype(str).str.strip()
                        
                        # Определяем главную причину
                        structure = defects.groupby('ОТСТУПЛЕНИЕ').agg(Сумма=('БАЛЛ', 'sum')).sort_values(by='Сумма', ascending=False)
                        if not structure.empty:
                            main_threat = structure.index[0]
                            
                        # Определяем критические метры
                        defects['Участок'] = (defects['М'] // 100) * 100
                        density = defects.groupby('Участок').agg(Баллы=('БАЛЛ', 'sum')).sort_values(by='Баллы', ascending=False)
                        if not density.empty:
                            m_start = int(density.index[0])
                            crit_meters_text = f"{m_start} - {m_start + 100}"
                        
                        reason_text = f"Хронический рост дефектов типа [{main_threat}] на интервале {crit_meters_text} м."
                    
                    # Заполняем 1 вкладку (Сводная)
                    summary_rows.append({
                        'Код направления': int(dkod),
                        'Дорожный мастер (ПД)': f"ПД-{int(row['ПД'])}",
                        'Номер пути': int(put),
                        'Километр': int(km),
                        'Исторический ср. балл': round(row['Ср_Балл'], 1),
                        'Пиковый балл в этот сезон': int(row['Макс_Балл']),
                        'Вероятность отказа (%)': row['Повторяемость_%'],
                        'Цифровой диагноз (Причина риска)': reason_text
                    })
                    
                    # Заполняем 2 вкладку (Адресные наряды для линии)
                    if not defects.empty:
                        # Показываем путейскую сортировку по метрам для нарядов
                        density_all = defects.groupby('Участок').agg(
                            Баллы=('БАЛЛ', 'sum'),
                            Кол=('БАЛЛ', 'count')
                        ).sort_index()
                        
                        for u_metr, d_row in density_all.iterrows():
                            # Ищем преобладающий дефект на этой сотне метров
                            sub_defects = defects[defects['Участок'] == u_metr]
                            loc_threat = sub_defects.groupby('ОТСТУПЛЕНИЕ').agg(s=('БАЛЛ', 'sum')).sort_values(by='s', ascending=False).index[0]
                            
                            assignment_rows.append({
                                'Исполнитель': f"ПД-{int(row['ПД'])}",
                                'Код направления': int(dkod),
                                'Путь': int(put),
                                'Километр': int(km),
                                'Интервал метров': f"{int(u_metr)} - {int(u_metr+100)}",
                                'Количество отступлений': int(d_row['Кол']),
                                'Сумма баллов на участке': int(d_row['Баллы']),
                                'Что устранить в первую очередь': loc_threat
                            })
                
                df_excel_sheet1 = pd.DataFrame(summary_rows)
                df_excel_sheet2 = pd.DataFrame(assignment_rows)
                
                # Запись в буфер памяти
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel_sheet1.to_excel(writer, sheet_name='Сводный отчет по ПЧ', index=False)
                    df_excel_sheet2.to_excel(writer, sheet_name='Адресные задания для ПД', index=False)
                
                st.success("🎉 Отчет успешно сформирован!")
                st.download_button(
                    label="📥 Скачать предотказную ведомость в Excel",
                    data=output.getvalue(),
                    file_name=f"Предотказы_ПЧ_месяц_{current_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("Вы можете использовать демонстрационные данные боковой панели или загрузить свои архивы Excel.")
