import streamlit as st
import pandas as pd
import numpy as np
import datetime

# Настройка страницы Streamlit
st.set_page_config(page_title="Комбайн-П: Предотказ", layout="wide")

st.title("🚂 Модуль «Комбайн-П: Предотказ»")
st.markdown("### Анализ сезонных рисков геометрии пути на основе листов «Оценка КМ» и «Отступления»")

# --- ФУНКЦИИ ГЕНЕРАЦИИ ДЕМО-ДАННЫХ ---
@st.cache_data
def get_demo_data():
    """Генерация тестовых данных, строго повторяющих структуру листов пользователя"""
    np.random.seed(42)
    
    # 1. Имитация листа "Оценка КМ"
    km_rows = []
    directions = [24602, 91022700, 91031600]
    months = [4, 5, 6]  # Апрель, Май, Июнь
    years = [2024, 2025, 2026]
    
    for yr in years:
        for mn in months:
            for day in [8, 22]:  # Два прохода в месяц
                for d_kod in directions:
                    for km in range(2320, 2335):
                        # Делаем 2328-й км на направлении 24602 хронически больным в мае
                        if km == 2328 and d_kod == 24602 and mn == 5:
                            ball = np.random.randint(95, 150)
                            rating = 'У'
                        else:
                            ball = np.random.randint(10, 35)
                            rating = 'Х'
                        
                        km_rows.append({
                            'КОДДОР': 96, 'ПЧ': 22, 'ГОД': yr, 'МЕСЯЦ': mn, 'ДЕНЬ': day,
                            'КОДНАПР': d_kod, 'ПУТЬ': 1, 'KM': km, 'БАЛЛ': ball, 'ОЦЕНКА': rating
                        })
    df_km = pd.DataFrame(km_rows)
    
    # 2. Имитация листа "Отступления"
    otst_rows = []
    types = ['Просадка', 'Уширение', 'Рихтовка', 'Перекос']
    for idx, row in df_km.iterrows():
        if row['БАЛЛ'] > 40:
            num_defects = int(row['БАЛЛ'] / 25)
            for _ in range(num_defects):
                otst_type = 'Просадка' if (row['KM'] == 2328 and row['МЕСЯЦ'] == 5) else np.random.choice(types)
                step = np.random.choice([2, 3], p=[0.8, 0.2])
                b_val = 20 if step == 2 else 50
                otst_rows.append({
                    'ГОД': row['ГОД'], 'МЕСЯЦ': row['МЕСЯЦ'], 'ДЕНЬ': row['ДЕНЬ'],
                    'КОДНАПРВ': row['КОДНАПР'], 'ПУТЬ': row['ПУТЬ'], 'KM': row['KM'],
                    'М': np.random.randint(100, 900), 'ОТСТУПЛЕНИЕ': otst_type,
                    'СТЕПЕНЬ': step, 'БАЛЛ': b_val
                })
    df_otst = pd.DataFrame(otst_rows)
    return df_km, df_otst

# --- БОКОВАЯ ПАНЕЛЬ С ВЫБОРОМ ОДНОГО ФАЙЛА EXCEL ---
st.sidebar.header("📁 Загрузка файла путеизмерителя")
uploaded_file = st.sidebar.file_uploader("Выберите файл Excel (.xlsx)", type=["xlsx"])

df_km_raw = None
df_otst_raw = None

if uploaded_file:
    try:
        # Сначала проверяем, какие листы есть в файле
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        # Поиск нужных листов (с игнорированием регистра букв и пробелов)
        target_km_sheet = next((s for s in sheet_names if s.strip().lower() == "оценка км"), None)
        target_otst_sheet = next((s for s in sheet_names if s.strip().lower() == "отступления"), None)
        
        if target_km_sheet and target_otst_sheet:
            df_km_raw = pd.read_excel(uploaded_file, sheet_name=target_km_sheet)
            df_otst_raw = pd.read_excel(uploaded_file, sheet_name=target_otst_sheet)
            st.sidebar.success("🎉 Успешно! Найдены листы 'Оценка КМ' и 'Отступления'")
        else:
            st.sidebar.error("❌ В файле отсутствуют листы 'Оценка КМ' или 'Отступления'.")
            st.sidebar.info("Доступные листы в файле: " + ", ".join(sheet_names))
            df_km_raw, df_otst_raw = get_demo_data()
            
    except Exception as e:
        st.sidebar.error(f"Ошибка чтения Excel: {e}. Загружены демо-данные.")
        df_km_raw, df_otst_raw = get_demo_data()
else:
    st.sidebar.info("Используются встроенные демонстрационные данные для ПЧ-22.")
    df_km_raw, df_otst_raw = get_demo_data()

# Настройки фильтрации
st.sidebar.header("⚙️ Настройки анализа")
current_month = st.sidebar.slider("Прогноз на месяц:", 1, 12, int(datetime.datetime.now().month))
threshold_ball = st.sidebar.number_input("Сигнальный порог баллов для КМ:", value=50, step=5)

# --- ОБРАБОТКА ДАННЫХ ---

# Очищаем заголовки колонок от пробелов и переводим в верхний регистр
df_km_raw.columns = df_km_raw.columns.str.strip().str.upper()
df_otst_raw.columns = df_otst_raw.columns.str.strip().str.upper()

# Унифицируем колонку направления (в одной таблице может быть КОДНАПР, в другой КОДНАПРВ)
if 'КОДНАПРВ' in df_otst_raw.columns:
    df_otst_raw = df_otst_raw.rename(columns={'КОДНАПРВ': 'КОДНАПР'})

# Принудительно переводим ключевые координаты в числа
for df in [df_km_raw, df_otst_raw]:
    for col in ['КОДНАПР', 'ПУТЬ', 'KM', 'МЕСЯЦ']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

# Фильтруем исторические данные по выбранному месяцу
hist_km = df_km_raw[df_km_raw['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM', 'БАЛЛ'])
hist_otst = df_otst_raw[df_otst_raw['МЕСЯЦ'] == current_month].dropna(subset=['КОДНАПР', 'ПУТЬ', 'KM'])

if hist_km.empty:
    st.warning(f"В базе данных нет исторических записей для месяца № {current_month}.")
else:
    # Строим профиль километров на основе таблицы "Оценка КМ"
    km_profile = hist_km.groupby(['КОДНАПР', 'ПУТЬ', 'KM']).agg(
        Ср_Балл=('БАЛЛ', 'mean'),
        Макс_Балл=('БАЛЛ', 'max'),
        Кол_Проверок=('БАЛЛ', 'count'),
        Превышений_Порога=('БАЛЛ', lambda x: (x >= threshold_ball).sum())
    ).reset_index()

    km_profile['Повторяемость_%'] = (km_profile['Превышений_Порога'] / km_profile['Кол_Проверок'] * 100).round(1)

    # Выделяем предотказные километры
    dangerous_kms = km_profile[
        (km_profile['Ср_Балл'] >= threshold_ball) | (km_profile['Превышений_Порога'] >= 2)
    ].sort_values(by='Ср_Балл', ascending=False)

    # --- ИНТЕРФЕЙС ---
    c1, c2 = st.columns(2)
    with c1:
        st.metric(label="Найдено км в зоне риска предотказа", value=len(dangerous_kms))
    with c2:
        st.success(f"Анализ проведен по архиву из {int(km_profile['Кол_Проверок'].max())} проверок.")

    if dangerous_kms.empty:
        st.success("🎉 Исторических аномалий для этого периода не обнаружено. Все участки стабильны!")
    else:
        st.subheader("📋 Сводная ведомость километров с высоким риском неисправностей")
        
        dangerous_kms['Ср_Балл'] = dangerous_kms['Ср_Балл'].round(1)
        
        st.dataframe(dangerous_kms[[
            'КОДНАПР', 'ПУТЬ', 'KM', 'Ср_Балл', 'Макс_Балл', 'Повторяемость_%'
        ]].rename(columns={
            'КОДНАПР': 'Код направления', 'ПУТЬ': 'Путь', 'KM': 'Километр',
            'Ср_Балл': 'Исторический ср. балл', 'Макс_Балл': 'Пиковый балл',
            'Повторяемость_%': 'Вероятность повторения %'
        }), use_container_width=True, hide_index=True)

        st.subheader("🔍 Что проверить бригаде на линии?")
        
        dangerous_kms['label'] = "Направление " + dangerous_kms['КОДНАПР'].astype(int).astype(str) + ", Путь " + dangerous_kms['ПУТЬ'].astype(int).astype(str) + ", Км " + dangerous_kms['KM'].astype(int).astype(str)
        selected_label = st.selectbox("Выберите километр из списка для детальной раскладки:", dangerous_kms['label'].unique())
        
        sel_row = dangerous_kms[dangerous_kms['label'] == selected_label].iloc[0]
        
        # Выгружаем точечные отступления для выбранного КМ
        km_defects = hist_otst[
            (hist_otst['КОДНАПР'] == sel_row['КОДНАПР']) & 
            (hist_otst['ПУТЬ'] == sel_row['ПУТЬ']) & 
            (hist_otst['KM'] == sel_row['KM'])
        ].copy()
        
        if km_defects.empty:
            st.info("💡 Для данного километра нет детализированных записей по отдельным отступлениям в листе 'Отступления'. Используйте общую статистику баллов.")
        else:
            if 'ОТСТУПЛЕНИЕ' in km_defects.columns:
                km_defects['ОТСТУПЛЕНИЕ'] = km_defects['ОТСТУПЛЕНИЕ'].astype(str).str.strip()
            
            dl, dr = st.columns(2)
            
            with dl:
                st.markdown("**Характерные виды неисправностей (по сумме баллов):**")
                structure = km_defects.groupby('ОТСТУПЛЕНИЕ').agg(
                    Суммарный_Балл=('БАЛЛ', 'sum'),
                    Количество_Случаев=('ОТСТУПЛЕНИЕ', 'count')
                ).sort_values(by='Суммарный_Балл', ascending=False).reset_index()
                st.dataframe(structure.rename(columns={'ОТСТУПЛЕНИЕ': 'Вид неисправности', 'Суммарный_Балл': 'Набрано баллов', 'Количество_Случаев': 'Кол-во случаев'}), use_container_width=True, hide_index=True)
                
            with dr:
                st.markdown("**Критические участки километра (Привязка к метрам):**")
                km_defects['Участок_Метры'] = (km_defects['М'] // 100) * 100
                density = km_defects.groupby('Участок_Метры').agg(
                    Баллы=('БАЛЛ', 'sum'),
                    Количество=('БАЛЛ', 'count')
                ).sort_values(by='Баллы', ascending=False).reset_index()
                
                density['Интервал (м)'] = density['Участок_Метры'].astype(int).astype(str) + " - " + (density['Участок_Метры'] + 100).astype(int).astype(str)
                st.dataframe(density[['Интервал (м)', 'Баллы', 'Количество']].rename(columns={'Баллы': 'Сумма баллов', 'Количество': 'Кол-во дефектов'}), use_container_width=True, hide_index=True)

            if not structure.empty and not density.empty:
                main_threat = structure.iloc[0]['Вид неисправности']
                crit_meters = density.iloc[0]['Интервал (м)']
                st.warning(f"📋 **Рекомендация для дорожного мастера:** На {int(sel_row['KM'])} км в этом месяце ожидается ухудшение пути. "
                           f"Выдайте задание бригадиру (ПДБ) проверить в первую очередь интервал **{crit_meters} метров**. "
                           f"Основное внимание обратить на поиск и устранение дефектов типа **{main_threat}**.")
