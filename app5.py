import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# --- 1. НАСТРОЙКИ СТРАНИЦЫ И ШРИФТ ---
st.set_page_config(page_title="HR Дашборд", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    html, body, [class*="css"] { font-family: 'Verdana', sans-serif; }
    .main { background-color: #f4f6f9; }
    h1, h2, h3 { color: #333333; font-family: 'Verdana', sans-serif; }
    .stAlert { background-color: #ffcccc; color: #990000; }
    
    /* ===== СТИЛИЗАЦИЯ БОКОВОЙ ПАНЕЛИ ===== */
    section[data-testid="stSidebar"] {
        background-color: #f4f6f9;
        padding-top: 20px;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 0; padding-left: 15px; padding-right: 15px;
    }
    section[data-testid="stSidebar"] label {
        font-family: 'Verdana', sans-serif;
        font-size: 15px; font-weight: 500;
        color: #333; margin-bottom: 5px;
    }
    section[data-testid="stSidebar"] .stSelectbox > div > div {
        background-color: #ffffff;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        font-family: 'Verdana', sans-serif;
        font-size: 14px; color: #333;
    }
    section[data-testid="stSidebar"] .stSelectbox > div > div:hover {
        border-color: #999;
    }
    section[data-testid="stSidebar"] h2 {
        font-size: 20px; color: #333; margin-bottom: 20px;
    }
    
    /* ===== Карточки подписей под диаграммой ===== */
    .legend-block { 
        text-align: left; padding: 2px 2px; 
        background-color: transparent; margin: 0;
        font-size: 14px; /* Увеличено на 2 кегля (было 12) */
    }
    .legend-title { font-weight: bold; margin-bottom: 0; display: inline; }
    .legend-value { font-weight: normal; display: inline; }
    
    /* ===== Индексы ===== */
    .index-value { 
        text-align: center; font-size: 26px; /* Увеличено на 2 кегля (было 24) */
        font-weight: bold; 
        color: #333; font-family: Verdana;
        margin-top: 5px; margin-bottom: 0;
    }
    .index-label { 
        text-align: center; font-size: 16px; /* Увеличено на 2 кегля (было 14) */
        color: #333; 
        font-family: Verdana; margin-top: 0; font-weight: bold;
    }
    
    /* ===== Иконки М/Ж ===== */
    .icon-container {
        text-align: center;
        margin-top: 5px;
    }
    .icon-svg {
        width: 80px;
        height: 80px;
    }
    .icon-pct {
        font-size: 24px; /* Увеличено на 2 кегля (было 22) */
        font-weight: bold;
        font-family: Verdana;
        margin-top: 5px;
    }

    /* ===== Блок с выводами ===== */
    .conclusion-box {
        background-color: #e8f0fe;
        border-left: 5px solid #3498db;
        padding: 15px 20px;
        border-radius: 5px;
        font-family: 'Verdana', sans-serif;
        font-size: 15px;
        color: #333;
        margin-top: 10px;
        margin-bottom: 20px;
    }
    .conclusion-title {
        font-weight: bold;
        color: #2980b9;
        margin-bottom: 5px;
        font-size: 16px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. ПАРОЛИ И РОЛИ ---
ADMIN_PASSWORD = "admin123"
USER_PASSWORD = "user123"

DATA_FILE_PATH = "data/data.xlsx"
os.makedirs("data", exist_ok=True)

# --- 3. АВТОРИЗАЦИЯ ---
if "role" not in st.session_state:
    st.session_state.role = None

def login_screen():
    st.title("Вход в дашборд") # Убрана иконка
    st.markdown("Выберите роль и введите пароль:")
    
    col1, col2 = st.columns(2)
    with col1:
        role_choice = st.radio("Роль:", ["Пользователь", "Администратор"])
    with col2:
        password_input = st.text_input("Пароль:", type="password")
    
    if st.button("Войти"):
        if role_choice == "Администратор" and password_input == ADMIN_PASSWORD:
            st.session_state.role = "admin"; st.rerun()
        elif role_choice == "Пользователь" and password_input == USER_PASSWORD:
            st.session_state.role = "user"; st.rerun()
        else:
            st.error("Неверный пароль. Попробуйте еще раз.") # Убрана иконка

if st.session_state.role is None:
    login_screen()
    st.stop()

# --- 4. РОЛЬ И ВЫХОД ---
st.sidebar.markdown("---")
if st.session_state.role == "admin":
    st.sidebar.success("Вы вошли как **Администратор**") # Убрана иконка
else:
    st.sidebar.info("Вы вошли как **Пользователь**") # Убрана иконка

if st.sidebar.button("Выйти"): # Убрана иконка
    st.session_state.role = None
    st.rerun()

# --- 5. ЗАГРУЗКА ДАННЫХ ---
df = None
df_conclusions = None 

def load_excel_data(file_path_or_buffer):
    try:
        xls = pd.ExcelFile(file_path_or_buffer)
        sheet_base = 'База для расчетов' if 'База для расчетов' in xls.sheet_names else xls.sheet_names[0]
        df_base = pd.read_excel(xls, sheet_name=sheet_base)
        
        if 'Выводы' in xls.sheet_names:
            df_conc = pd.read_excel(xls, sheet_name='Выводы')
        else:
            df_conc = pd.DataFrame(columns=['Подразделение', 'Вывод'])
            st.warning("Лист 'Выводы' не найден в файле. Блок с выводами будет пуст.") # Убрана иконка
            
        return df_base, df_conc
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}") # Убрана иконка
        st.stop()

if st.session_state.role == "admin":
    st.title("Аналитический дашборд HR (Администратор)") # Убрана иконка
    uploaded_file = st.file_uploader("Загрузите Excel-файл с данными", type=["xlsx"])
    
    if uploaded_file is not None:
        with open(DATA_FILE_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"Файл загружен и сохранен (все листы).") # Убрана иконка
        
        df, df_conclusions = load_excel_data(DATA_FILE_PATH)
    else:
        if os.path.exists(DATA_FILE_PATH):
            df, df_conclusions = load_excel_data(DATA_FILE_PATH)
            st.info("Используются данные из последней загрузки.") # Убрана иконка
        else:
            st.warning("Нет данных. Загрузите Excel-файл.") # Убрана иконка
            st.stop()

elif st.session_state.role == "user":
    st.title("Аналитический дашборд HR") # Убрана иконка
    if os.path.exists(DATA_FILE_PATH):
        df, df_conclusions = load_excel_data(DATA_FILE_PATH)
        st.caption("Данные предоставлены администратором.") # Убрана иконка
    else:
        st.warning("Данные еще не загружены администратором.") # Убрана иконка
        st.stop()

if df is None:
    st.stop()

# --- 6. ПРОВЕРКА КОЛОНОК ---
required_columns = [
    'Уникальные', 'Подразделение', 'Пол', 'Возраст', 'Стаж',
    'Готовность делать больше', 'Удовлетворение от работы', 'NPS компании',
    'Вероятность увольнения', 'Уровень стресса', 'Мотиваторы',
    'Удовлетворенность Мотиваторами', 'Рейтинг мотивов', 'Рейтинг барьеров',
    'NPS драйверы рейтинг', 'NPS антипричины рейтинг'
]

missing_cols = [col for col in required_columns if col not in df.columns]
if missing_cols:
    st.error(f"Ошибка: В файле отсутствуют обязательные колонки: {', '.join(missing_cols)}") # Убрана иконка
    st.stop()

df['Уникальные'] = pd.to_numeric(df['Уникальные'], errors='coerce').fillna(0)

# --- 7. ФИЛЬТРЫ ---
st.sidebar.markdown("## Фильтры") # Убрана иконка

dept_options = ["Все"] + sorted(df['Подразделение'].dropna().unique().tolist())
gender_options = ["Все"] + sorted(df['Пол'].dropna().unique().tolist())

try:
    age_options = ["Все"] + sorted([int(a) for a in df['Возраст'].dropna().unique()])
except:
    age_options = ["Все"] + sorted(df['Возраст'].dropna().unique().tolist())

exp_order = ['До 1 года', '1-3 года', '3-5 лет', '5-10 лет', '10-15 лет', 'более 15 лет']
exp_options = ["Все"] + [exp for exp in exp_order if exp in df['Стаж'].unique()]

selected_dept = st.sidebar.selectbox("Подразделение", options=dept_options)
selected_gender = st.sidebar.selectbox("Пол", options=gender_options)
selected_age = st.sidebar.selectbox("Возраст", options=age_options)
selected_exp = st.sidebar.selectbox("Стаж", options=exp_options)

# --- ПРИМЕНЯЕМ ФИЛЬТРЫ ---
filtered_df = df.copy()

if selected_dept != "Все":
    filtered_df = filtered_df[filtered_df['Подразделение'] == selected_dept]
if selected_gender != "Все":
    filtered_df = filtered_df[filtered_df['Пол'] == selected_gender]
if selected_age != "Все":
    try:
        filtered_df = filtered_df[filtered_df['Возраст'].astype(int) == int(selected_age)]
    except:
        filtered_df = filtered_df[filtered_df['Возраст'] == selected_age]
if selected_exp != "Все":
    filtered_df = filtered_df[filtered_df['Стаж'] == selected_exp]

# --- 8. ПРОВЕРКА МИНИМУМА ---
unique_count = filtered_df[filtered_df['Уникальные'] == 1].shape[0]

if unique_count < 10:
    st.warning("Отчет не сформирован в связи с недостаточным количеством ответов (менее 10 уникальных респондентов).") # Убрана иконка
    st.stop()

# --- 9. РАСЧЕТЫ ---
unique_df = filtered_df[filtered_df['Уникальные'] == 1]
total_unique = len(unique_df)
male_count = len(unique_df[unique_df['Пол'] == 'Мужской'])
female_count = len(unique_df[unique_df['Пол'] == 'Женский'])

male_pct = (male_count / total_unique * 100) if total_unique > 0 else 0
female_pct = (female_count / total_unique * 100) if total_unique > 0 else 0

def get_engagement_category(row):
    readiness = row['Готовность делать больше']
    satisfaction = row['Удовлетворение от работы']
    try: readiness = int(readiness)
    except: return None 
    if readiness in [1, 2]:
        if satisfaction in ['Всегда', 'Большую часть времени']: return 'Мотивированные'
        elif satisfaction in ['Иногда', 'Никогда', 'Половину времени']: return 'Разочарованные'
    elif readiness in [3, 4]:
        if satisfaction in ['Всегда', 'Большую часть времени']: return 'Довольные'
        elif satisfaction in ['Иногда', 'Никогда', 'Половину времени']: return 'Отчужденные'
    return None

filtered_df['Категория_вовлеченности'] = filtered_df.apply(get_engagement_category, axis=1)

engagement_counts = filtered_df['Категория_вовлеченности'].value_counts()
total_engagement = engagement_counts.sum()

motiv_pct = (engagement_counts.get('Мотивированные', 0) / total_engagement * 100) if total_engagement > 0 else 0
disap_pct = (engagement_counts.get('Разочарованные', 0) / total_engagement * 100) if total_engagement > 0 else 0
satisf_pct = (engagement_counts.get('Довольные', 0) / total_engagement * 100) if total_engagement > 0 else 0
alien_pct = (engagement_counts.get('Отчужденные', 0) / total_engagement * 100) if total_engagement > 0 else 0

engagement_index = motiv_pct + disap_pct

nps_counts = filtered_df['NPS компании'].value_counts()
total_nps = nps_counts.sum()

promoters = (nps_counts.get('Сторонники', 0) / total_nps * 100) if total_nps > 0 else 0
neutrals = (nps_counts.get('Нейтральные', 0) / total_nps * 100) if total_nps > 0 else 0
critics = (nps_counts.get('Критики', 0) / total_nps * 100) if total_nps > 0 else 0

enps_index = promoters - critics

exp_counts = filtered_df['Стаж'].value_counts(normalize=True) * 100
exp_data = []
for exp in exp_order:
    if exp in exp_counts:
        exp_data.append({'Стаж': exp, 'Доля': exp_counts[exp]})
exp_df = pd.DataFrame(exp_data)

# --- 10. ВИЗУАЛИЗАЦИЯ ---
tab1, tab2, tab3 = st.tabs(["Основные показатели", "Вкладка 2 (Заглушка)", "Вкладка 3 (Заглушка)"]) # Убраны иконки

with tab1:
    # === БЛОК С ВЫВОДАМИ ===
    if selected_dept != "Все" and not df_conclusions.empty:
        conclusion_row = df_conclusions[df_conclusions['Подразделение'] == selected_dept]
        if not conclusion_row.empty:
            text = conclusion_row.iloc[0]['Вывод']
            if pd.notna(text) and str(text).strip() != "":
                st.markdown(f"""
                <div class='conclusion-box'>
                    <div class='conclusion-title'>Вывод по подразделению: {selected_dept}</div>
                    {text}
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info(f"Для подразделения '{selected_dept}' вывод не заполнен.") # Убрана иконка
        else:
            st.info(f"Для подразделения '{selected_dept}' вывод не найден в файле.") # Убрана иконка
    elif selected_dept == "Все":
        st.info("Выберите конкретное подразделение в фильтре, чтобы увидеть аналитический вывод.") # Убрана иконка

    # === Блок 1: Состав респондентов ===
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("### Количество и состав респондентов")
        st.markdown(f"<h1 style='text-align: center; color: #333; font-family: Verdana; margin-bottom: 5px;'>{unique_count}</h1>", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        
        male_svg = """
        <div class='icon-container'>
            <svg class='icon-svg' viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
                <ellipse cx="50" cy="28" rx="18" ry="22" fill="#7da2ce"/>
                <path d="M 32 22 Q 30 10 50 8 Q 70 10 68 22 Q 65 15 50 14 Q 35 15 32 22 Z" fill="#5b7ba0"/>
                <path d="M 15 100 Q 15 65 35 58 L 65 58 Q 85 65 85 100 Z" fill="#7da2ce"/>
                <path d="M 42 58 L 50 75 L 58 58 Z" fill="#ffffff"/>
                <path d="M 48 65 L 50 72 L 52 65 L 50 60 Z" fill="#5b7ba0"/>
            </svg>
            <div class='icon-pct' style='color: #7da2ce;'>{:.1f}%</div>
        </div>
        """.format(male_pct)
        
        female_svg = """
        <div class='icon-container'>
            <svg class='icon-svg' viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
                <ellipse cx="50" cy="28" rx="17" ry="21" fill="#e89a9a"/>
                <path d="M 28 25 Q 25 8 50 6 Q 75 8 72 25 L 75 60 Q 70 45 68 30 Q 65 18 50 16 Q 35 18 32 30 Q 30 45 25 60 Z" fill="#c97a7a"/>
                <path d="M 15 100 Q 15 68 35 60 L 65 60 Q 85 68 85 100 Z" fill="#e89a9a"/>
                <path d="M 38 60 L 50 78 L 62 60 L 55 58 L 50 65 L 45 58 Z" fill="#ffffff"/>
            </svg>
            <div class='icon-pct' style='color: #e89a9a;'>{:.1f}%</div>
        </div>
        """.format(female_pct)
        
        with c1:
            st.markdown(male_svg, unsafe_allow_html=True)
        with c2:
            st.markdown(female_svg, unsafe_allow_html=True)

    with col2:
        st.markdown("### Стаж респондентов")
        if not exp_df.empty:
            fig_exp = px.bar(exp_df, x='Стаж', y='Доля', text='Доля', color_discrete_sequence=['#ff4b4b'])
            fig_exp.update_traces(texttemplate='%{text:.1f}%', textposition='outside', textfont_size=16) # Увеличен шрифт
            fig_exp.update_layout(height=300, margin=dict(t=20, b=20, l=0, r=0), showlegend=False, font_family="Verdana")
            st.plotly_chart(fig_exp, use_container_width=True)
        else:
            st.info("Нет данных по стажу")

    st.markdown("---")
    
    # --- Блок 2: КРУГОВЫЕ ДИАГРАММЫ (ПОНЧИКИ) ---
    col_g1, col_g2 = st.columns(2)

    # === ПРОФИЛЬ МОТИВАЦИИ ===
    with col_g1:
        st.markdown("<h3 style='text-align: center;'>Профиль мотивации</h3>", unsafe_allow_html=True)
        
        eng_labels = ['Вовлеченные', 'Разочарованные', 'Довольные', 'Отчужденные']
        eng_values = [motiv_pct, disap_pct, satisf_pct, alien_pct]
        eng_colors = ['#5d8c3a', '#ff1a1a', '#366cb3', '#3d3d3d']
        
        fig_eng = go.Figure(data=[go.Pie(
            labels=eng_labels,
            values=eng_values,
            hole=0.6,
            marker=dict(colors=eng_colors),
            textinfo='none',
            direction='clockwise',
            sort=False
        )])
        
        fig_eng.update_layout(
            height=250,
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=False
        )
        st.plotly_chart(fig_eng, use_container_width=True)
        
        # Подписи в одну строку под графиком (шрифт увеличен до 14px)
        st.markdown(f"""
        <div style="display: flex; justify-content: center; gap: 15px; font-size: 14px; font-family: Verdana; flex-wrap: wrap;">
            <div style="display: flex; align-items: center;"><span style="color: #5d8c3a; font-size: 18px; margin-right: 5px;">■</span> {eng_labels[0]}; {eng_values[0]:.0f}%</div>
            <div style="display: flex; align-items: center;"><span style="color: #ff1a1a; font-size: 18px; margin-right: 5px;">■</span> {eng_labels[1]}; {eng_values[1]:.0f}%</div>
            <div style="display: flex; align-items: center;"><span style="color: #366cb3; font-size: 18px; margin-right: 5px;">■</span> {eng_labels[2]}; {eng_values[2]:.0f}%</div>
            <div style="display: flex; align-items: center;"><span style="color: #3d3d3d; font-size: 18px; margin-right: 5px;">■</span> {eng_labels[3]}; {eng_values[3]:.0f}%</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"<div class='index-value'>{engagement_index:.0f}%</div>", unsafe_allow_html=True)
        st.markdown("<div class='index-label'>Индекс вовлеченности</div>", unsafe_allow_html=True)

    # === ПРОФИЛЬ eNPS ===
    with col_g2:
        st.markdown("<h3 style='text-align: center;'>Профиль eNPS</h3>", unsafe_allow_html=True)
        
        nps_labels = ['Сторонники', 'Нейтральные', 'Критики']
        nps_values = [promoters, neutrals, critics]
        nps_colors = ['#4b2e83', '#9ca6b4', '#d16b9b']
        
        fig_nps = go.Figure(data=[go.Pie(
            labels=nps_labels,
            values=nps_values,
            hole=0.6,
            marker=dict(colors=nps_colors),
            textinfo='none',
            direction='clockwise',
            sort=False
        )])
        
        fig_nps.update_layout(
            height=250,
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=False
        )
        st.plotly_chart(fig_nps, use_container_width=True)
        
        # Подписи в одну строку под графиком (шрифт увеличен до 14px)
        st.markdown(f"""
        <div style="display: flex; justify-content: center; gap: 15px; font-size: 14px; font-family: Verdana; flex-wrap: wrap;">
            <div style="display: flex; align-items: center;"><span style="color: #4b2e83; font-size: 18px; margin-right: 5px;">■</span> {nps_labels[0]}; {nps_values[0]:.0f}%</div>
            <div style="display: flex; align-items: center;"><span style="color: #9ca6b4; font-size: 18px; margin-right: 5px;">■</span> {nps_labels[1]}; {nps_values[1]:.0f}%</div>
            <div style="display: flex; align-items: center;"><span style="color: #d16b9b; font-size: 18px; margin-right: 5px;">■</span> {nps_labels[2]}; {nps_values[2]:.0f}%</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"<div class='index-value'>{enps_index:.0f}%</div>", unsafe_allow_html=True)
        st.markdown("<div class='index-label'>Индекс eNPS</div>", unsafe_allow_html=True)

# Обрати внимание: я поменял название вкладки на "Мотивы и барьеры"
tab1, tab2, tab3 = st.tabs(["Основные показатели", "Мотивы и барьеры", "Вкладка 3 (Заглушка)"])

with tab2:
    # Создаем сетку 2x2 для графиков
    col1, col2 = st.columns(2)
    
    # --- ГРАФИК 1: Рейтинг факторов мотивации (Левый верх) ---
    with col1:
        st.markdown("#### Рейтинг факторов мотивации")
        
        # Умный поиск мотивов
        MOTIVS_KEYS = {
            'Вознаграждение в виде премий за высокие результаты работы': 'Вознаграждение в виде премий',
            'Осознание значимости своей работы для Компании': 'Осознание значимости',
            'Ощущение командного духа и поддержки со стороны коллег': 'Ощущение командного духа',
            'Справедливая заработная плата за свой труд': 'Справедливая заработная плата',
            'Чёткое понимание своих рабочих задач': 'Чёткое понимание своих рабочих задач'
        }
        
        motiv_data = []
        for m_full, m_key in MOTIVS_KEYS.items():
            # Ищем вхождение ключевой фразы (без учета регистра)
            count = len(filtered_df[filtered_df['Рейтинг мотивов'].astype(str).str.contains(m_key, case=False, na=False)])
            pct = (count / total_unique * 100) if total_unique > 0 else 0
            motiv_data.append({'Мотив': m_full, 'Доля': pct})
        
        motiv_df_plot = pd.DataFrame(motiv_data)
        motiv_df_plot = motiv_df_plot.sort_values('Доля', ascending=True)
        
        fig_motiv = px.bar(
            motiv_df_plot, 
            x='Доля', 
            y='Мотив', 
            orientation='h',
            text='Доля',
            color_discrete_sequence=['#5b9bd5'] # Синий
        )
        fig_motiv.update_traces(texttemplate='%{text:.0f}%', textposition='inside', textfont_size=12)
        fig_motiv.update_layout(
            height=350,
            margin=dict(t=10, b=10, l=10, r=10),
            xaxis_title="",
            yaxis_title="",
            font_family="Verdana",
            plot_bgcolor='white',
            showlegend=False
        )
        fig_motiv.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#e0e0e0')
        st.plotly_chart(fig_motiv, use_container_width=True)

    # --- ГРАФИК 2: Удовлетворенность мотивами (Правый верх) ---
    with col2:
        st.markdown("#### Удовлетворенность факторами мотивации")
        
        sat_order = ['Согласен полностью', 'Скорее согласен, чем не согласен', 'Скорее не согласен, чем согласен', 'Вообще не согласен']
        
        sat_data = []
        for m_full, m_key in MOTIVS_KEYS.items():
            # Ищем вхождение ключевой фразы в колонке Мотиваторы
            m_df = filtered_df[filtered_df['Мотиваторы'].astype(str).str.contains(m_key, case=False, na=False)]
            row = {'Мотив': m_full}
            for ans in sat_order:
                count = len(m_df[m_df['Удовлетворенность Мотиваторами'] == ans])
                pct = (count / total_unique * 100) if total_unique > 0 else 0
                row[ans] = pct
            sat_data.append(row)
            
        motiv_sat_pct = pd.DataFrame(sat_data)
        # Выравниваем порядок мотивов как в первом графике (сверху вниз)
        motiv_sat_pct = motiv_sat_pct.set_index('Мотив').reindex(motiv_df_plot['Мотив']).reset_index()

        fig_sat = go.Figure()
        colors_sat = {
            'Согласен полностью': '#2e7d32', 
            'Скорее согласен, чем не согласен': '#81c784', 
            'Скорее не согласен, чем согласен': '#9e9e9e', 
            'Вообще не согласен': '#d32f2f' 
        }
        
        for col in sat_order:
            texts = [f"{val:.0f}%" if val > 5 else "" for val in motiv_sat_pct[col]]
            
            fig_sat.add_trace(go.Bar(
                name=col,
                y=motiv_sat_pct['Мотив'],
                x=motiv_sat_pct[col],
                orientation='h',
                marker_color=colors_sat[col],
                text=texts,
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(size=11, color='white')
            ))
            
        fig_sat.update_layout(
            barmode='stack',
            height=350,
            margin=dict(t=10, b=10, l=10, r=10),
            xaxis_title="",
            yaxis_title="",
            font_family="Verdana",
            plot_bgcolor='white',
            showlegend=False
        )
        st.plotly_chart(fig_sat, use_container_width=True)
        
        # === РИСУЕМ СВОЮ ЛЕГЕНДУ В ОДНУ СТРОКУ ПОД ГРАФИКОМ ===
        st.markdown(f"""
        <div style="display: flex; justify-content: center; gap: 15px; font-size: 12px; font-family: Verdana; margin-top: -10px; flex-wrap: wrap;">
            <div style="display: flex; align-items: center;"><span style="color: #2e7d32; font-size: 18px; margin-right: 5px;">■</span> Согласен полностью</div>
            <div style="display: flex; align-items: center;"><span style="color: #81c784; font-size: 18px; margin-right: 5px;">■</span> Скорее согласен, чем не согласен</div>
            <div style="display: flex; align-items: center;"><span style="color: #9e9e9e; font-size: 18px; margin-right: 5px;">■</span> Скорее не согласен, чем согласен</div>
            <div style="display: flex; align-items: center;"><span style="color: #d32f2f; font-size: 18px; margin-right: 5px;">■</span> Вообще не согласен</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- ГРАФИК 3: Барьеры эффективности (Левый низ) ---
    col3, col4 = st.columns(2)
    
    with col3:
        st.markdown("#### Барьеры эффективности")
        
        # Умный поиск барьеров
        BARRIERS_KEYS = {
            'Нет, я не ощущаю значительных препятствий для эффективной работы': 'Нет, я не ощущаю',
            'Непонятная система подсчета заработной платы и удержаний': 'Непонятная система подсчета',
            'Нехватка людей': 'Нехватка людей',
            'Неэффективная организация работы в отделе': 'Неэффективная организация',
            'Проблемы с коммуникацией и координацией между подразделениями и вертикально': 'Проблемы с коммуникацией'
        }
        
        barrier_data = []
        for b_full, b_key in BARRIERS_KEYS.items():
            # Ищем вхождение ключевой фразы
            count = len(filtered_df[filtered_df['Рейтинг барьеров'].astype(str).str.contains(b_key, case=False, na=False)])
            pct = (count / total_unique * 100) if total_unique > 0 else 0
            barrier_data.append({'Барьер': b_full, 'Доля': pct})
            
        barrier_df_plot = pd.DataFrame(barrier_data)
        barrier_df_plot = barrier_df_plot.sort_values('Доля', ascending=True)
        
        fig_barrier = px.bar(
            barrier_df_plot, 
            x='Доля', 
            y='Барьер', 
            orientation='h',
            text='Доля',
            color_discrete_sequence=['#ab47bc'] # Сиреневый
        )
        fig_barrier.update_traces(texttemplate='%{text:.0f}%', textposition='inside', textfont_size=12)
        fig_barrier.update_layout(
            height=350,
            margin=dict(t=10, b=10, l=10, r=10),
            xaxis_title="",
            yaxis_title="",
            font_family="Verdana",
            plot_bgcolor='white'
        )
        fig_barrier.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#e0e0e0')
        st.plotly_chart(fig_barrier, use_container_width=True)

    # --- ГРАФИК 4: Оценка уровня стресса (Правый низ) ---
    with col4:
        st.markdown("#### Оценка уровня стресса")
        
        stress_df = filtered_df[filtered_df['Уникальные'] == 1]['Уровень стресса'].dropna()
        stress_counts = stress_df.value_counts().reset_index()
        stress_counts.columns = ['Оценка', 'Количество']
        
        stress_counts['Оценка'] = stress_counts['Оценка'].astype(str).str.replace('.0', '')
        
        stress_counts['Доля'] = (stress_counts['Количество'] / total_unique * 100).round(0)
        
        stress_counts = stress_counts.sort_values('Оценка', ascending=False)
        
        stress_colors = {
            '1': '#43a047', 
            '2': '#a5d6a7', 
            '3': '#ffcc80', 
            '4': '#ff8a65', 
            '5': '#d84315' 
        }
        
        fig_stress = go.Figure()
        for _, row in stress_counts.iterrows():
            val = row['Оценка']
            color = stress_colors.get(val, '#9e9e9e')
            fig_stress.add_trace(go.Bar(
                y=[val],
                x=[row['Доля']],
                orientation='h',
                marker_color=color,
                text=f"{row['Доля']:.0f}%",
                textposition='inside',
                textfont=dict(size=14),
                showlegend=False
            ))
            
        fig_stress.update_layout(
            height=350,
            margin=dict(t=10, b=10, l=10, r=10),
            xaxis_title="",
            yaxis_title="",
            font_family="Verdana",
            plot_bgcolor='white',
            bargap=0.4
        )
        fig_stress.update_xaxes(showgrid=True, gridwidth=1, gridcolor='#e0e0e0')
        st.plotly_chart(fig_stress, use_container_width=True)
    
with tab3:
    st.info("Здесь будут таблицы с деталями (в разработке)")