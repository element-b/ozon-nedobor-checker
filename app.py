import hmac
import io
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# НАСТРОЙКА СТРАНИЦЫ
# ============================================================

st.set_page_config(
    page_title="Контроль недобора Ozon",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# НАСТРОЙКИ СТАНДАРТНОЙ СТРУКТУРЫ ФАЙЛОВ
# ============================================================

# Файл A — исходный файл МойСклад.
# Excel-колонки:
# D = индекс 3 в Python.
# H = индекс 7 в Python.
# Данные начинаются с Excel-строки 14, то есть index 13.
SOURCE_A_ARTICLE_COLUMN = 3
SOURCE_A_QUANTITY_COLUMN = 7
SOURCE_A_START_ROW = 13

# Файл B — файл Ozon.
# Excel-колонки:
# D = индекс 3 в Python.
# F = индекс 5 в Python.
# Данные начинаются с Excel-строки 2, то есть index 1.
SOURCE_B_ARTICLE_COLUMN = 3
SOURCE_B_QUANTITY_COLUMN = 5
SOURCE_B_START_ROW = 1


# ============================================================
# SESSION STATE
# ============================================================

def init_session_state() -> None:
    """Создаёт начальные переменные сессии."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if "comparison_result" not in st.session_state:
        st.session_state.comparison_result = None

    if "comparison_meta" not in st.session_state:
        st.session_state.comparison_meta = None


init_session_state()


# ============================================================
# СТИЛИ
# ============================================================

def apply_main_styles() -> None:
    """Основные стили авторизованной части приложения."""
    st.markdown(
        """
        <style>
            .stApp {
                background-color: #FFFFFF;
                color: #000000;
            }

            .main .block-container {
                padding: 2rem 3rem 3rem 3rem;
                max-width: 1450px;
            }

            .stApp h1,
            .stApp h2,
            .stApp h3 {
                color: #000000 !important;
                font-weight: 650;
            }

            .stApp p,
            .stApp span,
            .stApp div {
                color: #000000;
            }

            section[data-testid="stSidebar"] {
                background-color: #000000;
                border-right: 1px solid #262626;
                padding-top: 1rem;
            }

            section[data-testid="stSidebar"] * {
                color: #FFFFFF !important;
            }

            section[data-testid="stSidebar"] .stButton > button {
                width: 100%;
                background-color: transparent;
                color: #FFFFFF !important;
                border: 1px solid #4A4A4A;
                border-radius: 8px;
                padding: 11px 14px;
                font-size: 15px;
                font-weight: 500;
                transition: all 0.2s ease-in-out;
            }

            section[data-testid="stSidebar"] .stButton > button:hover {
                background-color: #202020;
                border-color: #777777;
                transform: translateX(2px);
            }

            .stButton > button {
                background-color: #009B77;
                color: #FFFFFF !important;
                border: none;
                border-radius: 7px;
                padding: 9px 16px;
                font-weight: 600;
                transition: all 0.2s ease-in-out;
            }

            .stButton > button:hover {
                background-color: #008268;
                box-shadow: 0 3px 10px rgba(0, 155, 119, 0.28);
            }

            .upload-card {
                border: 1px solid #E4E4E7;
                border-radius: 12px;
                background-color: #FAFAFA;
                padding: 20px;
                min-height: 220px;
            }

            .description-box {
                border-left: 4px solid #009B77;
                background: #F2FBF7;
                border-radius: 6px;
                padding: 14px 18px;
                margin-bottom: 24px;
                color: #1F2937;
            }

            .metric-card-note {
                color: #555555 !important;
                font-size: 0.9rem;
            }

            div[data-testid="metric-container"] {
                background-color: #FAFAFA;
                border: 1px solid #E5E7EB;
                border-radius: 10px;
                padding: 12px 16px;
            }

            [data-testid="stMetricValue"] {
                color: #000000 !important;
                font-size: 1.45rem !important;
                font-weight: 650 !important;
            }

            [data-testid="stMetricLabel"] {
                color: #444444 !important;
            }

            .stDownloadButton > button {
                width: 100%;
                background-color: #009B77 !important;
                color: #FFFFFF !important;
                border: none !important;
                border-radius: 7px !important;
                padding: 10px 16px !important;
                font-weight: 600 !important;
            }

            .stDownloadButton > button:hover {
                background-color: #008268 !important;
            }

            .streamlit-expanderHeader {
                background-color: #F8F9FA;
                color: #000000 !important;
                border-radius: 8px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def apply_login_styles() -> None:
    """Стили страницы входа в визуальном стиле исходного проекта."""
    st.markdown(
        """
        <style>
            header[data-testid="stHeader"],
            section[data-testid="stSidebar"] {
                display: none !important;
            }

            .stApp {
                background:
                    radial-gradient(
                        circle at top left,
                        rgba(80, 200, 120, 0.16),
                        transparent 35%
                    ),
                    radial-gradient(
                        circle at bottom right,
                        rgba(0, 155, 119, 0.12),
                        transparent 38%
                    ),
                    linear-gradient(
                        135deg,
                        #0B0915 0%,
                        #151226 55%,
                        #0B1220 100%
                    );
            }

            .main .block-container {
                padding: 0 !important;
                max-width: 100% !important;
            }

            .login-container div[data-testid="stForm"] {
                background-color: rgba(21, 18, 38, 0.96);
                padding: 2.5rem;
                border-radius: 14px;
                border: 1px solid rgba(255, 255, 255, 0.09);
                box-shadow: 0 18px 48px rgba(0, 0, 0, 0.4);
            }

            .login-container input {
                background-color: #151226 !important;
                border: 1px solid #34304B !important;
                color: #FFFFFF !important;
                border-radius: 7px !important;
            }

            .login-container input::placeholder {
                color: #A4A1B4 !important;
            }

            .login-container .stButton > button {
                width: 100%;
                background: transparent !important;
                color: #50C878 !important;
                border: 2px solid #50C878 !important;
                border-radius: 8px !important;
                padding: 11px !important;
                font-weight: 650 !important;
                transition: all 0.25s ease-in-out;
            }

            .login-container .stButton > button:hover {
                background: #50C878 !important;
                color: #FFFFFF !important;
                box-shadow: 0 5px 18px rgba(80, 200, 120, 0.35);
                transform: translateY(-1px);
            }

            .login-title {
                color: #FFFFFF !important;
                text-align: center;
                font-size: 1.9rem;
                font-weight: 700;
                margin-bottom: 0.4rem;
            }

            .login-subtitle {
                color: #B7B4C5 !important;
                text-align: center;
                margin-bottom: 1.8rem;
                font-size: 0.98rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# АВТОРИЗАЦИЯ
# ============================================================

def get_auth_settings() -> dict:
    """
    Возвращает параметры авторизации из Streamlit Secrets.

    Ожидаемые Secrets:

    AUTH_USERNAME = "operator"
    AUTH_PASSWORD = "ваш_пароль"
    AUTH_DISPLAY_NAME = "Оператор"
    """
    try:
        return {
            "username": str(st.secrets["AUTH_USERNAME"]),
            "password": str(st.secrets["AUTH_PASSWORD"]),
            "display_name": str(
                st.secrets.get("AUTH_DISPLAY_NAME", "Пользователь")
            ),
        }
    except Exception:
        return {}


def check_login(username: str, password: str) -> bool:
    """Безопасно сравнивает введённые логин и пароль со значениями Secrets."""
    settings = get_auth_settings()

    if not settings:
        return False

    username_ok = hmac.compare_digest(
        username.strip(),
        settings["username"],
    )

    password_ok = hmac.compare_digest(
        password,
        settings["password"],
    )

    return username_ok and password_ok


def logout() -> None:
    """Выход и очистка результатов текущей сессии."""
    st.session_state.authenticated = False
    st.session_state.comparison_result = None
    st.session_state.comparison_meta = None


def render_login_page() -> None:
    """Рендер страницы авторизации."""
    apply_login_styles()

    auth_settings = get_auth_settings()

    st.markdown("<br><br><br><br>", unsafe_allow_html=True)

    _, center_column, _ = st.columns([1, 1.05, 1])

    with center_column:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)

        st.markdown(
            '<div class="login-title">📦 Контроль недобора Ozon</div>',
            unsafe_allow_html=True,
        )

        st.markdown(
            (
                '<div class="login-subtitle">'
                "Загрузите файлы поставки и получите расчёт недобора."
                "</div>"
            ),
            unsafe_allow_html=True,
        )

        if not auth_settings:
            st.error(
                "Авторизация ещё не настроена. "
                "Добавьте `AUTH_USERNAME` и `AUTH_PASSWORD` "
                "в раздел Secrets приложения Streamlit."
            )
            st.stop()

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input(
                "Логин",
                placeholder="Логин",
                label_visibility="collapsed",
            )

            password = st.text_input(
                "Пароль",
                placeholder="Пароль",
                type="password",
                label_visibility="collapsed",
            )

            st.markdown("<br>", unsafe_allow_html=True)

            submitted = st.form_submit_button(
                "Войти",
                use_container_width=True,
            )

            if submitted:
                if check_login(username, password):
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Неверный логин или пароль.")

        st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# ЧТЕНИЕ ФАЙЛОВ
# ============================================================

@st.cache_data(show_spinner=False)
def read_uploaded_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Читает XLS, XLSX или CSV без автоматического заголовка.

    header=None важен, так как:
    - в файле A заголовки расположены на строке 13;
    - в файле B заголовки находятся на первой строке.
    """
    filename_lower = filename.lower()
    file_buffer = io.BytesIO(file_bytes)

    if filename_lower.endswith(".xlsx"):
        return pd.read_excel(
            file_buffer,
            header=None,
            dtype=object,
            engine="openpyxl",
        )

    if filename_lower.endswith(".xls"):
        return pd.read_excel(
            file_buffer,
            header=None,
            dtype=object,
            engine="xlrd",
        )

    if filename_lower.endswith(".csv"):
        encodings = ["utf-8-sig", "utf-8", "cp1251"]
        last_error = None

        for encoding in encodings:
            try:
                return pd.read_csv(
                    io.BytesIO(file_bytes),
                    header=None,
                    dtype=object,
                    keep_default_na=False,
                    encoding=encoding,
                    sep=None,
                    engine="python",
                )
            except Exception as error:
                last_error = error

        raise ValueError(
            f"Не удалось прочитать CSV-файл. Последняя ошибка: {last_error}"
        )

    raise ValueError(
        "Неподдерживаемый формат. Загрузите файл XLS, XLSX или CSV."
    )


# ============================================================
# ОБРАБОТКА ДАННЫХ
# ============================================================

def is_empty(value: object) -> bool:
    """Проверяет, является ли значение пустым."""
    if value is None:
        return True

    if isinstance(value, str):
        return value.strip() == ""

    try:
        return bool(pd.isna(value))
    except Exception:
        return False


def article_to_text(value: object) -> str:
    """Приводит артикул к строке без искажения целых чисел."""
    if is_empty(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()

    return str(value).strip()


def normalize_article(value: object) -> str:
    """
    Нормализует артикул для сопоставления.

    Оригинальное значение сохраняется в итоговой таблице,
    но сравнение выполняется без учёта регистра и лишних пробелов.
    """
    return article_to_text(value).replace("\u00A0", " ").strip().upper()


def parse_quantity(value: object) -> int:
    """
    Преобразует значение количества в целое неотрицательное число.

    Поддерживаются Excel-числа и значения формата:
    100, 100.0, 1 000, 1000.
    """
    if is_empty(value):
        raise ValueError("пустое количество")

    if isinstance(value, bool):
        raise ValueError("некорректное количество")

    if isinstance(value, int):
        quantity = value

    elif isinstance(value, float):
        if pd.isna(value):
            raise ValueError("пустое количество")

        if not value.is_integer():
            raise ValueError("количество должно быть целым числом")

        quantity = int(value)

    else:
        text = (
            str(value)
            .replace("\u00A0", "")
            .replace(" ", "")
            .strip()
            .replace(",", ".")
        )

        try:
            decimal_value = Decimal(text)
        except InvalidOperation:
            raise ValueError(f"некорректное число `{value}`")

        if decimal_value != decimal_value.to_integral_value():
            raise ValueError("количество должно быть целым числом")

        quantity = int(decimal_value)

    if quantity < 0:
        raise ValueError("количество не может быть отрицательным")

    return quantity


def collect_data_rows(
    raw_df: pd.DataFrame,
    article_column: int,
    quantity_column: int,
    start_row: int,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Извлекает строки с артикулом и количеством.

    Пустые и служебные строки без артикула игнорируются.
    """
    required_column = max(article_column, quantity_column)

    if raw_df.shape[1] <= required_column:
        return (
            pd.DataFrame(
                columns=["Артикул", "_article_key", "Количество"]
            ),
            [
                "В таблице недостаточно колонок "
                "для чтения артикула и количества."
            ],
        )

    rows = []
    errors = []

    for row_index in range(start_row, len(raw_df)):
        article_raw = raw_df.iloc[row_index, article_column]

        if is_empty(article_raw):
            continue

        article = article_to_text(article_raw)
        article_key = normalize_article(article)

        if not article_key:
            continue

        quantity_raw = raw_df.iloc[row_index, quantity_column]

        try:
            quantity = parse_quantity(quantity_raw)
        except ValueError as error:
            errors.append(
                f"Строка {row_index + 1}: "
                f"артикул `{article}` — {error}."
            )
            continue

        rows.append(
            {
                "Артикул": article,
                "_article_key": article_key,
                "Количество": quantity,
            }
        )

    return pd.DataFrame(rows), errors


def extract_source_a(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    Файл A — исходное количество из МойСклад.

    Используются:
    - D: артикул;
    - H: количество;
    - данные с 14 строки.
    """
    return collect_data_rows(
        raw_df=raw_df,
        article_column=SOURCE_A_ARTICLE_COLUMN,
        quantity_column=SOURCE_A_QUANTITY_COLUMN,
        start_row=SOURCE_A_START_ROW,
    )


def extract_source_b(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    Файл B — урезанный список из Ozon.

    Используются:
    - D: артикул;
    - F: количество в поставке;
    - данные со 2 строки.
    """
    return collect_data_rows(
        raw_df=raw_df,
        article_column=SOURCE_B_ARTICLE_COLUMN,
        quantity_column=SOURCE_B_QUANTITY_COLUMN,
        start_row=SOURCE_B_START_ROW,
    )


def aggregate_source_a(
    source_a: pd.DataFrame,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Суммирует дублирующиеся артикулы исходного файла A.

    Порядок сохраняется по первому появлению артикула в файле A.
    """
    if source_a.empty:
        return source_a, []

    duplicate_mask = source_a.duplicated(
        subset="_article_key",
        keep=False,
    )

    duplicate_articles = (
        source_a.loc[duplicate_mask, "Артикул"]
        .drop_duplicates()
        .tolist()
    )

    aggregated = (
        source_a.groupby(
            "_article_key",
            sort=False,
            as_index=False,
        )
        .agg(
            {
                "Артикул": "first",
                "Количество": "sum",
            }
        )
        .rename(
            columns={
                "Количество": "Исходное количество",
            }
        )
    )

    return aggregated, duplicate_articles


def aggregate_source_b(source_b: pd.DataFrame) -> pd.DataFrame:
    """Суммирует количество Ozon для повторяющихся артикулов."""
    if source_b.empty:
        return pd.DataFrame(
            columns=[
                "_article_key",
                "Количество в Ozon",
            ]
        )

    return (
        source_b.groupby(
            "_article_key",
            sort=False,
            as_index=False,
        )["Количество"]
        .sum()
        .rename(
            columns={
                "Количество": "Количество в Ozon",
            }
        )
    )


def compare_sources(
    source_a: pd.DataFrame,
    source_b: pd.DataFrame,
) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Сравнивает исходный список A и список Ozon B.

    Формула расчёта:

    Недобор = Исходное количество - Количество в Ozon

    Если результат меньше нуля, недобор считается равным 0.
    """
    source_a_aggregated, duplicate_articles_a = aggregate_source_a(
        source_a
    )

    source_b_aggregated = aggregate_source_b(source_b)

    result = source_a_aggregated.merge(
        source_b_aggregated,
        on="_article_key",
        how="left",
        sort=False,
    )

    result["Количество в Ozon"] = (
        result["Количество в Ozon"]
        .fillna(0)
        .astype(int)
    )

    result["Недобор"] = (
        result["Исходное количество"]
        - result["Количество в Ozon"]
    ).clip(lower=0)

    result["Недобор"] = result["Недобор"].astype(int)

    extra_articles_b = (
        source_b_aggregated[
            ~source_b_aggregated["_article_key"].isin(
                source_a_aggregated["_article_key"]
            )
        ]["_article_key"]
        .tolist()
    )

    result = result[
        [
            "Артикул",
            "Исходное количество",
            "Количество в Ozon",
            "Недобор",
        ]
    ]

    return result, duplicate_articles_a, extra_articles_b


# ============================================================
# СОЗДАНИЕ ФАЙЛОВ ДЛЯ СКАЧИВАНИЯ
# ============================================================

def create_csv_bytes(df: pd.DataFrame) -> bytes:
    """
    Готовит CSV с разделителем ; и UTF-8 BOM.

    Такой CSV корректно открывается в русской версии Excel.
    """
    return df.to_csv(
        index=False,
        sep=";",
    ).encode("utf-8-sig")


def create_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Готовит XLSX для скачивания."""
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Недобор",
        )

        worksheet = writer.sheets["Недобор"]

        # Закрепляет строку с заголовками.
        worksheet.freeze_panes = "A2"

        # Настраивает примерную ширину колонок.
        worksheet.column_dimensions["A"].width = 34
        worksheet.column_dimensions["B"].width = 24
        worksheet.column_dimensions["C"].width = 22
        worksheet.column_dimensions["D"].width = 15

        # Выделяет строки, где есть недобор.
        for row_index in range(2, worksheet.max_row + 1):
            shortage_value = worksheet.cell(
                row=row_index,
                column=4,
            ).value

            if shortage_value and shortage_value > 0:
                for column_index in range(1, 5):
                    worksheet.cell(
                        row=row_index,
                        column=column_index,
                    ).fill = __import__(
                        "openpyxl"
                    ).styles.PatternFill(
                        start_color="FFF2CC",
                        end_color="FFF2CC",
                        fill_type="solid",
                    )

    return output.getvalue()


# ============================================================
# ИНТЕРФЕЙС
# ============================================================

def render_sidebar() -> None:
    """Боковое меню авторизованной части."""
    settings = get_auth_settings()
    display_name = settings.get("display_name", "Пользователь")

    with st.sidebar:
        st.markdown("## 📦 Ozon")
        st.markdown(f"### {display_name}")
        st.caption("Контроль недобора поставок")

        st.divider()

        st.markdown("**Как пользоваться**")
        st.markdown(
            """
            1. Загрузите файл A из МойСклад.
            2. Загрузите файл B из Ozon.
            3. Нажмите «Сравнить файлы».
            4. Скачайте готовую таблицу.
            """
        )

        st.divider()

        if st.button("⍈  Выйти", use_container_width=True):
            logout()
            st.rerun()


def render_source_preview(
    raw_df: pd.DataFrame,
    title: str,
) -> None:
    """Диагностический предпросмотр загруженного файла."""
    with st.expander(title, expanded=False):
        st.caption(
            "Этот блок нужен только для проверки структуры "
            "загруженного файла."
        )

        st.dataframe(
            raw_df.head(20),
            use_container_width=True,
            hide_index=True,
        )


def render_main_page() -> None:
    """Главная страница приложения."""
    apply_main_styles()
    render_sidebar()

    st.title("📦 Контроль недобора поставки Ozon")

    st.markdown(
        """
        <div class="description-box">
            Загрузите два файла поставки. Приложение сопоставит позиции
            по артикулу, рассчитает недобор и подготовит итоговую таблицу
            в порядке исходного файла A.
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)

        st.subheader("A. Исходный файл")
        st.caption("Файл МойСклад или сборочный лист с исходным количеством.")
        st.caption("Артикул — D, количество — H, начало данных — строка 14.")

        source_file_a = st.file_uploader(
            "Загрузите файл A",
            type=["xls", "xlsx", "csv"],
            key="source_file_a",
            label_visibility="collapsed",
        )

        if source_file_a is not None:
            st.success(f"Загружен файл: `{source_file_a.name}`")

        st.markdown("</div>", unsafe_allow_html=True)

    with col_b:
        st.markdown('<div class="upload-card">', unsafe_allow_html=True)

        st.subheader("B. Файл Ozon")
        st.caption("Выгрузка Ozon с подтверждённым количеством в поставке.")
        st.caption("Артикул — D, количество — F, начало данных — строка 2.")

        source_file_b = st.file_uploader(
            "Загрузите файл B",
            type=["xls", "xlsx", "csv"],
            key="source_file_b",
            label_visibility="collapsed",
        )

        if source_file_b is not None:
            st.success(f"Загружен файл: `{source_file_b.name}`")

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    can_compare = (
        source_file_a is not None
        and source_file_b is not None
    )

    if st.button(
        "🔍 Сравнить файлы",
        type="primary",
        disabled=not can_compare,
    ):
        try:
            with st.spinner("Читаем файлы и рассчитываем недобор..."):
                raw_a = read_uploaded_file(
                    source_file_a.getvalue(),
                    source_file_a.name,
                )

                raw_b = read_uploaded_file(
                    source_file_b.getvalue(),
                    source_file_b.name,
                )

                data_a, errors_a = extract_source_a(raw_a)
                data_b, errors_b = extract_source_b(raw_b)

                validation_errors = errors_a + errors_b

                if data_a.empty:
                    st.error(
                        "В файле A не найдены товарные позиции. "
                        "Проверьте, что это исходный файл МойСклад "
                        "с артикулом в колонке D и количеством в H."
                    )
                    return

                if data_b.empty:
                    st.error(
                        "В файле B не найдены товарные позиции. "
                        "Проверьте, что это файл Ozon "
                        "с артикулом в колонке D и количеством в F."
                    )
                    return

                if validation_errors:
                    st.error(
                        "В файлах найдены строки с некорректным количеством. "
                        "Исправьте исходный файл и повторите загрузку."
                    )

                    for error in validation_errors[:20]:
                        st.write(f"• {error}")

                    if len(validation_errors) > 20:
                        st.write(
                            f"• Дополнительно найдено ошибок: "
                            f"{len(validation_errors) - 20}"
                        )

                    return

                result, duplicates_a, extra_articles_b = compare_sources(
                    data_a,
                    data_b,
                )

                st.session_state.comparison_result = result

                st.session_state.comparison_meta = {
                    "raw_a": raw_a,
                    "raw_b": raw_b,
                    "file_a_name": source_file_a.name,
                    "file_b_name": source_file_b.name,
                    "duplicates_a": duplicates_a,
                    "extra_articles_b": extra_articles_b,
                }

            st.success("Сравнение успешно выполнено.")

        except Exception as error:
            st.error(f"Ошибка при обработке файлов: {error}")
            return

    result = st.session_state.comparison_result
    meta = st.session_state.comparison_meta

    if result is None or meta is None:
        st.info(
            "Загрузите оба файла, затем нажмите кнопку «Сравнить файлы»."
        )
        return

    st.divider()
    st.subheader("📊 Итог сравнения")

    total_source_quantity = int(
        result["Исходное количество"].sum()
    )

    total_ozon_quantity = int(
        result["Количество в Ozon"].sum()
    )

    total_shortage = int(
        result["Недобор"].sum()
    )

    shortage_positions = int(
        (result["Недобор"] > 0).sum()
    )

    metric_1, metric_2, metric_3, metric_4 = st.columns(4)

    metric_1.metric(
        "Позиций в файле A",
        len(result),
    )

    metric_2.metric(
        "Исходное количество",
        f"{total_source_quantity:,}".replace(",", " "),
    )

    metric_3.metric(
        "Подтверждено Ozon",
        f"{total_ozon_quantity:,}".replace(",", " "),
    )

    metric_4.metric(
        "Общий недобор",
        f"{total_shortage:,}".replace(",", " "),
    )

    if shortage_positions > 0:
        st.warning(
            f"Недобор найден в позициях: {shortage_positions}."
        )
    else:
        st.success(
            "Недобора не найдено: количества по всем позициям совпадают."
        )

    duplicate_articles_a = meta["duplicates_a"]

    if duplicate_articles_a:
        examples = ", ".join(duplicate_articles_a[:8])

        st.warning(
            "В файле A обнаружены повторяющиеся артикулы. "
            "Их количества были просуммированы. "
            f"Примеры: {examples}"
        )

    extra_articles_b = meta["extra_articles_b"]

    if extra_articles_b:
        examples = ", ".join(extra_articles_b[:8])

        st.info(
            "В файле B есть артикулы, отсутствующие в файле A. "
            "Они не включены в итоговую таблицу. "
            f"Примеры: {examples}"
        )

    ozon_more_than_source = result[
        result["Количество в Ozon"] > result["Исходное количество"]
    ]

    if not ozon_more_than_source.empty:
        st.warning(
            "В некоторых позициях Ozon содержит количество больше исходного. "
            "Для этих строк недобор установлен в 0. "
            "Рекомендуется проверить исходные данные."
        )

    st.dataframe(
        result,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Артикул": st.column_config.TextColumn(
                "Артикул",
                width="large",
            ),
            "Исходное количество": st.column_config.NumberColumn(
                "Исходное количество",
                format="%d",
            ),
            "Количество в Ozon": st.column_config.NumberColumn(
                "Количество в Ozon",
                format="%d",
            ),
            "Недобор": st.column_config.NumberColumn(
                "Недобор",
                format="%d",
            ),
        },
    )

    st.markdown("<br>", unsafe_allow_html=True)

    download_col_xlsx, download_col_csv, _ = st.columns([1, 1, 2])

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    with download_col_xlsx:
        st.download_button(
            label="📥 Скачать XLSX",
            data=create_xlsx_bytes(result),
            file_name=f"ozon_nedobor_{timestamp}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    with download_col_csv:
        st.download_button(
            label="📥 Скачать CSV",
            data=create_csv_bytes(result),
            file_name=f"ozon_nedobor_{timestamp}.csv",
            mime="text/csv",
        )

    render_source_preview(
        meta["raw_a"],
        f"Предпросмотр файла A: {meta['file_a_name']}",
    )

    render_source_preview(
        meta["raw_b"],
        f"Предпросмотр файла B: {meta['file_b_name']}",
    )


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

def main() -> None:
    """Роутинг между авторизацией и главной страницей."""
    if not st.session_state.authenticated:
        render_login_page()
        return

    render_main_page()


if __name__ == "__main__":
    main()