# -*- coding: utf-8 -*-
"""
Дашборд по услугам-абонементам.

Пользователь загружает отчёт «Список абонементов по дате» (Excel) —
даты в отчёте могут быть любыми. Персональные данные (ФИО, телефоны,
номера карт) удаляются сразу при чтении файла, в дашборд попадают
только агрегированные сводки по услугам и месяцам.

Запуск: streamlit run app.py
"""

import io

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Дашборд абонементов",
    page_icon="📊",
    layout="wide",
)

# ------------------------------------------------------------------
# Чтение и очистка отчёта
# ------------------------------------------------------------------
PDN_KEYS = ["ФИО", "телефон", "карты"]  # колонки с персональными данными

REQUIRED_COLS = ["Дата создания", "Абонемент", "Стоимость абонемента",
                 "Оплачено", "Услуг оказано на сумму", "Стоимость - оказано"]


def _to_datetime_safe(s: pd.Series) -> pd.Series:
    """Даты могут прийти как datetime или как Excel-серийные числа."""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(s, errors="coerce", dayfirst=False)


@st.cache_data(show_spinner="Читаю отчёт…")
def parse_report(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Читает отчёт, находит строку заголовка, удаляет перс. данные."""
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)

    header_idx = None
    for i in range(min(20, len(raw))):
        if str(raw.iloc[i, 0]).strip().startswith("Номер карты"):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(
            "Не найдена строка заголовка («Номер карты»). "
            "Похоже, это не отчёт «Список абонементов по дате»."
        )

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, skiprows=header_idx)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"В отчёте не хватает колонок: {', '.join(missing)}")

    # --- удаление персональных данных ---
    pdn = [c for c in df.columns if any(k.lower() in str(c).lower() for k in PDN_KEYS)]
    df = df.drop(columns=pdn)

    df = df[df["Абонемент"].notna()].copy()

    df["Дата создания"] = _to_datetime_safe(df["Дата создания"])
    df["Дата начала"] = _to_datetime_safe(df["Дата начала"]) if "Дата начала" in df.columns else pd.NaT
    df["Дата окончания"] = _to_datetime_safe(df["Дата окончания"]) if "Дата окончания" in df.columns else pd.NaT

    for c in ["Стоимость абонемента", "Оплачено", "Услуг оказано на сумму", "Стоимость - оказано"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Месяц создания"] = df["Дата создания"].dt.to_period("M").astype(str)
    return df


def aggregate_services(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby("Абонемент").agg(
        Количество=("Абонемент", "size"),
        Стоимость_абонемента_сумма=("Стоимость абонемента", "sum"),
        Оплачено=("Оплачено", "sum"),
        Оказано_на_сумму=("Услуг оказано на сумму", "sum"),
        Остаток_стоимость_минус_оказано=("Стоимость - оказано", "sum"),
    ).reset_index()
    agg["Средняя_стоимость"] = agg["Стоимость_абонемента_сумма"] / agg["Количество"]
    denom = agg["Стоимость_абонемента_сумма"].where(agg["Стоимость_абонемента_сумма"] != 0)
    agg["Процент_оплаты"] = (agg["Оплачено"] / denom * 100).round(1)
    agg["Процент_потребления"] = (agg["Оказано_на_сумму"] / denom * 100).round(1)
    agg["Долг_по_оплате"] = agg["Стоимость_абонемента_сумма"] - agg["Оплачено"]
    return agg.fillna(0)


def aggregate_monthly(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("Месяц создания").agg(
        Количество=("Абонемент", "size"),
        Стоимость_сумма=("Стоимость абонемента", "sum"),
        Оплачено=("Оплачено", "sum"),
        Оказано_на_сумму=("Услуг оказано на сумму", "sum"),
    ).reset_index().sort_values("Месяц создания")


# ------------------------------------------------------------------
# Загрузка файла
# ------------------------------------------------------------------
st.sidebar.header("Загрузка отчёта")
uploaded = st.sidebar.file_uploader(
    "Отчёт «Список абонементов по дате» (.xlsx)",
    type=["xlsx", "xls"],
    help="Персональные данные удаляются автоматически при чтении файла.",
)

if uploaded is None:
    st.title("📊 Дашборд по услугам-абонементам")
    st.info("👈 Загрузите в боковой панели отчёт «Список абонементов по дате» (Excel). Даты в отчёте могут быть любыми — дашборд построится автоматически.")
    st.stop()

try:
    df = parse_report(uploaded.getvalue(), uploaded.name)
except Exception as e:
    st.error(f"Не удалось прочитать файл: {e}")
    st.stop()

data = aggregate_services(df)
monthly = aggregate_monthly(df)

date_from = df["Дата создания"].min()
date_to = df["Дата создания"].max()
period = f"{date_from:%d.%m.%Y} — {date_to:%d.%m.%Y}" if pd.notna(date_from) and pd.notna(date_to) else "период не определён"

COLS = {
    "Количество": "Количество",
    "Стоимость_абонемента_сумма": "Стоимость абонементов",
    "Оплачено": "Оплачено",
    "Оказано_на_сумму": "Оказано на сумму",
    "Долг_по_оплате": "Долг по оплате",
    "Остаток_стоимость_минус_оказано": "Стоимость − оказано",
    "Процент_оплаты": "% оплаты",
    "Процент_потребления": "% потребления",
}
fmt_rub = "{:,.0f} ₽"

# ------------------------------------------------------------------
# Боковая панель — фильтры
# ------------------------------------------------------------------
st.sidebar.header("Фильтры")
search = st.sidebar.text_input("Поиск по названию услуги", "")
only_debt = st.sidebar.checkbox("Только с долгом по оплате", value=False)

sort_by = st.sidebar.selectbox("Сортировка таблицы", options=list(COLS.keys()), format_func=lambda x: COLS[x], index=3)
sort_asc = st.sidebar.radio("Порядок", ["По убыванию", "По возрастанию"], horizontal=True) == "По возрастанию"

flt = data.copy()
if search.strip():
    flt = flt[flt["Абонемент"].str.contains(search.strip(), case=False, na=False)]
if only_debt:
    flt = flt[flt["Долг_по_оплате"] > 0]
flt = flt.sort_values(sort_by, ascending=sort_asc)

top_n = st.sidebar.slider("Топ-N услуг на графиках", min_value=5, max_value=max(5, len(flt)), value=min(10, len(flt)))

# ------------------------------------------------------------------
# Заголовок и KPI
# ------------------------------------------------------------------
st.title("📊 Дашборд по услугам-абонементам")
st.caption(f"Отчёт: **{uploaded.name}** · период создания абонементов: **{period}** · "
           "персональные данные удалены при загрузке")

total_cost = flt["Стоимость_абонемента_сумма"].sum()
total_paid = flt["Оплачено"].sum()
total_done = flt["Оказано_на_сумму"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Услуг", f"{len(flt)}")
c2.metric("Продано абонементов", f"{int(flt['Количество'].sum())}")
c3.metric("Стоимость проданных", fmt_rub.format(total_cost).replace(",", " "))
c4.metric("Оплачено", fmt_rub.format(total_paid).replace(",", " "))
c5.metric("Оказано на сумму", fmt_rub.format(total_done).replace(",", " "))

c6, c7 = st.columns(2)
c6.metric("Долг по оплате", fmt_rub.format(total_cost - total_paid).replace(",", " "))
c7.metric("Средний % потребления", f"{(total_done / total_cost * 100 if total_cost else 0):.1f}%",
          help="Оказано на сумму / Стоимость проданных абонементов × 100")

st.divider()

# ------------------------------------------------------------------
# Динамика по месяцам
# ------------------------------------------------------------------
st.subheader("Динамика по месяцам создания абонемента")
mon = monthly.copy()
mon_long = mon.melt(id_vars="Месяц создания", value_vars=["Оплачено", "Оказано_на_сумму"],
                    var_name="Показатель", value_name="Сумма")
mon_long["Показатель"] = mon_long["Показатель"].map({"Оплачено": "Оплачено", "Оказано_на_сумму": "Оказано"})

fig0 = px.bar(mon_long, x="Месяц создания", y="Сумма", color="Показатель", barmode="group",
              color_discrete_map={"Оплачено": "#2E86AB", "Оказано": "#F24236"},
              labels={"Сумма": "Сумма, ₽", "Месяц создания": "Месяц"})
fig0.add_scatter(x=mon["Месяц создания"], y=mon["Количество"], mode="lines+markers",
                 name="Кол-во абонементов", yaxis="y2", line=dict(color="#7C4DBE", width=2))
fig0.update_layout(height=420,
                   yaxis=dict(title="Сумма, ₽"),
                   yaxis2=dict(title="Кол-во", overlaying="y", side="right", showgrid=False),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right"))
st.plotly_chart(fig0, use_container_width=True)

st.divider()

# ------------------------------------------------------------------
# Графики по услугам
# ------------------------------------------------------------------
col_left, col_right = st.columns(2)

with col_left:
    st.subheader(f"Оплачено vs Оказано (топ-{top_n})")
    top = flt.nlargest(top_n, "Оказано_на_сумму")
    long = top.melt(id_vars="Абонемент", value_vars=["Оплачено", "Оказано_на_сумму"],
                    var_name="Показатель", value_name="Сумма")
    long["Показатель"] = long["Показатель"].map({"Оплачено": "Оплачено", "Оказано_на_сумму": "Оказано"})
    fig1 = px.bar(long, x="Сумма", y="Абонемент", color="Показатель", barmode="group", orientation="h",
                  color_discrete_map={"Оплачено": "#2E86AB", "Оказано": "#F24236"},
                  labels={"Сумма": "Сумма, ₽", "Абонемент": ""})
    fig1.update_layout(height=max(350, top_n * 40), yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.subheader("% потребления по услугам (топ по оказанным)")
    top_c = flt[flt["Стоимость_абонемента_сумма"] > 0].nlargest(top_n, "Оказано_на_сумму")
    if not top_c.empty:
        fig2 = px.bar(top_c.sort_values("Процент_потребления"), x="Процент_потребления", y="Абонемент",
                      orientation="h", color="Процент_потребления", color_continuous_scale="RdYlGn_r",
                      range_x=[0, max(130, top_c["Процент_потребления"].max() * 1.15)],
                      labels={"Процент_потребления": "% потребления", "Абонемент": ""},
                      text="Процент_потребления")
        fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
        fig2.update_layout(height=max(350, top_n * 40), coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Нет данных для расчёта процента потребления.")

st.subheader("Структура оплаченных средств по услугам")
pie_src = flt[flt["Оплачено"] > 0].nlargest(top_n, "Оплачено")
if not pie_src.empty:
    fig3 = px.pie(pie_src, values="Оплачено", names="Абонемент", hole=0.45,
                  labels={"Абонемент": "Услуга", "Оплачено": "Оплачено, ₽"})
    fig3.update_traces(textinfo="percent+label")
    fig3.update_layout(height=520)
    rest = total_paid - pie_src["Оплачено"].sum()
    if rest > 0:
        st.caption(f"Остальные услуги (вне топ-{top_n}): {fmt_rub.format(rest).replace(',', ' ')}")
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("Нет данных об оплаченных средствах.")

st.divider()

# ------------------------------------------------------------------
# Таблица
# ------------------------------------------------------------------
st.subheader("Сводная таблица по услугам")
show = flt.rename(columns={"Абонемент": "Услуга / абонемент", **COLS})
st.dataframe(
    show.style.format({
        "Количество": "{:.0f}", "Стоимость абонементов": "{:,.0f}", "Оплачено": "{:,.0f}",
        "Оказано на сумму": "{:,.0f}", "Долг по оплате": "{:,.0f}", "Стоимость − оказано": "{:,.0f}",
        "% оплаты": "{:.1f}%", "% потребления": "{:.1f}%",
    }),
    use_container_width=True,
    height=min(650, 60 + len(show) * 35),
)

st.download_button(
    "⬇️ Скачать агрегированные данные (CSV)",
    data=show.to_csv(index=False).encode("utf-8-sig"),
    file_name="abonementy_aggregated.csv",
    mime="text/csv",
)

st.divider()
st.caption(
    "«% потребления» = Оказано на сумму / Стоимость абонементов × 100; значения >100% означают, "
    "что услуги оказаны сверх стоимости проданных в периоде абонементов (списание остатков прошлых периодов). "
    "«Долг по оплате» = Стоимость − Оплачено."
)
