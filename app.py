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

def _extract_period(file_bytes: bytes) -> tuple:
    """Период из шапки отчёта (строки вида 'С: 01.01.2026' / 'ПО: 01.09.2026')."""
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    d_from = d_to = None
    for i in range(min(12, len(raw))):
        for v in raw.iloc[i].tolist():
            s = str(v).strip()
            if s.upper().startswith("С:"):
                d_from = s.split(":", 1)[1].strip()
            elif s.upper().startswith(("ПО:", "ПО ")) or s.startswith("По:"):
                d_to = s.split(":", 1)[1].strip() if ":" in s else s[2:].strip()
    return d_from, d_to


SERVICE_CATEGORIES = ["Абонементы", "Приёмы", "Анализы", "Прочие разовые услуги"]


def _classify_service(name: str, spec: str) -> str:
    name_u, spec_u = str(name).upper(), str(spec).upper()
    if spec_u == "АБОНЕМЕНТЫ" or "АБОНЕМЕНТ" in name_u:
        return "Абонементы"
    if spec_u == "ВНЕШНЯЯ ЛАБОРАТОРИЯ":
        return "Анализы"
    if name_u.strip().startswith("ПРИЕМ"):
        return "Приёмы"
    return "Прочие разовые услуги"


@st.cache_data(show_spinner="Читаю отчёт по услугам…")
def parse_services(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Читает отчёт «Количество выполненных услуг на сумму по убыванию»."""
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    header_idx = None
    for i in range(min(20, len(raw))):
        if str(raw.iloc[i, 0]).strip() == "USLCODE":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Не найдена строка заголовка ('USLCODE') — это не отчёт по услугам.")

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, skiprows=header_idx)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for c in ["Кол-во", "Актуальная цена", "кол-во * цена"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df = df[df["Услуга"].notna()].copy()
    df["Категория"] = [_classify_service(n, s) for n, s in zip(df["Услуга"], df["Специалитет"])]
    return df


def aggregate_service_categories(srv: pd.DataFrame) -> pd.DataFrame:
    agg = srv.groupby("Категория").agg(Количество=("Кол-во", "sum"), Сумма=("кол-во * цена", "sum"))
    agg = agg.reindex(SERVICE_CATEGORIES).fillna(0).reset_index()
    total_cnt = agg["Количество"].sum()
    total_sum = agg["Сумма"].sum()
    agg["Доля_кол_%"] = (agg["Количество"] / total_cnt * 100).round(2) if total_cnt else 0
    agg["Доля_суммы_%"] = (agg["Сумма"] / total_sum * 100).round(2) if total_sum else 0
    return agg



# ------------------------------------------------------------------
# Загрузка файла
# ------------------------------------------------------------------
st.sidebar.header("Загрузка отчёта")
uploaded = st.sidebar.file_uploader(
    "Отчёт «Список абонементов по дате» (.xlsx)",
    type=["xlsx", "xls"],
    help="Персональные данные удаляются автоматически при чтении файла.",
)

uploaded_services = st.sidebar.file_uploader(
    "Отчёт по всем услугам (.xlsx) — опционально",
    type=["xlsx", "xls"],
    help="Отчёт «Количество выполненных услуг на сумму по убыванию». "
         "Нужен для сравнения доли абонементов и разовых услуг.",
)

if uploaded is None:
    st.title("📊 Дашборд по услугам-абонементам")
    st.info("👈 Загрузите в боковой панели отчёт «Список абонементов по дате» (Excel). "
            "Даты в отчёте могут быть любыми — дашборд построится автоматически. "
            "Для сравнения с разовыми услугами загрузите также отчёт по всем услугам.")
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



# ------------------------------------------------------------------
# Сравнение: абонементы vs разовые услуги
# ------------------------------------------------------------------
if uploaded_services is not None:
    try:
        srv_df = parse_services(uploaded_services.getvalue(), uploaded_services.name)
        srv_period = _extract_period(uploaded_services.getvalue())
        srv_agg = aggregate_service_categories(srv_df)

        st.divider()
        st.subheader("🏥 Абонементы vs разовые услуги")
        st.caption(
            f"Отчёт по услугам: **{uploaded_services.name}** · период: "
            f"**{srv_period[0] or 'н/д'} — {srv_period[1] or 'н/д'}** · "
            f"всего оказано **{int(srv_df['Кол-во'].sum()):,} услуг на {srv_df['кол-во * цена'].sum():,.0f} ₽**".replace(",", " ")
        )

        st.warning(
            "⚠️ Периоды двух отчётов могут не совпадать: абонементный отчёт — по датам **создания** "
            "абонементов, отчёт по услугам — по датам **оказания**. Сравнение носит оценочный характер.",
            icon=None,
        )

        # --- доля по количеству ---
        fig_cnt = px.pie(
            srv_agg[srv_agg["Количество"] > 0],
            values="Количество", names="Категория", hole=0.45,
            color="Категория",
            color_discrete_map={
                "Абонементы": "#7C4DBE", "Приёмы": "#2E86AB",
                "Анализы": "#F6A21E", "Прочие разовые услуги": "#9AA5B1",
            },
            labels={"Количество": "Кол-во, шт."},
        )
        fig_cnt.update_traces(textinfo="percent+label", texttemplate="%{label}<br>%{percent:.1%}")
        fig_cnt.update_layout(height=430, legend=dict(orientation="h", yanchor="bottom", y=-0.15))

        # --- доля по выручке ---
        abon_paid = float(flt["Оплачено"].sum())
        oneoff_sum = float(srv_agg.loc[srv_agg["Категория"] != "Абонементы", "Сумма"].sum())
        rev_df = pd.DataFrame({
            "Источник": ["Абонементы (оплачено)", "Разовые услуги (оказано)"],
            "Сумма": [abon_paid, oneoff_sum],
        })
        fig_rev = px.pie(
            rev_df, values="Сумма", names="Источник", hole=0.45,
            color="Источник",
            color_discrete_map={"Абонементы (оплачено)": "#7C4DBE", "Разовые услуги (оказано)": "#2E86AB"},
        )
        fig_rev.update_traces(textinfo="percent+label", texttemplate="%{label}<br>%{percent:.1%}")
        fig_rev.update_layout(height=430, legend=dict(orientation="h", yanchor="bottom", y=-0.15))

        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown("**Доля по количеству оказанных услуг**")
            st.plotly_chart(fig_cnt, use_container_width=True)
        with rc2:
            st.markdown("**Доля в выручке**")
            st.plotly_chart(fig_rev, use_container_width=True)

        share_abon = abon_paid / (abon_paid + oneoff_sum) * 100 if (abon_paid + oneoff_sum) else 0
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Абонементы в выручке", f"{share_abon:.1f}%")
        m2.metric("Абонементы, оплачено", fmt_rub.format(abon_paid).replace(",", " "))
        m3.metric("Разовые услуги, оказано", fmt_rub.format(oneoff_sum).replace(",", " "))
        m4.metric("Доля приёмов (по кол-ву)",
                  f"{srv_agg.loc[srv_agg['Категория']=='Приёмы','Доля_кол_%'].iloc[0]:.1f}%")

        with st.expander("Детализация по категориям услуг"):
            show_srv = srv_agg.rename(columns={
                "Категория": "Категория", "Количество": "Кол-во, шт.",
                "Сумма": "Сумма, ₽", "Доля_кол_%": "Доля по кол-ву, %", "Доля_суммы_%": "Доля по сумме, %",
            })
            st.dataframe(
                show_srv.style.format({
                    "Кол-во, шт.": "{:,.0f}", "Сумма, ₽": "{:,.0f}",
                    "Доля по кол-ву, %": "{:.2f}", "Доля по сумме, %": "{:.2f}",
                }),
                use_container_width=True, hide_index=True,
            )
            st.caption(
                "Категории: **Абонементы** — услуги со специалитетом «Абонементы» или с «абонемент» в названии; "
                "**Приёмы** — наименования, начинающиеся со слова «Прием»; **Анализы** — внешняя лаборатория; "
                "**Прочие разовые услуги** — всё остальное (манипуляции, УЗИ, рентген, медикаменты, операции)."
            )
    except Exception as e:
        st.error(f"Не удалось прочитать отчёт по услугам: {e}")

st.divider()

# ------------------------------------------------------------------
# Главный график: популярность абонементов
# ------------------------------------------------------------------
st.subheader("Популярность абонементов: количество проданных и выручка")

pop_cnt = flt.sort_values("Количество", ascending=True)
pop_rev = flt[flt["Оплачено"] > 0].sort_values("Оплачено", ascending=True)

p1, p2 = st.columns(2)
with p1:
    fig_pop1 = px.bar(
        pop_cnt, x="Количество", y="Абонемент", orientation="h",
        color="Количество", color_continuous_scale="Blues",
        labels={"Количество": "Продано абонементов, шт.", "Абонемент": ""},
        text="Количество",
    )
    fig_pop1.update_traces(textposition="outside", cliponaxis=False)
    fig_pop1.update_layout(height=max(400, len(pop_cnt) * 40), coloraxis_showscale=False,
                           yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_pop1, use_container_width=True)

with p2:
    if not pop_rev.empty:
        fig_pop2 = px.bar(
            pop_rev, x="Оплачено", y="Абонемент", orientation="h",
            color="Оплачено", color_continuous_scale="Greens",
            labels={"Оплачено": "Выручка (оплачено), ₽", "Абонемент": ""},
            text="Оплачено",
        )
        fig_pop2.update_traces(texttemplate="%{text:,.0f} ₽", textposition="outside", cliponaxis=False)
        fig_pop2.update_layout(height=max(400, len(pop_rev) * 40), coloraxis_showscale=False,
                               yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_pop2, use_container_width=True)
    else:
        st.info("Нет данных об оплатах.")

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
    st.subheader("Оплачено vs Оказано (все услуги)")
    top = flt.sort_values("Оказано_на_сумму", ascending=False)
    long = top.melt(id_vars="Абонемент", value_vars=["Оплачено", "Оказано_на_сумму"],
                    var_name="Показатель", value_name="Сумма")
    long["Показатель"] = long["Показатель"].map({"Оплачено": "Оплачено", "Оказано_на_сумму": "Оказано"})
    fig1 = px.bar(long, x="Сумма", y="Абонемент", color="Показатель", barmode="group", orientation="h",
                  color_discrete_map={"Оплачено": "#2E86AB", "Оказано": "#F24236"},
                  labels={"Сумма": "Сумма, ₽", "Абонемент": ""})
    fig1.update_layout(height=max(350, len(top) * 40), yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.subheader("% потребления по услугам")
    top_c = flt[flt["Стоимость_абонемента_сумма"] > 0]
    if not top_c.empty:
        fig2 = px.bar(top_c.sort_values("Процент_потребления"), x="Процент_потребления", y="Абонемент",
                      orientation="h", color="Процент_потребления", color_continuous_scale="RdYlGn_r",
                      range_x=[0, max(130, top_c["Процент_потребления"].max() * 1.15)],
                      labels={"Процент_потребления": "% потребления", "Абонемент": ""},
                      text="Процент_потребления")
        fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
        fig2.update_layout(height=max(350, len(top_c) * 40), coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Нет данных для расчёта процента потребления.")

st.subheader("Структура оплаченных средств по услугам")
pie_src = flt[flt["Оплачено"] > 0]
if not pie_src.empty:
    fig3 = px.pie(pie_src, values="Оплачено", names="Абонемент", hole=0.45,
                  labels={"Абонемент": "Услуга", "Оплачено": "Оплачено, ₽"})
    fig3.update_traces(textinfo="percent+label")
    fig3.update_layout(height=520)
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
