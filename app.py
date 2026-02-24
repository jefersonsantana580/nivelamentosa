# -*- coding: utf-8 -*-
# App: Programação de Produção 2º Sem/2026 (balanceamento diário por MODELO)
# Autor: M365 Copilot p/ Jeferson Santana
# Como rodar local:
#   pip install -r requirements.txt
#   streamlit run app.py

import io
from datetime import date, datetime, timedelta
from calendar import monthrange
from dateutil.relativedelta import relativedelta

import pandas as pd
import streamlit as st

# =========================
# Config e UI
# =========================
st.set_page_config(page_title="Programação 2S2026", page_icon="🏭", layout="wide")
st.title("🏭 Programação de Produção – 2026")

with st.expander("📎 Instruções (resumo)", expanded=False):
    st.markdown("""
**Entrada**: Excel com aba **Planilha1**. Colunas:
- **MODELO** (coluna B)
- **PRODUTO** (coluna D)
- **MERCADO** (coluna F)
- **G–R**: meses **jul/26, ago/26, set/26, out/26, nov/26, dez/26** (ou variação com ano completo).
Pode haver linhas repetidas (PRODUTO/MERCADO/MODELO): serão **somadas**.

**Regras** (principais):
- Seg–Sex, sem feriados (feriados opcionais via lista), **capacidade por dia útil = `capacidade_dia_util`** (padrão 50).
- **Dias úteis anteriores**: é possível iniciar a produção do mês **N** até `dias_uteis_anteriores` **dias úteis ANTES** do 1º dia útil do mês **N** (ex.: julho inicia em 25/jun se o offset for 4 dias úteis).
- Prioridade diária: mercados **≠ "MERCADO INTERNO"** antes de MI.
- Cumpre **exatamente** o volume mensal por **PRODUTO**.
- Excedente do mês vai para **sábados do mesmo mês** (máx. por sábado = `teto_sabado`), ignorando **sábados que forem feriado**.
- Balanceamento diário por **MODELO** via **cotas proporcionais e maiores restos** + **round-robin**.
- IDs globais únicos: `fila 1 ... fila N` (ordenados por data, MODELO, PRODUTO).

**Saídas**:
- Aba **Programacao_2026**
- Aba **Relatorio_2026** (programado por **mês de referência** – o mês da demanda, mesmo que a produção inicie dias úteis antes em outro mês)
- Aba **Base_Original**
    """)

st.sidebar.header("⚙️ Parâmetros")
uploaded = st.sidebar.file_uploader("Carregar arquivo Excel (base)", type=["xlsx"])

# Parâmetros variáveis ([PARAM])
limite_diario_por_modelo = st.sidebar.number_input(
    "limite_diario_por_modelo (opcional)", min_value=0, step=1, value=0,
    help="0 = desativado. Se >0, nenhum MODELO ultrapassa esse número por dia."
)
capacidade_dia_util = st.sidebar.number_input(
    "capacidade_dia_util (todos os MODELOS juntos)", min_value=1, step=1, value=50,
    help="Capacidade total por dia útil (padrão 50)."
)
dias_uteis_anteriores = st.sidebar.number_input(
    "dias_uteis_anterior ", min_value=0, step=1, value=0,
    help="Quantidade de dias úteis ANTERIORES ao 1º dia útil do mês de referência. Ex.: Canoas=4, Mogi=6."
)
teto_sabado = st.sidebar.number_input(
    "teto_sabado", min_value=1, max_value=1000, value=50,
    help="Capacidade máxima por sábado usado para excedente mensal."
)
nomenclatura_arquivo = st.sidebar.text_input(
    "nomenclatura_arquivo (opcional)",
    value="programacao_2S2026_balanceada.xlsx",
    help="Se quiser, troque o nome do arquivo de saída."
)
# Intervalo padrão: Jul–Dez/2026
default_start = date(2026, 7, 1)
default_end = date(2026, 12, 31)
intervalo = st.sidebar.date_input(
    "intervalo_meses (início e fim)",
    value=(default_start, default_end),
    help="Use para trocar o semestre, mantendo regras. Considera meses completos entre as duas datas."
)

# =========================
# Feriados (opcional)
# =========================
st.sidebar.subheader("📅 Feriados (opcional)")
feriados_text = st.sidebar.text_area(
    "Lista de feriados (um por linha)",
    value="",
    placeholder="Exemplos válidos:\n2026-09-07\n07/09/2026\n2026-12-25",
    help="Datas no formato YYYY-MM-DD ou DD/MM/YYYY; um por linha."
)
feriados_csv = st.sidebar.file_uploader(
    "ou envie CSV com coluna 'data'",
    type=["csv"],
    help="CSV com cabeçalho 'data' contendo as datas de feriado."
)

def parse_feriados(text: str, csv_file) -> set:
    datas = set()
    # Text area
    for raw in (text or '').replace(";", "\n").splitlines():
        s = raw.strip()
        if not s:
            continue
        dt = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(s, fmt).date()
                break
            except Exception:
                continue
        if dt:
            datas.add(dt)
    # CSV
    if csv_file is not None:
        try:
            d = pd.read_csv(csv_file)
            if 'data' in d.columns:
                for x in d['data'].tolist():
                    try:
                        dt = pd.to_datetime(x, dayfirst=True).date()
                        datas.add(dt)
                    except Exception:
                        pass
        except Exception:
            pass
    return datas

feriados_set = parse_feriados(feriados_text, feriados_csv)
st.sidebar.caption(f"Feriados carregados: {len(feriados_set)}")

# Botão principal
gerar = st.sidebar.button("🚀 Gerar Programação", type="primary", use_container_width=True)

# =========================
# Utilitários de calendário/meses
# =========================
PT_BR_MONTHS = {
    1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun",
    7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"
}

def first_business_day(year: int, month: int, feriados:set) -> date:
    d = date(year, month, 1)
    while d.weekday() >= 5 or d in feriados:
        d += timedelta(days=1)
    return d

def business_days_in_month(year: int, month: int, feriados: set):
    last_day = monthrange(year, month)[1]
    days = []
    for day in range(1, last_day + 1):
        d = date(year, month, day)
        if d.weekday() < 5 and d not in feriados:  # Mon-Fri e não feriado
            days.append(d)
    return days

def saturdays_in_month(year: int, month: int, feriados: set):
    last_day = monthrange(year, month)[1]
    days = []
    for day in range(1, last_day + 1):
        d = date(year, month, day)
        if d.weekday() == 5 and d not in feriados:  # Saturday e não feriado
            days.append(d)
    return days

def previous_business_days(year:int, month:int, n:int, feriados:set):
    if n <= 0:
        return []
    d = first_business_day(year, month, feriados)
    prev = []
    cur = d - timedelta(days=1)
    while len(prev) < n:
        if cur.weekday() < 5 and cur not in feriados:
            prev.append(cur)
        cur -= timedelta(days=1)
    prev.sort()  # cronológico ascendente
    return prev

def month_label_pt_br(d: date) -> str:
    return f"{PT_BR_MONTHS[d.month].capitalize()}/{d.year}"

def month_key_text(year:int, month:int):
    # chaves aceitas no Excel, ex.: "jul/26", "Jul/2026", "JUL/26", etc.
    short = f"{PT_BR_MONTHS[month]}/{str(year)[2:]}"
    full = f"{PT_BR_MONTHS[month]}/{year}"
    alt_full_cap = f"{PT_BR_MONTHS[month].capitalize()}/{year}"
    alt_short_cap = f"{PT_BR_MONTHS[month].capitalize()}/{str(year)[2:]}"
    return {short, full, alt_full_cap, alt_short_cap}

def enumerate_months(start: date, end: date):
    # retorna lista de (year, month)
    mlist = []
    d = date(start.year, start.month, 1)
    stop = date(end.year, end.month, 1)
    while d <= stop:
        mlist.append((d.year, d.month))
        d += relativedelta(months=1)
    return mlist

# =========================
# Leitura e preparação da base
# =========================
def read_source_excel(file) -> pd.DataFrame:
    # Lê Planilha1
    df_src = pd.read_excel(file, sheet_name="Planilha1", engine="openpyxl")
    df_src.columns = [str(c).strip() for c in df_src.columns]
    return df_src

def map_month_columns(df: pd.DataFrame, months):
    """Retorna dict {(year,month): col_name} mapeando as colunas de mês encontradas.
       Aceita "jul/26" ... ou "Jul/2026" etc. Se não encontrar, assume 0."""
    colmap = {}
    existing = set(df.columns)
    for (y, m) in months:
        variants = month_key_text(y, m)
        found = None
        for v in variants:
            for col in existing:
                if str(col).strip().lower() == v.lower():
                    found = col
                    break
            if found:
                break
        colmap[(y, m)] = found  # pode ser None
    return colmap

def aggregate_by_product_market_model(df: pd.DataFrame, colmap):
    # Verifica colunas obrigatórias exatas
    req = ["MODELO", "PRODUTO", "MERCADO"]
    missing = [c for c in req if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes em Planilha1: {', '.join(missing)}")

    work = df.copy()
    # Completa meses não encontrados com 0; normaliza os demais para int
    for (y, m), cname in colmap.items():
        label = f"__{y}-{m:02d}__"
        if cname is None:
            work[label] = 0
        else:
            work[label] = pd.to_numeric(work[cname], errors="coerce").fillna(0).astype(int)

    # Agrega por PRODUTO/MERCADO/MODELO somando meses
    agg_cols = [c for c in work.columns if c.startswith("__")]
    group_cols = ["PRODUTO", "MERCADO", "MODELO"]
    agg = work.groupby(group_cols, dropna=False)[agg_cols].sum().reset_index()

    # Renomeia para chaves (y,m)
    renamed = {}
    for c in agg_cols:
        parts = c.strip("_").split("-")
        renamed[c] = (int(parts[0]), int(parts[1]))
    agg = agg.rename(columns=renamed)
    return agg

# =========================
# Núcleo: Programação mensal -> diária
# =========================
def proportional_quotas(balance_by_model: dict, S: int, limit_per_model: int|None):
    """Retorna dict model->cota do dia usando floor e maiores restos; respeita limit_per_model e saldo."""
    models = list(balance_by_model.keys())
    total = sum(balance_by_model[m] for m in models)
    if total == 0 or S == 0:
        return {m: 0 for m in models}

    # base floor
    floors = {m: int((balance_by_model[m] * S) // total) for m in models}
    # aplica teto diário por modelo, se houver (cap provisório)
    if limit_per_model is not None:
        floors = {m: min(floors[m], limit_per_model, balance_by_model[m]) for m in models}
    else:
        floors = {m: min(floors[m], balance_by_model[m]) for m in models}

    allocated = sum(floors.values())
    remainder = S - allocated
    # maiores restos
    if remainder > 0:
        rema_list = []
        for m in models:
            quota_raw = balance_by_model[m] * S / total
            rest = quota_raw - floors[m]
            rema_list.append((rest, m))
        rema_list.sort(reverse=True)
        for _, m in rema_list:
            if remainder <= 0:
                break
            cap_m = balance_by_model[m]
            if limit_per_model is not None:
                if floors[m] >= min(limit_per_model, cap_m):
                    continue
                floors[m] += 1
                remainder -= 1
            else:
                if floors[m] >= cap_m:
                    continue
                floors[m] += 1
                remainder -= 1
    return floors


def schedule_month(month_df: pd.DataFrame, year:int, month:int,
                   limite_diario_por_modelo:int|None, teto_sabado:int,
                   capacidade_dia_util:int, feriados:set,
                   dias_uteis_anteriores:int):
    """
    Gera a lista de produções do mês respeitando:
    - Capacidade 'capacidade_dia_util' por dia útil; excedente em sábados (até 'teto_sabado')
    - Pode iniciar até 'dias_uteis_anteriores' **dias úteis antes** do mês de referência
    - Balanceamento por MODELO com cotas proporcionais + maiores restos
    - Prioridade de mercados != MERCADO INTERNO antes de MI
    - Cumprimento exato do volume por PRODUTO no mês
    - Ignora feriados (seg-sex e sábados que estiverem na lista)
    Retorna lista de dicts: [{date, produto, modelo, mercado, mes_ref_year, mes_ref_month}, ...]
    """
    WEEKDAY_CAP = capacidade_dia_util

    # Saldos por (PRODUTO, MERCADO, MODELO)
    saldo_prod = {}
    for _, r in month_df.iterrows():
        key = (r["PRODUTO"], r["MERCADO"], r["MODELO"])
        saldo_prod[key] = int(r["qty"]) if pd.notnull(r["qty"]) else 0

    def total_necessario():
        return sum(saldo_prod.values())

    necessario_inicial = total_necessario()

    # Calendário do mês + dias úteis anteriores
    prev_days = previous_business_days(year, month, dias_uteis_anteriores, feriados)
    bdays_cur = business_days_in_month(year, month, feriados)
    all_weekdays = prev_days + bdays_cur  # cronológico ascendente

    capacidade_weekdays_total = len(all_weekdays) * WEEKDAY_CAP

    used_weekdays = []
    saturday_plan = []  # (date, capacity_for_that_saturday)

    if necessario_inicial <= capacidade_weekdays_total:
        need_days = (necessario_inicial + WEEKDAY_CAP - 1) // WEEKDAY_CAP if necessario_inicial > 0 else 0
        used_weekdays = all_weekdays[:need_days]
    else:
        used_weekdays = all_weekdays
        excedente = necessario_inicial - capacidade_weekdays_total
        sats = saturdays_in_month(year, month, feriados)
        for s in sats:
            if excedente <= 0:
                break
            cap_sab = min(teto_sabado, excedente)
            saturday_plan.append((s, cap_sab))
            excedente -= cap_sab
        if excedente > 0:
            raise ValueError(
                f"Excedente mensal de {excedente} não cabe nos sábados (não feriados) de {month:02d}/{year}. Aumente 'teto_sabado', 'capacidade_dia_util', 'dias_uteis_anteriores' ou revise a demanda.")

    def allocate_day(day_date: date, S: int, results: list):
        # Monta saldos por MODELO
        balance_by_model = {}
        for (prod, merc, mod), q in saldo_prod.items():
            if q > 0:
                balance_by_model[mod] = balance_by_model.get(mod, 0) + q
        if not balance_by_model or S <= 0:
            return

        # Limite diário
        limit = limite_diario_por_modelo if (limite_diario_por_modelo and limite_diario_por_modelo > 0) else None

        # Cotas por modelo
        quotas = proportional_quotas(balance_by_model, S, limit)

        # Constrói filas internas por MODELO: prioridade (mercado != MI) e MI
        per_model_priority = {m: [] for m in quotas.keys()}
        per_model_mi = {m: [] for m in quotas.keys()}
        for (prod, merc, mod), q in list(saldo_prod.items()):
            if q <= 0:
                continue
            if str(merc).strip().upper() == "MERCADO INTERNO":
                per_model_mi.setdefault(mod, []).append((prod, merc))
            else:
                per_model_priority.setdefault(mod, []).append((prod, merc))

        produced = 0
        model_order = [m for m in quotas.keys() if quotas[m] > 0]
        idx = 0
        guard = 0
        while produced < S and any(q > 0 for q in quotas.values()) and guard < 50000:
            guard += 1
            if not model_order:
                break
            m = model_order[idx % len(model_order)]
            if quotas[m] <= 0:
                idx += 1
                continue
            chosen_list = per_model_priority.get(m, []) or per_model_mi.get(m, [])
            if not chosen_list:
                quotas[m] = 0
                idx += 1
                continue
            prod, merc = chosen_list[0]
            key = (prod, merc, m)
            if saldo_prod.get(key, 0) <= 0:
                # remove e tenta próximo
                chosen_list.pop(0)
                continue
            # consome 1
            saldo_prod[key] -= 1
            quotas[m] -= 1
            produced += 1
            results.append({
                "date": day_date,
                "PRODUTO": prod,
                "MODELO": m,
                "MERCADO": merc,
                "mes_ref_year": year,
                "mes_ref_month": month,
            })
            # move par ao final da lista respectiva para round-robin interno
            chosen_list.append(chosen_list.pop(0))
            idx += 1

    results_month = []

    # Aloca nos dias úteis (inclui dias úteis anteriores), depois sábados do mês
    remaining_to_allocate = necessario_inicial
    for d in used_weekdays:
        rem = sum(1 for v in saldo_prod.values() if v > 0)
        if remaining_to_allocate <= 0:
            break
        cap = min(WEEKDAY_CAP, remaining_to_allocate)
        allocate_day(d, cap, results_month)
        remaining_to_allocate = sum(saldo_prod.values())

    for s, cap_sab in saturday_plan:
        if sum(saldo_prod.values()) <= 0:
            break
        allocate_day(s, cap_sab, results_month)

    if sum(saldo_prod.values()) != 0:
        raise ValueError(f"Falha ao alocar todo o volume de {year}-{month:02d}. Sobraram {sum(saldo_prod.values())} unidades.")

    return results_month

# =========================
# Pipeline completo
# =========================
def build_schedule(df_agg: pd.DataFrame, months, limite_diario_por_modelo:int|None, teto_sabado:int, capacidade_dia_util:int, feriados:set, dias_uteis_anteriores:int):
    all_rows = []
    relatorios = []
    for (y, m) in months:
        qty_col = (y, m)
        if qty_col not in df_agg.columns:
            df_agg[qty_col] = 0
        month_df = df_agg[["PRODUTO", "MERCADO", "MODELO", qty_col]].copy()
        month_df = month_df.rename(columns={qty_col: "qty"})
        month_df["qty"] = pd.to_numeric(month_df["qty"], errors="coerce").fillna(0).astype(int)
        necessario = int(month_df["qty"].sum())

        bdays = business_days_in_month(y, m, feriados)
        dias_uteis = len(bdays)
        capacidade = dias_uteis * capacidade_dia_util

        if necessario == 0:
            relatorios.append({
                "year": y, "month": m,
                "dias_uteis": dias_uteis,
                "capacidade": capacidade,
                "necessario": 0,
                "programado": 0,
                "unid_dia_extra": 0,
                "desvio": 0,
                "utilizacao": 0.0
            })
            continue

        month_plan = schedule_month(month_df, y, m,
                                    limite_diario_por_modelo=limite_diario_por_modelo,
                                    teto_sabado=teto_sabado,
                                    capacidade_dia_util=capacidade_dia_util,
                                    feriados=feriados,
                                    dias_uteis_anteriores=dias_uteis_anteriores)
        dfm = pd.DataFrame(month_plan)
        dfm["is_sat"] = dfm["date"].apply(lambda d: 1 if d.weekday()==5 else 0)
        programado_total = len(dfm)  # por mês de referência
        unid_extra = int(dfm["is_sat"].sum())
        relatorios.append({
            "year": y, "month": m,
            "dias_uteis": dias_uteis,
            "capacidade": capacidade,
            "necessario": necessario,
            "programado": programado_total,  # contado por mês de referência
            "unid_dia_extra": unid_extra,
            "desvio": programado_total - necessario,
            "utilizacao": (programado_total / capacidade) if capacidade > 0 else 0.0
        })
        all_rows.extend(month_plan)

    return all_rows, relatorios


def finalize_output(all_rows, relatorios):
    prog = pd.DataFrame(all_rows)
    if prog.empty:
        prog = pd.DataFrame(columns=["date","PRODUTO","MODELO","MERCADO","mes_ref_year","mes_ref_month"])

    # Conversão para datetime64[ns]
    if "date" in prog.columns:
        prog["date"] = pd.to_datetime(prog["date"])

    prog = prog.sort_values(by=["date","MODELO","PRODUTO"], kind="stable").reset_index(drop=True)

    # IDs globais sequenciais (fila 1...N)
    prog["ID"] = ["fila " + str(i) for i in range(1, len(prog)+1)]

    # Mês/ano (rótulo pt-BR)
    prog["mes_ano"] = prog["date"].apply(lambda d: month_label_pt_br(d.date() if hasattr(d, "date") else d))
    prog["mes"] = prog["date"].dt.month
    prog["ano"] = prog["date"].dt.year

    prog_out = prog.rename(columns={
        "date": "dt_producao",
        "mes_ano": "mes_ano_producao",
        "PRODUTO": "produto",
        "MODELO": "modelo",
        "ID": "ID",
        "MERCADO": "mercado",
        "mes": "mes_producao",
        "ano": "ano_producao"
    })[["dt_producao","mes_ano_producao","produto","modelo","ID","mercado","mes_producao","ano_producao"]]

    rel = pd.DataFrame(relatorios)
    if not rel.empty:
        rel["mes_ano"] = rel.apply(lambda r: f"{PT_BR_MONTHS[r['month']].capitalize()}/{int(r['year'])}", axis=1)
        rel = rel[["mes_ano","dias_uteis","capacidade","necessario","programado","unid_dia_extra","desvio","utilizacao"]]

    return prog_out, rel

# =========================
# Execução (UI)
# =========================
if gerar:
    if uploaded is None:
        st.error("Envie o Excel de entrada (aba 'Planilha1').")
        st.stop()

    try:
        df_src = read_source_excel(uploaded)
        months = enumerate_months(intervalo[0], intervalo[1])
        colmap = map_month_columns(df_src, months)
        df_agg = aggregate_by_product_market_model(df_src, colmap)

        # Diagnóstico opcional (ajuda a validar mapeamento e totais)
        with st.expander("🔎 Diagnóstico de mapeamento e totais lidos", expanded=False):
            mapped = {f"{y}-{m:02d}": (colmap[(y,m)] if colmap[(y,m)] is not None else "(não encontrado)") for (y,m) in months}
            st.write("**Mapeamento de colunas (ano-mês → coluna):**")
            st.json(mapped, expanded=False)
            totais = []
            for (y,m) in months:
                col = (y,m)
                if col not in df_agg.columns:
                    df_agg[col] = 0
                total_mes = int(pd.to_numeric(df_agg[col], errors="coerce").fillna(0).sum())
                totais.append({"ano": y, "mes": m, "total_lido": total_mes})
            df_tot = pd.DataFrame(totais)
            df_tot["mes_ano"] = df_tot.apply(lambda r: f"{PT_BR_MONTHS[r['mes']].capitalize()}/{r['ano']}", axis=1)
            st.write("**Necessário por mês (somado do Excel):**")
            st.dataframe(df_tot[["mes_ano","total_lido"]], use_container_width=True)
            st.write("**Soma total lida:**", int(df_tot["total_lido"].sum()))
            st.write("**Feriados considerados:**", ", ".join(sorted([d.isoformat() for d in feriados_set])) or "(nenhum)")

        lmt = limite_diario_por_modelo if limite_diario_por_modelo > 0 else None
        rows, relatorios = build_schedule(df_agg, months, lmt, teto_sabado, capacidade_dia_util, feriados_set, dias_uteis_anteriores)
        prog_out, rel_out = finalize_output(rows, relatorios)

        por_mercado = prog_out.groupby("mercado").size().reset_index(name="programado")
        por_modelo = prog_out.groupby("modelo").size().reset_index(name="programado")

        st.success("Programação gerada com sucesso!")
        st.subheader("📄 Programação (amostra)")
        st.dataframe(prog_out.head(50), use_container_width=True)
        st.subheader("📊 Relatório (mensal)")
        st.dataframe(rel_out, use_container_width=True)
        st.caption("Obs.: 'programado' por mês no relatório considera o **mês de referência da demanda**, mesmo quando parte da produção inicia em dias úteis do mês anterior.")
        st.subheader("📊 Programado por MERCADO")
        st.dataframe(por_mercado, use_container_width=True)
        st.subheader("📊 Programado por MODELO")
        st.dataframe(por_modelo, use_container_width=True)

        out_name = nomenclatura_arquivo.strip() if nomenclatura_arquivo.strip() else "programacao_2S2026_balanceada.xlsx"
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            prog_out.to_excel(writer, index=False, sheet_name="Programacao_2S2026")
            rel_out.to_excel(writer, index=False, sheet_name="Relatorio_2S2026")
            # Copia base original
            df_src.to_excel(writer, index=False, sheet_name="Base_Original")
            # Extras
            por_mercado.to_excel(writer, index=False, sheet_name="Programado_por_MERCADO")
            por_modelo.to_excel(writer, index=False, sheet_name="Programado_por_MODELO")

        st.download_button(
            "⬇️ Baixar arquivo final",
            data=buffer.getvalue(),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        st.exception(e)
