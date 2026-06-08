import streamlit as st
import pandas as pd
import altair as alt
from collections import defaultdict
from datetime import date
import gspread
from google.oauth2.service_account import Credentials
import json, os

SHEET_ID = "1LR63NFna6y1z88aI1HxRKq8kyNGaPN_3-dLzX3kn5Ys"

_MES_NUM = {
    'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4,
    'Mai': 5, 'Jun': 6, 'Jul': 7, 'Ago': 8,
    'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12,
}

st.set_page_config(
    page_title="Financas 2026",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .block-container { padding-top: 0.8rem; padding-bottom: 1rem; max-width: 500px; }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: 700; }
    div[data-testid="stMetricLabel"] { font-size: 0.75rem; color: #888; }
    thead tr th { font-size: 0.75rem !important; }
    tbody tr td { font-size: 0.8rem !important; }
</style>
""", unsafe_allow_html=True)


def fmt(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _ordem_mes(aba):
    return _MES_NUM.get(aba.split('.')[0], 99)


def _get_creds():
    # Streamlit Cloud: credenciais em st.secrets
    # Local: lê do arquivo JSON
    try:
        info = dict(st.secrets["gcp_service_account"])
        return Credentials.from_service_account_info(info, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ])
    except Exception:
        # fallback local
        local = os.path.join(os.path.dirname(__file__), "credenciais_google.json")
        return Credentials.from_service_account_file(local, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ])


@st.cache_data(ttl=300, show_spinner="Carregando dados...")
def carregar():
    creds = _get_creds()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)

    # Despesas
    rows_d = sh.worksheet("Despesas").get_all_records()
    despesas   = defaultdict(list)
    total_mes  = defaultdict(float)
    for r in rows_d:
        if not r.get("Mes") or not r.get("Valor"):
            continue
        try:
            valor = float(r["Valor"])
        except (ValueError, TypeError):
            continue
        mes = r["Mes"]
        despesas[mes].append({
            "data":      str(r.get("Data", "")),
            "descricao": str(r.get("Descricao", "")),
            "valor":     valor,
            "cartao":    str(r.get("Cartao", "")),
            "categoria": str(r.get("Categoria", "")),
            "parcela":   str(r.get("Parcela", "")),
            "tipo":      str(r.get("Tipo", "variavel")),
        })
        total_mes[mes] += valor

    # Recebimentos
    rows_r = sh.worksheet("Recebimentos").get_all_records()
    receitas = []
    for r in rows_r:
        if not r.get("Origem") or not r.get("Valor"):
            continue
        try:
            valor = float(r["Valor"])
        except (ValueError, TypeError):
            continue
        receitas.append({
            "origem":    str(r.get("Origem", "")),
            "categoria": str(r.get("Categoria", "")),
            "referencia":str(r.get("Referencia", "")),
            "data":      str(r.get("Data", "")),
            "tipo":      str(r.get("Tipo", "")),
            "valor":     valor,
        })

    return dict(despesas), dict(total_mes), receitas


# ── Header ────────────────────────────────────────────────
st.header("Financas 2026")
if st.button("Atualizar dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

despesas, total_mes, receitas = carregar()
meses = sorted(despesas.keys(), key=_ordem_mes)

pagina = st.radio(
    "nav", ["Resumo", "Despesas", "Recebimentos"],
    horizontal=True, label_visibility="collapsed"
)
st.divider()


# ──────────────────────────────────────────
# RESUMO
# ──────────────────────────────────────────
if pagina == "Resumo":
    total_desp = sum(total_mes.values())
    total_rec  = sum(r['valor'] for r in receitas if '2026' in r['data'])
    saldo      = total_rec - total_desp
    media_mes  = total_desp / len(meses) if meses else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Receitas",  fmt(total_rec))
    c2.metric("Despesas",  fmt(total_desp))
    c3.metric("Saldo",     fmt(saldo))

    st.divider()

    if meses:
        hoje = date.today()
        mes_abrev = list(_MES_NUM.keys())[hoje.month - 1]
        mes_atual_key = f"{mes_abrev}.{hoje.year}"
        if mes_atual_key in meses:
            atual = mes_atual_key
        else:
            candidatos = [m for m in meses if _ordem_mes(m) <= hoje.month]
            atual = max(candidatos, key=_ordem_mes) if candidatos else meses[0]

        t_atual = total_mes[atual]
        t_fixo  = sum(t['valor'] for t in despesas[atual] if t.get('tipo') == 'fixo')
        t_var   = sum(t['valor'] for t in despesas[atual] if t.get('tipo') == 'variavel')

        st.subheader(f"Mes atual — {atual}")

        if len(meses) >= 2:
            anterior  = meses[meses.index(atual) - 1] if meses.index(atual) > 0 else None
            if anterior:
                t_ant     = total_mes[anterior]
                delta     = t_atual - t_ant
                delta_pct = delta / t_ant * 100 if t_ant else 0
                sinal     = "+" if delta >= 0 else ""
                delta_label = f"{sinal}{fmt(delta)} ({sinal}{delta_pct:.1f}% vs {anterior})"
            else:
                delta_label = None
        else:
            delta_label = None

        c1, c2, c3 = st.columns(3)
        c1.metric("Total",     fmt(t_atual), delta=delta_label, delta_color="inverse")
        c2.metric("Fixos",     fmt(t_fixo))
        c3.metric("Variaveis", fmt(t_var))

        por_cat = defaultdict(float)
        for t in despesas[atual]:
            por_cat[t['categoria']] += t['valor']
        top3 = sorted(por_cat.items(), key=lambda x: -x[1])[:3]

        st.markdown("**Top 3 categorias**")
        for cat, val in top3:
            pct = val / t_atual * 100 if t_atual else 0
            st.markdown(f"- **{cat}**: {fmt(val)} — {pct:.0f}%")

        mes_caro = max(total_mes, key=total_mes.get)
        st.markdown(f"**Mes mais caro:** {mes_caro} — {fmt(total_mes[mes_caro])}")
        st.markdown(f"**Media mensal:** {fmt(media_mes)}")

    st.divider()

    st.subheader("Despesas por mes")
    df_bar = pd.DataFrame(
        [(m, round(total_mes[m], 2)) for m in meses],
        columns=["Mes", "Total"]
    )
    chart = alt.Chart(df_bar).mark_bar().encode(
        x=alt.X("Mes:N", sort=meses, axis=alt.Axis(labelAngle=-45, title=None)),
        y=alt.Y("Total:Q", title="R$"),
        tooltip=["Mes", "Total"],
    ).properties(height=280)
    st.altair_chart(chart, use_container_width=True)

    with st.expander("Resumo por mes"):
        rows = []
        for m in meses:
            fix = sum(t['valor'] for t in despesas[m] if t.get('tipo') == 'fixo')
            var = sum(t['valor'] for t in despesas[m] if t.get('tipo') == 'variavel')
            rows.append({"Mes": m, "Total": fmt(total_mes[m]),
                         "Fixos": fmt(fix), "Variaveis": fmt(var)})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ──────────────────────────────────────────
# DESPESAS
# ──────────────────────────────────────────
elif pagina == "Despesas":
    hoje = date.today()
    mes_abrev = list(_MES_NUM.keys())[hoje.month - 1]
    mes_atual_key = f"{mes_abrev}.{hoje.year}"
    idx_default = meses.index(mes_atual_key) if mes_atual_key in meses else len(meses) - 1
    mes = st.selectbox("Mes", meses, index=idx_default)

    if mes in despesas:
        lista = list(despesas[mes])

        t_fixo = sum(t['valor'] for t in lista if t.get('tipo') == 'fixo')
        t_var  = sum(t['valor'] for t in lista if t.get('tipo') == 'variavel')

        c1, c2, c3 = st.columns(3)
        c1.metric("Total",     fmt(total_mes[mes]))
        c2.metric("Fixos",     fmt(t_fixo))
        c3.metric("Variaveis", fmt(t_var))

        tipo_sel = st.radio("Tipo", ["Todos", "Variaveis", "Fixos"], horizontal=True)
        if tipo_sel == "Variaveis":
            lista = [t for t in lista if t.get('tipo') == 'variavel']
        elif tipo_sel == "Fixos":
            lista = [t for t in lista if t.get('tipo') == 'fixo']

        cats = ["Todas"] + sorted(set(t['categoria'] for t in lista if t['categoria']))
        cat = st.selectbox("Categoria", cats)
        if cat != "Todas":
            lista = [t for t in lista if t['categoria'] == cat]

        busca = st.text_input("Buscar", placeholder="ex: Uber, Pizza, BB...")
        if busca:
            lista = [t for t in lista if busca.lower() in t['descricao'].lower()]

        subtotal = sum(t['valor'] for t in lista)
        st.caption(f"{len(lista)} transacoes  |  subtotal: {fmt(subtotal)}")

        rows = []
        for t in lista:
            d = t['data']
            if len(d) >= 10 and d[4] == '-':
                d = d[8:10] + "/" + d[5:7]
            else:
                d = d[:5]
            rows.append({
                "Data":      d,
                "Descricao": t['descricao'][:32],
                "Valor":     fmt(t['valor']),
                "Categoria": t['categoria'],
                "Tipo":      "Fixo" if t.get('tipo') == 'fixo' else "Var",
                "Cartao":    t['cartao'],
            })

        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        else:
            st.info("Nenhuma transacao encontrada.")


# ──────────────────────────────────────────
# RECEBIMENTOS
# ──────────────────────────────────────────
elif pagina == "Recebimentos":
    anos = sorted(set(r['data'][:4] for r in receitas if len(r['data']) >= 4), reverse=True)
    if not anos:
        st.info("Nenhum recebimento encontrado.")
    else:
        ano = st.selectbox("Ano", anos)

        filtrado = sorted(
            [r for r in receitas if r['data'].startswith(ano)],
            key=lambda x: x['data'], reverse=True
        )

        total_ano = sum(r['valor'] for r in filtrado)
        st.metric(f"Total {ano}", fmt(total_ano))

        por_cat = defaultdict(float)
        for r in filtrado:
            por_cat[r['categoria']] += r['valor']

        with st.expander("Ver por categoria"):
            df_cat = pd.DataFrame(
                sorted(por_cat.items(), key=lambda x: -x[1]),
                columns=["Categoria", "Valor"]
            )
            df_cat["Valor"] = df_cat["Valor"].apply(fmt)
            st.dataframe(df_cat, hide_index=True, use_container_width=True)

        rows = []
        for r in filtrado:
            d = r['data']
            if len(d) >= 10 and d[4] == '-':
                d = d[8:10] + "/" + d[5:7] + "/" + d[:4]
            rows.append({
                "Data":      d[:10],
                "Origem":    r['origem'],
                "Categoria": r['categoria'],
                "Tipo":      r['tipo'],
                "Valor":     fmt(r['valor']),
            })

        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
