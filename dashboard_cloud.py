import streamlit as st
import pandas as pd
import altair as alt
from collections import defaultdict
from datetime import date
import gspread
from google.oauth2.service_account import Credentials
import os

SHEET_ID = "1LR63NFna6y1z88aI1HxRKq8kyNGaPN_3-dLzX3kn5Ys"

_MES_NUM = {
    'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4,
    'Mai': 5, 'Jun': 6, 'Jul': 7, 'Ago': 8,
    'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12,
}

CAT_COR = {
    'MORADIA': '#4c8bf5', 'CARRO': '#94a3b8', 'MOTO': '#f97316',
    'BARCO': '#06b6d4', 'D. PESSOAL': '#a855f7', 'INVEST.': '#10b981',
    'THALI': '#ec4899', 'JULI': '#eab308',
    'V2PA': '#6366f1', 'SAFO ENG.': '#6366f1',
}
CAT_ICONE = {
    'MORADIA': '🏠', 'CARRO': '🚗', 'MOTO': '🏍️', 'BARCO': '⛵',
    'D. PESSOAL': '👤', 'INVEST.': '📈', 'THALI': '💝', 'JULI': '👶',
    'V2PA': '🏢', 'SAFO ENG.': '🏢',
}

# ── Page config (deve ser o primeiro comando Streamlit) ──────────────
st.set_page_config(
    page_title="Finanças 2026",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .block-container { padding-top: 0.8rem; padding-bottom: 1rem; max-width: 520px; }
    div[data-testid="stMetricValue"] { font-size: 1.25rem !important; font-weight: 700 !important; }
    div[data-testid="stMetricLabel"] { font-size: 0.7rem !important; color: #94a3b8 !important; }
    div[data-testid="stMetricDelta"] { font-size: 0.72rem !important; }
    .stRadio > div { gap: 4px; }
    .stTextInput input { border-radius: 10px; }
    /* Destaque mês atual no selectbox */
    .st-emotion-cache-1fttcpj { font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)

# ── Autenticação ─────────────────────────────────────────────────────
if not st.session_state.get("autenticado"):
    st.markdown("<style>.block-container{max-width:340px;padding-top:4rem;}</style>",
                unsafe_allow_html=True)
    st.markdown("## 💰 Finanças 2026")
    senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")
    if st.button("Entrar", use_container_width=True, type="primary"):
        try:
            correta = st.secrets["app_password"]
        except Exception:
            try:
                correta = st.secrets["gcp_service_account"]["app_password"]
            except Exception:
                correta = ""
        if senha == correta:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()


# ── Helpers ──────────────────────────────────────────────────────────
def fmt(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_k(v):
    if abs(v) >= 1000:
        return f"R$ {v/1000:.1f}k".replace(".", ",")
    return fmt(v)

def _ordem_mes(aba):
    return _MES_NUM.get(aba.split('.')[0], 99)

def _parse_brl(v):
    """Converte string pt-BR (ex: '17547,78' ou '17.547,78') para float."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    # Remove separador de milhar e troca decimal
    s = s.replace('.', '').replace(',', '.')
    return float(s)

def _get_creds():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        info = dict(st.secrets["gcp_service_account"])
        return Credentials.from_service_account_info(info, scopes=scopes)
    except Exception:
        local = os.path.join(os.path.dirname(__file__), "credenciais_google.json")
        return Credentials.from_service_account_file(local, scopes=scopes)

def _get_sheet():
    gc = gspread.authorize(_get_creds())
    return gc.open_by_key(SHEET_ID)


# ── Carregamento de dados ────────────────────────────────────────────
def _ws_to_dicts(ws):
    """Lê worksheet como lista de dicts usando get_all_values() (strings brutas)."""
    rows = ws.get_all_values()
    if not rows:
        return []
    headers = rows[0]
    return [dict(zip(headers, row)) for row in rows[1:]]

@st.cache_data(ttl=300, show_spinner="Carregando dados...")
def carregar():
    sh = _get_sheet()

    despesas  = defaultdict(list)
    total_mes = defaultdict(float)
    for r in _ws_to_dicts(sh.worksheet("Despesas")):
        if not r.get("Mes") or not r.get("Valor"):
            continue
        try:
            valor = _parse_brl(r["Valor"])
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

    receitas = []
    for r in _ws_to_dicts(sh.worksheet("Recebimentos")):
        if not r.get("Origem") or not r.get("Valor"):
            continue
        try:
            valor = _parse_brl(r["Valor"])
        except (ValueError, TypeError):
            continue
        receitas.append({
            "origem":     str(r.get("Origem", "")),
            "categoria":  str(r.get("Categoria", "")),
            "referencia": str(r.get("Referencia", "")),
            "data":       str(r.get("Data", "")),
            "tipo":       str(r.get("Tipo", "")),
            "valor":      valor,
        })

    return dict(despesas), dict(total_mes), receitas


# ── Solicitar sync (escreve flag no Google Sheets) ───────────────────
def solicitar_sync():
    try:
        sh = _get_sheet()
        try:
            ws = sh.worksheet("Config")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet("Config", 10, 2)
            ws.update([["Key", "Value"], ["sync_pending", "FALSE"], ["last_sync", "-"]])
        cell = ws.find("sync_pending")
        ws.update_cell(cell.row, 2, "TRUE")
        return True
    except Exception:
        return False

def status_sync():
    try:
        sh = _get_sheet()
        ws = sh.worksheet("Config")
        data = {r["Key"]: r["Value"] for r in ws.get_all_records() if r.get("Key")}
        return data.get("sync_pending", "?"), data.get("last_sync", "-")
    except Exception:
        return "?", "-"


# ── Contexto para o chat ─────────────────────────────────────────────
def montar_contexto(despesas, total_mes, receitas, meses):
    linhas = ["=== PLANILHA FINANCEIRA 2026 — DADOS COMPLETOS ===\n"]
    linhas.append("## DESPESAS POR MÊS")
    for m in meses:
        total = total_mes.get(m, 0)
        fix = sum(t['valor'] for t in despesas.get(m, []) if t.get('tipo') == 'fixo')
        var = sum(t['valor'] for t in despesas.get(m, []) if t.get('tipo') == 'variavel')
        linhas.append(f"  {m}: Total {fmt(total)} | Fixos {fmt(fix)} | Variáveis {fmt(var)}")
        por_cat = defaultdict(float)
        for t in despesas.get(m, []):
            por_cat[t['categoria']] += t['valor']
        for cat, val in sorted(por_cat.items(), key=lambda x: -x[1]):
            linhas.append(f"    {cat}: {fmt(val)}")

    linhas.append("\n## RECEITAS 2026")
    total_rec = sum(r['valor'] for r in receitas if '2026' in r.get('data', ''))
    linhas.append(f"  Total: {fmt(total_rec)}")
    por_cat_rec = defaultdict(float)
    for r in receitas:
        if '2026' in r.get('data', ''):
            por_cat_rec[r['categoria']] += r['valor']
    for cat, val in sorted(por_cat_rec.items(), key=lambda x: -x[1]):
        linhas.append(f"    {cat}: {fmt(val)}")

    return "\n".join(linhas)


# ── Carrega dados ────────────────────────────────────────────────────
despesas, total_mes, receitas = carregar()
meses = sorted(despesas.keys(), key=_ordem_mes)

hoje       = date.today()
mes_abrev  = list(_MES_NUM.keys())[hoje.month - 1]
mes_key    = f"{mes_abrev}.{hoje.year}"
if mes_key in meses:
    mes_atual = mes_key
else:
    cands     = [m for m in meses if _ordem_mes(m) <= hoje.month]
    mes_atual = max(cands, key=_ordem_mes) if cands else (meses[0] if meses else "")

total_desp = sum(total_mes.values())
total_rec  = sum(r['valor'] for r in receitas if '2026' in r.get('data', ''))
saldo      = total_rec - total_desp


# ── Header ───────────────────────────────────────────────────────────
col_t, col_r = st.columns([4, 1])
with col_t:
    st.markdown("### 💰 Finanças 2026")
    cor = "#10b981" if saldo >= 0 else "#ef4444"
    st.markdown(
        f"<span style='color:{cor};font-weight:600'>Saldo anual: {fmt(saldo)}</span>",
        unsafe_allow_html=True,
    )
with col_r:
    if st.button("🔄", help="Atualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

pagina = st.radio(
    "nav", ["Início", "Despesas", "Receitas", "Chat IA", "Dados"],
    horizontal=True, label_visibility="collapsed",
)
st.divider()


# ════════════════════════════════════════════════════════════════════
# INÍCIO
# ════════════════════════════════════════════════════════════════════
if pagina == "Início":
    t_atual = total_mes.get(mes_atual, 0)
    t_fixo  = sum(t['valor'] for t in despesas.get(mes_atual, []) if t.get('tipo') == 'fixo')
    t_var   = sum(t['valor'] for t in despesas.get(mes_atual, []) if t.get('tipo') == 'variavel')

    st.markdown(f"**{mes_atual}**")
    c1, c2, c3 = st.columns(3)

    idx_atual = meses.index(mes_atual) if mes_atual in meses else -1
    if idx_atual > 0:
        t_ant     = total_mes.get(meses[idx_atual - 1], 0)
        delta_pct = (t_atual - t_ant) / t_ant * 100 if t_ant else 0
        sinal     = "+" if delta_pct >= 0 else ""
        delta_lbl = f"{sinal}{delta_pct:.1f}% vs {meses[idx_atual-1]}"
    else:
        delta_lbl = None

    c1.metric("Despesas",   fmt_k(t_atual), delta=delta_lbl, delta_color="inverse")
    c2.metric("Fixos",      fmt_k(t_fixo))
    c3.metric("Variáveis",  fmt_k(t_var))

    # Receita e saldo do mês
    mes_str = f"{hoje.month:02d}"
    ano_str = str(hoje.year)
    rec_mes = sum(
        r['valor'] for r in receitas
        if r.get('data', '')[3:5] == mes_str
        and r.get('data', '')[6:10] == ano_str
    )
    if rec_mes:
        saldo_mes = rec_mes - t_atual
        c1b, c2b = st.columns(2)
        c1b.metric("Receitas do mês", fmt_k(rec_mes))
        c2b.metric("Saldo do mês", fmt_k(saldo_mes))

    st.divider()

    # Categorias com barras horizontais (Altair)
    st.markdown("**Gastos por categoria**")
    por_cat = defaultdict(float)
    for t in despesas.get(mes_atual, []):
        por_cat[t['categoria']] += t['valor']

    if por_cat:
        df_cat = pd.DataFrame(
            [(cat, val, CAT_COR.get(cat.upper(), '#6366f1'),
              CAT_ICONE.get(cat.upper(), '💳'))
             for cat, val in sorted(por_cat.items(), key=lambda x: -x[1])[:10]],
            columns=["Categoria", "Valor", "Cor", "Icone"],
        )
        df_cat["Label"] = df_cat.apply(
            lambda r: f"{r['Icone']} {r['Categoria']}", axis=1
        )
        chart_cat = alt.Chart(df_cat).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            y=alt.Y("Label:N", sort="-x", axis=alt.Axis(title=None, labelFontSize=11)),
            x=alt.X("Valor:Q", axis=alt.Axis(title=None, labelFontSize=10)),
            color=alt.Color("Cor:N", scale=None, legend=None),
            tooltip=["Categoria", alt.Tooltip("Valor:Q", format=",.2f")],
        ).properties(height=min(300, len(df_cat) * 32 + 40))
        st.altair_chart(chart_cat, use_container_width=True)

    st.divider()

    # Evolução mensal
    st.markdown("**Evolução mensal**")
    df_bar = pd.DataFrame(
        [(m, round(total_mes[m], 2), m == mes_atual) for m in meses],
        columns=["Mes", "Total", "Atual"],
    )
    chart_bar = alt.Chart(df_bar).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("Mes:N", sort=meses, axis=alt.Axis(labelAngle=-45, title=None, labelFontSize=10)),
        y=alt.Y("Total:Q", axis=alt.Axis(title=None, labelFontSize=10)),
        color=alt.condition(
            alt.datum.Atual,
            alt.value("#7c3aed"),
            alt.value("#3d4270"),
        ),
        tooltip=["Mes", alt.Tooltip("Total:Q", format=",.2f")],
    ).properties(height=220)
    st.altair_chart(chart_bar, use_container_width=True)

    with st.expander("📊 Tabela resumo anual"):
        rows = []
        for m in meses:
            fix = sum(t['valor'] for t in despesas[m] if t.get('tipo') == 'fixo')
            var = sum(t['valor'] for t in despesas[m] if t.get('tipo') == 'variavel')
            rows.append({"Mês": m, "Total": fmt(total_mes[m]),
                         "Fixos": fmt(fix), "Variáveis": fmt(var)})
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# DESPESAS
# ════════════════════════════════════════════════════════════════════
elif pagina == "Despesas":
    idx_def = meses.index(mes_atual) if mes_atual in meses else len(meses) - 1
    mes = st.selectbox("Mês", meses, index=idx_def)

    if mes in despesas:
        lista  = list(despesas[mes])
        t_tot  = total_mes[mes]
        t_fixo = sum(t['valor'] for t in lista if t.get('tipo') == 'fixo')
        t_var  = sum(t['valor'] for t in lista if t.get('tipo') == 'variavel')

        c1, c2, c3 = st.columns(3)
        c1.metric("Total",     fmt_k(t_tot))
        c2.metric("Fixos",     fmt_k(t_fixo))
        c3.metric("Variáveis", fmt_k(t_var))

        ca, cb = st.columns(2)
        with ca:
            tipo_sel = st.selectbox("Tipo", ["Todos", "Variáveis", "Fixos"])
        with cb:
            cats = ["Todas"] + sorted(set(t['categoria'] for t in lista if t['categoria']))
            cat  = st.selectbox("Categoria", cats)

        busca = st.text_input("🔍 Buscar", placeholder="Uber, Supermercado...")

        if tipo_sel == "Variáveis":
            lista = [t for t in lista if t.get('tipo') == 'variavel']
        elif tipo_sel == "Fixos":
            lista = [t for t in lista if t.get('tipo') == 'fixo']
        if cat != "Todas":
            lista = [t for t in lista if t['categoria'] == cat]
        if busca:
            lista = [t for t in lista if busca.lower() in t['descricao'].lower()]

        subtotal = sum(t['valor'] for t in lista)
        st.caption(f"{len(lista)} transações · subtotal: {fmt(subtotal)}")

        rows = []
        for t in sorted(lista, key=lambda x: x['data'], reverse=True):
            d = t['data']
            if len(d) >= 10 and d[4] == '-':
                d = d[8:10] + "/" + d[5:7]
            else:
                d = d[:5]
            icone = CAT_ICONE.get(t['categoria'].upper(), '💳')
            rows.append({
                "Data":      d,
                "":          icone,
                "Descrição": t['descricao'][:30],
                "Valor":     fmt(t['valor']),
                "Cat.":      t['categoria'],
                "Tipo":      "Fix" if t.get('tipo') == 'fixo' else "Var",
            })
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        else:
            st.info("Nenhuma transação encontrada.")


# ════════════════════════════════════════════════════════════════════
# RECEITAS
# ════════════════════════════════════════════════════════════════════
elif pagina == "Receitas":
    anos = sorted(set(r['data'][:4] for r in receitas if len(r['data']) >= 4), reverse=True)
    if not anos:
        st.info("Nenhum recebimento encontrado.")
        st.stop()

    ano      = st.selectbox("Ano", anos)
    filtrado = sorted(
        [r for r in receitas if r['data'].startswith(ano)],
        key=lambda x: x['data'], reverse=True,
    )
    total_ano = sum(r['valor'] for r in filtrado)
    st.metric(f"Total recebido em {ano}", fmt(total_ano))

    # Breakdown por categoria (barras)
    por_cat = defaultdict(float)
    for r in filtrado:
        por_cat[r['categoria']] += r['valor']

    if por_cat:
        df_rec = pd.DataFrame(
            sorted(por_cat.items(), key=lambda x: -x[1]),
            columns=["Categoria", "Valor"],
        )
        chart_rec = alt.Chart(df_rec).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
            y=alt.Y("Categoria:N", sort="-x", axis=alt.Axis(title=None, labelFontSize=11)),
            x=alt.X("Valor:Q", axis=alt.Axis(title=None)),
            color=alt.value("#10b981"),
            tooltip=["Categoria", alt.Tooltip("Valor:Q", format=",.2f")],
        ).properties(height=min(280, len(df_rec) * 34 + 40))
        st.altair_chart(chart_rec, use_container_width=True)

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


# ════════════════════════════════════════════════════════════════════
# CHAT IA
# ════════════════════════════════════════════════════════════════════
elif pagina == "Chat IA":
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", "") or st.secrets.get("gemini_api_key", "")
    except Exception:
        api_key = ""

    if not api_key:
        st.warning("⚙️ Configure a chave Gemini para ativar o Chat IA.")
        st.code('GEMINI_API_KEY = "sua-chave-aqui"', language="toml")
        st.caption("Adicione essa linha nos Secrets do Streamlit Cloud (share.streamlit.io → seu app → Settings → Secrets)")
        st.stop()

    from google import genai as genai_new
    from google.genai import types as genai_types

    if "chat_historico" not in st.session_state or st.session_state.get("chat_mes") != mes_atual:
        ctx = montar_contexto(despesas, total_mes, receitas, meses)
        system = f"""Você é um assistente financeiro pessoal inteligente e direto do Vinicius.
Você tem acesso completo aos dados financeiros dele para 2026.
Dê insights práticos, compare meses quando relevante, identifique padrões.
Use valores em R$ com vírgula decimal. Responda em português brasileiro.
Seja objetivo — máximo 3 parágrafos por resposta, a não ser que seja pedido mais detalhe.

{ctx}"""
        try:
            client = genai_new.Client(api_key=api_key)
            st.session_state.gemini_client = client
            st.session_state.gemini_system = system
            st.session_state.gemini_history = []
            st.session_state.chat_historico = []
            st.session_state.chat_mes = mes_atual
        except Exception as e:
            st.error(f"Erro ao iniciar Gemini: {e}")
            st.stop()

    st.markdown("**🤖 Assistente Financeiro**")
    st.caption("Pergunte sobre gastos, tendências e onde economizar.")

    # Sugestões rápidas
    sugestoes = [
        "Como estão meus gastos este mês?",
        "Qual categoria gasto mais no ano?",
        "Compare os últimos 3 meses",
        "Onde posso economizar?",
        "Meu saldo está saudável?",
        "Quais são meus maiores custos fixos?",
    ]
    cols = st.columns(2)
    for i, s in enumerate(sugestoes):
        if cols[i % 2].button(s, use_container_width=True, key=f"sug_{i}"):
            st.session_state.chat_historico.append({"role": "user", "content": s})
            try:
                client = st.session_state.gemini_client
                hist   = st.session_state.gemini_history
                hist.append(genai_types.Content(role="user", parts=[genai_types.Part(text=s)]))
                resp = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=hist,
                    config=genai_types.GenerateContentConfig(system_instruction=st.session_state.gemini_system),
                )
                texto = resp.text
                hist.append(genai_types.Content(role="model", parts=[genai_types.Part(text=texto)]))
                st.session_state.chat_historico.append({"role": "assistant", "content": texto})
            except Exception as e:
                st.session_state.chat_historico.append({"role": "assistant", "content": f"Erro: {e}"})
            st.rerun()

    st.divider()

    for msg in st.session_state.chat_historico:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Pergunte sobre suas finanças..."):
        st.session_state.chat_historico.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("Analisando..."):
                try:
                    client = st.session_state.gemini_client
                    hist = st.session_state.gemini_history
                    hist.append(genai_types.Content(role="user", parts=[genai_types.Part(text=prompt)]))
                    resp = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=hist,
                        config=genai_types.GenerateContentConfig(system_instruction=st.session_state.gemini_system),
                    )
                    texto = resp.text
                    hist.append(genai_types.Content(role="model", parts=[genai_types.Part(text=texto)]))
                    st.markdown(texto)
                    st.session_state.chat_historico.append({"role": "assistant", "content": texto})
                except Exception as e:
                    st.error(f"Erro: {e}")

    if st.session_state.chat_historico:
        if st.button("🗑️ Limpar conversa"):
            st.session_state.chat_historico = []
            st.session_state.pop("gemini_model", None)
            st.rerun()


# ════════════════════════════════════════════════════════════════════
# DADOS
# ════════════════════════════════════════════════════════════════════
elif pagina == "Dados":
    st.markdown("**📡 Sincronização**")
    st.caption("Solicita ao PC que envie os dados mais recentes da planilha Excel para o Google Sheets. O PC precisa estar ligado.")

    pending, last_sync = status_sync()
    if pending == "TRUE":
        st.warning("⏳ Sincronização em andamento...")
    elif pending == "FALSE" and last_sync != "-":
        st.success(f"✅ Última sincronização: {last_sync}")

    ca, cb = st.columns(2)
    if ca.button("📤 Sincronizar agora", use_container_width=True, type="primary"):
        with st.spinner("Enviando solicitação..."):
            ok = solicitar_sync()
        if ok:
            st.success("Solicitação enviada! Dados atualizados em ~2 minutos se o PC estiver ligado.")
        else:
            st.error("Erro ao enviar. Verifique a conexão.")

    if cb.button("🔄 Recarregar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.markdown("**📊 Status**")
    total_tx = sum(len(v) for v in despesas.values())
    col1, col2 = st.columns(2)
    col1.metric("Meses com dados", len(meses))
    col2.metric("Total transações", total_tx)
    col1.metric("Receitas", len(receitas))
    col2.metric("Último mês", meses[-1] if meses else "-")

    st.divider()
    if st.button("🔒 Sair", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
