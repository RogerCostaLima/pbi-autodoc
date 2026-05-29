"""
Power BI Doc Fácil — Streamlit App
Documentação técnica para .pbix, sem abrir o Power BI Desktop.
"""

import streamlit as st
import sys
import os
import time
import json
import zipfile
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from core.extractor import extract_pbix
from core.docgen import generate_docx
from core.ai_providers import generate_ai_overview

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PBI Doc Fácil",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=DM+Sans:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1100px; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #1A1A2E !important;
}
section[data-testid="stSidebar"] * { color: #C8C8E8 !important; }
section[data-testid="stSidebar"] .stRadio label { font-size: 0.95rem !important; }

/* Botão primário */
.stButton > button {
    background: #F0A500 !important;
    color: #1A1A2E !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.5rem 1.5rem !important;
}
.stButton > button:hover { background: #D49200 !important; }

/* Cards de métricas */
.metric-card {
    background: #F8F7F4;
    border: 0.5px solid #DDDDCE;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    text-align: center;
}
.metric-card .label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 0.05em; }
.metric-card .value { font-size: 28px; font-weight: 700; color: #1A1A2E; }

/* Success / warning boxes */
.info-box {
    background: #E8F5EC;
    border-left: 4px solid #22A06B;
    border-radius: 6px;
    padding: 0.75rem 1rem;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}
.warn-box {
    background: #FFF8E6;
    border-left: 4px solid #F0A500;
    border-radius: 6px;
    padding: 0.75rem 1rem;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}
.err-box {
    background: #FDECEC;
    border-left: 4px solid #E24B4A;
    border-radius: 6px;
    padding: 0.75rem 1rem;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}

/* Tabelas */
.dataframe thead th { background: #1A1A2E !important; color: white !important; }
.dataframe tbody tr:nth-child(even) { background: #F8F7F4; }

/* Logo text */
.logo-text {
    font-family: 'DM Sans', sans-serif;
    font-size: 1.2rem;
    font-weight: 700;
    color: #F0A500;
    letter-spacing: -0.02em;
}
.logo-sub {
    font-size: 0.7rem;
    color: #8888AA;
    margin-top: -4px;
}
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "history" not in st.session_state:
    st.session_state.history = []
if "ai_configs" not in st.session_state:
    st.session_state.ai_configs = {}
if "last_results" not in st.session_state:
    st.session_state.last_results = []


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="logo-text">⚡ PBI Doc Fácil</div>', unsafe_allow_html=True)
    st.markdown('<div class="logo-sub">Documentação técnica para Power BI</div>', unsafe_allow_html=True)
    st.markdown("---")

    page = st.radio(
        "Navegação",
        ["📄 Gerar Doc", "📊 Análises", "🤖 Provedores de IA", "⚙️ Configurações", "🏠 Início"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown(
        '<div style="font-size:0.75rem; color:#555577; text-align:center;">v1.0.0 · 2026<br>'
        '<a href="https://www.powerbidocfacil.com.br" style="color:#F0A500;">powerbidocfacil.com.br</a></div>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# 🏠  INÍCIO
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Início":
    st.markdown("## Dashboard")

    history = st.session_state.history
    total_tables = sum(r.get("tables", 0) for r in history)
    total_cols = sum(r.get("columns", 0) for r in history)
    total_measures = sum(r.get("measures", 0) for r in history)
    total_rels = sum(r.get("relationships", 0) for r in history)

    cols = st.columns(4)
    for col, label, val in zip(
        cols,
        ["Relatórios Gerados", "Tabelas Documentadas", "Medidas DAX", "Relacionamentos"],
        [len(history), total_tables, total_measures, total_rels],
    ):
        col.markdown(
            f'<div class="metric-card"><div class="label">{label}</div><div class="value">{val}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    if history:
        st.markdown("### Execuções recentes")
        for run in reversed(history[-10:]):
            st.markdown(
                f'<div class="info-box">📄 <b>{run["file"]}</b> · '
                f'{run["date"]} · {run["tables"]} tabelas · '
                f'{run["measures"]} medidas · '
                f'{"🤖 IA: " + run["provider"] if run.get("provider") else "Sem IA"}</div>',
                unsafe_allow_html=True,
            )
    else:
        st.info("Nenhuma documentação gerada ainda. Vá em **Gerar Doc** para começar.")


# ══════════════════════════════════════════════════════════════════════════════
# 📄  GERAR DOC
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📄 Gerar Doc":
    st.markdown("## Gerar Documentação")
    st.markdown("Faça upload dos seus arquivos `.pbix` e gere o relatório Word em segundos.")

    # Upload
    uploaded_files = st.file_uploader(
        "Selecione um ou mais arquivos .pbix",
        type=["pbix"],
        accept_multiple_files=True,
        help="Você pode selecionar vários arquivos de uma vez.",
    )

    if not uploaded_files:
        st.markdown(
            '<div class="warn-box">📁 Nenhum arquivo selecionado. Arraste os arquivos .pbix acima.</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown("### Opções de geração")

    col1, col2 = st.columns(2)

    with col1:
        use_ai = st.checkbox("🤖 Usar IA para gerar visão geral (overview)", value=False)
        if use_ai:
            configs = st.session_state.ai_configs
            provider_options = list(configs.keys()) if configs else []
            if provider_options:
                selected_provider = st.selectbox("Provedor de IA", provider_options)
            else:
                st.markdown(
                    '<div class="warn-box">Nenhum provedor configurado. '
                    'Vá em <b>Provedores de IA</b> para adicionar.</div>',
                    unsafe_allow_html=True,
                )
                selected_provider = None
                use_ai = False

    with col2:
        if use_ai:
            context_mode = st.radio(
                "Contexto enviado para IA",
                ["Amostra (econômico)", "Completo (mais detalhado)"],
                help="'Amostra' envia as primeiras 5 tabelas. 'Completo' envia todos os metadados.",
            )
            full_context = "Completo" in context_mode
        logo_path_cfg = st.session_state.get("logo_path", "")

    extra_instructions = st.text_area(
        "Instruções adicionais para IA (opcional)",
        placeholder="Ex: Este relatório é voltado para a área financeira. Enfatize as medidas de DRE e os filtros de período.",
        height=80,
    )

    st.markdown("---")

    if st.button("▶  Gerar Relatório", disabled=not uploaded_files):
        results = []
        progress = st.progress(0, text="Iniciando extração...")
        total = len(uploaded_files)

        for idx, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            progress.progress((idx) / total, text=f"Processando {file_name}...")

            try:
                t0 = time.time()

                # Salva temp
                tmp_path = Path(f"/tmp/{file_name}")
                tmp_path.write_bytes(uploaded_file.getvalue())

                # Extração
                t_extract_start = time.time()
                model = extract_pbix(tmp_path)
                t_extract = time.time() - t_extract_start

                # IA
                ai_overview = ""
                t_ai = 0
                if use_ai and selected_provider:
                    cfg = st.session_state.ai_configs[selected_provider]
                    t_ai_start = time.time()
                    ai_overview = generate_ai_overview(
                        model=model,
                        provider=cfg["type"],
                        api_key=cfg.get("api_key", ""),
                        model_name=cfg.get("model_name", ""),
                        full_context=full_context,
                        extra_instructions=extra_instructions,
                        base_url=cfg.get("base_url", ""),
                    )
                    t_ai = time.time() - t_ai_start

                # Gera DOCX
                t_doc_start = time.time()
                docx_bytes = generate_docx(
                    model=model,
                    ai_overview=ai_overview,
                    logo_path=logo_path_cfg or None,
                )
                t_doc = time.time() - t_doc_start
                t_total = time.time() - t0

                summary = model.summary
                results.append({
                    "file_name": file_name,
                    "model": model,
                    "docx_bytes": docx_bytes,
                    "ai_overview": ai_overview,
                    "t_extract": t_extract,
                    "t_ai": t_ai,
                    "t_doc": t_doc,
                    "t_total": t_total,
                    "summary": summary,
                    "warnings": model.extraction_warnings,
                    "ok": True,
                })

                # Histórico
                st.session_state.history.append({
                    "file": file_name,
                    "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "provider": selected_provider if use_ai else "",
                    **summary,
                })

            except Exception as e:
                results.append({
                    "file_name": file_name,
                    "ok": False,
                    "error": str(e),
                })

        progress.progress(1.0, text="Concluído!")
        st.session_state.last_results = results

        # ── Resultados ─────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### Resultados")

        for res in results:
            if res["ok"]:
                summ = res["summary"]
                ia_part = f" · IA {res['t_ai']:.1f}s" if res["t_ai"] else ""
                st.markdown(
                    f'<div class="info-box">✅ <b>{res["file_name"]}</b> · '
                    f'Total: {res["t_total"]:.1f}s '
                    f'(extração {res["t_extract"]:.1f}s{ia_part}'
                    f' · Word {res["t_doc"]:.1f}s) · '
                    f'{summ["tables"]} tabelas · {summ["measures"]} medidas</div>',
                    unsafe_allow_html=True,
                )

                if res["warnings"]:
                    for w in res["warnings"]:
                        st.markdown(f'<div class="warn-box">⚠️ {w}</div>', unsafe_allow_html=True)

                dl_name = res["file_name"].replace(".pbix", "_documentacao.docx")
                st.download_button(
                    label=f"⬇️  Baixar {dl_name}",
                    data=res["docx_bytes"],
                    file_name=dl_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{res['file_name']}",
                )
            else:
                st.markdown(
                    f'<div class="err-box">❌ <b>{res["file_name"]}</b> — Erro: {res["error"]}</div>',
                    unsafe_allow_html=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# 📊  ANÁLISES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📊 Análises":
    st.markdown("## Análises do Modelo")

    results = st.session_state.last_results
    ok_results = [r for r in results if r.get("ok")]

    if not ok_results:
        st.info("Gere uma documentação primeiro em **Gerar Doc** para ver as análises aqui.")
        st.stop()

    names = [r["file_name"] for r in ok_results]
    selected = st.selectbox("Selecione o relatório", names)
    res = next(r for r in ok_results if r["file_name"] == selected)
    model = res["model"]

    # KPIs
    summ = model.summary
    cols = st.columns(6)
    for col, label, val in zip(
        cols,
        ["Tabelas", "Colunas", "Medidas", "Relacionamentos", "Fontes", "Páginas"],
        [summ["tables"], summ["columns"], summ["measures"], summ["relationships"], summ["data_sources"], summ["report_pages"]],
    ):
        col.markdown(
            f'<div class="metric-card"><div class="label">{label}</div><div class="value">{val}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs(["Tabelas e Colunas", "Medidas DAX", "Relacionamentos", "Fontes"])

    with tab1:
        import pandas as pd
        rows = []
        for t in model.tables:
            for c in t.columns:
                rows.append({
                    "Tabela": t.name,
                    "Coluna": c.name,
                    "Tipo": c.data_type,
                    "Calculada": "Sim" if c.is_calculated else "Não",
                    "Oculta": "Sim" if c.is_hidden else "Não",
                    "Descrição": c.description or "—",
                })
        if rows:
            df = pd.DataFrame(rows)
            search = st.text_input("🔍 Filtrar por tabela ou coluna")
            if search:
                mask = df.apply(lambda row: search.lower() in str(row).lower(), axis=1)
                df = df[mask]
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma coluna extraída. Verifique os avisos na aba de geração.")

    with tab2:
        measures = model.all_measures
        if measures:
            search_m = st.text_input("🔍 Filtrar medidas", key="search_meas")
            for m in measures:
                if search_m and search_m.lower() not in m.name.lower() and search_m.lower() not in m.expression.lower():
                    continue
                with st.expander(f"📐 {m.table} · **{m.name}**"):
                    if m.description:
                        st.markdown(f"*{m.description}*")
                    if m.format_string:
                        st.markdown(f"**Formato:** `{m.format_string}`")
                    if m.expression:
                        st.code(m.expression, language="text")
                    else:
                        st.markdown("*(sem expressão extraída)*")
        else:
            st.info("Nenhuma medida encontrada no modelo.")

    with tab3:
        if model.relationships:
            import pandas as pd
            rel_rows = [
                {
                    "De (Tabela)": r.from_table,
                    "De (Coluna)": r.from_column,
                    "Para (Tabela)": r.to_table,
                    "Para (Coluna)": r.to_column,
                    "Cardinalidade": r.cardinality,
                    "Cross Filter": r.cross_filter,
                    "Ativo": "Sim" if r.is_active else "Não",
                }
                for r in model.relationships
            ]
            st.dataframe(pd.DataFrame(rel_rows), use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum relacionamento encontrado.")

    with tab4:
        if model.data_sources:
            for ds in model.data_sources:
                st.markdown(f"- **{ds.name}** — Tipo: `{ds.kind}`")
        else:
            st.info("Nenhuma fonte de dados identificada.")


# ══════════════════════════════════════════════════════════════════════════════
# 🤖  PROVEDORES DE IA
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🤖 Provedores de IA":
    st.markdown("## Provedores de IA")
    st.markdown(
        "Cadastre suas chaves de API para habilitar a geração de visão geral por IA. "
        "As chaves ficam **apenas nesta sessão** — não são enviadas a nenhum servidor externo além do próprio provedor."
    )

    with st.form("add_provider"):
        st.markdown("### Adicionar / Atualizar provedor")
        col1, col2 = st.columns(2)
        with col1:
            prov_label = st.text_input("Nome do perfil", placeholder="Ex: Claude Sonnet")
            prov_type = st.selectbox(
                "Provedor",
                ["claude", "openai", "gemini", "openrouter", "local"],
            )
            api_key_input = st.text_input("API Key", type="password", placeholder="sk-... / AIza... / etc.")
        with col2:
            model_name_input = st.text_input(
                "Nome do modelo (opcional)",
                placeholder="Ex: claude-sonnet-4-20250514 / gpt-4o-mini",
            )
            base_url_input = st.text_input(
                "URL base (somente para local)",
                placeholder="http://localhost:11434",
            )

        submitted = st.form_submit_button("💾 Salvar provedor")
        if submitted and prov_label:
            st.session_state.ai_configs[prov_label] = {
                "type": prov_type,
                "api_key": api_key_input,
                "model_name": model_name_input,
                "base_url": base_url_input,
            }
            st.success(f"Provedor '{prov_label}' salvo com sucesso!")

    st.markdown("---")
    st.markdown("### Provedores configurados")

    configs = st.session_state.ai_configs
    if configs:
        for name, cfg in list(configs.items()):
            with st.expander(f"🔑 {name}  ({cfg['type']})"):
                st.markdown(f"**Tipo:** `{cfg['type']}`")
                st.markdown(f"**Modelo:** `{cfg.get('model_name') or '(padrão)'}`")
                st.markdown(f"**API Key:** `{'*' * 8 + cfg.get('api_key', '')[-4:] if cfg.get('api_key') else 'não definida'}`")
                if st.button(f"🗑️  Remover '{name}'", key=f"rm_{name}"):
                    del st.session_state.ai_configs[name]
                    st.rerun()
    else:
        st.info("Nenhum provedor cadastrado.")

    st.markdown("---")
    st.markdown("### Modelos sugeridos por provedor")
    st.markdown("""
| Provedor | Modelo recomendado | Notas |
|---|---|---|
| Claude | `claude-sonnet-4-20250514` | Excelente qualidade, custo moderado |
| OpenAI | `gpt-4o-mini` | Rápido e econômico |
| Gemini | `gemini-2.0-flash` | Gratuito até certo limite |
| OpenRouter | `anthropic/claude-3-haiku` | Acesso multi-provedor |
| Local | `llama3` | Privacidade total, sem custo de API |
""")


# ══════════════════════════════════════════════════════════════════════════════
# ⚙️  CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Configurações":
    st.markdown("## Configurações")

    st.markdown("### Logo da organização")
    uploaded_logo = st.file_uploader("Imagem PNG ou JPG (máx. 200x200px recomendado)", type=["png", "jpg", "jpeg"])
    if uploaded_logo:
        tmp_logo = Path(f"/tmp/logo_{uploaded_logo.name}")
        tmp_logo.write_bytes(uploaded_logo.getvalue())
        st.session_state["logo_path"] = str(tmp_logo)
        st.image(str(tmp_logo), width=150)
        st.success("Logo salva e será incluída no cabeçalho do Word.")

    st.markdown("---")
    st.markdown("### Sobre o aplicativo")
    st.markdown("""
**Power BI Doc Fácil** lê arquivos `.pbix` (que são ZIPs), extrai os metadados do modelo semântico
(tabelas, colunas, medidas DAX, relacionamentos e fontes de dados) e gera um documento Word completo.

**Como funciona a extração:**
- Arquivos `.pbix` contêm um arquivo `DataModel` ou `model.bim` com o modelo semântico em JSON/ABF
- O app lê esse arquivo diretamente, sem precisar do Power BI Desktop instalado
- Para melhor extração, salve o relatório como **`.pbit` (Power BI Template)** antes de fazer upload — o `.pbit` sempre contém o JSON completo

**Limitações conhecidas:**
- Arquivos `.pbix` com DataModel em formato binário (ABF puro) podem ter extração parcial
- Modelos muito grandes (>500 tabelas) podem demorar mais para processar
- A IA não altera o modelo, apenas descreve o que foi extraído

**Desenvolvido como projeto open source** inspirado em [powerbidocfacil.com.br](https://www.powerbidocfacil.com.br).
""")

    st.markdown("---")
    if st.button("🗑️  Limpar histórico de sessão"):
        st.session_state.history = []
        st.session_state.last_results = []
        st.success("Histórico limpo.")
