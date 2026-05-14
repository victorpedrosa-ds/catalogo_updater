"""
main.py — PMPF Updater (Protótipo — Atualização de Preços)
Execute com: py -m streamlit run main.py
"""

import streamlit as st
from pathlib import Path

from extractor import extrair_precos
from comparator import carregar_catalogo, comparar_precos
from applier import aplicar_mudancas

st.set_page_config(page_title='PMPF Updater', page_icon='📋', layout='wide')
DIR_DATA = Path(__file__).parent / 'data'
DIR_DATA.mkdir(exist_ok=True)
ITENS_POR_PAGINA = 50


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title('📋 PMPF Updater')
    st.caption('Atualização de preços — Portaria PMPF')
    st.divider()

    pdf_upload  = st.file_uploader('PDF da Portaria (novo)', type=['pdf'])
    xlsx_upload = st.file_uploader('Catálogo atual (.xlsx)', type=['xlsx'])

    st.divider()
    processar = st.button('🔍 Processar', use_container_width=True, type='primary')


# ── PROCESSAMENTO ─────────────────────────────────────────────────────────────
if processar:
    if not pdf_upload or not xlsx_upload:
        st.error('Faça upload do PDF e do Excel antes de processar.')
        st.stop()

    caminho_pdf  = DIR_DATA / 'portaria_temp.pdf'
    caminho_xlsx = DIR_DATA / 'catalogo_temp.xlsx'
    caminho_pdf.write_bytes(pdf_upload.read())
    caminho_xlsx.write_bytes(xlsx_upload.read())

    with st.spinner('Extraindo preços do PDF...'):
        try:
            df_pdf = extrair_precos(str(caminho_pdf))
        except Exception as e:
            st.error(f'Erro na extração do PDF: {e}')
            st.stop()

    with st.spinner('Comparando com o catálogo...'):
        try:
            catalogo = carregar_catalogo(str(caminho_xlsx))
            mudancas = comparar_precos(df_pdf, catalogo)
        except Exception as e:
            st.error(f'Erro ao comparar com o catálogo: {e}')
            st.stop()

    st.session_state['mudancas']     = mudancas
    st.session_state['caminho_xlsx'] = str(caminho_xlsx)
    st.session_state['processado']   = True
    st.session_state['aprovacoes']   = {i: True for i in range(len(mudancas))}
    st.session_state['pag']          = 0


# ── TELA INICIAL ──────────────────────────────────────────────────────────────
if not st.session_state.get('processado'):
    st.title('📋 PMPF Updater')
    st.markdown("""
    Protótipo de atualização de preços da portaria PMPF.

    ### Como usar:
    1. Faça upload do **PDF da portaria** (novo) na barra lateral
    2. Faça upload do **catálogo Excel** atual
    3. Clique em **Processar**
    4. Revise as mudanças de preço detectadas
    5. Aprove ou rejeite cada item
    6. Clique em **Aplicar aprovados** para atualizar o catálogo
    """)
    st.stop()


# ── INTERFACE DE REVISÃO ──────────────────────────────────────────────────────
mudancas   = st.session_state['mudancas']
aprovacoes = st.session_state['aprovacoes']

st.title('📋 Mudanças de Preço Detectadas')

if not mudancas:
    st.success('✅ Nenhuma mudança de preço detectada. O catálogo está atualizado.')
    st.stop()

# Métricas
aumentos  = sum(1 for m in mudancas if m['preco_novo'] > m['preco_atual'])
reducoes  = sum(1 for m in mudancas if m['preco_novo'] < m['preco_atual'])
aprovados = sum(1 for v in aprovacoes.values() if v)

c1, c2, c3, c4 = st.columns(4)
c1.metric('Total de mudanças', len(mudancas))
c2.metric('Aumentos de preço', aumentos, delta=f'+{aumentos}', delta_color='inverse')
c3.metric('Reduções de preço', reducoes, delta=f'-{reducoes}', delta_color='normal')
c4.metric('Aprovadas', aprovados)

st.divider()

# Botões de aprovação em bloco
c1, c2, _ = st.columns([1, 1, 4])
if c1.button('✅ Aprovar tudo'):
    for i in range(len(mudancas)):
        st.session_state['aprovacoes'][i] = True
    st.rerun()
if c2.button('❌ Rejeitar tudo'):
    for i in range(len(mudancas)):
        st.session_state['aprovacoes'][i] = False
    st.rerun()

aprovados_atual = sum(1 for v in st.session_state['aprovacoes'].values() if v)
st.caption(f'**{aprovados_atual}** de **{len(mudancas)}** mudanças aprovadas.')

# Paginação
total_pags = max(1, (len(mudancas) - 1) // ITENS_POR_PAGINA + 1)
pag = st.session_state.get('pag', 0)

col_prev, col_info, col_next = st.columns([1, 2, 1])
if col_prev.button('◀ Anterior', disabled=(pag == 0)):
    st.session_state['pag'] = pag - 1
    st.rerun()
col_info.markdown(
    f"<div style='text-align:center;padding-top:8px'>Página {pag+1} de {total_pags}</div>",
    unsafe_allow_html=True
)
if col_next.button('Próxima ▶', disabled=(pag >= total_pags - 1)):
    st.session_state['pag'] = pag + 1
    st.rerun()

st.divider()

# Lista de mudanças
inicio = pag * ITENS_POR_PAGINA
fim    = min(inicio + ITENS_POR_PAGINA, len(mudancas))

for i in range(inicio, fim):
    m     = mudancas[i]
    aprov = st.session_state['aprovacoes'].get(i, True)
    delta = m['preco_novo'] - m['preco_atual']
    sinal = '📈' if delta > 0 else '📉'

    col_card, col_chk = st.columns([5, 1])
    with col_card:
        if delta > 0:
            st.warning(
                f"{sinal} **{m['nome']}** | GTIN: `{m['gtin']}`\n\n"
                f"R$ {m['preco_atual']:.2f} → **R$ {m['preco_novo']:.2f}** "
                f"(+R$ {delta:.2f}) | Vigência: {m['vigencia']}"
            )
        else:
            st.success(
                f"{sinal} **{m['nome']}** | GTIN: `{m['gtin']}`\n\n"
                f"R$ {m['preco_atual']:.2f} → **R$ {m['preco_novo']:.2f}** "
                f"(R$ {delta:.2f}) | Vigência: {m['vigencia']}"
            )
    with col_chk:
        novo_val = st.checkbox('Aprovar', value=aprov, key=f'chk_{i}')
        if novo_val != aprov:
            st.session_state['aprovacoes'][i] = novo_val

    st.divider()


# ── BOTÃO APLICAR ─────────────────────────────────────────────────────────────
n_aprov = sum(1 for v in st.session_state['aprovacoes'].values() if v)
col_btn, col_info = st.columns([2, 4])
col_info.info(f'**{n_aprov}** mudança(s) serão aplicadas na aba **PRECO-VIGENCIA**.')

if col_btn.button('💾 Aplicar e salvar catálogo', type='primary', use_container_width=True):
    aprovadas = [mudancas[i] for i, v in st.session_state['aprovacoes'].items() if v]

    if not aprovadas:
        st.warning('Nenhuma mudança aprovada.')
    else:
        with st.spinner('Aplicando mudanças e salvando...'):
            caminho_saida = aplicar_mudancas(
                caminho_catalogo=st.session_state['caminho_xlsx'],
                mudancas_aprovadas=aprovadas,
            )

        st.success(f'✅ {len(aprovadas)} preço(s) atualizado(s) com sucesso!')
        st.code(caminho_saida)

        with open(caminho_saida, 'rb') as f:
            st.download_button(
                label='⬇️ Baixar catálogo atualizado',
                data=f,
                file_name=Path(caminho_saida).name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )
