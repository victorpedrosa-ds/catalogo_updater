"""
main.py — PMPF Updater
Execute com: py -m streamlit run main.py
"""

import streamlit as st
from pathlib import Path

from extractor  import extrair_precos
from comparator import carregar_catalogo, comparar_precos
from applier    import aplicar_todas_mudancas

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

    with st.spinner('Extraindo dados do PDF...'):
        try:
            df_pdf = extrair_precos(str(caminho_pdf))
        except Exception as e:
            st.error(f'Erro na extração do PDF: {e}')
            st.stop()

    with st.spinner('Comparando com o catálogo...'):
        try:
            catalogo  = carregar_catalogo(str(caminho_xlsx))
            resultado = comparar_precos(df_pdf, catalogo)
        except Exception as e:
            st.error(f'Erro ao comparar com o catálogo: {e}')
            st.stop()

    mudancas   = resultado['mudancas']
    novos      = resultado['novos']
    removidos  = resultado['removidos']
    descricoes = resultado['atualizacoes_descricao']

    st.session_state.update({
        'mudancas':            mudancas,
        'novos':               novos,
        'removidos':           removidos,
        'descricoes':          descricoes,
        'caminho_xlsx':        str(caminho_xlsx),
        'processado':          True,
        # Aprovações — preços/novos/descrições: True por padrão; removidos: False
        'aprov_precos':    {i: True  for i in range(len(mudancas))},
        'aprov_novos':     {i: True  for i in range(len(novos))},
        'aprov_removidos': {i: False for i in range(len(removidos))},
        'aprov_descricoes':{i: True  for i in range(len(descricoes))},
        'pag_precos':    0,
        'pag_novos':     0,
        'pag_removidos': 0,
        'pag_descricoes':0,
    })


# ── TELA INICIAL ──────────────────────────────────────────────────────────────
if not st.session_state.get('processado'):
    st.title('📋 PMPF Updater')
    st.markdown("""
    Ferramenta de atualização do catálogo PMPF (SEFA-PA).

    ### Como usar:
    1. Faça upload do **PDF da portaria** (novo) na barra lateral
    2. Faça upload do **catálogo Excel** atual
    3. Clique em **Processar**
    4. Revise e aprove cada tipo de mudança nas abas
    5. Clique em **Aplicar e salvar** para gerar o catálogo atualizado
    """)
    st.stop()


# ── DADOS ─────────────────────────────────────────────────────────────────────
mudancas   = st.session_state['mudancas']
novos      = st.session_state['novos']
removidos  = st.session_state['removidos']
descricoes = st.session_state['descricoes']

n_aprov_precos    = sum(1 for v in st.session_state['aprov_precos'].values()    if v)
n_aprov_novos     = sum(1 for v in st.session_state['aprov_novos'].values()     if v)
n_aprov_removidos = sum(1 for v in st.session_state['aprov_removidos'].values() if v)
n_aprov_descricoes= sum(1 for v in st.session_state['aprov_descricoes'].values()if v)

# ── MÉTRICAS GLOBAIS ──────────────────────────────────────────────────────────
st.title('📋 Resultado do Processamento')

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric('🔄 Mudanças de preço',   len(mudancas))
c2.metric('🆕 Produtos novos',      len(novos))
c3.metric('🗑️ Removidos',           len(removidos))
c4.metric('✏️ Atualizações descrição', len(descricoes))
c5.metric('✅ Total aprovados',
          n_aprov_precos + n_aprov_novos + n_aprov_removidos + n_aprov_descricoes)

st.divider()

# ── ABAS ──────────────────────────────────────────────────────────────────────
aba_precos, aba_novos, aba_removidos, aba_descricoes = st.tabs([
    f'🔄 Mudanças de Preço ({len(mudancas)})',
    f'🆕 Produtos Novos ({len(novos)})',
    f'🗑️ Removidos da Portaria ({len(removidos)})',
    f'✏️ Atualização de Descrição ({len(descricoes)})',
])


# ════════════════════════════════════════════════════════════════════════════
# FUNÇÃO GENÉRICA DE PAGINAÇÃO + CARDS
# ════════════════════════════════════════════════════════════════════════════

def _renderizar_aba(items, chave_aprov, chave_pag, fn_card, default_aprov=True):
    if not items:
        return

    n          = len(items)
    aprovacoes = st.session_state[chave_aprov]

    # Botões em bloco
    col_a, col_b, _ = st.columns([1, 1, 4])
    if col_a.button('✅ Aprovar tudo', key=f'btn_aprov_{chave_pag}'):
        for i in range(n):
            st.session_state[chave_aprov][i] = True
        st.rerun()
    if col_b.button('❌ Rejeitar tudo', key=f'btn_rejeit_{chave_pag}'):
        for i in range(n):
            st.session_state[chave_aprov][i] = False
        st.rerun()

    aprovados_n = sum(1 for v in aprovacoes.values() if v)
    st.caption(f'**{aprovados_n}** de **{n}** aprovados.')

    # Paginação
    total_pags = max(1, (n - 1) // ITENS_POR_PAGINA + 1)
    pag = st.session_state.get(chave_pag, 0)

    col_prev, col_info, col_next = st.columns([1, 2, 1])
    if col_prev.button('◀ Anterior', key=f'prev_{chave_pag}', disabled=(pag == 0)):
        st.session_state[chave_pag] = pag - 1
        st.rerun()
    col_info.markdown(
        f"<div style='text-align:center;padding-top:8px'>"
        f"Página {pag+1} de {total_pags}</div>",
        unsafe_allow_html=True,
    )
    if col_next.button('Próxima ▶', key=f'next_{chave_pag}',
                        disabled=(pag >= total_pags - 1)):
        st.session_state[chave_pag] = pag + 1
        st.rerun()

    st.divider()

    inicio = pag * ITENS_POR_PAGINA
    fim    = min(inicio + ITENS_POR_PAGINA, n)

    for i in range(inicio, fim):
        item  = items[i]
        aprov = aprovacoes.get(i, default_aprov)

        col_card, col_chk = st.columns([5, 1])
        with col_card:
            fn_card(item)
        with col_chk:
            novo_val = st.checkbox('Aprovar', value=aprov, key=f'{chave_pag}_chk_{i}')
            if novo_val != aprov:
                st.session_state[chave_aprov][i] = novo_val

        st.divider()


# ════════════════════════════════════════════════════════════════════════════
# ABA 1 — MUDANÇAS DE PREÇO
# ════════════════════════════════════════════════════════════════════════════
with aba_precos:
    if not mudancas:
        st.success('✅ Nenhuma mudança de preço detectada.')
    else:
        def card_preco(m):
            delta = m['preco_novo'] - m['preco_atual']
            sinal = '📈' if delta > 0 else '📉'
            texto = (
                f"{sinal} **{m['nome']}** | GTIN: `{m['gtin']}`\n\n"
                f"R$ {m['preco_atual']:.2f} → **R$ {m['preco_novo']:.2f}** "
                f"({'%+.2f' % delta}) | Vigência: {m['vigencia']}"
            )
            if delta > 0:
                st.warning(texto)
            else:
                st.success(texto)

        _renderizar_aba(mudancas, 'aprov_precos', 'pag_precos', card_preco)


# ════════════════════════════════════════════════════════════════════════════
# ABA 2 — PRODUTOS NOVOS
# ════════════════════════════════════════════════════════════════════════════
with aba_novos:
    if not novos:
        st.success('✅ Nenhum produto novo detectado.')
    else:
        st.info(
            f'**{len(novos)}** GTIN(s) presentes na portaria mas ausentes no catálogo. '
            f'Os aprovados serão adicionados às abas **GTIN**, **PRODUTOS** e **PRECO-VIGENCIA**.'
        )

        def card_novo(p):
            sug  = p.get('sugestao')
            vol  = p.get('volume', '')
            emb  = p.get('embalagem', '')
            fab  = p.get('fabricante', '')
            nome = p.get('nome_pdf') or '_(nome não disponível)_'

            linhas = [
                f"🆕 **{nome}** | GTIN: `{p['gtin']}`",
                f"Fabricante: {fab or '—'} | Embalagem: {emb or '—'} | Volume: {vol or '—'} ml",
                f"Preço: **R$ {p['preco']:.2f}** | Vigência: {p.get('vigencia') or '—'}",
            ]
            if sug:
                cor = '🟢' if sug['score'] >= 85 else '🟡'
                linhas.append(
                    f"{cor} Produto similar no catálogo: *{sug['nome']}* "
                    f"(ID {sug['id_produto']}, {sug['score']}%)"
                )
            st.info('\n\n'.join(linhas))

        _renderizar_aba(novos, 'aprov_novos', 'pag_novos', card_novo)


# ════════════════════════════════════════════════════════════════════════════
# ABA 3 — REMOVIDOS DA PORTARIA
# ════════════════════════════════════════════════════════════════════════════
with aba_removidos:
    if not removidos:
        st.success('✅ Todos os produtos do catálogo estão presentes na portaria.')
    else:
        st.warning(
            f'**{len(removidos)}** produto(s) do catálogo **não aparecem** na portaria atual.\n\n'
            f'Os aprovados serão **marcados** (em vermelho) na aba GTIN do catálogo. '
            f'⚠️ O padrão é **não aprovar** — selecione apenas os que realmente foram removidos da pauta.'
        )

        def card_removido(r):
            preco_str = f"R$ {r['ultimo_preco']:.2f}" if r['ultimo_preco'] else '—'
            st.warning(
                f"🗑️ **{r['nome']}** | GTIN: `{r['gtin']}`\n\n"
                f"ID Produto: `{r['id_produto']}` | "
                f"Último preço: **{preco_str}** | Última vigência: {r.get('ultima_vigencia') or '—'}"
            )

        _renderizar_aba(removidos, 'aprov_removidos', 'pag_removidos',
                        card_removido, default_aprov=False)


# ════════════════════════════════════════════════════════════════════════════
# ABA 4 — ATUALIZAÇÃO DE DESCRIÇÃO
# ════════════════════════════════════════════════════════════════════════════
with aba_descricoes:
    if not descricoes:
        st.success('✅ Nenhuma atualização de descrição detectada.')
    else:
        st.info(
            f'**{len(descricoes)}** produto(s) com descrição diferente entre o catálogo '
            f'e a portaria atual. Os aprovados terão a coluna '
            f'**MARCA/DESCRIÇÃO** atualizada na aba PRODUTOS.'
        )

        def card_descricao(d):
            score = d.get('similaridade', 0)
            cor   = '🟡' if score >= 70 else '🔴'
            st.info(
                f"✏️ **{d.get('nome_exibicao', d['gtin'])}** | "
                f"GTIN: `{d['gtin']}` | ID: `{d['id_produto']}`\n\n"
                f"**Catálogo atual:** {d.get('nome_atual') or '—'}\n\n"
                f"**Portaria (novo):** {d.get('nome_novo') or '—'}\n\n"
                f"{cor} Similaridade: **{score:.0f}%**"
            )

        _renderizar_aba(descricoes, 'aprov_descricoes', 'pag_descricoes', card_descricao)


# ════════════════════════════════════════════════════════════════════════════
# PAINEL GLOBAL DE APLICAÇÃO
# ════════════════════════════════════════════════════════════════════════════
st.divider()
st.subheader('💾 Aplicar mudanças aprovadas')

n_p = sum(1 for v in st.session_state['aprov_precos'].values()     if v)
n_n = sum(1 for v in st.session_state['aprov_novos'].values()      if v)
n_r = sum(1 for v in st.session_state['aprov_removidos'].values()  if v)
n_d = sum(1 for v in st.session_state['aprov_descricoes'].values() if v)
total_aprov = n_p + n_n + n_r + n_d

col_info, col_btn = st.columns([4, 2])
col_info.info(
    f"**{total_aprov}** mudança(s) serão aplicadas no catálogo Excel:\n\n"
    f"🔄 **{n_p}** preços · "
    f"🆕 **{n_n}** novos produtos · "
    f"🗑️ **{n_r}** removidos · "
    f"✏️ **{n_d}** descrições"
)

if col_btn.button('💾 Aplicar e salvar catálogo',
                   type='primary', use_container_width=True,
                   disabled=(total_aprov == 0)):

    aprovadas_precos    = [mudancas[i]   for i, v in st.session_state['aprov_precos'].items()     if v]
    aprovados_novos     = [novos[i]      for i, v in st.session_state['aprov_novos'].items()      if v]
    aprovados_removidos = [removidos[i]  for i, v in st.session_state['aprov_removidos'].items()  if v]
    aprovadas_descricoes= [descricoes[i] for i, v in st.session_state['aprov_descricoes'].items() if v]

    with st.spinner('Aplicando mudanças e salvando catálogo...'):
        try:
            caminho_saida = aplicar_todas_mudancas(
                caminho_catalogo     = st.session_state['caminho_xlsx'],
                mudancas_aprovadas   = aprovadas_precos,
                novos_aprovados      = aprovados_novos,
                removidos_aprovados  = aprovados_removidos,
                descricoes_aprovadas = aprovadas_descricoes,
            )
        except Exception as e:
            st.error(f'Erro ao aplicar mudanças: {e}')
            st.stop()

    st.success(
        f'✅ Catálogo atualizado com sucesso! '
        f'{n_p} preços · {n_n} novos · {n_r} removidos · {n_d} descrições.'
    )
    st.code(caminho_saida)

    with open(caminho_saida, 'rb') as f:
        st.download_button(
            label='⬇️ Baixar catálogo atualizado',
            data=f,
            file_name=Path(caminho_saida).name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )
