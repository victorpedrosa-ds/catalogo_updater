"""
comparator.py
-------------
Compara PDF com catálogo e detecta 4 tipos de mudança:
  - mudancas           : preço alterado
  - novos              : GTIN do PDF ausente no catálogo
  - removidos          : GTIN do catálogo ausente na portaria
  - atualizacoes_descricao : mesma GTIN com descrição diferente
"""

import re
import pandas as pd

ABA_GTIN     = 'GTIN'
ABA_PRECO    = 'PRECO-VIGENCIA'
ABA_PRODUTOS = 'PRODUTOS'

FUZZY_THRESHOLD       = 70   # limiar para sugestão de produto similar (novos)
DESCRICAO_THRESHOLD   = 90   # abaixo disso → flag de atualização de descrição


def carregar_catalogo(caminho_xlsx: str) -> dict:
    xl = pd.ExcelFile(caminho_xlsx)
    print(f"[Comparador] Abas: {xl.sheet_names}")

    # ── GTIN → ID_PRODUTO ────────────────────────────────────────────────────
    df_gtin = pd.read_excel(caminho_xlsx, sheet_name=ABA_GTIN, dtype=str)
    df_gtin.columns = [str(c).strip() for c in df_gtin.columns]

    col_gtin_val = _achar_col(df_gtin, ['GTIN'])
    col_id_prod  = _achar_col(df_gtin, ['ID PRODUTO'])
    col_valido   = _achar_col(df_gtin, ['VÁLIDO', 'VALIDO'])

    gtin_para_id   = {}   # {gtin: id_produto} — apenas GTINs com VÁLIDO=True
    gtins_invalidos = set()  # GTINs já marcados como FALSO no catálogo

    for _, row in df_gtin.iterrows():
        gtin = re.sub(r'\D', '', str(row[col_gtin_val]))
        id_p = str(row[col_id_prod]).strip()
        if len(gtin) < 8 or not id_p.isdigit():
            continue

        # Verifica se o GTIN está ativo (VÁLIDO = True/VERDADEIRO)
        valido = True
        if col_valido:
            val = str(row.get(col_valido, '')).strip().upper()
            if val in ('FALSE', 'FALSO', '0', 'NÃO', 'NAO', 'NO'):
                valido = False

        if valido:
            gtin_para_id[gtin] = int(id_p)
        else:
            gtins_invalidos.add(gtin)

    print(f"[Comparador] {len(gtin_para_id)} GTINs ativos mapeados "
          f"({len(gtins_invalidos)} já marcados como inválidos ignorados).")

    # ── Preço mais recente por ID_PRODUTO ─────────────────────────────────────
    df_preco = pd.read_excel(caminho_xlsx, sheet_name=ABA_PRECO, dtype=str)
    df_preco.columns = [str(c).strip() for c in df_preco.columns]

    col_id   = _achar_col(df_preco, ['ID PRODUTO'])
    col_vig  = _achar_col(df_preco, ['Vigência a partir de', 'VIGENCIA', 'Vigencia'])
    col_prec = _achar_col(df_preco, ['Preço portaria', 'PRECO', 'Preço'])

    if not col_id or not col_prec:
        raise ValueError(
            f"Colunas necessárias não encontradas em PRECO-VIGENCIA. "
            f"Colunas: {list(df_preco.columns)}"
        )

    preco_atual = {}
    for _, row in df_preco.iterrows():
        id_p_str = str(row[col_id]).strip()
        if not id_p_str.isdigit():
            continue
        preco = _limpar_preco(row[col_prec])
        if preco is None:
            continue
        id_p = int(id_p_str)
        vig  = str(row[col_vig]).strip() if col_vig else ''
        preco_atual[id_p] = {'preco': preco, 'vigencia': vig}

    print(f"[Comparador] {len(preco_atual)} preços mais recentes carregados.")

    # ── Nome e descrição dos produtos ─────────────────────────────────────────
    df_prod = pd.read_excel(caminho_xlsx, sheet_name=ABA_PRODUTOS, dtype=str)
    df_prod.columns = [str(c).strip() for c in df_prod.columns]

    col_id_p   = _achar_col(df_prod, ['ID'])
    col_concat = _achar_col(df_prod, ['CONCATENAR GTIN'])
    # Para comparação usamos PORTARIA (o que estava na última portaria processada)
    col_desc_portaria = _achar_col(df_prod, ['MARCA/DESCRIÇÃO PORTARIA', 'MARCA/DESCRICAO PORTARIA'])
    # Para exibição usamos CATÁLOGO (o nome escolhido pelo usuário)
    col_desc_catalogo = _achar_col(df_prod, ['MARCA/DESCRIÇÃO CATÁLOGO', 'MARCA/DESCRICAO CATALOGO'])

    nome_produto      = {}   # nome para exibição (CONCATENAR GTIN ou CATÁLOGO)
    descricao_catalog = {}   # descrição PORTARIA anterior (para comparar com novo PDF)

    for _, row in df_prod.iterrows():
        id_p = str(row.get(col_id_p, '')).strip()
        if id_p.isdigit():
            id_int = int(id_p)
            nome = str(row.get(col_concat) or row.get(col_desc_catalogo) or '').strip()
            nome_produto[id_int] = nome
            # Usa a descrição portaria (se existir) para comparar com o PDF atual
            desc_port = str(row.get(col_desc_portaria) or '').strip() if col_desc_portaria else ''
            descricao_catalog[id_int] = desc_port

    return {
        'gtin_para_id':      gtin_para_id,
        'gtins_invalidos':   gtins_invalidos,
        'preco_atual':       preco_atual,
        'nome_produto':      nome_produto,
        'descricao_catalog': descricao_catalog,
    }


def comparar_precos(df_pdf: pd.DataFrame, catalogo: dict) -> dict:
    """
    Retorna dict com 4 listas:
      mudancas              — preço alterado (para aprovação)
      novos                 — GTIN do PDF ausente no catálogo
      removidos             — GTIN do catálogo ausente na portaria
      atualizacoes_descricao — mesmo GTIN com descrição diferente
    """
    gtin_para_id      = catalogo['gtin_para_id']
    preco_atual       = catalogo['preco_atual']
    nome_produto      = catalogo['nome_produto']
    descricao_catalog = catalogo.get('descricao_catalog', {})

    mudancas              = []
    novos                 = []
    atualizacoes_descricao = []
    gtins_pdf             = set()
    sem_preco             = 0

    for _, row in df_pdf.iterrows():
        gtin       = str(row['GTIN']).strip()
        preco_novo = row['PRECO']
        vigencia       = str(row.get('VIGENCIA',      '') or '').strip()
        nome_pdf       = str(row.get('NOME',          '') or '').strip()
        fabricante     = str(row.get('FABRICANTE',    '') or '').strip()
        embalagem      = str(row.get('EMBALAGEM',     '') or '').strip()
        material       = str(row.get('MATERIAL',      '') or '').strip()
        volume         = str(row.get('VOLUME',        '') or '').strip()
        ret_desc       = str(row.get('RET_DESC',      '') or '').strip()
        tipo_portaria  = str(row.get('TIPO_PORTARIA', '') or '').strip()

        gtins_pdf.add(gtin)

        if gtin not in gtin_para_id:
            sugestao = _sugerir_produto_similar(nome_pdf, nome_produto)
            novos.append({
                'gtin':          gtin,
                'nome_pdf':      nome_pdf,
                'fabricante':    fabricante,
                'embalagem':     embalagem,
                'material':      material,
                'volume':        volume,
                'ret_desc':      ret_desc,
                'tipo_portaria': tipo_portaria,
                'preco':         preco_novo,
                'vigencia':      vigencia,
                'sugestao':      sugestao,
            })
            continue

        id_prod = gtin_para_id[gtin]
        nome_exibicao = nome_produto.get(id_prod, f'ID {id_prod}')

        # ── Detecta atualização de descrição ─────────────────────────────────
        desc_cat = descricao_catalog.get(id_prod, '').strip()
        if nome_pdf and desc_cat:
            score = _similaridade(nome_pdf, desc_cat)
            if score < DESCRICAO_THRESHOLD:
                atualizacoes_descricao.append({
                    'gtin':        gtin,
                    'id_produto':  id_prod,
                    'nome_atual':  desc_cat,
                    'nome_exibicao': nome_exibicao,
                    'nome_novo':   nome_pdf,
                    'vigencia':    vigencia,
                    'similaridade': score,
                })

        # ── Detecta mudança de preço ─────────────────────────────────────────
        if id_prod not in preco_atual:
            sem_preco += 1
            continue

        preco_cat = preco_atual[id_prod]['preco']

        if round(preco_novo, 2) != round(preco_cat, 2):
            mudancas.append({
                'gtin':        gtin,
                'id_produto':  id_prod,
                'nome':        nome_exibicao,
                'vigencia':    vigencia,
                'preco_atual': preco_cat,
                'preco_novo':  preco_novo,
            })

    # ── Removidos da portaria ─────────────────────────────────────────────────
    # Considera apenas GTINs com VÁLIDO=True no catálogo.
    # GTINs já marcados como FALSO são ignorados — já foram removidos anteriormente.
    removidos = []
    for gtin, id_prod in gtin_para_id.items():
        if gtin not in gtins_pdf:
            nome       = nome_produto.get(id_prod, f'ID {id_prod}')
            info_preco = preco_atual.get(id_prod, {})
            removidos.append({
                'gtin':            gtin,
                'id_produto':      id_prod,
                'nome':            nome,
                'ultimo_preco':    info_preco.get('preco'),
                'ultima_vigencia': info_preco.get('vigencia', ''),
            })

    print(
        f"[Comparador] {len(mudancas)} mudança(s) de preço | "
        f"{len(novos)} produto(s) novo(s) | "
        f"{len(removidos)} produto(s) removido(s) | "
        f"{len(atualizacoes_descricao)} atualização(ões) de descrição | "
        f"{sem_preco} sem preço cadastrado."
    )

    return {
        'mudancas':               mudancas,
        'novos':                  novos,
        'removidos':              removidos,
        'atualizacoes_descricao': atualizacoes_descricao,
    }


# ── Helpers ───────────────────────────────────────────────────────────────────

def _similaridade(a: str, b: str) -> float:
    try:
        from rapidfuzz import fuzz
        return fuzz.token_sort_ratio(a.upper(), b.upper())
    except ImportError:
        return 100.0 if a.upper() == b.upper() else 0.0


def _sugerir_produto_similar(nome_pdf: str, nome_produto: dict) -> dict | None:
    if not nome_pdf:
        return None
    try:
        from rapidfuzz import process, fuzz
    except ImportError:
        return None

    nomes_validos = {id_p: nome for id_p, nome in nome_produto.items() if nome}
    if not nomes_validos:
        return None

    ids       = list(nomes_validos.keys())
    candidatos = [nomes_validos[i] for i in ids]
    resultado  = process.extractOne(nome_pdf, candidatos, scorer=fuzz.token_sort_ratio)

    if resultado is None or resultado[1] < FUZZY_THRESHOLD:
        return None

    return {
        'id_produto': ids[resultado[2]],
        'nome':       resultado[0],
        'score':      resultado[1],
    }


def _achar_col(df: pd.DataFrame, candidatos: list) -> str | None:
    """Busca coluna por nome exato, sem acentos, para máxima compatibilidade."""
    import unicodedata
    def _sem_acento(s: str) -> str:
        nfd = unicodedata.normalize('NFD', str(s))
        return ''.join(c for c in nfd if not unicodedata.combining(c)).strip().upper()
    for nome in candidatos:
        nome_n = _sem_acento(nome)
        for col in df.columns:
            if _sem_acento(col) == nome_n:
                return col
    return None


def _limpar_preco(valor) -> float | None:
    if pd.isna(valor):
        return None
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    try:
        return float(texto)
    except ValueError:
        pass
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None
