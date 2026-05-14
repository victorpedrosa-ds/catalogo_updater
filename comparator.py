"""
comparator.py
-------------
Compara preços do novo PDF com o preço MAIS RECENTE do catálogo.

Fluxo:
  PDF (GTIN + preço)
    → aba GTIN  (GTIN → ID_PRODUTO)
    → aba PRECO-VIGENCIA (último preço por ID_PRODUTO)
    → mudanças onde preço novo ≠ último preço cadastrado
"""

import re
import pandas as pd


ABA_GTIN     = 'GTIN'
ABA_PRECO    = 'PRECO-VIGENCIA'
ABA_PRODUTOS = 'PRODUTOS'


def carregar_catalogo(caminho_xlsx: str) -> dict:
    xl = pd.ExcelFile(caminho_xlsx)
    print(f"[Comparador] Abas: {xl.sheet_names}")

    # ── GTIN → ID_PRODUTO ────────────────────────────────────────────────────
    df_gtin = pd.read_excel(caminho_xlsx, sheet_name=ABA_GTIN, dtype=str)
    df_gtin.columns = [str(c).strip() for c in df_gtin.columns]

    col_gtin_val = _achar_col(df_gtin, ['GTIN'])
    col_id_prod  = _achar_col(df_gtin, ['ID PRODUTO'])

    gtin_para_id = {}
    for _, row in df_gtin.iterrows():
        gtin = re.sub(r'\D', '', str(row[col_gtin_val]))
        id_p = str(row[col_id_prod]).strip()
        if len(gtin) >= 8 and id_p.isdigit():
            gtin_para_id[gtin] = int(id_p)

    print(f"[Comparador] {len(gtin_para_id)} GTINs mapeados.")

    # ── Preço mais recente por ID_PRODUTO ─────────────────────────────────────
    # A aba PRECO-VIGENCIA é histórica — linhas mais abaixo = mais recentes.
    # Iterando em ordem e sempre sobrescrevendo, o último valor fica salvo.
    df_preco = pd.read_excel(caminho_xlsx, sheet_name=ABA_PRECO, dtype=str)
    df_preco.columns = [str(c).strip() for c in df_preco.columns]

    col_id   = _achar_col(df_preco, ['ID PRODUTO'])
    col_vig  = _achar_col(df_preco, ['Vigência a partir de', 'VIGENCIA', 'Vigencia'])
    col_prec = _achar_col(df_preco, ['Preço portaria', 'PRECO', 'Preço'])

    if not col_id or not col_prec:
        raise ValueError(f"Colunas necessárias não encontradas em PRECO-VIGENCIA. Colunas: {list(df_preco.columns)}")

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
        # Sempre sobrescreve — garante que fica o ÚLTIMO registro (mais recente)
        preco_atual[id_p] = {'preco': preco, 'vigencia': vig}

    print(f"[Comparador] {len(preco_atual)} preços mais recentes carregados.")

    # ── Nome do produto ───────────────────────────────────────────────────────
    df_prod = pd.read_excel(caminho_xlsx, sheet_name=ABA_PRODUTOS, dtype=str)
    df_prod.columns = [str(c).strip() for c in df_prod.columns]

    col_id_p   = _achar_col(df_prod, ['ID'])
    col_concat = _achar_col(df_prod, ['CONCATENAR GTIN'])
    col_nome   = _achar_col(df_prod, ['MARCA/DESCRIÇÃO CATÁLOGO', 'MARCA/DESCRIÇÃO PORTARIA'])

    nome_produto = {}
    for _, row in df_prod.iterrows():
        id_p = str(row.get(col_id_p, '')).strip()
        if id_p.isdigit():
            nome = str(row.get(col_concat) or row.get(col_nome) or '').strip()
            nome_produto[int(id_p)] = nome

    return {
        'gtin_para_id': gtin_para_id,
        'preco_atual':  preco_atual,
        'nome_produto': nome_produto,
    }


def comparar_precos(df_pdf: pd.DataFrame, catalogo: dict) -> list[dict]:
    gtin_para_id = catalogo['gtin_para_id']
    preco_atual  = catalogo['preco_atual']
    nome_produto = catalogo['nome_produto']

    mudancas        = []
    nao_encontrados = 0
    sem_preco       = 0

    for _, row in df_pdf.iterrows():
        gtin       = str(row['GTIN']).strip()
        preco_novo = row['PRECO']
        vigencia   = str(row.get('VIGENCIA', '') or '').strip()

        if gtin not in gtin_para_id:
            nao_encontrados += 1
            continue

        id_prod = gtin_para_id[gtin]

        if id_prod not in preco_atual:
            sem_preco += 1
            continue

        preco_cat = preco_atual[id_prod]['preco']
        nome      = nome_produto.get(id_prod, f'ID {id_prod}')

        if round(preco_novo, 2) != round(preco_cat, 2):
            mudancas.append({
                'gtin':        gtin,
                'id_produto':  id_prod,
                'nome':        nome,
                'vigencia':    vigencia,
                'preco_atual': preco_cat,
                'preco_novo':  preco_novo,
            })

    print(
        f"[Comparador] {len(mudancas)} mudança(s) de preço | "
        f"{nao_encontrados} GTINs não encontrados no catálogo | "
        f"{sem_preco} sem preço cadastrado."
    )
    return mudancas


def _achar_col(df: pd.DataFrame, candidatos: list) -> str | None:
    for nome in candidatos:
        for col in df.columns:
            if col.strip().upper() == nome.strip().upper():
                return col
    return None


def _limpar_preco(valor) -> float | None:
    if pd.isna(valor):
        return None
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    # Tenta converter direto (Excel salva como "6.09" com ponto decimal)
    try:
        return float(texto)
    except ValueError:
        pass
    # Fallback: formato brasileiro "6,09" ou "1.234,56"
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None