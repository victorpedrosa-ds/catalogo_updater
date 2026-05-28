"""
extractor.py
------------
Le o PDF da portaria PMPF e retorna todas as colunas relevantes.
Detecta o tipo de cada produto diretamente do titulo do anexo no PDF
(ex: 'REFRIGERANTES', 'ENERGETICOS E ISOFONICOS', ou qualquer outro futuro).
"""

import re
import unicodedata
import numpy as np
import pandas as pd
import camelot
from normalizer import normalizar_campos_produto


def extrair_precos(caminho_pdf: str) -> pd.DataFrame:
    print("[Extrator] Lendo PDF com Camelot...")
    tabelas = camelot.read_pdf(caminho_pdf, pages='all', line_scale=40)

    if tabelas.n == 0:
        raise ValueError("Nenhuma tabela encontrada no PDF.")

    tipos_por_pagina = _detectar_tipos_por_pagina(caminho_pdf)

    partes = []
    for t in tabelas:
        df_t = t.df.copy()
        df_t['_TIPO_PORTARIA'] = tipos_por_pagina.get(t.page, '')
        partes.append(df_t)

    df = pd.concat(partes, ignore_index=True)

    df.columns = df.iloc[0].tolist()[:-1] + ['_TIPO_PORTARIA']
    df = df[1:].reset_index(drop=True)
    df.replace('', np.nan, inplace=True)

    novos_cols = []
    for c in df.columns:
        if c == '_TIPO_PORTARIA':
            novos_cols.append(c)
        else:
            novos_cols.append(re.sub(r'\s+', ' ', str(c)).strip())
    df.columns = novos_cols

    col_gtin       = _encontrar_coluna(df, ['GTIN / EAN', 'GTIN/EAN', 'GTIN'])
    col_preco      = _encontrar_coluna(df, ['PRECO', 'PRECO (R$)'])
    col_vigencia   = _encontrar_coluna(df, ['EFEITOS A PARTIR DE', 'EFEITOS A PARTIR DE:'])
    col_nome       = _encontrar_coluna(df, [
        'MARCA / DESCRICAO', 'MARCA/DESCRICAO', 'DESCRICAO', 'PRODUTO', 'NOME',
    ])
    col_fabricante = _encontrar_coluna(df, ['FABRICANTE'])
    col_embalagem  = _encontrar_coluna(df, ['EMBALAGEM'])
    col_material   = _encontrar_coluna(df, ['MATERIAL'])
    col_volume     = _encontrar_coluna(df, ['VOLUME (ML)', 'VOLUME', 'ML'])
    col_ret_desc   = _encontrar_coluna(df, [
        'RETORNAVEL / DESCARTAVEL', 'RETORNAVEL/DESCARTAVEL', 'DESCARTAVEL',
    ])

    if not col_gtin or not col_preco:
        raise ValueError(
            f"Colunas GTIN ou PRECO nao encontradas. "
            f"Colunas disponiveis: {list(df.columns)}"
        )

    df_out = pd.DataFrame({
        'GTIN':          df[col_gtin],
        'PRECO':         df[col_preco],
        'VIGENCIA':      df[col_vigencia]   if col_vigencia   else None,
        'NOME':          df[col_nome]       if col_nome       else None,
        'FABRICANTE':    df[col_fabricante] if col_fabricante else None,
        'EMBALAGEM':     df[col_embalagem]  if col_embalagem  else None,
        'MATERIAL':      df[col_material]   if col_material   else None,
        'VOLUME':        df[col_volume]     if col_volume     else None,
        'RET_DESC':      df[col_ret_desc]   if col_ret_desc   else None,
        'TIPO_PORTARIA': df['_TIPO_PORTARIA'],
    })

    df_out['GTIN'] = df_out['GTIN'].apply(
        lambda v: re.sub(r'\D', '', str(v)) if pd.notna(v) else ''
    )
    df_out = df_out[df_out['GTIN'].str.len() >= 8].copy()
    df_out['PRECO'] = df_out['PRECO'].apply(_limpar_preco)
    df_out = df_out[df_out['PRECO'].notna()].copy()
    df_out = df_out[df_out['GTIN'].str.isnumeric()].reset_index(drop=True)
    df_out = df_out.drop_duplicates(subset='GTIN', keep='first').reset_index(drop=True)

    for col in ['NOME', 'FABRICANTE', 'EMBALAGEM', 'MATERIAL', 'VOLUME', 'RET_DESC']:
        if col in df_out.columns:
            df_out[col] = df_out[col].apply(
                lambda v: str(v).strip()
                if pd.notna(v) and str(v).strip() not in ('', 'nan') else ''
            )

    # Normaliza MATERIAL, EMBALAGEM e RET/DESC para valores canônicos.
    # TIPO_PORTARIA é mantido com o valor exato do anexo (ex: 'REFRIGERANTES',
    # 'ENERGÉTICOS E ISOTÔNICOS', 'REFRIGERANTES/ISOTÔNICOS' em portarias futuras).
    df_out = df_out.apply(
        lambda row: pd.Series(normalizar_campos_produto(row.to_dict())),
        axis=1
    )

    tipos = df_out['TIPO_PORTARIA'].value_counts().to_dict()
    print(f"[Extrator] {len(df_out)} produtos extraidos. Tipos do PDF: {tipos}")
    print(f"[Extrator] Colunas: NOME={'sim' if col_nome else 'nao'} | "
          f"RET/DESC={'sim' if col_ret_desc else 'NAO ENCONTRADO'}")
    return df_out


# ── Deteccao de tipo por pagina ───────────────────────────────────────────────

def _detectar_tipos_por_pagina(caminho_pdf: str) -> dict:
    """
    Percorre o PDF e detecta o titulo da secao de cada pagina a partir
    do texto que aparece apos 'ANEXO I', 'ANEXO II', etc.
    Retorna {num_pagina: titulo_da_secao}.
    Funciona para qualquer titulo presente no PDF, sem mapeamentos fixos.
    """
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTTextContainer
    except ImportError:
        print("[Extrator] pdfminer nao disponivel — TIPO_PORTARIA ficara em branco.")
        return {}

    tipos      = {}
    tipo_atual = None

    for num_pag, layout in enumerate(extract_pages(caminho_pdf), start=1):
        texto = ''.join(
            elem.get_text()
            for elem in layout
            if isinstance(elem, LTTextContainer)
        )

        novo_tipo = _extrair_titulo_secao(texto)
        if novo_tipo:
            tipo_atual = novo_tipo

        if tipo_atual:
            tipos[num_pag] = tipo_atual

    return tipos


def _extrair_titulo_secao(texto: str) -> str | None:
    """
    Procura por 'ANEXO I', 'ANEXO II', etc. no texto da pagina e retorna
    o titulo da secao que aparece logo apos (ex: 'REFRIGERANTES',
    'ENERGETICOS E ISOFONICOS', ou qualquer outro titulo futuro).

    Estrategia: pega a primeira linha nao-vazia apos a linha 'ANEXO N'
    que nao seja texto juridico/prefacio (ex: 'Redacao dada...').
    """
    linhas = texto.split('\n')
    for i, linha in enumerate(linhas):
        # Detecta linha que e APENAS "ANEXO I", "ANEXO II", etc.
        if re.match(r'^\s*ANEXO\s+[IVX]+\s*$', linha, re.IGNORECASE):
            # Procura a proxima linha relevante apos o cabecalho do anexo
            for j in range(i + 1, min(i + 15, len(linhas))):
                candidato = linhas[j].strip()
                if not candidato:
                    continue
                # Ignora linhas que sao texto juridico/descricao (tem virgula, ponto final,
                # palavras como "Redacao", "efeitos", "portaria", "art.")
                norm = _normalizar(candidato)
                if re.search(r'\b(REDACAO|EFEITOS|PORTARIA|DECRETO|ART|REVOGAD|VIGENCIA)\b', norm):
                    continue
                # Aceita como titulo: linha com pelo menos 4 caracteres
                if len(candidato) >= 4:
                    return candidato.upper()
    return None


def _normalizar(texto: str) -> str:
    """Remove acentos e converte para maiusculo para comparacao robusta."""
    nfkd = unicodedata.normalize('NFKD', texto)
    sem_acento = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.upper()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _encontrar_coluna(df: pd.DataFrame, candidatos: list) -> str | None:
    for nome in candidatos:
        nome_n = _normalizar(nome)
        for col in df.columns:
            if col == '_TIPO_PORTARIA':
                continue
            if _normalizar(str(col)) == nome_n:
                return col
    return None


def _limpar_preco(valor) -> float | None:
    if pd.isna(valor):
        return None
    texto = str(valor).strip().replace('R$', '').replace(' ', '')
    try:
        return float(texto)
    except ValueError:
        pass
    texto = texto.replace('.', '').replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None
