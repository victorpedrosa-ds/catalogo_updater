"""
extractor.py
------------
Lê o PDF da portaria PMPF e retorna todas as colunas relevantes:
  GTIN, PRECO, VIGENCIA, NOME, FABRICANTE, EMBALAGEM, MATERIAL, VOLUME
Usa Camelot com line_scale=40.
"""

import re
import numpy as np
import pandas as pd
import camelot


def extrair_precos(caminho_pdf: str) -> pd.DataFrame:
    print("[Extrator] Lendo PDF com Camelot...")
    tabelas = camelot.read_pdf(caminho_pdf, pages='all', line_scale=40)

    if tabelas.n == 0:
        raise ValueError("Nenhuma tabela encontrada no PDF.")

    df = pd.concat([t.df for t in tabelas], ignore_index=True)

    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df.replace('', np.nan, inplace=True)
    df.columns = df.columns.str.replace(r'\s+', ' ', regex=True).str.strip()

    col_gtin       = _encontrar_coluna(df, ['GTIN / EAN', 'GTIN/EAN', 'GTIN'])
    col_preco      = _encontrar_coluna(df, ['PREÇO', 'PRECO', 'PREÇO (R$)'])
    col_vigencia   = _encontrar_coluna(df, ['EFEITOS A PARTIR DE', 'EFEITOS A PARTIR DE:'])
    col_nome       = _encontrar_coluna(df, [
        'MARCA / DESCRIÇÃO', 'MARCA/DESCRIÇÃO',
        'MARCA / DESCRICAO', 'MARCA/DESCRICAO',
        'DESCRIÇÃO', 'DESCRICAO', 'PRODUTO', 'NOME',
    ])
    col_fabricante = _encontrar_coluna(df, ['FABRICANTE'])
    col_embalagem  = _encontrar_coluna(df, ['EMBALAGEM'])
    col_material   = _encontrar_coluna(df, ['MATERIAL'])
    col_volume     = _encontrar_coluna(df, ['VOLUME (ML)', 'VOLUME', 'ML'])

    if not col_gtin or not col_preco:
        raise ValueError(
            f"Colunas GTIN ou PREÇO não encontradas. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    df_out = pd.DataFrame({
        'GTIN':       df[col_gtin],
        'PRECO':      df[col_preco],
        'VIGENCIA':   df[col_vigencia]   if col_vigencia   else None,
        'NOME':       df[col_nome]       if col_nome       else None,
        'FABRICANTE': df[col_fabricante] if col_fabricante else None,
        'EMBALAGEM':  df[col_embalagem]  if col_embalagem  else None,
        'MATERIAL':   df[col_material]   if col_material   else None,
        'VOLUME':     df[col_volume]     if col_volume     else None,
    })

    df_out['GTIN'] = df_out['GTIN'].apply(
        lambda v: re.sub(r'\D', '', str(v)) if pd.notna(v) else ''
    )
    df_out = df_out[df_out['GTIN'].str.len() >= 8].copy()

    df_out['PRECO'] = df_out['PRECO'].apply(_limpar_preco)
    df_out = df_out[df_out['PRECO'].notna()].copy()

    df_out = df_out[df_out['GTIN'].str.isnumeric()].reset_index(drop=True)
    df_out = df_out.drop_duplicates(subset='GTIN', keep='first').reset_index(drop=True)

    for col in ['NOME', 'FABRICANTE', 'EMBALAGEM', 'MATERIAL', 'VOLUME']:
        if col in df_out.columns:
            df_out[col] = df_out[col].apply(
                lambda v: str(v).strip()
                if pd.notna(v) and str(v).strip() not in ('', 'nan') else ''
            )

    print(f"[Extrator] {len(df_out)} produtos extraídos do PDF.")
    print(f"[Extrator] Colunas: NOME={'sim' if col_nome else 'não'} | "
          f"FABRICANTE={'sim' if col_fabricante else 'não'} | "
          f"VOLUME={'sim' if col_volume else 'não'}")
    return df_out


def _encontrar_coluna(df: pd.DataFrame, candidatos: list) -> str | None:
    for nome in candidatos:
        for col in df.columns:
            if col.strip().upper() == nome.upper():
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
