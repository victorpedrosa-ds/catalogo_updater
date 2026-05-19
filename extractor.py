"""
extractor.py
------------
Lê o PDF da portaria PMPF e retorna apenas GTIN, preço e vigência.
Usa Camelot com line_scale=40.
"""

import re
import numpy as np
import pandas as pd
import camelot


def extrair_precos(caminho_pdf: str) -> pd.DataFrame:
    """
    Extrai do PDF apenas as colunas relevantes para atualização de preço:
      GTIN, PRECO, VIGENCIA

    Retorna DataFrame com essas 3 colunas, uma linha por produto.
    """
    print("[Extrator] Lendo PDF com Camelot...")
    tabelas = camelot.read_pdf(caminho_pdf, pages='all', line_scale=40)

    if tabelas.n == 0:
        raise ValueError("Nenhuma tabela encontrada no PDF.")

    df = pd.concat([t.df for t in tabelas], ignore_index=True)

    # Primeira linha como cabeçalho
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df.replace('', np.nan, inplace=True)
    df.columns = df.columns.str.replace(r'\s+', ' ', regex=True).str.strip()

    # Identifica colunas pelo nome (flexível)
    col_gtin    = _encontrar_coluna(df, ['GTIN / EAN', 'GTIN/EAN', 'GTIN'])
    col_preco   = _encontrar_coluna(df, ['PREÇO', 'PRECO', 'PREÇO (R$)'])
    col_vigencia = _encontrar_coluna(df, ['EFEITOS A PARTIR DE', 'EFEITOS A PARTIR DE:'])

    if not col_gtin or not col_preco:
        raise ValueError(
            f"Colunas GTIN ou PREÇO não encontradas. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    df_out = pd.DataFrame({
        'GTIN':     df[col_gtin],
        'PRECO':    df[col_preco],
        'VIGENCIA': df[col_vigencia] if col_vigencia else None,
    })

    # Limpa GTIN: só dígitos, mínimo 8
    df_out['GTIN'] = df_out['GTIN'].apply(
        lambda v: re.sub(r'\D', '', str(v)) if pd.notna(v) else ''
    )
    df_out = df_out[df_out['GTIN'].str.len() >= 8].copy()

    # Limpa preço: converte para float
    df_out['PRECO'] = df_out['PRECO'].apply(_limpar_preco)
    df_out = df_out[df_out['PRECO'].notna()].copy()

    # Remove linhas de cabeçalho repetidas (onde GTIN não é numérico)
    df_out = df_out[df_out['GTIN'].str.isnumeric()].reset_index(drop=True)

    # Remove duplicatas: mantém primeiro registro por GTIN
    df_out = df_out.drop_duplicates(subset='GTIN', keep='first').reset_index(drop=True)

    print(f"[Extrator] {len(df_out)} produtos extraídos do PDF.")
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
    # Tenta direto — PDF via Camelot retorna "6.09" com ponto decimal
    try:
        return float(texto)
    except ValueError:
        pass
    # Fallback para formato brasileiro: "6,09" ou "1.234,56"
    texto = texto.replace('.', '').replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None