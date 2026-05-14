"""
applier.py
----------
Insere novas linhas de preço na aba PRECO-VIGENCIA do catálogo.
O catálogo mantém histórico — nunca substitui, sempre insere.
"""

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ABA_PRECO     = 'PRECO-VIGENCIA'
ABA_RELATORIO = 'Relatório'


def aplicar_mudancas(
    caminho_catalogo: str,
    mudancas_aprovadas: list[dict],
) -> str:
    """
    Para cada mudança aprovada, insere uma nova linha na aba PRECO-VIGENCIA.
    Salva como novo arquivo com sufixo de data. Retorna o caminho do arquivo salvo.
    """
    if not mudancas_aprovadas:
        raise ValueError("Nenhuma mudança aprovada para aplicar.")

    # Copia o arquivo original
    caminho = Path(caminho_catalogo)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_saida = caminho.parent / f"{caminho.stem}_atualizado_{ts}{caminho.suffix}"
    shutil.copy2(str(caminho), str(caminho_saida))

    wb = load_workbook(str(caminho_saida))
    ws = wb[ABA_PRECO]

    # Descobre o próximo ID (coluna A) e a última linha com dados
    ultima_linha = ws.max_row
    while ultima_linha > 1 and ws.cell(ultima_linha, 1).value is None:
        ultima_linha -= 1

    ultimo_id = ws.cell(ultima_linha, 1).value
    try:
        proximo_id = int(ultimo_id) + 1
    except (TypeError, ValueError):
        proximo_id = ultima_linha  # fallback

    # Insere uma linha por mudança aprovada
    for mudanca in mudancas_aprovadas:
        nova_linha = ultima_linha + 1
        ultima_linha = nova_linha

        id_produto = mudanca['id_produto']
        vigencia   = mudanca['vigencia']   # string dd/mm/aaaa
        preco_novo = mudanca['preco_novo']

        # Converte vigência para datetime se possível
        try:
            from datetime import datetime as dt
            vig_dt = dt.strptime(vigencia, '%d/%m/%Y')
        except Exception:
            vig_dt = vigencia  # mantém como string se não conseguir converter

        # Colunas da aba PRECO-VIGENCIA:
        # A: ID | B: ID PRODUTO | C: Vigência a partir de | D: Data inclusão obs
        # E: Preço portaria | F: NOME PRODUTO (fórmula) | G: concat (fórmula) | H: cont (fórmula)
        ws.cell(nova_linha, 1).value = proximo_id
        ws.cell(nova_linha, 2).value = id_produto
        ws.cell(nova_linha, 3).value = vig_dt
        ws.cell(nova_linha, 4).value = None
        ws.cell(nova_linha, 5).value = preco_novo

        # Replica as fórmulas das colunas F, G, H da linha anterior
        for col in [6, 7, 8]:
            formula_ref = ws.cell(nova_linha - 1, col).value
            if formula_ref and str(formula_ref).startswith('='):
                nova_formula = _ajustar_formula(str(formula_ref), nova_linha - 1, nova_linha)
                ws.cell(nova_linha, col).value = nova_formula

        proximo_id += 1
        print(f"[Applier] GTIN {mudanca['gtin']} — {mudanca['nome']}: "
              f"R$ {mudanca['preco_atual']:.2f} → R$ {mudanca['preco_novo']:.2f}")

    # Gera aba Relatório
    _gerar_relatorio(wb, mudancas_aprovadas)

    wb.save(str(caminho_saida))
    print(f"[Applier] Arquivo salvo: {caminho_saida}")
    return str(caminho_saida)


def _ajustar_formula(formula: str, linha_origem: int, linha_destino: int) -> str:
    """Substitui referências de linha na fórmula (ex: B2 → B3)."""
    import re
    def substituir(m):
        col_letra = m.group(1)
        num_linha = int(m.group(2))
        if num_linha == linha_origem:
            return f"{col_letra}{linha_destino}"
        return m.group(0)
    return re.sub(r'([A-Z]+)(\d+)', substituir, formula)


def _gerar_relatorio(wb, mudancas_aprovadas: list[dict]):
    """Cria ou substitui a aba Relatório com as mudanças aplicadas."""
    if ABA_RELATORIO in wb.sheetnames:
        del wb[ABA_RELATORIO]
    ws = wb.create_sheet(ABA_RELATORIO)

    ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    # Título
    ws.merge_cells('A1:G1')
    cel = ws['A1']
    cel.value = f'Relatório de Atualização de Preços PMPF — {ts}'
    cel.font = Font(bold=True, size=12)
    cel.alignment = Alignment(horizontal='center')

    # Cabeçalho
    cabecalhos = ['GTIN', 'ID PRODUTO', 'PRODUTO', 'VIGÊNCIA', 'PREÇO ANTERIOR (R$)', 'PREÇO NOVO (R$)', 'VARIAÇÃO (R$)']
    for col_idx, cab in enumerate(cabecalhos, start=1):
        cel = ws.cell(row=3, column=col_idx, value=cab)
        cel.font = Font(bold=True, color='FFFFFF')
        cel.fill = PatternFill('solid', fgColor='2F4F4F')
        cel.alignment = Alignment(horizontal='center')

    # Dados
    for linha, m in enumerate(mudancas_aprovadas, start=4):
        variacao = m['preco_novo'] - m['preco_atual']
        cor = 'FFC7CE' if variacao > 0 else 'C6EFCE'  # vermelho = aumento, verde = redução

        valores = [
            m['gtin'], m['id_produto'], m['nome'],
            m['vigencia'],
            round(m['preco_atual'], 2),
            round(m['preco_novo'], 2),
            round(variacao, 2),
        ]
        for col_idx, val in enumerate(valores, start=1):
            cel = ws.cell(row=linha, column=col_idx, value=val)
            if col_idx >= 5:
                cel.fill = PatternFill('solid', fgColor=cor)

    # Larguras
    for col_idx, larg in enumerate([18, 12, 40, 14, 22, 20, 16], start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = larg

    # Resumo
    linha_res = len(mudancas_aprovadas) + 5
    ws.cell(linha_res, 1, 'TOTAL DE ATUALIZAÇÕES:').font = Font(bold=True)
    ws.cell(linha_res, 2, len(mudancas_aprovadas))