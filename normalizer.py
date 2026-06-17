"""
normalizer.py
-------------
Padroniza os valores de texto antes de gravar no catálogo Excel.
Todos os mapeamentos usam o valor CANÔNICO já existente no catálogo como destino.

Uso:
    from normalizer import normalizar_tipo, normalizar_material,
                           normalizar_embalagem, normalizar_ret_desc
"""

import unicodedata
import re


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sem_acento(s: str) -> str:
    """Remove acentos e retorna em maiúsculas para comparação."""
    nfd = unicodedata.normalize('NFD', str(s))
    return ''.join(c for c in nfd if not unicodedata.combining(c)).strip().upper()


def _normalizar_por_mapa(valor: str, mapa: dict) -> str:
    """
    Tenta casar `valor` (sem acento, maiúsculo) com as chaves do mapa.
    Retorna o valor canônico correspondente, ou o original se não encontrar.
    """
    if not valor or not valor.strip():
        return valor
    chave = _sem_acento(valor)
    return mapa.get(chave, valor.strip())


# ── TIPO (aba PRODUTOS col 3) ─────────────────────────────────────────────────
# Canônicos confirmados no catálogo: CERVEJA | ENERGETICO | ISOTONICO | REFRIGERANTE
# O PDF usa plural e variações acentuadas — mapeamos tudo para os valores exatos do catálogo.

# Chaves são os valores JÁ SEM ACENTO (output de _sem_acento),
# valores são os canônicos do catálogo.
_MAPA_TIPO = {
    # Refrigerantes
    'REFRIGERANTE':                    'REFRIGERANTE',
    'REFRIGERANTES':                   'REFRIGERANTE',

    # Energéticos — variações com e sem acento são todas mapeadas
    'ENERGETICO':                      'ENERGETICO',
    'ENERGETICOS':                     'ENERGETICO',
    'ENERGETICO/ISOBONICO':            'ENERGETICO',
    'ENERGETICO E ISOBONICO':          'ENERGETICO',
    'ENERGETICOS E ISOFONICOS':        'ENERGETICO',   # sem acento no PDF
    'ENERGETICOS E ISOTONICOS':        'ENERGETICO',   # sem acento via _sem_acento()
    'ENERGETICOS E ISOFONICOS':        'ENERGETICO',

    # Isotônicos
    'ISOBONICO':                       'ISOTONICO',
    'ISOFONICOS':                      'ISOTONICO',
    'ISOTONICOS':                      'ISOTONICO',
    'ISOFONICOS E ENERGETICOS':        'ISOTONICO',
    'ISOTONICOS E ENERGETICOS':        'ISOTONICO',

    # Cerveja
    'CERVEJA':                         'CERVEJA',
    'CERVEJAS':                        'CERVEJA',
    'CERVEJAS ARTESANAIS':             'CERVEJA',
}

def normalizar_tipo(valor: str) -> str:
    """Normaliza o TIPO para os valores canônicos do catálogo."""
    return _normalizar_por_mapa(valor, _MAPA_TIPO)


# ── TIPO GTIN (aba GTIN col 4) ────────────────────────────────────────────────git status

def normalizar_tipo_gtin(tipo_portaria: str) -> str:
    """
    Constrói o TIPO GTIN a partir do tipo da portaria normalizado.
    Ex: 'REFRIGERANTES' → 'GTIN PORTARIA - REFRIGERANTE'
    """
    tipo_norm = normalizar_tipo(tipo_portaria)
    return f'GTIN PORTARIA - {tipo_norm}'


# ── MATERIAL (aba PRODUTOS col 9) ─────────────────────────────────────────────
# Canônicos: ALUMÍNIO | BARRIL | PET | PLÁSTICO | VIDRO

_MAPA_MATERIAL = {
    'ALUMINIO':   'ALUMÍNIO',
    'ALUMÍNIO':   'ALUMÍNIO',
    'ALUM':       'ALUMÍNIO',
    'PLASTICO':   'PLÁSTICO',
    'PLÁSTICO':   'PLÁSTICO',
    'PLASTIC':    'PLÁSTICO',
    'VIDRO':      'VIDRO',
    'PET':        'PET',
    'BARRIL':     'BARRIL',
    'KEG':        'BARRIL',
}

def normalizar_material(valor: str) -> str:
    return _normalizar_por_mapa(valor, _MAPA_MATERIAL)


# ── EMBALAGEM (aba PRODUTOS col 7) ────────────────────────────────────────────
# Canônicos: GARRAFA | LATA | BARRIL(KEG)
# Multpacks seguem o padrão 'GARRAFA Multpack Xund' / 'LATA Multpack Xund'

_MAPA_EMBALAGEM_BASE = {
    'GARRAFA': 'GARRAFA',
    'LATA':    'LATA',
    'BARRIL':  'BARRIL(KEG)',
    'KEG':     'BARRIL(KEG)',
    'CAN':     'LATA',
    'LATA ALUMÍNIO': 'LATA',
    'LATA ALUMINIO': 'LATA',
}

def normalizar_embalagem(valor: str) -> str:
    """
    Normaliza embalagem preservando informações de Multpack.
    Ex: 'LATA Multpack 6und' → 'LATA Multpack 6und' (mantém)
        'GARRAFA PLÁSTICO' → 'GARRAFA'
    """
    if not valor or not valor.strip():
        return valor

    # Detecta se é Multpack — preserva como está após normalizar a base
    multpack_match = re.search(r'(Multpack\s+\d+\s*und)', valor, re.IGNORECASE)

    chave = _sem_acento(valor)
    # Tenta match exato primeiro
    if chave in _MAPA_EMBALAGEM_BASE:
        base = _MAPA_EMBALAGEM_BASE[chave]
    else:
        # Tenta match parcial pela primeira palavra
        primeira = chave.split()[0] if chave.split() else chave
        base = _MAPA_EMBALAGEM_BASE.get(primeira, valor.strip())

    if multpack_match:
        return f"{base} {multpack_match.group(1)}"
    return base


# ── RETORNÁVEL/DESCARTÁVEL (aba PRODUTOS col 10) ──────────────────────────────
# Canônicos: DESCARTÁVEL | RETORNÁVEL | DESCARTÁVEL/RETORNÁVEL

_MAPA_RET_DESC = {
    'DESCARTAVEL':           'DESCARTÁVEL',
    'DESCARTÁVEL':           'DESCARTÁVEL',
    'RETORNAVEL':            'RETORNÁVEL',
    'RETORNÁVEL':            'RETORNÁVEL',
    'DESCARTAVEL/RETORNAVEL': 'DESCARTÁVEL/RETORNÁVEL',
    'DESCARTÁVEL/RETORNÁVEL': 'DESCARTÁVEL/RETORNÁVEL',
    'RETORNAVEL/DESCARTAVEL': 'DESCARTÁVEL/RETORNÁVEL',
    'RETORNÁVEL/DESCARTÁVEL': 'DESCARTÁVEL/RETORNÁVEL',
    'AMBOS':                 'DESCARTÁVEL/RETORNÁVEL',
}

def normalizar_ret_desc(valor: str) -> str:
    return _normalizar_por_mapa(valor, _MAPA_RET_DESC)


# ── Função unificada ──────────────────────────────────────────────────────────

def normalizar_campos_produto(dados: dict) -> dict:
    """
    Normaliza todos os campos de um produto de uma vez.
    Aceita e retorna um dicionário com as chaves do extractor.
    """
    resultado = dict(dados)
    # TIPO_PORTARIA é preservado como vem do PDF (ex: 'REFRIGERANTES',
    # 'ENERGÉTICOS E ISOTÔNICOS', 'REFRIGERANTES/ISOTÔNICOS'). Não normalizar.
    if 'MATERIAL' in resultado:
        resultado['MATERIAL'] = normalizar_material(resultado['MATERIAL'])
    if 'EMBALAGEM' in resultado:
        resultado['EMBALAGEM'] = normalizar_embalagem(resultado['EMBALAGEM'])
    if 'RET_DESC' in resultado:
        resultado['RET_DESC'] = normalizar_ret_desc(resultado['RET_DESC'])
    return resultado
