"""
Lógica do cruzamento Razão × SPED (Triple Check + CNAE + CST + Pendências).
Port da implementação client-side JS para Python — mantém todos os comportamentos:

  • Extração de NF do Complemento Histórico (com fallback pra coluna 15)
  • Mapa Nome→CNPJ derivado dos SPEDs
  • Enriquecimento Razão Social oficial + 1º/2º CNAE via BrasilAPI/CNPJa
  • Filtro CST 50–67 em A100 e F100 (faixas "Direito a Crédito" e "Crédito Presumido")
  • Triple Check: NF + CNPJ + Período (C100-EFD/A100), CNPJ + Período + Valor (F100)
  • Tolerância de R$ 100 para considerar valor batido
  • Coluna ANÁLISE DO CRUZAMENTO (texto enxuto: bloco encontrado ou valor da diferença)
  • Aba PENDÊNCIAS DE CRUZAMENTO (consolidação por CNPJ, ordenada DESC)
  • Padrão visual EFCT: Exo 2 + paleta navy/lime + sem gridlines
"""
import io
import re
import difflib
import logging
import unicodedata
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import SheetView


log = logging.getLogger(__name__)

# ============================================================
# CONSTANTES
# ============================================================
TOLERANCIA_VALOR = 100.00
CNAE_CONCURRENCY = 5
CNAE_TIMEOUT = 12  # segundos por request HTTP
USER_AGENT = 'Mozilla/5.0 EFCT-Hub/1.0'

# Cache global em memória (vive enquanto o processo Flask estiver de pé)
CNAE_CACHE: dict[str, dict] = {}

# Cores EFCT (paleta oficial)
COR_NAVY = '0C2B38'
COR_LIME = 'B3BC2B'
COR_AMBAR = '92400E'
COR_VERMELHO = '991B1B'
COR_BRANCO = 'FFFFFF'
FONT_BASE = 'Exo 2'

# ============================================================
# ESTILOS EFCT (openpyxl)
# ============================================================
def make_styles():
    """Cria os objetos de estilo openpyxl. Chamado uma vez por execução."""
    return {
        # Cabeçalho cols originais (1-15): fundo navy + texto lime
        'header': {
            'font': Font(name=FONT_BASE, size=12, color=COR_LIME),
            'fill': PatternFill('solid', fgColor=COR_NAVY),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        },
        # Cabeçalho cols adicionadas (16+): fundo lime + texto navy
        'new_header': {
            'font': Font(name=FONT_BASE, size=12, color=COR_NAVY),
            'fill': PatternFill('solid', fgColor=COR_LIME),
            'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        },
        # Info text (NF, Razão Social, CNPJ, CNAEs) — branco + navy
        'info': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        },
        # Valor OK — branco + navy (alinhado à direita, formato número)
        'value_ok': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'number_format': '#,##0.00',
        },
        # Valor divergente — branco + âmbar bold
        'value_warn': {
            'font': Font(name=FONT_BASE, size=11, color=COR_AMBAR, bold=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'number_format': '#,##0.00',
        },
        # Alerta (ambíguo / não identificado) — branco + vermelho italic
        'alert': {
            'font': Font(name=FONT_BASE, size=11, color=COR_VERMELHO, italic=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        },
        # Análise OK
        'analise_ok': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        },
        # Análise divergente — só o valor
        'analise_warn': {
            'font': Font(name=FONT_BASE, size=11, color=COR_AMBAR, bold=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='right', vertical='center'),
        },
        # Análise não localizada
        'analise_err': {
            'font': Font(name=FONT_BASE, size=11, color=COR_VERMELHO, italic=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        },
        # Padrão das células de dados (todas em Exo 2 navy)
        'default': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(vertical='center'),
        },
        # Valor sem crédito (NF presente mas em CST fora de 50-67) — italic gray-navy
        'value_info': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY, italic=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='right', vertical='center'),
            'number_format': '#,##0.00',
        },
        # Análise sem crédito — italic navy
        'analise_info': {
            'font': Font(name=FONT_BASE, size=11, color=COR_NAVY, italic=True),
            'fill': PatternFill('solid', fgColor=COR_BRANCO),
            'alignment': Alignment(horizontal='left', vertical='center', wrap_text=True),
        },
    }


def apply_style(cell, style):
    """Aplica um dicionário de atributos a uma célula openpyxl."""
    if 'font' in style:
        cell.font = style['font']
    if 'fill' in style:
        cell.fill = style['fill']
    if 'alignment' in style:
        cell.alignment = style['alignment']
    if 'number_format' in style:
        cell.number_format = style['number_format']


# ============================================================
# UTILITÁRIOS DE NORMALIZAÇÃO
# ============================================================
def normalize_nf(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).strip()
    if not s:
        return ''
    # Limpa decimais herdados do pandas (ex: "40273.0")
    if re.match(r'^\d+\.0+$', s):
        s = s.split('.')[0]
    if s.isdigit():
        return str(int(s))
    return s


def normalize_name(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ''
    nfd = unicodedata.normalize('NFD', str(s))
    no_accent = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^A-Z0-9]', '', no_accent.upper())


def normalize_cnpj(c) -> str:
    if c is None or (isinstance(c, float) and pd.isna(c)):
        return ''
    digits = re.sub(r'\D', '', str(c))
    if not digits:
        return ''
    return digits.zfill(14)


def format_cnpj(c) -> str:
    d = normalize_cnpj(c)
    if len(d) != 14:
        return d
    return f'{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}'


def to_month_year(val) -> str:
    """Normaliza qualquer formato de período para 'MM/YYYY'."""
    if val is None or val == '' or (isinstance(val, float) and pd.isna(val)):
        return ''
    if isinstance(val, (datetime, pd.Timestamp)):
        return f'{val.month:02d}/{val.year}'
    s = str(val).strip()
    # DD/MM/YYYY
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return f'{m.group(2).zfill(2)}/{m.group(3)}'
    # YYYY-MM-DD
    m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f'{m.group(2).zfill(2)}/{m.group(1)}'
    # Excel serial (número)
    try:
        n = float(s.replace(',', '.'))
        d = datetime(1899, 12, 30) + timedelta(days=n)
        return f'{d.month:02d}/{d.year}'
    except (ValueError, OverflowError):
        pass
    return s


def to_number(v) -> float:
    if v is None or v == '' or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # Trata números BR ("1.234,56") e EN ("1234.56")
    if ',' in s and '.' in s:
        # 1.234,56
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    s = re.sub(r'[^\d\-\.]', '', s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def _most_common_cnpj(values) -> str | None:
    """Devolve o CNPJ válido (14 dígitos) mais frequente numa lista (ou None)."""
    counts = {}
    for v in values:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        c = normalize_cnpj(v)
        if len(c) == 14:
            counts[c] = counts.get(c, 0) + 1
    if not counts:
        return None
    return max(counts.items(), key=lambda x: x[1])[0]


def is_credit_cst(cst) -> bool:
    """Faixa CST 50–67 (Direito a Crédito + Crédito Presumido)."""
    if cst is None or cst == '' or (isinstance(cst, float) and pd.isna(cst)):
        return False
    s = str(cst).strip()
    if not s:
        return False
    try:
        # Pode vir como "50", "50.0", "50 - ...", etc.
        n = int(float(re.match(r'\s*(\d+)', s).group(1)))
    except (ValueError, AttributeError):
        return False
    return 50 <= n <= 67


# ============================================================
# EXTRAÇÃO DE HISTÓRICO
# ============================================================
# Padrões de extração de NF do Complemento Histórico — testados em ordem de confiança.
# Cada padrão captura o número que vem após uma palavra-chave fiscal.
NF_PATTERNS = [
    # 1. VlrrefNF<digits> — padrão específico de alguns sistemas contábeis
    re.compile(r'Vlr\s*ref\s*N[\.\s]*F[\.\s]*(\d{2,10})', re.IGNORECASE),
    # 2. N(F) / Nota (F)iscal / NFe / NF-e — aceita separadores ENTRE N e F também
    #    (cobre "N F 4480", "N.F 4480", "NF 4480", "Nota Fiscal 4480", "NFe 4480")
    re.compile(
        r'\bN(?:ota)?[\s\.\-]*F(?:iscal)?[\s\.\-/#:]*(?:e[\s\.\-/]*)?(?:n[ºo°]?\.?[\s]*)?(\d{2,10})\b',
        re.IGNORECASE,
    ),
    # 3. N° / N. / Nº (sem o F) — quando o histórico abrevia
    re.compile(r'\bN(?:[oOºo°]|\.)\s*[\s\.\-/#:]*(\d{2,10})\b', re.IGNORECASE),
    # 4. Doc / Documento / Doc Fiscal
    re.compile(
        r'\bDoc(?:umento)?(?:\s*fiscal)?[\s\.\-/#:]*(?:n[ºo°]?\.?\s*)?(\d{2,10})\b',
        re.IGNORECASE,
    ),
    # 5. Fatura / Duplicata
    re.compile(
        r'\b(?:Fatura|Duplicata|Dupl?)[\s\.\-/#:]*(?:n[ºo°]?\.?\s*)?(\d{2,10})\b',
        re.IGNORECASE,
    ),
    # 6. Boleto / Bol — comum quando o histórico é "Pagto boleto NF 12345"
    re.compile(r'\bBol(?:eto)?[\s\.\-/#:]*(?:n[ºo°]?\.?\s*)?(\d{2,10})\b', re.IGNORECASE),
    # 7. Número solto no INÍCIO do histórico seguido de nome (ex: "32029  ALLFOR SOLUCOES")
    #    Vários históricos não têm prefixo "NF" — só o número e o nome do fornecedor.
    re.compile(r'^\s*(\d{2,10})\s+[A-Za-zÀ-ÿ]'),
]

# CNPJ formatado (12.345.678/0001-90) ou 14 dígitos seguidos
RE_CNPJ_FORMATADO = re.compile(r'\b(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\b')
RE_CNPJ_RAW = re.compile(r'(?<!\d)(\d{14})(?!\d)')


def _looks_like_nf_number(digits: str) -> bool:
    """NF típica tem 2-10 dígitos. Acima disso geralmente é CPF/CNPJ/CEP/conta."""
    return bool(digits) and digits.isdigit() and 2 <= len(digits) <= 10


def extract_nf_from_historico(complemento, col15_raw=None) -> dict:
    """Extrai NF do Complemento Histórico testando múltiplos padrões em ordem de confiança.

    Estratégia:
      1. Tenta cada padrão (VlrrefNF → NF → N. → Doc → Fatura → Boleto)
      2. Coleta TODOS os candidatos plausíveis (2–10 dígitos, descartando CPF/CNPJ)
      3. Se há col 15 (NF original da Razão) e bate com um candidato → usa esse
      4. Se houver só 1 candidato distinto → CONFIDENT
      5. Se houver múltiplos sem resolução → AMBIGUOUS
      6. Se nenhum candidato no histórico mas col 15 tem valor → FROM_COL15 (fallback)
      7. Caso contrário → NOT_FOUND
    """
    s = str(complemento) if complemento and not (isinstance(complemento, float) and pd.isna(complemento)) else ''
    candidates: list[str] = []

    if s:
        for pat in NF_PATTERNS:
            for m in pat.findall(s):
                if isinstance(m, tuple):
                    m = next((g for g in m if g), '')
                if _looks_like_nf_number(m):
                    try:
                        candidates.append(str(int(m)))
                    except (ValueError, OverflowError):
                        pass

    unique = list(dict.fromkeys(candidates))
    c15 = normalize_nf(col15_raw)

    # Caso 1: nenhum candidato no histórico
    if not unique:
        if c15:
            return {'value': c15, 'status': 'FROM_COL15', 'candidates': [c15]}
        return {
            'value': None,
            'status': ('EMPTY' if not s else 'NOT_FOUND'),
            'candidates': [],
        }

    # Caso 2: um candidato único
    if len(unique) == 1:
        return {'value': unique[0], 'status': 'CONFIDENT', 'candidates': unique}

    # Caso 3: múltiplos — tenta resolver com col 15
    if c15 and c15 in unique:
        return {'value': c15, 'status': 'AMBIGUOUS_RESOLVED', 'candidates': unique}

    # Caso 4: ambiguidade não resolvida
    return {'value': None, 'status': 'AMBIGUOUS', 'candidates': unique}


def extract_cnpj_from_historico(complemento) -> str | None:
    """Se o histórico já cita o CNPJ (formatado ou 14 dígitos), extrai direto.
    Mais confiável que tentar adivinhar pelo nome do fornecedor."""
    if not complemento or (isinstance(complemento, float) and pd.isna(complemento)):
        return None
    s = str(complemento)
    m = RE_CNPJ_FORMATADO.search(s)
    if m:
        return normalize_cnpj(m.group(1))
    m = RE_CNPJ_RAW.search(s)
    if m:
        return normalize_cnpj(m.group(1))
    return None


def extract_supplier_name(complemento) -> str:
    """Limpa o Complemento Histórico tentando isolar a Razão Social.

    Remove tokens conhecidos (VlrrefNF, NF/N F, Nota Fiscal, Doc, Fatura, Duplicata,
    Boleto, número solto no início, CNPJ/CPF, datas, palavras de ação contábil) e
    devolve o que sobra. Para o casamento real com o mapa SPED a função
    `lookup_cnpj` faz fuzzy match (Levenshtein), então mesmo uma limpeza
    imperfeita ainda permite achar o CNPJ.
    """
    if not complemento or (isinstance(complemento, float) and pd.isna(complemento)):
        return ''
    s = str(complemento)
    # Tokens fiscais com número (mesmos padrões da extração de NF — aceita espaços entre N e F)
    s = re.sub(r'\bVlr\s*ref\s*N[\.\s\-]*F[\.\s\-]*\d+', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\bN(?:ota)?[\s\.\-]*F(?:iscal)?[\s\.\-/#:]*(?:e[\s\.\-/]*)?(?:n[ºo°]?\.?[\s]*)?\d{2,}', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\bDoc(?:umento)?(?:\s*fiscal)?[\s\.\-/#:]*(?:n[ºo°]?\.?\s*)?\d{2,}', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:Fatura|Duplicata|Dupl?|Bol(?:eto)?)[\s\.\-/#:]*(?:n[ºo°]?\.?\s*)?\d{2,}', ' ', s, flags=re.IGNORECASE)
    # CNPJ e CPF (formatado e raw)
    s = re.sub(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b', ' ', s)
    s = re.sub(r'\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b', ' ', s)
    # Datas DD/MM/YY[YY] ou DD.MM.YY[YY] ou DD-MM-YY[YY]
    s = re.sub(r'\b\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}\b', ' ', s)
    # Número solto no INÍCIO (caso "25443 A GAZETA DO ESPIRITO SANTO RADIO E TV")
    # Apenas se seguido de espaço + letra (não come datas/CPF que não casaram acima)
    s = re.sub(r'^\s*\d{2,10}\s+(?=[A-Za-zÀ-ÿ])', ' ', s)
    # Palavras de ação contábil
    s = re.sub(
        r'\b(?:Pagto|Pagamento|Refer[eê]nte|Refer|Ref|Pgto|Vlr|Valor|Total|Pago|Aquisi[cç][aã]o|Compra|Devolu[cç][aã]o|pgmto)\b',
        ' ', s, flags=re.IGNORECASE,
    )
    # Restos de tokens "n.", "nº", "no" soltos
    s = re.sub(r'\bn[ºo°][\.\s]*\d*', ' ', s, flags=re.IGNORECASE)
    # Normaliza espaços múltiplos
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def build_nf_only_index(agg_nf_map: dict) -> dict:
    """Constrói um índice <NF> → lista de (CNPJ, entry) a partir do agg_map de
    C100-EFD/A100 (cujas chaves são "<NF>|<CNPJ>"). Permite buscar uma NF nos
    blocos sem precisar do CNPJ — usado como último fallback quando a Razão
    não dá pistas suficientes pra resolver o CNPJ pelo nome do fornecedor.
    """
    idx: dict[str, list] = {}
    for key, entry in agg_nf_map.items():
        if '|' not in key:
            continue
        nf, cnpj = key.split('|', 1)
        idx.setdefault(nf, []).append((cnpj, entry))
    return idx


def lookup_cnpj_by_nf_in_speds(nf_value, vlr_partida, nf_idx_list) -> dict | None:
    """Quarta camada de fallback: dado um número de NF e uma lista de índices
    NF-only (C100-EFD, A100 com crédito, A100 sem crédito, etc.), tenta achar
    qual CNPJ tem essa NF nos próprios SPEDs.

      • Se 1 único CNPJ tem essa NF em algum bloco → usa direto (via NF_SPED_UNICO).
      • Se múltiplos CNPJs tem → usa Vlr Partida como desempate (mais próximo
        dentro da tolerância → via NF_SPED_VALOR).
      • Caso contrário → None (não é seguro inferir CNPJ).

    nf_idx_list: lista de dicts construídos por build_nf_only_index, em ordem
    de prioridade (C100-EFD primeiro, A100 depois).
    """
    if not nf_value or not nf_idx_list:
        return None

    for idx in nf_idx_list:
        candidates = idx.get(nf_value)
        if not candidates:
            continue
        # Dedup CNPJs (uma mesma NF+CNPJ não aparece duas vezes no mesmo idx,
        # mas se aparecer em vários idx, o loop externo trata cada um)
        if len(candidates) == 1:
            cnpj_found, _entry = candidates[0]
            return {'cnpj': cnpj_found, 'via': 'NF_SPED_UNICO'}
        # Múltiplos → tenta desambiguar pelo valor mais próximo da Vlr Partida
        best_cnpj, best_entry = min(
            candidates, key=lambda c: abs(c[1]['total'] - vlr_partida)
        )
        if abs(best_entry['total'] - vlr_partida) <= TOLERANCIA_VALOR:
            return {'cnpj': best_cnpj, 'via': 'NF_SPED_VALOR'}
        # Ambíguo: vários CNPJs com essa NF e nenhum bate com Vlr Partida.
        # Não retorna nada — preferimos deixar sem CNPJ a inferir errado.
    return None


def lookup_cnpj_by_historico_scan(complemento, cnpj_map) -> dict | None:
    """Reverse lookup: varre o Complemento Histórico procurando nomes de fornecedores
    conhecidos (do mapa SPED) como substring. Funciona mesmo quando a extração de
    nome falha — basta que o nome do fornecedor esteja em algum lugar do histórico.

    Retorna o match de maior comprimento (mais específico) ou None.
    """
    if not complemento or not cnpj_map:
        return None
    norm_hist = normalize_name(complemento)
    if len(norm_hist) < 8:
        return None
    best_cnpj = None
    best_name = ''
    for sup_name, cnpj in cnpj_map.items():
        if len(sup_name) < 8:
            continue  # nomes muito curtos podem causar falso positivo
        if sup_name in norm_hist and len(sup_name) > len(best_name):
            best_cnpj = cnpj
            best_name = sup_name
    if best_cnpj:
        return {'cnpj': best_cnpj, 'via': 'HISTORICO_SCAN', 'matched_name': best_name}
    return None


# Mantida só para compatibilidade — a lógica agora está dentro de extract_nf_from_historico.
def resolve_nf_extraction(extracted: dict, col15_raw) -> dict:
    return extracted


def lookup_cnpj(extracted_name: str, cnpj_map: dict) -> dict | None:
    """Procura CNPJ a partir do nome extraído. 4 camadas de match em ordem:
      1. EXACT — match exato após normalização
      2. STRIP_DIGITS — remove CPF/CNPJ no fim do nome e tenta de novo
      3. PREFIX — chave do mapa é prefixo do nome extraído (ou vice-versa)
      4. FUZZY — similaridade Levenshtein ≥ 85% (cobre variações como "LTDA" vs "LTDA ME")
    """
    if not extracted_name:
        return None
    norm = normalize_name(extracted_name)
    if not norm:
        return None
    # 1. Exato
    if norm in cnpj_map:
        return {'cnpj': cnpj_map[norm], 'via': 'EXACT'}
    # 2. Strip CPF/CNPJ no fim
    stripped = re.sub(r'\d{11,14}$', '', norm)
    if stripped and stripped != norm and stripped in cnpj_map:
        return {'cnpj': cnpj_map[stripped], 'via': 'STRIP_DIGITS'}
    # 3. Prefixo (cobre abreviações)
    if len(stripped) >= 10:
        for k, v in cnpj_map.items():
            if k == stripped:
                return {'cnpj': v, 'via': 'EXACT_STRIPPED'}
            if (k.startswith(stripped) or stripped.startswith(k)) and min(len(k), len(stripped)) >= 10:
                return {'cnpj': v, 'via': 'PREFIX'}
    # 4. Fuzzy match (Levenshtein) — cobre "EMPRESA LTDA" vs "EMPRESA LTDA ME", etc.
    # Só dispara pra nomes longos o suficiente pra evitar falso positivo
    needle = stripped if len(stripped) >= 10 else norm
    if len(needle) >= 10 and cnpj_map:
        matches = difflib.get_close_matches(needle, list(cnpj_map.keys()), n=1, cutoff=0.85)
        if matches:
            return {'cnpj': cnpj_map[matches[0]], 'via': 'FUZZY', 'matched_name': matches[0]}
    return None


# ============================================================
# CNAE LOOKUP — BrasilAPI primário + CNPJa fallback, com cache
# ============================================================
def fetch_cnae_once(cnpj: str) -> dict | None:
    if cnpj in CNAE_CACHE:
        return CNAE_CACHE[cnpj]
    headers = {'User-Agent': USER_AGENT}
    # BrasilAPI primário
    try:
        r = requests.get(
            f'https://brasilapi.com.br/api/cnpj/v1/{cnpj}',
            headers=headers, timeout=CNAE_TIMEOUT
        )
        if r.ok:
            d = r.json()
            result = {
                'razao_social': d.get('razao_social') or d.get('nome_fantasia') or '',
                'cnae1_desc': d.get('cnae_fiscal_descricao') or '',
                'cnae2_desc': '',
                'source': 'BrasilAPI',
            }
            sec = d.get('cnaes_secundarios') or []
            if sec and isinstance(sec, list) and sec[0]:
                result['cnae2_desc'] = sec[0].get('descricao') or ''
            CNAE_CACHE[cnpj] = result
            return result
    except Exception as e:
        log.debug('BrasilAPI %s falhou: %s', cnpj, e)
    # CNPJa fallback
    try:
        r = requests.get(
            f'https://open.cnpja.com/office/{cnpj}',
            headers=headers, timeout=CNAE_TIMEOUT
        )
        if r.ok:
            d = r.json()
            main = d.get('mainActivity') or {}
            sides = d.get('sideActivities') or []
            side = sides[0] if sides else {}
            company = d.get('company') or {}
            result = {
                'razao_social': company.get('name') or d.get('alias') or '',
                'cnae1_desc': main.get('text') or '',
                'cnae2_desc': side.get('text') or '',
                'source': 'CNPJa',
            }
            CNAE_CACHE[cnpj] = result
            return result
    except Exception as e:
        log.debug('CNPJa %s falhou: %s', cnpj, e)
    return {'razao_social': '', 'cnae1_desc': '', 'cnae2_desc': '', 'source': 'ERROR'}


def fetch_all_cnaes(cnpjs: list[str]) -> dict[str, dict]:
    """Busca paralela com cache global. Falhas não bloqueiam o processo."""
    results = {}
    to_fetch = [c for c in cnpjs if c not in CNAE_CACHE]
    for c in cnpjs:
        if c in CNAE_CACHE:
            results[c] = CNAE_CACHE[c]
    log.info('CNAE: %s no cache, %s a buscar', len(cnpjs) - len(to_fetch), len(to_fetch))
    if not to_fetch:
        return results
    with ThreadPoolExecutor(max_workers=CNAE_CONCURRENCY) as ex:
        futures = {ex.submit(fetch_cnae_once, c): c for c in to_fetch}
        for fut in as_completed(futures):
            c = futures[fut]
            try:
                results[c] = fut.result() or {
                    'razao_social': '', 'cnae1_desc': '', 'cnae2_desc': '', 'source': 'ERROR'
                }
            except Exception as e:
                log.warning('CNAE fetch erro %s: %s', c, e)
                results[c] = {'razao_social': '', 'cnae1_desc': '', 'cnae2_desc': '', 'source': 'ERROR'}
    return results


# ============================================================
# LEITURA E AGREGAÇÃO DOS SPEDs
# ============================================================
# Schema dos blocos (1-indexed igual ao usado no front-end)
# Col 0 (1-indexed: 1) sempre é o CNPJ da EMPRESA QUE REPORTA (a empresa-raiz que está sendo auditada).
# Os outros CNPJs (col 'cnpj' aqui) são dos PARTICIPANTES (fornecedores/clientes).
SPED_SCHEMA = {
    'c100efd': {
        'cols_zero_indexed': [0, 2, 6, 8, 20, 55],  # 1-indexed: 1, 3, 7, 9, 21, 56
        'names': ['cnpj_empresa', 'periodo', 'cnpj', 'nome', 'nf', 'valor'],
        'apply_cst': False,
        'kind': 'nf',
    },
    'a100': {
        'cols_zero_indexed': [0, 1, 5, 7, 11, 41, 42],  # 1-indexed: 1, 2, 6, 8, 12, 42, 43
        'names': ['cnpj_empresa', 'periodo', 'cnpj', 'nome', 'nf', 'cst', 'valor'],
        'apply_cst': True,
        'kind': 'nf',
    },
    'f100': {
        'cols_zero_indexed': [0, 1, 4, 6, 15, 22, 23],  # 1-indexed: 1, 2, 5, 7, 16, 23, 24
        'names': ['cnpj_empresa', 'periodo', 'cnpj', 'nome', 'vlr_operacao', 'cst', 'valor'],
        'apply_cst': True,
        'kind': 'f100',
    },
}


def _safe_read_excel(file_stream, friendly_name, **kwargs):
    """
    Envolve pd.read_excel num try/except amigável que diz QUAL arquivo falhou
    e por quê (xlsx inválido, arquivo vazio, formato xls antigo etc).
    """
    try:
        if hasattr(file_stream, 'seek'):
            try:
                file_stream.seek(0)
            except Exception:
                pass
        return pd.read_excel(file_stream, engine='openpyxl', **kwargs)
    except Exception as e:
        msg = str(e).lower()
        if 'not a zip file' in msg or 'badzipfile' in msg:
            raise ValueError(
                f'Não foi possível abrir "{friendly_name}". O arquivo não parece ser '
                f'um .xlsx válido (pode estar vazio, corrompido, em formato .xls antigo, '
                f'ou ser um .csv renomeado). Reabra no Excel e salve novamente como '
                f'"Pasta de Trabalho do Excel (*.xlsx)".'
            )
        raise ValueError(f'Erro ao ler "{friendly_name}": {e}')


def read_sped(file_stream, schema, friendly_name='SPED') -> pd.DataFrame:
    """
    Lê um SPED com `usecols` para carregar APENAS as colunas que precisamos.
    Reduz drasticamente o uso de memória (de 88 cols para 5-6 cols).
    """
    df = _safe_read_excel(
        file_stream,
        friendly_name=friendly_name,
        header=0,
        usecols=schema['cols_zero_indexed'],
        dtype=object,  # mantém tudo como objeto pra normalizar depois
    )
    df.columns = schema['names'][:len(df.columns)]
    return df


def process_sped_block(file_stream, schema, cnpj_map: dict, friendly_name='SPED'):
    """
    Lê + agrega um SPED em um passo só, populando o mapa Nome→CNPJ.

    Para A100/F100 (apply_cst=True) construímos DOIS aggregates:
      • agg          → CST 50-67 (com crédito — alvo principal do cruzamento)
      • agg_no_cred  → CST FORA de 50-67 (sem crédito — informativo)
    Pra C100-EFD (apply_cst=False), só `agg` (sem distinção de CST).

    Retorna dict com 'agg', 'agg_no_cred', 'skipped', 'total', 'cnpj_added', 'company_cnpj'.
    """
    df = read_sped(file_stream, schema, friendly_name=friendly_name)
    apply_cst = schema['apply_cst']
    kind = schema['kind']

    agg_map = {}
    agg_no_cred = {} if apply_cst else None
    skipped = 0  # linhas que foram pro aggregate sem crédito
    cnpj_added = 0
    company_cnpj_count = {}

    def _push_nf(target, nf_n, cnpj_val, value):
        key = f'{nf_n}|{cnpj_val}'
        entry = target.get(key)
        if entry:
            entry['total'] += value
            entry['items'] += 1
        else:
            target[key] = {'total': value, 'items': 1}

    def _push_f100(target, cnpj_val, per, op, bc):
        key = f'{cnpj_val}|{per}'
        if target.get(key) is None:
            target[key] = [{'op': op, 'bc': bc}]
        else:
            target[key].append({'op': op, 'bc': bc})

    for row in df.itertuples(index=False):
        # Rastreia CNPJ da empresa-raiz (col 1 = cnpj_empresa)
        ce_raw = getattr(row, 'cnpj_empresa', None)
        if ce_raw is not None:
            ce_norm = normalize_cnpj(ce_raw)
            if len(ce_norm) == 14:
                company_cnpj_count[ce_norm] = company_cnpj_count.get(ce_norm, 0) + 1

        # Decide pra qual aggregate vai a linha
        has_credit = True if not apply_cst else is_credit_cst(row.cst)
        target = agg_map if has_credit else agg_no_cred
        if not has_credit:
            skipped += 1
        if target is None:
            continue

        # Build Nome→CNPJ map sempre (independente de ter crédito)
        name_key = normalize_name(row.nome)
        cnpj_val = normalize_cnpj(row.cnpj)
        if name_key and len(cnpj_val) == 14 and name_key not in cnpj_map:
            cnpj_map[name_key] = cnpj_val
            cnpj_added += 1

        if kind == 'nf':
            nf_n = normalize_nf(row.nf)
            if not nf_n or not cnpj_val:
                continue
            _push_nf(target, nf_n, cnpj_val, to_number(row.valor))
        elif kind == 'f100':
            per = to_month_year(row.periodo)
            if not cnpj_val or not per:
                continue
            _push_f100(target, cnpj_val, per,
                       to_number(row.vlr_operacao), to_number(row.valor))

    total = len(df)
    del df  # libera memória imediatamente
    company_cnpj = (
        max(company_cnpj_count.items(), key=lambda x: x[1])[0]
        if company_cnpj_count else None
    )
    return {
        'agg': agg_map,
        'agg_no_cred': agg_no_cred,
        'skipped': skipped,
        'total': total,
        'cnpj_added': cnpj_added,
        'company_cnpj': company_cnpj,
    }


def find_f100_match(f100_map, cnpj, periodo, vlr_partida):
    """Para F100, procura operações com Vlr Operação ≈ Vlr Partida (tolerância)."""
    key = f'{cnpj}|{periodo}'
    lst = f100_map.get(key)
    if not lst:
        return None
    matches = [x for x in lst if abs(x['op'] - vlr_partida) <= TOLERANCIA_VALOR]
    if matches:
        total = sum(x['bc'] for x in matches)
        return {'total': total, 'count': len(matches), 'exact_match': True}
    # Sem match exato: ainda assim retorna a soma do CNPJ+período (informativo)
    total = sum(x['bc'] for x in lst)
    return {'total': total, 'count': len(lst), 'exact_match': False}


# ============================================================
# ANÁLISE DO CRUZAMENTO
# ============================================================
def _fmt_brl(value):
    """Formata número como BR currency 'R$ 1.234,56'."""
    s = f'{abs(value):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return f'R$ {s}'


def build_analise(vlr_partida, val_efd, val_a100, val_f100,
                  partidas_match_efd=False, partidas_match_a100=False,
                  total_partidas=None, partidas_count=1,
                  a100_no_credit=False, f100_no_credit=False,
                  efd_via_nf_only=False):
    """Devolve os 3 campos finais (separados pra facilitar filtros):
      - analise_text / analise_style  → status enxuto na col ANÁLISE
      - soma_partidas_text            → texto da col "Fecha somando partidas" (ou vazio)
      - diferenca_value               → número da col "Valor da diferença" (ou None)

    Quando `a100_no_credit` ou `f100_no_credit` é True, a NF foi localizada
    no bloco mas em CST fora de 50-67 (sem crédito).
    Quando `efd_via_nf_only` é True, a NF foi localizada no C100-EFD por NF only
    (CNPJ não foi confirmado pelo cruzamento estrito).
    """
    # Blocos com match confirmado (NF + CNPJ batendo) e com crédito
    blocos_cred = []
    if val_efd is not None and not efd_via_nf_only:
        blocos_cred.append('C100-EFD')
    if val_a100 is not None and not a100_no_credit:
        blocos_cred.append('A100-CONTRI')
    if val_f100 is not None and not f100_no_credit:
        blocos_cred.append('F100-CONTRI')

    # Blocos onde achou mas sem crédito (CST fora 50-67)
    blocos_nc = []
    if a100_no_credit:
        blocos_nc.append('A100-CONTRI')
    if f100_no_credit:
        blocos_nc.append('F100-CONTRI')

    # Blocos onde achou só pela NF (informativo — CNPJ não confirmado)
    blocos_nf_only = []
    if val_efd is not None and efd_via_nf_only:
        blocos_nf_only.append('C100-EFD')

    def _fmt_blocos(lst):
        return lst[0] if len(lst) == 1 else ', '.join(lst[:-1]) + ' e ' + lst[-1]

    # Nada encontrado em lugar nenhum
    if not blocos_cred and not blocos_nc and not blocos_nf_only:
        return {
            'analise_text': 'Não localizado nos blocos de crédito',
            'analise_style': 'analise_err',
            'soma_partidas_text': '',
            'diferenca_value': None,
        }

    # Achou só por NF (CNPJ não confirmado) — informativo, evita duplicidade
    if not blocos_cred and not blocos_nc and blocos_nf_only:
        return {
            'analise_text': f'Encontrado em {_fmt_blocos(blocos_nf_only)} (NF localizada, CNPJ não confirmado)',
            'analise_style': 'analise_info',
            'soma_partidas_text': '',
            'diferenca_value': None,
        }

    # Achou só sem crédito (nada com crédito)
    if not blocos_cred and blocos_nc:
        sufixo_nf = f' · também em {_fmt_blocos(blocos_nf_only)} (NF localizada)' if blocos_nf_only else ''
        return {
            'analise_text': f'Encontrado em {_fmt_blocos(blocos_nc)} sem crédito{sufixo_nf}',
            'analise_style': 'analise_info',
            'soma_partidas_text': '',
            'diferenca_value': None,
        }

    # Achou com crédito — somar só os valores que TÊM crédito (ignora os no_credit)
    soma_sped = 0.0
    if val_efd is not None:
        soma_sped += val_efd
    if val_a100 is not None and not a100_no_credit:
        soma_sped += val_a100
    if val_f100 is not None and not f100_no_credit:
        soma_sped += val_f100
    diff = vlr_partida - soma_sped

    blocos_txt = _fmt_blocos(blocos_cred)
    # Sufixo opcional pra blocos sem crédito e/ou NF-only quando também há crédito
    sufixos = []
    if blocos_nc:
        sufixos.append(f'também em {_fmt_blocos(blocos_nc)} sem crédito')
    if blocos_nf_only:
        sufixos.append(f'NF também em {_fmt_blocos(blocos_nf_only)}')
    sufixo_nc = f' ({"; ".join(sufixos)})' if sufixos else ''

    # 1) Match individual (Vlr Partida ≈ Σ SPED com crédito)
    if abs(diff) <= TOLERANCIA_VALOR:
        return {
            'analise_text': f'Encontrado em {blocos_txt}{sufixo_nc}',
            'analise_style': 'analise_ok',
            'soma_partidas_text': '',
            'diferenca_value': None,
        }

    # 2) Match por soma das partidas da Razão
    if (partidas_match_efd or partidas_match_a100) and partidas_count > 1 and total_partidas is not None:
        return {
            'analise_text': f'Encontrado em {blocos_txt}{sufixo_nc}',
            'analise_style': 'analise_ok',
            'soma_partidas_text': f'Fecha somando as {partidas_count} partidas ({_fmt_brl(total_partidas)})',
            'diferenca_value': None,
        }

    # 3) Divergência real
    return {
        'analise_text': f'Encontrado em {blocos_txt} com diferença{sufixo_nc}',
        'analise_style': 'analise_warn',
        'soma_partidas_text': '',
        'diferenca_value': round(diff, 2),
    }


# ============================================================
# CONSTRUÇÃO DO XLSX DE SAÍDA (com padrão visual EFCT)
# ============================================================
COLS_ORIGINAIS = 15  # Cols 1-15 da Razão original (até a NF)
NOVAS_COLS = [
    'Número da Nota',           # 16
    'Razão Social',             # 17
    'CNPJ',                     # 18
    '1º CNAE',                  # 19
    '2º CNAE',                  # 20
    'C100 - EFD FISCAL',        # 21
    'A100 - CONTRI',            # 22
    'F100 - CONTRI',            # 23
    'Fecha somando partidas',   # 24 — preenchida só quando match é por soma agregada
    'Valor da diferença',       # 25 — preenchida só quando há divergência (com sinal)
    'ANÁLISE DO CRUZAMENTO',    # 26 — status enxuto pra filtrar
]
TOTAL_COLS = COLS_ORIGINAIS + len(NOVAS_COLS)  # 26

COL_WIDTHS = [
    16, 9, 9, 11, 13, 9, 10, 38, 10, 13, 5, 9, 9, 38, 10,         # originais 1-15
    14, 38, 20, 48, 48, 16, 16, 16, 44, 18, 38                     # novas 16-26
]


def build_workbook(df_razao, enriched_list, agg_efd, agg_a100, agg_f100, cnae_results,
                   razao_partidas_sum=None, razao_partidas_count=None,
                   agg_a100_nc=None, agg_f100_nc=None,
                   nf_idx_efd=None):
    """Constrói o workbook com 2 abas (Razão cruzada + Pendências por CNPJ).

    `agg_a100_nc` e `agg_f100_nc` são aggregates de NFs SEM crédito (CST fora 50-67).
    Quando uma linha não acha match com crédito mas acha sem crédito, a célula é
    preenchida em estilo informativo (italic navy) e a Análise indica "sem crédito".

    `nf_idx_efd` é o índice NF-only do C100-EFD — usado pra mostrar a presença da NF
    no Fiscal mesmo quando o cruzamento estrito (NF+CNPJ) falhou. Evita que a usuária
    conte essa NF como benefício em duas análises diferentes (duplicidade).
    """
    razao_partidas_sum = razao_partidas_sum or {}
    razao_partidas_count = razao_partidas_count or {}
    agg_a100_nc = agg_a100_nc or {}
    agg_f100_nc = agg_f100_nc or {}
    nf_idx_efd = nf_idx_efd or {}
    wb = Workbook()
    ws = wb.active
    ws.title = 'RAZÃO CONTABIL'
    styles = make_styles()

    # ----- Cabeçalho -----
    header_src = list(df_razao.columns)[:COLS_ORIGINAIS]
    # Garante 15 cols mesmo se o original tiver menos
    while len(header_src) < COLS_ORIGINAIS:
        header_src.append('')
    full_header = header_src + NOVAS_COLS
    for col_idx, name in enumerate(full_header, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        if col_idx <= COLS_ORIGINAIS:
            apply_style(cell, styles['header'])
        else:
            apply_style(cell, styles['new_header'])

    # ----- Stats e pendências -----
    stats = {
        'linhas_razao': 0,
        'match_efd': 0, 'match_a100': 0, 'match_f100': 0,
        'match_any': 0, 'no_match': 0,
        'div_efd': 0, 'div_a100': 0, 'div_f100': 0,
        'sem_divergencia': 0,
        'amb_unresolved': 0, 'not_found': 0, 'cnpj_missing': 0,
        'amb_resolved': 0,
    }
    pendencias = {}  # cnpj → {cnpj, nome, cnae1, cnae2, total, count}

    # ----- Linhas de dados -----
    for i, (row_data, meta) in enumerate(zip(df_razao.itertuples(index=False), enriched_list), start=2):
        stats['linhas_razao'] += 1
        # Cols 1-15 originais
        for c in range(COLS_ORIGINAIS):
            val = row_data[c] if c < len(row_data) else None
            if val is not None and isinstance(val, float) and pd.isna(val):
                val = None
            cell = ws.cell(row=i, column=c + 1, value=val)
            apply_style(cell, styles['default'])

        # Resolve display dos campos enriquecidos
        nf_ext = meta['nf_ext']
        st = nf_ext['status']
        if st in ('CONFIDENT', 'CONFIDENT_ALT', 'AMBIGUOUS_RESOLVED', 'FROM_COL15'):
            nf_display = nf_ext['value']
            nf_style = 'info'
            nf_value = nf_ext['value']
        elif st == 'AMBIGUOUS':
            nf_display = '⚠ AMBÍGUO: ' + '/'.join(nf_ext['candidates'])
            nf_style = 'alert'
            nf_value = None
            stats['amb_unresolved'] += 1
        elif st == 'NOT_FOUND':
            nf_display = '⚠ NÃO IDENTIFICADO'
            nf_style = 'alert'
            nf_value = None
            stats['not_found'] += 1
        else:
            nf_display = ''
            nf_style = 'info'
            nf_value = None
        if st == 'AMBIGUOUS_RESOLVED':
            stats['amb_resolved'] += 1
        if st == 'FROM_COL15':
            stats['from_col15'] = stats.get('from_col15', 0) + 1

        cnpj = meta['cnpj']
        cnpj_via = meta.get('cnpj_via') or ''
        if cnpj_via.startswith('NF_SPED'):
            stats['cnpj_via_nf_sped'] = stats.get('cnpj_via_nf_sped', 0) + 1
        elif cnpj_via == 'HISTORICO_SCAN':
            stats['cnpj_via_scan'] = stats.get('cnpj_via_scan', 0) + 1
        elif cnpj_via == 'DIRETO_HISTORICO':
            stats['cnpj_via_direto'] = stats.get('cnpj_via_direto', 0) + 1
        cnae_data = cnae_results.get(cnpj, {}) if cnpj else {}
        # Razão Social: prioridade BrasilAPI > extraído
        razao_social_display = (
            cnae_data.get('razao_social') or meta['supplier_name'] or ''
        )
        cnpj_display = format_cnpj(cnpj) if cnpj else '⚠ CNPJ NÃO ENCONTRADO' if meta else ''
        if not cnpj:
            stats['cnpj_missing'] += 1
        cnae1 = cnae_data.get('cnae1_desc') or ''
        cnae2 = cnae_data.get('cnae2_desc') or ''

        # Col 16: NF Extraída
        cell = ws.cell(row=i, column=16, value=nf_display)
        apply_style(cell, styles[nf_style])

        # Col 17: Razão Social
        cell = ws.cell(row=i, column=17, value=razao_social_display)
        apply_style(cell, styles['info'])

        # Col 18: CNPJ
        cell = ws.cell(row=i, column=18, value=cnpj_display)
        apply_style(cell, styles['info'] if cnpj else styles['alert'])

        # Col 19: 1º CNAE (descrição completa)
        cell = ws.cell(row=i, column=19, value=cnae1)
        apply_style(cell, styles['info'])

        # Col 20: 2º CNAE
        cell = ws.cell(row=i, column=20, value=cnae2)
        apply_style(cell, styles['info'])

        # Triple Check
        vlr_partida = to_number(row_data[9]) if len(row_data) > 9 else 0  # col 10 Vlr Partida (0-indexed 9)
        periodo = to_month_year(row_data[1]) if len(row_data) > 1 else ''  # col 2 Período

        # Soma de partidas da Razão para esta NF+CNPJ (pra detectar match por soma)
        partida_key = f'{nf_value}|{cnpj}' if (nf_value and cnpj) else None
        total_partidas = razao_partidas_sum.get(partida_key, vlr_partida) if partida_key else vlr_partida
        partidas_count = razao_partidas_count.get(partida_key, 1) if partida_key else 1

        # C100-EFD: match por NF + CNPJ — soma itens da nota.
        # Se Σ itens bate com Vlr Partida → match individual (verde).
        # Se Σ itens não bate com partida mas bate com Σ TODAS as partidas dessa NF → match por soma (verde, sinalizado).
        val_efd = None
        diverge_efd = False
        partidas_match_efd = False
        efd_via_nf_only = False  # NF está no C100-EFD mas CNPJ não foi confirmado
        if nf_value and cnpj:
            hit = agg_efd.get(f'{nf_value}|{cnpj}')
            if hit:
                val_efd = round(hit['total'], 2)
                diff_ind = abs(val_efd - vlr_partida)
                diff_tot = abs(val_efd - total_partidas)
                if diff_ind <= TOLERANCIA_VALOR:
                    diverge_efd = False
                elif partidas_count > 1 and diff_tot <= TOLERANCIA_VALOR:
                    diverge_efd = False
                    partidas_match_efd = True
                else:
                    diverge_efd = True
                stats['match_efd'] += 1
                if diverge_efd:
                    stats['div_efd'] += 1
                if partidas_match_efd:
                    stats['partidas_match_efd'] = stats.get('partidas_match_efd', 0) + 1

        # FALLBACK FINAL DO C100-EFD: se o match estrito não pegou MAS a NF existe
        # no Fiscal, traz o valor informativo. Evita que essa NF apareça como
        # "Não localizada" aqui e seja contada como benefício em outra análise.
        if val_efd is None and nf_value:
            candidates = nf_idx_efd.get(nf_value, [])
            if candidates:
                # Soma TODAS as ocorrências dessa NF no C100-EFD (qualquer CNPJ).
                val_efd = round(sum(e['total'] for _, e in candidates), 2)
                efd_via_nf_only = True
                stats['efd_via_nf_only'] = stats.get('efd_via_nf_only', 0) + 1

        # A100: mesma lógica de C100-EFD. Se não achar com crédito, tenta sem crédito.
        val_a100 = None
        diverge_a100 = False
        partidas_match_a100 = False
        a100_no_credit = False  # NF está no A100 mas em CST fora 50-67
        if nf_value and cnpj:
            hit = agg_a100.get(f'{nf_value}|{cnpj}')
            if hit:
                val_a100 = round(hit['total'], 2)
                diff_ind = abs(val_a100 - vlr_partida)
                diff_tot = abs(val_a100 - total_partidas)
                if diff_ind <= TOLERANCIA_VALOR:
                    diverge_a100 = False
                elif partidas_count > 1 and diff_tot <= TOLERANCIA_VALOR:
                    diverge_a100 = False
                    partidas_match_a100 = True
                else:
                    diverge_a100 = True
                stats['match_a100'] += 1
                if diverge_a100:
                    stats['div_a100'] += 1
                if partidas_match_a100:
                    stats['partidas_match_a100'] = stats.get('partidas_match_a100', 0) + 1
            else:
                # Fallback: NF presente no A100 sem crédito (CST fora 50-67)
                hit_nc = agg_a100_nc.get(f'{nf_value}|{cnpj}')
                if hit_nc:
                    val_a100 = round(hit_nc['total'], 2)
                    a100_no_credit = True
                    stats['no_credit_a100'] = stats.get('no_credit_a100', 0) + 1

        val_f100 = None
        diverge_f100 = False
        f100_no_credit = False
        if cnpj and periodo:
            m = find_f100_match(agg_f100, cnpj, periodo, vlr_partida)
            if m and m['exact_match']:
                val_f100 = round(m['total'], 2)
                diverge_f100 = abs(val_f100 - vlr_partida) > TOLERANCIA_VALOR
                stats['match_f100'] += 1
                if diverge_f100:
                    stats['div_f100'] += 1
            else:
                # Fallback: F100 sem crédito (CST fora 50-67)
                m_nc = find_f100_match(agg_f100_nc, cnpj, periodo, vlr_partida)
                if m_nc and m_nc['exact_match']:
                    val_f100 = round(m_nc['total'], 2)
                    f100_no_credit = True
                    stats['no_credit_f100'] = stats.get('no_credit_f100', 0) + 1

        # Col 21: C100 - EFD FISCAL
        cell = ws.cell(row=i, column=21, value=val_efd)
        if efd_via_nf_only:
            apply_style(cell, styles['value_info'])  # NF localizada sem confirmação de CNPJ
        else:
            apply_style(cell, styles['value_warn'] if diverge_efd else styles['value_ok'])

        # Col 22: A100 - CONTRI (italic se for sem crédito)
        cell = ws.cell(row=i, column=22, value=val_a100)
        if a100_no_credit:
            apply_style(cell, styles['value_info'])
        else:
            apply_style(cell, styles['value_warn'] if diverge_a100 else styles['value_ok'])

        # Col 23: F100 - CONTRI (italic se for sem crédito)
        cell = ws.cell(row=i, column=23, value=val_f100)
        if f100_no_credit:
            apply_style(cell, styles['value_info'])
        else:
            apply_style(cell, styles['value_warn'] if diverge_f100 else styles['value_ok'])

        analise = build_analise(
            vlr_partida, val_efd, val_a100, val_f100,
            partidas_match_efd=partidas_match_efd,
            partidas_match_a100=partidas_match_a100,
            total_partidas=total_partidas,
            partidas_count=partidas_count,
            a100_no_credit=a100_no_credit,
            f100_no_credit=f100_no_credit,
            efd_via_nf_only=efd_via_nf_only,
        )

        # Col 24: Fecha somando partidas (só preenchida quando match foi via soma)
        cell = ws.cell(row=i, column=24, value=analise['soma_partidas_text'] or None)
        apply_style(cell, styles['info'])

        # Col 25: Valor da diferença (só preenchida quando há divergência)
        diff_val = analise['diferenca_value']
        cell = ws.cell(row=i, column=25, value=diff_val if diff_val is not None else None)
        if diff_val is not None:
            apply_style(cell, styles['value_warn'])
            cell.number_format = 'R$ #,##0.00;-R$ #,##0.00'
        else:
            apply_style(cell, styles['default'])

        # Col 26: ANÁLISE DO CRUZAMENTO (status enxuto pra filtrar)
        cell = ws.cell(row=i, column=26, value=analise['analise_text'])
        apply_style(cell, styles[analise['analise_style']])

        # Stats e pendências
        if analise['analise_style'] == 'analise_ok':
            stats['sem_divergencia'] += 1
        if val_efd is not None or val_a100 is not None or val_f100 is not None:
            stats['match_any'] += 1
        else:
            stats['no_match'] += 1
            # Pendência: agrega por CNPJ (ou marker '__NO_CNPJ__')
            pkey = cnpj or '__NO_CNPJ__'
            p = pendencias.get(pkey)
            if p is None:
                p = {
                    'cnpj': cnpj or '',
                    'nome': razao_social_display,
                    'cnae1': cnae1,
                    'cnae2': cnae2,
                    'total': 0.0,
                    'count': 0,
                }
                pendencias[pkey] = p
            p['total'] += vlr_partida
            p['count'] += 1
            if not p['nome'] and razao_social_display:
                p['nome'] = razao_social_display

    # ----- Ajustes finais da aba RAZÃO -----
    # Larguras
    for col_idx, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    # Altura do cabeçalho
    ws.row_dimensions[1].height = 32
    # Sem gridlines
    ws.sheet_view.showGridLines = False
    # Congelar cabeçalho
    ws.freeze_panes = 'A2'
    # Autofiltro
    ws.auto_filter.ref = f'A1:{get_column_letter(TOTAL_COLS)}{stats["linhas_razao"] + 1}'

    # ----- Aba PENDÊNCIAS DE CRUZAMENTO -----
    pend_total_geral = 0.0
    pend_list = sorted(pendencias.values(), key=lambda x: x['total'], reverse=True)
    if pend_list:
        ws2 = wb.create_sheet('PENDÊNCIAS DE CRUZAMENTO')
        pend_header = [
            'CNPJ', 'Razão Social', '1º CNAE', '2º CNAE',
            'Valor Total Não Encontrado', 'Qtd Lançamentos',
        ]
        for col_idx, name in enumerate(pend_header, start=1):
            cell = ws2.cell(row=1, column=col_idx, value=name)
            apply_style(cell, styles['new_header'])

        for i, p in enumerate(pend_list, start=2):
            ws2.cell(row=i, column=1, value=format_cnpj(p['cnpj']) if p['cnpj'] else '⚠ Sem CNPJ')
            ws2.cell(row=i, column=2, value=p['nome'] or '(sem identificação)')
            ws2.cell(row=i, column=3, value=p['cnae1'] or '')
            ws2.cell(row=i, column=4, value=p['cnae2'] or '')
            ws2.cell(row=i, column=5, value=round(p['total'], 2))
            ws2.cell(row=i, column=6, value=p['count'])
            pend_total_geral += p['total']

            apply_style(ws2.cell(row=i, column=1), styles['info'] if p['cnpj'] else styles['alert'])
            apply_style(ws2.cell(row=i, column=2), styles['info'])
            apply_style(ws2.cell(row=i, column=3), styles['info'])
            apply_style(ws2.cell(row=i, column=4), styles['info'])
            apply_style(ws2.cell(row=i, column=5), styles['value_warn'])
            apply_style(ws2.cell(row=i, column=6), styles['info'])

        # Linha de total
        total_row = len(pend_list) + 2
        ws2.cell(row=total_row, column=4, value='TOTAL GERAL →')
        ws2.cell(row=total_row, column=5, value=round(pend_total_geral, 2))
        ws2.cell(row=total_row, column=6, value=sum(p['count'] for p in pend_list))
        for c in [4, 5, 6]:
            apply_style(ws2.cell(row=total_row, column=c), styles['new_header'])
        ws2.cell(row=total_row, column=5).number_format = '#,##0.00'

        # Larguras / vista
        for col_idx, w in enumerate([22, 42, 48, 48, 24, 16], start=1):
            ws2.column_dimensions[get_column_letter(col_idx)].width = w
        ws2.row_dimensions[1].height = 32
        ws2.sheet_view.showGridLines = False
        ws2.freeze_panes = 'A2'
        ws2.auto_filter.ref = f'A1:F{len(pend_list) + 1}'

    stats['pendencias'] = len(pend_list)
    stats['pendencias_total'] = round(pend_total_geral, 2)
    return wb, stats


# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def processar_cruzamento(
    razao_stream, c100efd_stream, a100_stream=None, f100_stream=None, skip_cnae=False
):
    """
    Entrypoint chamado pelo Flask.
    `a100_stream` e `f100_stream` são OPCIONAIS — se a empresa não tem esses blocos,
    o processamento segue só com Razão + C100-EFD (aggregates A100/F100 ficam vazios).
    Retorna (bytes_xlsx, dict_stats).
    """
    log.info('Lendo Razão Contábil…')
    df_razao = _safe_read_excel(razao_stream, friendly_name='Razão Contábil',
                                header=0, dtype=object)
    log.info('Razão: %s linhas', len(df_razao))

    # CNPJ raiz da Razão (col 1, normalmente o mesmo em todas as linhas)
    razao_company_cnpj = _most_common_cnpj(df_razao.iloc[:, 0].tolist())

    cnpj_map: dict[str, str] = {}

    log.info('Processando C100-EFD FISCAL…')
    r_efd = process_sped_block(c100efd_stream, SPED_SCHEMA['c100efd'], cnpj_map,
                               friendly_name='C100-EFD FISCAL')
    agg_efd_raw = r_efd['agg']
    efd_company_cnpj = r_efd['company_cnpj']
    log.info('C100-EFD: %s linhas · %s chaves agregadas', r_efd['total'], len(agg_efd_raw))

    if a100_stream is not None:
        log.info('Processando A100-CONTRI…')
        r_a100 = process_sped_block(a100_stream, SPED_SCHEMA['a100'], cnpj_map,
                                    friendly_name='A100-CONTRI')
        agg_a100_raw = r_a100['agg']
        agg_a100_nc = r_a100['agg_no_cred'] or {}
        a100_company_cnpj = r_a100['company_cnpj']
        log.info('A100: %s linhas · %s sem crédito (CST fora 50-67) · %s chaves com crédito · %s chaves sem crédito',
                 r_a100['total'], r_a100['skipped'], len(agg_a100_raw), len(agg_a100_nc))
    else:
        log.info('A100-CONTRI: arquivo não fornecido — pulando bloco.')
        agg_a100_raw = {}
        agg_a100_nc = {}
        a100_company_cnpj = None

    if f100_stream is not None:
        log.info('Processando F100-CONTRI…')
        r_f100 = process_sped_block(f100_stream, SPED_SCHEMA['f100'], cnpj_map,
                                    friendly_name='F100-CONTRI')
        agg_f100_raw = r_f100['agg']
        agg_f100_nc = r_f100['agg_no_cred'] or {}
        f100_company_cnpj = r_f100['company_cnpj']
        log.info('F100: %s linhas · %s sem crédito · %s chaves com crédito · %s chaves sem crédito',
                 r_f100['total'], r_f100['skipped'], len(agg_f100_raw), len(agg_f100_nc))
    else:
        log.info('F100-CONTRI: arquivo não fornecido — pulando bloco.')
        agg_f100_raw = {}
        agg_f100_nc = {}
        f100_company_cnpj = None

    log.info('Mapa CNPJ: %s fornecedores', len(cnpj_map))
    # Avisos proativos quando o mapa pode estar "pobre"
    if a100_stream is None:
        log.warning('A100 não foi enviado — fornecedores de serviços (advogados, '
                    'consultorias, prestadores) provavelmente não estarão no mapa CNPJ. '
                    'Esses casos podem aparecer como "CNPJ NÃO ENCONTRADO".')
    if f100_stream is None:
        log.warning('F100 não foi enviado — operações como depreciação, frete, '
                    'receitas financeiras não serão validadas.')

    # ──────────────────────────────────────────────────────────────────────
    # VALIDAÇÃO: CNPJ raiz da empresa precisa bater entre todos os 4 arquivos.
    # Evita que o usuário importe acidentalmente um A100 de outra empresa.
    # ──────────────────────────────────────────────────────────────────────
    cnpjs_por_arquivo = {
        'Razão Contábil': razao_company_cnpj,
        'C100-EFD FISCAL': efd_company_cnpj,
        'A100-CONTRI': a100_company_cnpj,
        'F100-CONTRI': f100_company_cnpj,
    }
    log.info('CNPJs da empresa-raiz por arquivo: %s', cnpjs_por_arquivo)
    cnpjs_validos = {k: v for k, v in cnpjs_por_arquivo.items() if v}
    cnpjs_distintos = set(cnpjs_validos.values())
    if len(cnpjs_distintos) > 1:
        # Identifica qual(is) divergem do CNPJ majoritário
        majoritario = max(
            set(cnpjs_validos.values()),
            key=lambda c: list(cnpjs_validos.values()).count(c),
        )
        divergentes = [
            f'{nome}={format_cnpj(c)}'
            for nome, c in cnpjs_validos.items() if c != majoritario
        ]
        msg = (
            f'⚠ CNPJ da empresa difere entre os arquivos. '
            f'Esperado (majoritário): {format_cnpj(majoritario)}. '
            f'Divergente(s): {" · ".join(divergentes)}. '
            f'Confira se você não importou um SPED de outra empresa por engano.'
        )
        raise ValueError(msg)
    elif len(cnpjs_distintos) == 0:
        log.warning('Nenhum CNPJ da empresa-raiz pôde ser identificado em nenhum arquivo.')
    else:
        log.info('✓ CNPJ raiz validado em todos os arquivos: %s',
                 format_cnpj(next(iter(cnpjs_distintos))))

    # Índices NF-only pros blocos com NF — usados como último fallback pra inferir CNPJ
    # quando a Razão não fornece pistas suficientes (sem CNPJ direto, sem nome no histórico, etc).
    # Ordem de prioridade: C100-EFD (mais reliable) → A100 com crédito → A100 sem crédito.
    nf_idx_list = [
        build_nf_only_index(agg_efd_raw),
        build_nf_only_index(agg_a100_raw),
        build_nf_only_index(agg_a100_nc),
    ]
    log.info('Índices NF-only: C100-EFD %s NFs · A100 com cr %s · A100 s/ cr %s',
             len(nf_idx_list[0]), len(nf_idx_list[1]), len(nf_idx_list[2]))

    # Pré-processa Razão (extração de NF, lookup CNPJ) + agrega partidas por NF+CNPJ
    log.info('Extraindo NF e CNPJ da Razão…')
    enriched_list = []
    unique_cnpjs = set()
    razao_partidas_sum = {}    # chave NF|CNPJ → soma das Vlr Partida da Razão
    razao_partidas_count = {}  # chave NF|CNPJ → quantas partidas da Razão dão nesse documento
    # Convert NaN to None na Razão pra evitar problemas no openpyxl
    df_razao = df_razao.where(pd.notna(df_razao), None)
    for row in df_razao.itertuples(index=False):
        # row[13] = Complemento Histórico (col 14, 0-indexed 13)
        # row[14] = NF original (col 15, 0-indexed 14)
        complemento = row[13] if len(row) > 13 else None
        col15 = row[14] if len(row) > 14 else None
        vlr_partida_row = to_number(row[9]) if len(row) > 9 else 0  # col 10 Vlr Partida

        # NF: extração multi-padrão + fallback automático pra col 15
        nf_ext = extract_nf_from_historico(complemento, col15_raw=col15)
        supplier_name = extract_supplier_name(complemento)

        # CNPJ — cascata de estratégias (mais confiável primeiro):
        #   1. CNPJ direto no histórico (formatado ou 14 dígitos)
        #   2. Nome limpo extraído + lookup_cnpj (exato → fuzzy)
        #   3. Reverse lookup: varre histórico procurando nomes do mapa SPED
        #   4. NF-only fallback: se temos a NF, busca nos blocos SPED por essa NF
        #      (único CNPJ tem ela? usa. Múltiplos? desempata por Vlr Partida.)
        lookup = None
        direct_cnpj = extract_cnpj_from_historico(complemento)
        if direct_cnpj and len(direct_cnpj) == 14:
            lookup = {'cnpj': direct_cnpj, 'via': 'DIRETO_HISTORICO'}
        else:
            lookup = lookup_cnpj(supplier_name, cnpj_map)
            if not lookup:
                lookup = lookup_cnpj_by_historico_scan(complemento, cnpj_map)
                if lookup and lookup.get('matched_name') and not supplier_name:
                    supplier_name = lookup['matched_name']
            if not lookup:
                # Última tentativa: usa a NF (se identificada) pra buscar nos blocos SPED
                nf_v = nf_ext.get('value')
                if nf_v:
                    lookup = lookup_cnpj_by_nf_in_speds(nf_v, vlr_partida_row, nf_idx_list)

        cnpj = lookup['cnpj'] if lookup else None
        if lookup:
            via = lookup.get('via', '')
            if via.startswith('NF_SPED'):
                # Estatística de quantas foram salvas por esse fallback
                # (não conta DIRETO_HISTORICO nem HISTORICO_SCAN — só o último recurso)
                pass  # contador feito no build_workbook
        if cnpj:
            unique_cnpjs.add(cnpj)
        # Acumula partidas por NF+CNPJ (apenas se a NF foi identificada com confiança)
        nf_val = nf_ext['value'] if nf_ext['status'] in (
            'CONFIDENT', 'CONFIDENT_ALT', 'AMBIGUOUS_RESOLVED'
        ) else None
        if nf_val and cnpj:
            partida_key = f'{nf_val}|{cnpj}'
            razao_partidas_sum[partida_key] = razao_partidas_sum.get(partida_key, 0.0) + vlr_partida_row
            razao_partidas_count[partida_key] = razao_partidas_count.get(partida_key, 0) + 1
        enriched_list.append({
            'nf_ext': nf_ext,
            'supplier_name': supplier_name,
            'cnpj': cnpj,
            'cnpj_via': lookup.get('via') if lookup else None,
        })
    log.info('CNPJs únicos a buscar CNAE: %s', len(unique_cnpjs))

    # Busca CNAEs
    cnae_results = {}
    if not skip_cnae and unique_cnpjs:
        log.info('Buscando CNAEs via BrasilAPI/CNPJa…')
        cnae_results = fetch_all_cnaes(list(unique_cnpjs))
        log.info('CNAEs: %s respondidos', len(cnae_results))

    # Monta o workbook
    log.info('Montando workbook de saída…')
    wb, stats = build_workbook(
        df_razao, enriched_list,
        agg_efd_raw, agg_a100_raw, agg_f100_raw,
        cnae_results,
        razao_partidas_sum, razao_partidas_count,
        agg_a100_nc=agg_a100_nc, agg_f100_nc=agg_f100_nc,
        nf_idx_efd=nf_idx_list[0],  # índice NF-only do C100-EFD (já calculado acima)
    )
    stats['cnaes_buscados'] = len(cnae_results)

    # Serializa
    buf = io.BytesIO()
    wb.save(buf)
    out_bytes = buf.getvalue()
    log.info('Saída: %s bytes', len(out_bytes))

    return out_bytes, stats
