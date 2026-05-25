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
RE_VLRREFNF = re.compile(r'VlrrefNF\s*(\d+)', re.IGNORECASE)
RE_NF_ALT = re.compile(r'\bNF[\s\.:#]*(\d{2,})', re.IGNORECASE)


def extract_nf_from_historico(complemento) -> dict:
    """Extrai NF do Complemento Histórico, classificando confiança."""
    if not complemento or (isinstance(complemento, float) and pd.isna(complemento)):
        return {'value': None, 'status': 'EMPTY', 'candidates': []}
    s = str(complemento)
    matches = [str(int(m)) for m in RE_VLRREFNF.findall(s)]
    if matches:
        unique = list(dict.fromkeys(matches))
        if len(unique) == 1:
            return {'value': unique[0], 'status': 'CONFIDENT', 'candidates': unique}
        return {'value': None, 'status': 'AMBIGUOUS', 'candidates': unique}
    alt = [str(int(m)) for m in RE_NF_ALT.findall(s)]
    if not alt:
        return {'value': None, 'status': 'NOT_FOUND', 'candidates': []}
    unique = list(dict.fromkeys(alt))
    if len(unique) == 1:
        return {'value': unique[0], 'status': 'CONFIDENT_ALT', 'candidates': unique}
    return {'value': None, 'status': 'AMBIGUOUS', 'candidates': unique}


def extract_supplier_name(complemento) -> str:
    """Extrai a Razão Social do Complemento (tira os tokens VlrrefNF e CPF/CNPJ no fim)."""
    if not complemento or (isinstance(complemento, float) and pd.isna(complemento)):
        return ''
    cleaned = re.sub(r'VlrrefNF\s*\d+', '', str(complemento)).strip()
    cleaned = re.sub(r'\d{11,14}$', '', cleaned).strip()
    return cleaned


def resolve_nf_extraction(extracted: dict, col15_raw) -> dict:
    """Tenta resolver ambiguidade comparando com a coluna NF original (col 15)."""
    st = extracted['status']
    if st in ('CONFIDENT', 'CONFIDENT_ALT'):
        return extracted
    if st == 'AMBIGUOUS':
        c15 = normalize_nf(col15_raw)
        if c15 and c15 in extracted['candidates']:
            return {
                'value': c15,
                'status': 'AMBIGUOUS_RESOLVED',
                'candidates': extracted['candidates'],
            }
    return extracted


def lookup_cnpj(extracted_name: str, cnpj_map: dict) -> dict | None:
    """Procura CNPJ a partir do nome extraído (exato → strip dígitos → prefixo)."""
    if not extracted_name:
        return None
    norm = normalize_name(extracted_name)
    if not norm:
        return None
    if norm in cnpj_map:
        return {'cnpj': cnpj_map[norm], 'via': 'EXACT'}
    stripped = re.sub(r'\d{11,14}$', '', norm)
    if stripped and stripped != norm and stripped in cnpj_map:
        return {'cnpj': cnpj_map[stripped], 'via': 'STRIP_DIGITS'}
    # Prefixo (cobre abreviações)
    if len(stripped) >= 10:
        for k, v in cnpj_map.items():
            if k == stripped:
                return {'cnpj': v, 'via': 'EXACT_STRIPPED'}
            if (k.startswith(stripped) or stripped.startswith(k)) and min(len(k), len(stripped)) >= 10:
                return {'cnpj': v, 'via': 'PREFIX'}
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
SPED_SCHEMA = {
    'c100efd': {
        'cols_zero_indexed': [2, 6, 8, 20, 55],  # 1-indexed: 3, 7, 9, 21, 56
        'names': ['periodo', 'cnpj', 'nome', 'nf', 'valor'],
        'apply_cst': False,
        'kind': 'nf',
    },
    'a100': {
        'cols_zero_indexed': [1, 5, 7, 11, 41, 42],  # 1-indexed: 2, 6, 8, 12, 42, 43
        'names': ['periodo', 'cnpj', 'nome', 'nf', 'cst', 'valor'],
        'apply_cst': True,
        'kind': 'nf',
    },
    'f100': {
        'cols_zero_indexed': [1, 4, 6, 15, 22, 23],  # 1-indexed: 2, 5, 7, 16, 23, 24
        'names': ['periodo', 'cnpj', 'nome', 'vlr_operacao', 'cst', 'valor'],
        'apply_cst': True,
        'kind': 'f100',
    },
}


def read_sped(file_stream, schema) -> pd.DataFrame:
    """
    Lê um SPED com `usecols` para carregar APENAS as colunas que precisamos.
    Reduz drasticamente o uso de memória (de 88 cols para 5-6 cols).
    """
    df = pd.read_excel(
        file_stream,
        header=0,
        usecols=schema['cols_zero_indexed'],
        dtype=object,  # mantém tudo como objeto pra normalizar depois
        engine='openpyxl',
    )
    df.columns = schema['names'][:len(df.columns)]
    return df


def process_sped_block(file_stream, schema, cnpj_map: dict):
    """
    Lê + agrega um SPED em um passo só, populando o mapa Nome→CNPJ.
    Retorna (agg_map, skipped_by_cst, total_rows, cnpj_added).
    """
    df = read_sped(file_stream, schema)
    apply_cst = schema['apply_cst']
    kind = schema['kind']

    agg_map = {}
    skipped = 0
    cnpj_added = 0

    for row in df.itertuples(index=False):
        # CST filter (A100/F100)
        if apply_cst and not is_credit_cst(row.cst):
            skipped += 1
            continue
        # Build Nome→CNPJ map
        name_key = normalize_name(row.nome)
        cnpj_val = normalize_cnpj(row.cnpj)
        if name_key and len(cnpj_val) == 14 and name_key not in cnpj_map:
            cnpj_map[name_key] = cnpj_val
            cnpj_added += 1
        # Aggregate
        per = to_month_year(row.periodo)
        if kind == 'nf':
            nf_n = normalize_nf(row.nf)
            if not nf_n or not cnpj_val or not per:
                continue
            key = f'{nf_n}|{cnpj_val}|{per}'
            v = to_number(row.valor)
            entry = agg_map.get(key)
            if entry:
                entry['total'] += v
                entry['items'] += 1
            else:
                agg_map[key] = {'total': v, 'items': 1}
        elif kind == 'f100':
            if not cnpj_val or not per:
                continue
            key = f'{cnpj_val}|{per}'
            op = to_number(row.vlr_operacao)
            bc = to_number(row.valor)
            lst = agg_map.get(key)
            if lst is None:
                agg_map[key] = [{'op': op, 'bc': bc}]
            else:
                lst.append({'op': op, 'bc': bc})

    total = len(df)
    del df  # libera memória imediatamente
    return agg_map, skipped, total, cnpj_added


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
def build_analise(vlr_partida, val_efd, val_a100, val_f100):
    """Gera o texto + estilo da coluna ANÁLISE DO CRUZAMENTO."""
    blocos = []
    if val_efd is not None:
        blocos.append('C100-EFD')
    if val_a100 is not None:
        blocos.append('A100-CONTRI')
    if val_f100 is not None:
        blocos.append('F100-CONTRI')

    if not blocos:
        return {'text': 'Não localizado nos blocos de crédito', 'style': 'analise_err'}

    soma_sped = (val_efd or 0) + (val_a100 or 0) + (val_f100 or 0)
    diff = vlr_partida - soma_sped
    blocos_txt = (
        blocos[0] if len(blocos) == 1
        else ', '.join(blocos[:-1]) + ' e ' + blocos[-1]
    )
    if abs(diff) <= TOLERANCIA_VALOR:
        return {'text': f'Encontrado em {blocos_txt}', 'style': 'analise_ok'}
    # Divergência: só o valor com sinal
    sign = '-' if diff < 0 else ''
    txt = f'{sign}R$ {abs(diff):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return {'text': txt, 'style': 'analise_warn'}


# ============================================================
# CONSTRUÇÃO DO XLSX DE SAÍDA (com padrão visual EFCT)
# ============================================================
COLS_ORIGINAIS = 15  # Cols 1-15 da Razão original (até a NF)
NOVAS_COLS = [
    'Número da Nota',
    'Razão Social',
    'CNPJ',
    '1º CNAE',
    '2º CNAE',
    'C100 - EFD FISCAL',
    'A100 - CONTRI',
    'F100 - CONTRI',
    'ANÁLISE DO CRUZAMENTO',
]
TOTAL_COLS = COLS_ORIGINAIS + len(NOVAS_COLS)  # 24

COL_WIDTHS = [
    16, 9, 9, 11, 13, 9, 10, 38, 10, 13, 5, 9, 9, 38, 10,        # originais 1-15
    14, 38, 20, 48, 48, 16, 16, 16, 38                            # novas 16-24
]


def build_workbook(df_razao, enriched_list, agg_efd, agg_a100, agg_f100, cnae_results):
    """Constrói o workbook com 2 abas (Razão cruzada + Pendências por CNPJ)."""
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
        if st in ('CONFIDENT', 'CONFIDENT_ALT', 'AMBIGUOUS_RESOLVED'):
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

        cnpj = meta['cnpj']
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

        val_efd = None
        diverge_efd = False
        if nf_value and cnpj and periodo:
            hit = agg_efd.get(f'{nf_value}|{cnpj}|{periodo}')
            if hit:
                val_efd = round(hit['total'], 2)
                diverge_efd = abs(val_efd - vlr_partida) > TOLERANCIA_VALOR
                stats['match_efd'] += 1
                if diverge_efd:
                    stats['div_efd'] += 1

        val_a100 = None
        diverge_a100 = False
        if nf_value and cnpj and periodo:
            hit = agg_a100.get(f'{nf_value}|{cnpj}|{periodo}')
            if hit:
                val_a100 = round(hit['total'], 2)
                diverge_a100 = abs(val_a100 - vlr_partida) > TOLERANCIA_VALOR
                stats['match_a100'] += 1
                if diverge_a100:
                    stats['div_a100'] += 1

        val_f100 = None
        diverge_f100 = False
        if cnpj and periodo:
            m = find_f100_match(agg_f100, cnpj, periodo, vlr_partida)
            if m and m['exact_match']:
                val_f100 = round(m['total'], 2)
                diverge_f100 = abs(val_f100 - vlr_partida) > TOLERANCIA_VALOR
                stats['match_f100'] += 1
                if diverge_f100:
                    stats['div_f100'] += 1

        # Col 21: C100 - EFD FISCAL
        cell = ws.cell(row=i, column=21, value=val_efd)
        apply_style(cell, styles['value_warn'] if diverge_efd else styles['value_ok'])

        # Col 22: A100 - CONTRI
        cell = ws.cell(row=i, column=22, value=val_a100)
        apply_style(cell, styles['value_warn'] if diverge_a100 else styles['value_ok'])

        # Col 23: F100 - CONTRI
        cell = ws.cell(row=i, column=23, value=val_f100)
        apply_style(cell, styles['value_warn'] if diverge_f100 else styles['value_ok'])

        # Col 24: ANÁLISE DO CRUZAMENTO
        analise = build_analise(vlr_partida, val_efd, val_a100, val_f100)
        cell = ws.cell(row=i, column=24, value=analise['text'])
        apply_style(cell, styles[analise['style']])

        # Stats e pendências
        if analise['style'] == 'analise_ok':
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
    razao_stream, c100efd_stream, a100_stream, f100_stream, skip_cnae=False
):
    """
    Entrypoint chamado pelo Flask.
    Retorna (bytes_xlsx, dict_stats).
    """
    log.info('Lendo Razão Contábil…')
    df_razao = pd.read_excel(razao_stream, header=0, dtype=object, engine='openpyxl')
    log.info('Razão: %s linhas', len(df_razao))

    cnpj_map: dict[str, str] = {}

    log.info('Processando C100-EFD FISCAL…')
    agg_efd_raw, _, total_efd, cnpj_add_efd = process_sped_block(
        c100efd_stream, SPED_SCHEMA['c100efd'], cnpj_map
    )
    log.info('C100-EFD: %s linhas · %s chaves agregadas', total_efd, len(agg_efd_raw))

    log.info('Processando A100-CONTRI…')
    agg_a100_raw, skip_a100, total_a100, cnpj_add_a100 = process_sped_block(
        a100_stream, SPED_SCHEMA['a100'], cnpj_map
    )
    log.info('A100: %s linhas · %s fora CST · %s chaves', total_a100, skip_a100, len(agg_a100_raw))

    log.info('Processando F100-CONTRI…')
    agg_f100_raw, skip_f100, total_f100, cnpj_add_f100 = process_sped_block(
        f100_stream, SPED_SCHEMA['f100'], cnpj_map
    )
    log.info('F100: %s linhas · %s fora CST · %s chaves', total_f100, skip_f100, len(agg_f100_raw))
    log.info('Mapa CNPJ: %s fornecedores', len(cnpj_map))

    # Pré-processa Razão (extração de NF, lookup CNPJ)
    log.info('Extraindo NF e CNPJ da Razão…')
    enriched_list = []
    unique_cnpjs = set()
    # Convert NaN to None na Razão pra evitar problemas no openpyxl
    df_razao = df_razao.where(pd.notna(df_razao), None)
    for row in df_razao.itertuples(index=False):
        # row[13] = Complemento Histórico (col 14, 0-indexed 13)
        # row[14] = NF original (col 15, 0-indexed 14)
        complemento = row[13] if len(row) > 13 else None
        col15 = row[14] if len(row) > 14 else None
        nf_ext = resolve_nf_extraction(extract_nf_from_historico(complemento), col15)
        supplier_name = extract_supplier_name(complemento)
        lookup = lookup_cnpj(supplier_name, cnpj_map)
        cnpj = lookup['cnpj'] if lookup else None
        if cnpj:
            unique_cnpjs.add(cnpj)
        enriched_list.append({
            'nf_ext': nf_ext,
            'supplier_name': supplier_name,
            'cnpj': cnpj,
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
        df_razao, enriched_list, agg_efd_raw, agg_a100_raw, agg_f100_raw, cnae_results
    )
    stats['cnaes_buscados'] = len(cnae_results)

    # Serializa
    buf = io.BytesIO()
    wb.save(buf)
    out_bytes = buf.getvalue()
    log.info('Saída: %s bytes', len(out_bytes))

    return out_bytes, stats
