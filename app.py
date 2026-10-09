import os
from pathlib import Path

import src.app as _backend

_backend.app.template_folder = str(Path(__file__).parent / "templates")


def __getattr__(name):
    return getattr(_backend, name)


if __name__ == "__main__":
    _backend.app.run(
        debug=os.getenv('FLASK_DEBUG', 'false').lower() == 'true' and _backend.APP_ENV != 'production',
        use_reloader=False,
        host='0.0.0.0',
        port=int(os.getenv('PORT', '5000')),
    )# ============================================================
# app.py - Sistema de Análise de Dados com Flask
# ============================================================

import os
import re
import hmac
import base64
import hashlib
import json
import secrets
import uuid
import unicodedata
from difflib import SequenceMatcher
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
import bcrypt
from dotenv import load_dotenv
from flask import Flask, request, jsonify, render_template, redirect, send_file, session, url_for
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

try:
    from pyodide.ffi import run_sync
except ImportError:
    run_sync = None

# Em desenvolvimento, o .env do projeto prevalece sobre variáveis herdadas
# pelo terminal ou pelo depurador do VS Code. Em produção, o ambiente prevalece.
load_dotenv(override=os.getenv('NEXUS_ENV', 'development').strip().lower() == 'development')

APP_ENV = os.getenv('NEXUS_ENV', 'development').strip().lower()
if APP_ENV not in {'development', 'production'}:
    raise RuntimeError('NEXUS_ENV deve ser development ou production.')

ADMIN_USER = os.getenv('ADMIN_USER', '').strip()
ADMIN_PASS_HASH = os.getenv('ADMIN_PASS_HASH', '')
SECRET_KEY = os.getenv('SECRET_KEY')
if APP_ENV == 'production':
    if not SECRET_KEY or len(SECRET_KEY) < 32:
        raise RuntimeError('Configure SECRET_KEY com pelo menos 32 caracteres no ambiente de produção.')
    if not ADMIN_USER or not ADMIN_PASS_HASH:
        raise RuntimeError('Configure ADMIN_USER e ADMIN_PASS_HASH no ambiente de produção.')

app = Flask(__name__)
app.secret_key = SECRET_KEY or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=(
        APP_ENV == 'production'
        or os.getenv('SESSION_COOKIE_SECURE', 'false').lower() == 'true'
    ),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
)
if APP_ENV == 'production':
    app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1)

app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024
ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
UPLOAD_CHUNK_BYTES = 256 * 1024


def _cloudflare_env():
    try:
        return request.environ.get('workers.env')
    except (RuntimeError, AttributeError):
        return None


def _d1():
    env = _cloudflare_env()
    return getattr(env, 'NEXA_DB', None) if env is not None else None


def _d1_rows(sql, params=()):
    db = _d1()
    if db is None or run_sync is None:
        return None
    statement = db.prepare(sql)
    if params:
        statement = statement.bind(*params)
    result = run_sync(statement.all())
    try:
        return result.results.to_py()
    except AttributeError:
        return []


def _d1_run(sql, params=()):
    db = _d1()
    if db is None or run_sync is None:
        return False
    statement = db.prepare(sql)
    if params:
        statement = statement.bind(*params)
    run_sync(statement.run())
    return True


def _d1_first_value(sql, params=()):
    rows = _d1_rows(sql, params)
    if not rows:
        return None
    row = rows[0]
    try:
        return next(iter(row.values()))
    except AttributeError:
        return None


def _store_spreadsheet(filename, content):
    content_hash = hashlib.sha256(content).hexdigest()
    db = _d1()
    if db is not None:
        if run_sync is None:
            raise RuntimeError('D1 está disponível, mas a ponte assíncrona não foi carregada.')
        existing = _d1_rows(
            'SELECT id, filename, size_bytes, sha256, uploaded_at FROM uploaded_spreadsheets '
            'WHERE sha256 = ? ORDER BY uploaded_at DESC LIMIT 1',
            (content_hash,),
        )
        if existing:
            return existing[0]

    archive = Path(app.root_path) / 'uploads'
    if db is None and archive.exists():
        for metadata_path in archive.glob('*.json'):
            try:
                existing = json.loads(metadata_path.read_text(encoding='utf-8'))
            except (OSError, json.JSONDecodeError):
                continue
            if existing.get('sha256') == content_hash:
                return existing

    spreadsheet_id = uuid.uuid4().hex
    safe_name = secure_filename(filename) or 'planilha.xlsx'
    uploaded_at = datetime.now().isoformat(timespec='seconds')
    metadata = {
        'id': spreadsheet_id,
        'filename': safe_name,
        'size_bytes': len(content),
        'sha256': content_hash,
        'uploaded_at': uploaded_at,
    }
    if db is not None:
        chunk_count = (len(content) + UPLOAD_CHUNK_BYTES - 1) // UPLOAD_CHUNK_BYTES
        statement = db.prepare(
            'INSERT INTO uploaded_spreadsheets '
            '(id, filename, size_bytes, sha256, uploaded_at, chunk_count) '
            'VALUES (?, ?, ?, ?, ?, ?)'
        ).bind(
            spreadsheet_id, safe_name, len(content), metadata['sha256'], uploaded_at, chunk_count
        )
        run_sync(statement.run())
        for chunk_index, start in enumerate(range(0, len(content), UPLOAD_CHUNK_BYTES)):
            encoded = base64.b64encode(content[start:start + UPLOAD_CHUNK_BYTES]).decode('ascii')
            chunk_statement = db.prepare(
                'INSERT INTO uploaded_spreadsheet_chunks (spreadsheet_id, chunk_index, data_b64) '
                'VALUES (?, ?, ?)'
            ).bind(spreadsheet_id, chunk_index, encoded)
            run_sync(chunk_statement.run())
        return metadata

    archive.mkdir(exist_ok=True)
    file_path = archive / f'{spreadsheet_id}_{safe_name}'
    file_path.write_bytes(content)
    (archive / f'{spreadsheet_id}.json').write_text(json.dumps(metadata), encoding='utf-8')
    return metadata


def _list_spreadsheets():
    if _d1() is not None and run_sync is None:
        raise RuntimeError('D1 está disponível, mas a ponte assíncrona não foi carregada.')
    rows = _d1_rows(
        'SELECT id, filename, size_bytes, sha256, uploaded_at FROM uploaded_spreadsheets '
        'ORDER BY uploaded_at DESC'
    )
    if rows is not None:
        return rows
    archive = Path(app.root_path) / 'uploads'
    if not archive.exists():
        return []
    metadata = []
    for file_path in archive.glob('*.json'):
        try:
            metadata.append(json.loads(file_path.read_text(encoding='utf-8')))
        except (OSError, json.JSONDecodeError):
            continue
    return sorted(metadata, key=lambda item: item.get('uploaded_at', ''), reverse=True)


def _read_spreadsheet(spreadsheet_id):
    if _d1() is not None and run_sync is None:
        raise RuntimeError('D1 está disponível, mas a ponte assíncrona não foi carregada.')
    if _d1() is not None:
        metadata = _d1_rows(
            'SELECT chunk_count FROM uploaded_spreadsheets WHERE id = ?',
            (spreadsheet_id,),
        )
        if not metadata:
            return None
        chunks = []
        for chunk_index in range(metadata[0]['chunk_count']):
            rows = _d1_rows(
                'SELECT data_b64 FROM uploaded_spreadsheet_chunks '
                'WHERE spreadsheet_id = ? AND chunk_index = ?',
                (spreadsheet_id, chunk_index),
            )
            if not rows:
                return None
            chunks.append(base64.b64decode(rows[0]['data_b64']))
        return b''.join(chunks)
    archive = Path(app.root_path) / 'uploads'
    matches = list(archive.glob(f'{spreadsheet_id}_*')) if archive.exists() else []
    return matches[0].read_bytes() if matches else None


@app.before_request
def exigir_login():
    if request.endpoint in {'login', 'static'}:
        return

    if ADMIN_USER and session.get('usuario') == ADMIN_USER:
        token = session.get('token_sessao')
        if token:
            token_d1 = _d1_first_value('SELECT token FROM sessoes WHERE usuario = ?', (ADMIN_USER,))
            if token_d1 is None and _d1() is not None:
                session.clear()
            elif token_d1 is None or hmac.compare_digest(str(token_d1), str(token)):
                return
        elif _d1() is None:
            return

    session.clear()
    if request.path.startswith('/api/'):
        return jsonify({'erro': 'Faça login para continuar.'}), 401
    return redirect(url_for('login'))
ALIASES_COLUNAS = {
    'data': ('data', 'date', 'transaction date', 'posting date', 'fecha', 'fecha de operacion', 'datum', 'date de transaction', 'date comptable'),
    'ano': ('ano', 'year', 'ejercicio', 'jahr', 'annee'),
    'mes': ('mes', 'month', 'monat', 'mois'),
    'categoria': ('categoria', 'category', 'expense category', 'cost center', 'cost centre', 'account', 'rubro', 'nature', 'kategorie', 'kostenart', 'department', 'division', 'business unit', 'supplier', 'vendor', 'supplier name', 'vendor name', 'proveedor', 'fornecedor', 'fournisseur', 'lieferant'),
    'valor': ('valor', 'valor r$', 'valor total', 'total', 'amount', 'amount actual', 'actual amount', 'expense', 'expenses', 'spend', 'expenditure', 'cost', 'costs', 'importe', 'monto', 'gasto', 'montant', 'depense', 'depenses', 'betrag', 'ausgabe', 'ausgaben'),
    'budget': ('budget', 'budget r$', 'budget amount', 'budgeted amount', 'planned amount', 'plan amount', 'presupuesto', 'orçamento', 'orcamento', 'montant budget', 'budget alloue', 'planwert', 'soll'),
    'tipo': ('tipo', 'type', 'kind', 'nature type', 'typ', 'classe'),
}


def normalizar_cabecalho(valor):
    texto = unicodedata.normalize('NFKD', str(valor).casefold())
    texto = ''.join(caractere for caractere in texto if not unicodedata.combining(caractere))
    return re.sub(r'[^a-z0-9]+', ' ', texto).strip()


def mapear_colunas(colunas):
    aliases = {
        campo: {normalizar_cabecalho(alias) for alias in nomes}
        for campo, nomes in ALIASES_COLUNAS.items()
    }
    candidatas = []
    for coluna in colunas:
        nome = normalizar_cabecalho(coluna)
        if not nome:
            continue
        for campo, nomes in aliases.items():
            similaridade = max(
                1.0 if nome == alias else
                0.98 if len(alias.split()) > 1 and alias in nome else
                0.93 if len(alias) >= 5 and alias in nome.split() else
                SequenceMatcher(None, nome, alias).ratio()
                for alias in nomes
            )
            if similaridade >= 0.84:
                candidatas.append((similaridade, campo, coluna))

    mapeadas = {}
    colunas_usadas = set()
    for _, campo, coluna in sorted(candidatas, reverse=True, key=lambda item: item[0]):
        if campo not in mapeadas and coluna not in colunas_usadas:
            mapeadas[campo] = coluna
            colunas_usadas.add(coluna)
    return mapeadas


def ler_abas_flexivel(arquivo):
    excel = pd.ExcelFile(BytesIO(arquivo))
    abas = {}
    for nome_aba in excel.sheet_names:
        previa = pd.read_excel(excel, sheet_name=nome_aba, header=None, nrows=12)
        melhor_linha = None
        melhor_pontuacao = 0
        for indice, linha in previa.iterrows():
            mapeamento = mapear_colunas(linha.dropna().tolist())
            pontuacao = len(mapeamento)
            if pontuacao > melhor_pontuacao:
                melhor_linha = indice
                melhor_pontuacao = pontuacao
        cabecalho = melhor_linha if melhor_pontuacao >= 2 else 0
        abas[nome_aba] = pd.read_excel(excel, sheet_name=nome_aba, header=cabecalho)
    return abas


def converter_numero_local(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    negativo = texto.startswith('(') and texto.endswith(')')
    if negativo:
        texto = texto[1:-1]
    texto = re.sub(r'(?i)(USD|EUR|BRL|CAD|AUD|GBP|CHF|JPY|CNY|INR|US\$|R\$|[$€£¥])', '', texto)
    texto = re.sub(r'\s+', '', texto)
    if re.search(r'[^0-9,.+\-]', texto):
        return None
    if not texto or texto in ('-', '+'):
        return None
    if ',' in texto and '.' in texto:
        separador_decimal = ',' if texto.rfind(',') > texto.rfind('.') else '.'
        separador_milhar = '.' if separador_decimal == ',' else ','
        texto = texto.replace(separador_milhar, '').replace(separador_decimal, '.')
    elif ',' in texto:
        casas = len(texto) - texto.rfind(',') - 1
        texto = texto.replace(',', '.') if casas in (1, 2) else texto.replace(',', '')
    elif texto.count('.') > 1:
        partes = texto.split('.')
        texto = ''.join(partes[:-1]) + '.' + partes[-1]
    elif '.' in texto and len(texto) - texto.rfind('.') - 1 == 3:
        texto = texto.replace('.', '')
    try:
        numero = float(texto)
        return -numero if negativo else numero
    except ValueError:
        return None


def converter_mes_local(valor):
    numero = converter_numero_local(valor)
    if numero is not None:
        return int(numero) if 1 <= numero <= 12 else None
    mes = normalizar_cabecalho(valor)
    aliases = {
        'janeiro': 1, 'january': 1, 'enero': 1, 'janvier': 1, 'januar': 1,
        'fevereiro': 2, 'february': 2, 'febrero': 2, 'fevrier': 2, 'februar': 2,
        'marco': 3, 'march': 3, 'marzo': 3, 'mars': 3, 'marz': 3,
        'abril': 4, 'april': 4, 'avril': 4,
        'maio': 5, 'may': 5, 'mayo': 5, 'mai': 5,
        'junho': 6, 'june': 6, 'junio': 6, 'juin': 6, 'juni': 6,
        'julho': 7, 'july': 7, 'julio': 7, 'juillet': 7, 'juli': 7,
        'agosto': 8, 'august': 8, 'aout': 8,
        'setembro': 9, 'september': 9, 'septiembre': 9, 'septembre': 9,
        'outubro': 10, 'october': 10, 'octubre': 10, 'octobre': 10, 'oktober': 10,
        'novembro': 11, 'november': 11, 'noviembre': 11, 'novembre': 11,
        'dezembro': 12, 'december': 12, 'diciembre': 12, 'decembre': 12, 'dezember': 12,
    }
    return aliases.get(mes)


def serializar(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime('%d/%m/%Y')
    if hasattr(valor, 'item'):
        return valor.item()
    return valor


def encontrar_coluna_valor(df):
    coluna_mapeada = mapear_colunas(df.columns).get('valor')
    if coluna_mapeada is not None:
        valores_mapeados = df[coluna_mapeada].map(converter_numero_local)
        if valores_mapeados.notna().mean() >= 0.35:
            return coluna_mapeada, valores_mapeados

    palavras = ('valor', 'preço', 'preco', 'custo', 'total', 'amount', 'bill', 'despesa', 'gasto')
    candidatas = []
    for coluna in df.columns:
        serie = df[coluna].map(converter_numero_local)
        taxa_numerica = serie.notna().mean()
        nome = normalizar_cabecalho(coluna)
        pontuacao = (3 if any(palavra in nome for palavra in palavras) else 0) + taxa_numerica
        if taxa_numerica >= 0.35:
            candidatas.append((pontuacao, coluna, serie))
    if not candidatas:
        return None, None
    _, coluna, serie = max(candidatas, key=lambda item: item[0])
    return coluna, serie


def encontrar_coluna_eletricidade(colunas, campo):
    aliases = {
        'consumo_kwh': (
            'consumo kwh', 'consumo de kwh', 'kwh consumido', 'kwh consumidos',
            'energia consumida kwh', 'consumption kwh', 'kwh consumption',
            'energy consumption kwh', 'kwh used', 'consumo de energia',
            'consumo energia', 'consumo eletrico',
            'consumo de eletricidade', 'consumo',
        ),
        'total': (
            'valor total', 'total a pagar', 'valor a pagar', 'total da fatura',
            'valor da fatura', 'valor da conta', 'valor cobrado', 'total da conta',
            'total faturado', 'conta de energia', 'conta de luz', 'total amount',
            'amount due', 'invoice total', 'bill amount', 'amount', 'valor',
            'total',
        ),
    }
    nomes = [(coluna, normalizar_cabecalho(coluna)) for coluna in colunas]
    nomes = [(coluna, nome) for coluna, nome in nomes if nome]
    aliases_normalizados = [normalizar_cabecalho(alias) for alias in aliases[campo]]
    for alias in aliases_normalizados:
        coluna = next((coluna for coluna, nome in nomes if nome == alias), None)
        if coluna is not None:
            return coluna

    if campo == 'consumo_kwh':
        candidatas = [
            (coluna, nome) for coluna, nome in nomes
            if 'kwh' in nome and not any(
                termo in nome for termo in ('tarifa', 'preco', 'price', 'rate', 'valor')
            )
        ]
        if candidatas:
            return candidatas[0][0]
    else:
        candidatas = [
            (coluna, nome) for coluna, nome in nomes
            if any(termo in nome for termo in ('total', 'fatura', 'cobrado', 'amount', 'bill', 'valor'))
            and 'kwh' not in nome
            and not any(termo in nome for termo in ('por kwh', 'unitario', 'unit price', 'rate', 'tarifa'))
        ]
        if candidatas:
            return candidatas[0][0]
    return None


def ler_abas_eletricidade(arquivo):
    excel = pd.ExcelFile(BytesIO(arquivo))
    abas = {}
    for nome_aba in excel.sheet_names:
        previa = pd.read_excel(excel, sheet_name=nome_aba, header=None, nrows=12)
        melhor_linha = None
        melhor_pontuacao = 0
        for indice, linha in previa.iterrows():
            colunas = linha.dropna().tolist()
            consumo = encontrar_coluna_eletricidade(colunas, 'consumo_kwh')
            total = encontrar_coluna_eletricidade(colunas, 'total')
            pontuacao = int(consumo is not None) + int(total is not None)
            if pontuacao > melhor_pontuacao:
                melhor_linha = indice
                melhor_pontuacao = pontuacao
        cabecalho = melhor_linha if melhor_pontuacao else 0
        abas[nome_aba] = pd.read_excel(excel, sheet_name=nome_aba, header=cabecalho)
    return abas


def analisar_eletricidade(arquivo):
    total_geral = 0.0
    consumo_geral = 0.0
    linhas_analisadas = 0
    abas_analisadas = 0

    for dados in ler_abas_eletricidade(arquivo).values():
        coluna_consumo = encontrar_coluna_eletricidade(dados.columns, 'consumo_kwh')
        coluna_total = encontrar_coluna_eletricidade(dados.columns, 'total')
        if coluna_consumo is None or coluna_total is None:
            continue

        consumo = dados[coluna_consumo].map(
            lambda valor: converter_numero_local(
                re.sub(r'(?i)\s*kwh\s*$', '', str(valor).strip())
                if isinstance(valor, str) else valor
            )
        )
        total = dados[coluna_total].map(converter_numero_local)
        linhas_validas = consumo.notna() & total.notna()
        if not linhas_validas.any():
            continue

        consumo_geral += float(consumo[linhas_validas].sum())
        total_geral += float(total[linhas_validas].sum())
        linhas_analisadas += int(linhas_validas.sum())
        abas_analisadas += 1

    if not linhas_analisadas:
        raise ValueError('Não encontrei colunas de valor total e consumo em kWh na planilha.')

    consumo_arredondado = round(consumo_geral, 3)
    total_arredondado = round(total_geral, 2)
    return {
        'total': total_arredondado,
        'kwh': consumo_arredondado,
        'valor_por_kwh': round(total_arredondado / consumo_arredondado, 6)
        if consumo_arredondado else None,
        'linhas': linhas_analisadas,
        'abas': abas_analisadas,
    }


def analisar_planilha(arquivo, nome_arquivo='planilha.xlsx'):
    abas = ler_abas_flexivel(arquivo)
    mapas_abas = {nome: mapear_colunas(df.columns) for nome, df in abas.items()}
    abas_norm = {normalizar_cabecalho(nome): nome for nome in abas}
    nomes_budget = {'budget', 'presupuesto', 'orcamento', 'budget table', 'budget annuel'}
    nome_budget = next((
        nome for nome, mapeamento in mapas_abas.items()
        if {'ano', 'mes', 'categoria', 'budget'} <= mapeamento.keys() and 'valor' not in mapeamento
    ), None)
    if nome_budget is None:
        nome_budget = next((nome for normalizado, nome in abas_norm.items() if normalizado in nomes_budget), None)
    budget_df = abas.get(nome_budget) if nome_budget else None
    nomes_historico = {'historico anual', 'annual history', 'historial anual', 'historique annuel', 'annual summary'}
    nomes_premissas = {'premissas', 'assumptions', 'premisas', 'hypotheses', 'pramissen'}
    abas = {
        nome: dados for nome, dados in abas.items()
        if normalizar_cabecalho(nome) not in nomes_historico | nomes_premissas
        and nome != nome_budget
        and not (
            {'ano', 'categoria', 'valor'} <= mapas_abas[nome].keys()
            and 'mes' not in mapas_abas[nome]
        )
    }
    registros = []
    resumo_abas = []
    for nome_aba, df in abas.items():
        df = df.dropna(how='all').copy()
        colunas_mapeadas = mapear_colunas(df.columns)
        if not {'data', 'categoria', 'mes', 'ano', 'tipo'}.intersection(colunas_mapeadas):
            resumo_abas.append({'nome': nome_aba, 'linhas': len(df), 'coluna_valor': None, 'ignorada': True})
            continue
        coluna_valor, valores = encontrar_coluna_valor(df)
        if coluna_valor is None:
            resumo_abas.append({'nome': nome_aba, 'linhas': len(df), 'coluna_valor': None, 'ignorada': True})
            continue
        data_coluna = colunas_mapeadas.get('data')
        categoria_coluna = colunas_mapeadas.get('categoria')
        validos = valores.dropna()
        for indice, valor in validos.items():
            linha = {str(coluna): serializar(df.loc[indice, coluna]) for coluna in df.columns}
            linha['_valor'] = round(float(valor), 2)
            linha['_aba'] = nome_aba
            data = pd.to_datetime(df.loc[indice, data_coluna], errors='coerce') if data_coluna else pd.NaT
            linha['_ano'] = int(data.year) if not pd.isna(data) else None
            linha['_mes'] = int(data.month) if not pd.isna(data) else None
            linha['_categoria'] = str(df.loc[indice, categoria_coluna]).strip() if categoria_coluna else None
            registros.append(linha)
        resumo_abas.append({'nome': nome_aba, 'linhas': len(df), 'coluna_valor': str(coluna_valor), 'ignorada': False})

    if not registros:
        raise ValueError('Nenhuma aba possui uma coluna numérica que pareça representar valores.')

    budget_lookup = {}
    if budget_df is not None:
        budget_colunas = mapear_colunas(budget_df.columns)
        ano_coluna = budget_colunas.get('ano')
        mes_coluna = budget_colunas.get('mes')
        budget_categoria = budget_colunas.get('categoria')
        budget_valor = budget_colunas.get('budget') or budget_colunas.get('valor')
        if ano_coluna and mes_coluna and budget_categoria and budget_valor:
            for _, linha in budget_df.iterrows():
                ano = converter_numero_local(linha[ano_coluna])
                mes = converter_mes_local(linha[mes_coluna])
                valor = converter_numero_local(linha[budget_valor])
                if not pd.isna(ano) and not pd.isna(mes) and not pd.isna(valor):
                    budget_lookup[(int(ano), int(mes), str(linha[budget_categoria]).strip())] = float(valor)
    for registro in registros:
        registro['_budget'] = round(budget_lookup.get((registro['_ano'], registro['_mes'], registro['_categoria']), 0), 2)

    valores = pd.Series([registro['_valor'] for registro in registros], dtype='float64')
    media = float(valores.mean())
    mediana = float(valores.median())
    q1, q3 = valores.quantile([0.25, 0.75])
    iqr = float(q3 - q1)
    limite_alto = float(q3 + 1.5 * iqr)
    limite_baixo = max(0.0, float(q1 - 1.5 * iqr))
    for registro in registros:
        registro['classificacao'] = 'acima_da_media' if registro['_valor'] > media else 'abaixo_da_media'
        registro['fora_do_padrao'] = registro['_valor'] > limite_alto or registro['_valor'] < limite_baixo

    maior = max(registros, key=lambda item: item['_valor'])
    menor = min(registros, key=lambda item: item['_valor'])
    acima = sum(registro['_valor'] > media for registro in registros)
    fora_padrao = [registro for registro in registros if registro['fora_do_padrao']]
    budget_grupos = {}
    for registro in registros:
        chave = (registro['_ano'], registro['_mes'], registro['_categoria'])
        grupo = budget_grupos.setdefault(chave, {'realizado': 0.0, 'budget': registro['_budget']})
        grupo['realizado'] += registro['_valor']
    budget_alertas = []
    for (ano, mes, categoria), grupo in budget_grupos.items():
        if not ano or not mes or not grupo['budget']:
            continue
        realizado = round(grupo['realizado'], 2)
        budget = round(grupo['budget'], 2)
        diferenca = round(realizado - budget, 2)
        percentual = round(abs(diferenca) / budget * 100, 1)
        budget_alertas.append({'ano': ano, 'mes': mes, 'categoria': categoria, 'realizado': realizado, 'budget': budget, 'diferenca': diferenca, 'percentual': percentual, 'status': 'acima' if diferenca > 0 else 'economia'})
    budget_total = round(sum(grupo['budget'] for grupo in budget_grupos.values() if grupo['budget']), 2)
    realizado_total = round(float(valores.sum()), 2)
    budget_diferenca = round(realizado_total - budget_total, 2) if budget_total else None
    budget_percentual = round(abs(budget_diferenca) / budget_total * 100, 1) if budget_total else None
    budget_status = 'acima' if budget_diferenca and budget_diferenca > 0 else 'economia' if budget_diferenca is not None else 'indisponivel'

    ano_atual, mes_atual = datetime.now().year, datetime.now().month

    def resumir_budget_periodo(mes_periodo=None):
        categorias_periodo = {}
        for (ano_registro, mes_registro, nome_categoria), grupo in budget_grupos.items():
            if ano_registro != ano_atual or (mes_periodo is not None and mes_registro != mes_periodo):
                continue
            categoria = categorias_periodo.setdefault(nome_categoria, {'realizado': 0.0, 'budget': 0.0})
            categoria['realizado'] += grupo['realizado']
            categoria['budget'] += grupo['budget']

        realizado_periodo = round(sum(item['realizado'] for item in categorias_periodo.values()), 2)
        budget_periodo = round(sum(item['budget'] for item in categorias_periodo.values()), 2)
        diferenca_periodo = round(realizado_periodo - budget_periodo, 2)
        categorias_resumidas = []
        for nome_categoria, valores_categoria in categorias_periodo.items():
            realizado_categoria = round(valores_categoria['realizado'], 2)
            budget_categoria = round(valores_categoria['budget'], 2)
            diferenca_categoria = round(realizado_categoria - budget_categoria, 2)
            categorias_resumidas.append({
                'categoria': str(nome_categoria),
                'realizado': realizado_categoria,
                'budget': budget_categoria,
                'diferenca': diferenca_categoria,
                'percentual': round(abs(diferenca_categoria) / budget_categoria * 100, 1) if budget_categoria else None,
                'status': 'acima' if diferenca_categoria > 0 else 'economia' if budget_categoria else 'indisponivel',
            })
        categorias_resumidas.sort(key=lambda item: abs(item['diferenca']), reverse=True)
        return {
            'ano': ano_atual,
            'mes': mes_periodo,
            'realizado': realizado_periodo,
            'budget': budget_periodo,
            'diferenca': diferenca_periodo,
            'percentual': round(abs(diferenca_periodo) / budget_periodo * 100, 1) if budget_periodo else None,
            'status': 'acima' if diferenca_periodo > 0 else 'economia' if budget_periodo else 'indisponivel',
            'categorias': categorias_resumidas,
        }

    budget_periodos = {
        'ano': resumir_budget_periodo(),
        'mes': resumir_budget_periodo(mes_atual),
    }
    budget_alertas.sort(key=lambda item: abs(item['diferenca']), reverse=True)
    campos = [campo for campo in registros[0] if not campo.startswith('_') and campo not in ('classificacao', 'fora_do_padrao')]
    categoria = next((campo for campo in campos if mapear_colunas([campo]).get('categoria') == campo), None)
    grupos = []
    if categoria:
        agrupado = pd.DataFrame(registros).groupby(categoria)['_valor'].agg(['sum', 'count']).sort_values('sum', ascending=False).head(8)
        grupos = [{'nome': str(indice), 'total': round(float(linha['sum']), 2), 'quantidade': int(linha['count'])} for indice, linha in agrupado.iterrows()]

    def limpar(registro):
        return {chave: valor for chave, valor in registro.items() if not chave.startswith('_') or chave == '_aba'}

    return {
        'arquivo': secure_filename(nome_arquivo),
        'abas': resumo_abas,
        'metricas': {'total': realizado_total, 'media': round(media, 2), 'mediana': round(mediana, 2), 'quantidade': len(registros), 'acima_media': acima, 'fora_padrao': len(fora_padrao), 'budget_total': budget_total, 'budget_diferenca': budget_diferenca, 'budget_percentual': budget_percentual, 'budget_status': budget_status},
        'budget_periodos': budget_periodos,
        'budget_alertas': budget_alertas[:12],
        'maior': {'valor': maior['_valor'], 'dados': limpar(maior)},
        'menor': {'valor': menor['_valor'], 'dados': limpar(menor)},
        'fora_padrao': [{'valor': registro['_valor'], 'dados': limpar(registro)} for registro in sorted(fora_padrao, key=lambda item: item['_valor'], reverse=True)[:12]],
        'grupos': grupos,
        'colunas': campos,
        'registros': [limpar(registro) for registro in sorted(registros, key=lambda item: item['_valor'], reverse=True)[:20]],
    }

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if ADMIN_USER and session.get('usuario') == ADMIN_USER:
        return redirect(url_for('index'))
    if request.method == 'GET':
        aviso = None if ADMIN_USER and ADMIN_PASS_HASH else 'Configure ADMIN_USER e ADMIN_PASS_HASH no arquivo .env.'
        return render_template('login.html', erro=None, aviso=aviso)
    if not ADMIN_USER or not ADMIN_PASS_HASH:
        return render_template(
            'login.html',
            erro='O login ainda não foi configurado no servidor.',
            aviso=None,
        ), 503

    usuario = request.form.get('usuario', '').strip()
    senha = request.form.get('senha', '')
    try:
        senha_valida = bcrypt.checkpw(senha.encode('utf-8'), ADMIN_PASS_HASH.encode('utf-8'))
    except (TypeError, ValueError):
        senha_valida = False

    if not hmac.compare_digest(usuario, ADMIN_USER) or not senha_valida:
        return render_template('login.html', erro='Usuário ou senha inválidos.', aviso=None), 401

    session.clear()
    session['usuario'] = ADMIN_USER
    session['token_sessao'] = secrets.token_urlsafe(32)
    session.permanent = True
    _d1_run(
        'INSERT INTO sessoes (usuario, token, atualizado_em) VALUES (?, ?, CURRENT_TIMESTAMP) '
        'ON CONFLICT(usuario) DO UPDATE SET token=excluded.token, atualizado_em=excluded.atualizado_em',
        (ADMIN_USER, session['token_sessao']),
    )
    return redirect(url_for('index'))


@app.route('/logout', methods=['POST'])
def logout():
    usuario = session.get('usuario')
    if usuario:
        _d1_run('DELETE FROM sessoes WHERE usuario = ?', (usuario,))
    session.clear()
    return redirect(url_for('login'))


@app.route('/api/analisar', methods=['POST'])
def analisar_upload():
    arquivo = request.files.get('arquivo')
    extensao = os.path.splitext(arquivo.filename)[1].lower() if arquivo else ''
    if not arquivo or not arquivo.filename:
        return jsonify({'erro': 'Escolha um arquivo Excel para analisar.'}), 400
    if extensao not in ALLOWED_EXTENSIONS:
        return jsonify({'erro': 'Envie um arquivo .xlsx ou .xls.'}), 400
    try:
        conteudo = arquivo.read()
        resultado = analisar_planilha(conteudo, arquivo.filename)
        resultado['arquivo_salvo'] = _store_spreadsheet(arquivo.filename, conteudo)
        return jsonify(resultado)
    except ValueError as erro:
        return jsonify({'erro': str(erro)}), 422
    except Exception as erro:
        print(f'Erro ao analisar planilha: {erro}')
        return jsonify({'erro': 'Não foi possível ler essa planilha. Confira se ela não está corrompida.'}), 422


@app.route('/api/comparar-anos', methods=['POST'])
def comparar_anos_upload():
    arquivo = request.files.get('arquivo')
    extensao = os.path.splitext(arquivo.filename)[1].lower() if arquivo else ''
    if not arquivo or not arquivo.filename:
        return jsonify({'erro': 'Escolha um arquivo Excel para comparar.'}), 400
    if extensao not in ALLOWED_EXTENSIONS:
        return jsonify({'erro': 'Envie um arquivo .xlsx ou .xls.'}), 400
    try:
        conteudo = arquivo.read()
        abas = ler_abas_flexivel(conteudo)
        historicos = []
        for nome_aba, dados_aba in abas.items():
            colunas = mapear_colunas(dados_aba.columns)
            if {'ano', 'categoria', 'valor'} <= colunas.keys() and 'mes' not in colunas:
                historicos.append((nome_aba, dados_aba, colunas))
        if not historicos:
            raise ValueError('Não encontrei uma aba anual com colunas de ano, categoria e total.')
        nome_aba_historico, quadro, colunas = max(historicos, key=lambda item: len(item[2]))
        coluna_ano = colunas['ano']
        coluna_total = colunas['valor']
        coluna_tipo = colunas.get('tipo')
        coluna_categoria = colunas['categoria']
        quadro = quadro.copy()
        quadro['__ano'] = quadro[coluna_ano].map(converter_numero_local)
        quadro['__total'] = quadro[coluna_total].map(converter_numero_local)
        quadro = quadro.dropna(subset=['__ano', '__total', coluna_categoria])
        if quadro.empty:
            raise ValueError('A aba Historico anual não possui valores válidos.')
        quadro['__ano'] = quadro['__ano'].astype(int)
        total_por_ano = quadro.groupby('__ano')['__total'].sum().sort_index()
        tipos_por_ano = {}
        for ano, grupo in quadro.groupby('__ano'):
            tipo_valor = normalizar_cabecalho(grupo[coluna_tipo].iloc[0]) if coluna_tipo else ''
            tipos_por_ano[int(ano)] = (
                'Projecao'
                if tipo_valor.startswith(('pro', 'forecast', 'project', 'plan', 'estim'))
                else 'Realizado'
            )

        budget_por_ano = {}
        budget_por_ano_categoria = {}
        nomes_abas_budget = {'budget', 'orcamento', 'orcamentos', 'presupuesto', 'budget table'}
        for nome_budget, dados_budget in abas.items():
            colunas_budget = mapear_colunas(dados_budget.columns)
            coluna_budget_ano = colunas_budget.get('ano')
            coluna_budget_categoria = colunas_budget.get('categoria')
            coluna_budget_valor = colunas_budget.get('budget')
            if (
                coluna_budget_valor is None
                and normalizar_cabecalho(nome_budget) in nomes_abas_budget
            ):
                coluna_budget_valor = colunas_budget.get('valor')
            if not all((coluna_budget_ano, coluna_budget_categoria, coluna_budget_valor)):
                continue

            quadro_budget = dados_budget.copy()
            quadro_budget['__ano'] = quadro_budget[coluna_budget_ano].map(converter_numero_local)
            quadro_budget['__budget'] = quadro_budget[coluna_budget_valor].map(converter_numero_local)
            quadro_budget = quadro_budget.dropna(subset=['__ano', '__budget', coluna_budget_categoria])
            for _, linha in quadro_budget.iterrows():
                ano = int(linha['__ano'])
                categoria = str(linha[coluna_budget_categoria]).strip()
                valor_budget = float(linha['__budget'])
                chave = (ano, categoria)
                budget_por_ano_categoria[chave] = budget_por_ano_categoria.get(chave, 0.0) + valor_budget
                budget_por_ano[ano] = budget_por_ano.get(ano, 0.0) + valor_budget
            break

        anos = []
        anterior = None
        for ano, total in total_por_ano.items():
            variacao = None if anterior in (None, 0) else round((float(total) - anterior) / anterior * 100, 1)
            ano_int = int(ano)
            budget_ano = budget_por_ano.get(ano_int)
            percentual_budget = round(float(total) / budget_ano * 100, 1) if budget_ano else None
            anos.append({
                'ano': ano_int,
                'tipo': tipos_por_ano[ano_int],
                'total': round(float(total), 2),
                'budget': round(budget_ano, 2) if budget_ano is not None else None,
                'percentual_budget': percentual_budget,
                'variacao_percentual': variacao,
            })
            anterior = float(total)

        anos_realizados_com_budget = [
            ano for ano in anos
            if ano['tipo'] == 'Realizado' and ano['budget'] is not None
        ]
        resumo = {
            'total_gastos': round(sum(ano['total'] for ano in anos_realizados_com_budget), 2),
            'total_budget': round(sum(ano['budget'] for ano in anos_realizados_com_budget), 2),
            'percentual_budget': None,
            'anos_comparados': len(anos_realizados_com_budget),
        }
        if resumo['total_budget']:
            resumo['percentual_budget'] = round(
                resumo['total_gastos'] / resumo['total_budget'] * 100,
                1,
            )

        por_categoria = quadro.pivot_table(index=coluna_categoria, columns='__ano', values='__total', aggfunc='sum', fill_value=0)
        categorias = [
            {
                'categoria': str(categoria),
                'valores': {str(int(ano)): round(float(valor), 2) for ano, valor in linha.items()},
                'budget_por_ano': {
                    str(ano): round(budget_por_ano_categoria.get((ano, str(categoria).strip()), 0.0), 2)
                    for ano in total_por_ano.index
                },
            }
            for categoria, linha in por_categoria.iterrows()
        ]
        return jsonify({
            'arquivo_salvo': _store_spreadsheet(arquivo.filename, conteudo),
            'arquivo': secure_filename(arquivo.filename),
            'aba': nome_aba_historico,
            'anos': anos,
            'resumo': resumo,
            'categorias': categorias,
        })
    except ValueError as erro:
        return jsonify({'erro': str(erro)}), 422
    except Exception as erro:
        print(f'Erro ao comparar anos: {erro}')
        return jsonify({'erro': 'Não foi possível ler a aba Historico anual.'}), 422


@app.route('/api/comparar-planilhas', methods=['POST'])
def comparar_planilhas_upload():
    arquivos = [arquivo for arquivo in request.files.getlist('arquivos') if arquivo and arquivo.filename]
    if not arquivos:
        return jsonify({'erro': 'Escolha pelo menos duas planilhas para comparar.'}), 400
    if len(arquivos) > 12:
        return jsonify({'erro': 'Envie no máximo 12 planilhas por comparação.'}), 400
    resultados = []
    for arquivo in arquivos:
        extensao = os.path.splitext(arquivo.filename)[1].lower()
        if extensao not in ALLOWED_EXTENSIONS:
            resultados.append({'arquivo': secure_filename(arquivo.filename), 'erro': 'Formato não aceito.'})
            continue
        try:
            conteudo = arquivo.read()
            analise = analisar_planilha(conteudo, arquivo.filename)
            arquivo_salvo = _store_spreadsheet(arquivo.filename, conteudo)
            resultados.append({
                'arquivo': analise['arquivo'],
                'arquivo_salvo': arquivo_salvo,
                'metricas': analise['metricas'],
            })
        except Exception as erro:
            resultados.append({'arquivo': secure_filename(arquivo.filename), 'erro': str(erro)})
    validos = [resultado for resultado in resultados if 'metricas' in resultado]
    if len(validos) < 2:
        return jsonify({'erro': 'Não foi possível analisar pelo menos duas planilhas.', 'resultados': resultados}), 422
    return jsonify({'resultados': resultados})


@app.route('/api/analisar-eletricidade', methods=['POST'])
def analisar_eletricidade_upload():
    arquivos = [
        arquivo for arquivo in request.files.getlist('arquivos')
        if arquivo and arquivo.filename
    ]
    if not arquivos:
        arquivo = request.files.get('arquivo')
        if arquivo and arquivo.filename:
            arquivos = [arquivo]
    if not arquivos:
        return jsonify({'erro': 'Escolha pelo menos uma planilha de eletricidade.'}), 400
    if len(arquivos) > 12:
        return jsonify({'erro': 'Envie no máximo 12 planilhas por análise.'}), 400

    resultados = []
    for arquivo in arquivos:
        nome_arquivo = secure_filename(arquivo.filename) or 'planilha.xlsx'
        if os.path.splitext(nome_arquivo)[1].lower() not in ALLOWED_EXTENSIONS:
            resultados.append({'arquivo': nome_arquivo, 'erro': 'Formato não aceito.'})
            continue
        try:
            metricas = analisar_eletricidade(arquivo.read())
            resultados.append({'arquivo': nome_arquivo, **metricas})
        except ValueError as erro:
            resultados.append({'arquivo': nome_arquivo, 'erro': str(erro)})
        except Exception as erro:
            print(f'Erro ao analisar eletricidade em {nome_arquivo}: {erro}')
            resultados.append({
                'arquivo': nome_arquivo,
                'erro': 'Não foi possível ler essa planilha. Confira se ela não está corrompida.',
            })

    validos = [resultado for resultado in resultados if 'total' in resultado]
    if not validos:
        return jsonify({
            'erro': 'Nenhuma planilha pôde ser analisada.',
            'resultados': resultados,
        }), 422

    total_geral = round(sum(item['total'] for item in validos), 2)
    consumo_geral = round(sum(item['kwh'] for item in validos), 3)
    return jsonify({
        'resultados': resultados,
        'metricas': {
            'total': total_geral,
            'kwh': consumo_geral,
            'valor_por_kwh': round(total_geral / consumo_geral, 6) if consumo_geral else None,
        },
    })


@app.route('/api/planilhas', methods=['GET'])
def listar_planilhas_salvas():
    return jsonify({'planilhas': _list_spreadsheets()})


@app.route('/api/planilhas/<spreadsheet_id>/download', methods=['GET'])
def baixar_planilha_salva(spreadsheet_id):
    if not re.fullmatch(r'[a-f0-9]{32}', spreadsheet_id):
        return jsonify({'erro': 'Planilha não encontrada.'}), 404
    metadata = next((item for item in _list_spreadsheets() if item['id'] == spreadsheet_id), None)
    conteudo = _read_spreadsheet(spreadsheet_id)
    if metadata is None or conteudo is None:
        return jsonify({'erro': 'Planilha não encontrada.'}), 404
    return send_file(
        BytesIO(conteudo),
        as_attachment=True,
        download_name=metadata['filename'],
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(
        debug=os.getenv('FLASK_DEBUG', 'false').lower() == 'true' and APP_ENV != 'production',
        use_reloader=False,
        host='0.0.0.0',
        port=int(os.getenv('PORT', '5000')),
    )
