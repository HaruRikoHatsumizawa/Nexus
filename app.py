# ============================================================
# app.py - Sistema de Análise de Dados com Flask
# ============================================================

import os
import uuid
from io import BytesIO
from datetime import datetime
import pandas as pd
import requests
import bcrypt
from dotenv import load_dotenv
from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename

# Carrega as variáveis do arquivo .env
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'fallback-chave-insegura')
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024
ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
LEITURAS_ENERGIA = []


def serializar(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime('%d/%m/%Y')
    if hasattr(valor, 'item'):
        return valor.item()
    return valor


def encontrar_coluna_valor(df):
    palavras = ('valor', 'preço', 'preco', 'custo', 'total', 'amount', 'bill', 'despesa', 'gasto')
    candidatas = []
    for coluna in df.columns:
        serie = pd.to_numeric(df[coluna], errors='coerce')
        taxa_numerica = serie.notna().mean()
        nome = str(coluna).lower()
        pontuacao = (3 if any(palavra in nome for palavra in palavras) else 0) + taxa_numerica
        if taxa_numerica >= 0.35:
            candidatas.append((pontuacao, coluna, serie))
    if not candidatas:
        return None, None
    _, coluna, serie = max(candidatas, key=lambda item: item[0])
    return coluna, serie


def analisar_planilha(arquivo, nome_arquivo='planilha.xlsx'):
    abas = pd.read_excel(BytesIO(arquivo), sheet_name=None)
    budget_df = abas.get('Budget')
    if 'Historico anual' in abas:
        nome_atual = next((nome for nome in abas if str(nome).lower().startswith('despesas 2026')), None)
        if nome_atual:
            abas = {nome_atual: abas[nome_atual]}
        else:
            abas = {nome: dados for nome, dados in abas.items() if nome not in ('Historico anual', 'Premissas')}
    registros = []
    data_coluna = None
    categoria_coluna = None
    resumo_abas = []
    for nome_aba, df in abas.items():
        df = df.dropna(how='all').copy()
        coluna_valor, valores = encontrar_coluna_valor(df)
        if coluna_valor is None:
            resumo_abas.append({'nome': nome_aba, 'linhas': len(df), 'coluna_valor': None, 'ignorada': True})
            continue
        if data_coluna is None:
            data_coluna = next((coluna for coluna in df.columns if str(coluna).lower() in ('data', 'date', 'mes')), None)
        if categoria_coluna is None:
            categoria_coluna = next((coluna for coluna in df.columns if any(p in str(coluna).lower() for p in ('categoria', 'tipo', 'descrição', 'descricao'))), None)
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
        budget_colunas = {str(coluna).strip().lower(): coluna for coluna in budget_df.columns}
        ano_coluna = budget_colunas.get('ano')
        mes_coluna = budget_colunas.get('mes')
        budget_categoria = budget_colunas.get('categoria')
        budget_valor = budget_colunas.get('budget (r$)') or budget_colunas.get('budget')
        if ano_coluna and mes_coluna and budget_categoria and budget_valor:
            for _, linha in budget_df.iterrows():
                ano = pd.to_numeric(linha[ano_coluna], errors='coerce')
                mes = pd.to_numeric(linha[mes_coluna], errors='coerce')
                valor = pd.to_numeric(linha[budget_valor], errors='coerce')
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
    budget_alertas.sort(key=lambda item: abs(item['diferenca']), reverse=True)
    campos = [campo for campo in registros[0] if not campo.startswith('_') and campo not in ('classificacao', 'fora_do_padrao')]
    categoria = next((campo for campo in campos if any(p in campo.lower() for p in ('categoria', 'tipo', 'descrição', 'descricao', 'fornecedor'))), None)
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


@app.route('/api/analisar', methods=['POST'])
def analisar_upload():
    arquivo = request.files.get('arquivo')
    extensao = os.path.splitext(arquivo.filename)[1].lower() if arquivo else ''
    if not arquivo or not arquivo.filename:
        return jsonify({'erro': 'Escolha um arquivo Excel para analisar.'}), 400
    if extensao not in ALLOWED_EXTENSIONS:
        return jsonify({'erro': 'Envie um arquivo .xlsx ou .xls.'}), 400
    try:
        return jsonify(analisar_planilha(arquivo.read(), arquivo.filename))
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
        quadro = pd.read_excel(BytesIO(arquivo.read()), sheet_name='Historico anual')
        colunas = {str(coluna).strip().lower(): coluna for coluna in quadro.columns}
        coluna_ano = colunas.get('ano')
        coluna_total = colunas.get('total (r$)') or colunas.get('total')
        coluna_tipo = colunas.get('tipo')
        coluna_categoria = colunas.get('categoria')
        if not coluna_ano or not coluna_total or not coluna_categoria:
            raise ValueError('A aba Historico anual precisa ter as colunas Ano, Categoria e Total (R$).')
        quadro = quadro.copy()
        quadro['__ano'] = pd.to_numeric(quadro[coluna_ano], errors='coerce')
        quadro['__total'] = pd.to_numeric(quadro[coluna_total], errors='coerce')
        quadro = quadro.dropna(subset=['__ano', '__total', coluna_categoria])
        if quadro.empty:
            raise ValueError('A aba Historico anual não possui valores válidos.')
        quadro['__ano'] = quadro['__ano'].astype(int)
        total_por_ano = quadro.groupby('__ano')['__total'].sum().sort_index()
        anos = []
        anterior = None
        for ano, total in total_por_ano.items():
            variacao = None if anterior in (None, 0) else round((float(total) - anterior) / anterior * 100, 1)
            tipo = 'Projecao' if coluna_tipo and str(quadro.loc[quadro['__ano'] == ano, coluna_tipo].iloc[0]).lower().startswith('pro') else 'Realizado'
            anos.append({'ano': int(ano), 'tipo': tipo, 'total': round(float(total), 2), 'variacao_percentual': variacao})
            anterior = float(total)
        por_categoria = quadro.pivot_table(index=coluna_categoria, columns='__ano', values='__total', aggfunc='sum', fill_value=0)
        categorias = [
            {'categoria': str(categoria), 'valores': {str(int(ano)): round(float(valor), 2) for ano, valor in linha.items()}}
            for categoria, linha in por_categoria.iterrows()
        ]
        return jsonify({'arquivo': secure_filename(arquivo.filename), 'anos': anos, 'categorias': categorias})
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
            analise = analisar_planilha(arquivo.read(), arquivo.filename)
            resultados.append({'arquivo': analise['arquivo'], 'metricas': analise['metricas']})
        except Exception as erro:
            resultados.append({'arquivo': secure_filename(arquivo.filename), 'erro': str(erro)})
    validos = [resultado for resultado in resultados if 'metricas' in resultado]
    if len(validos) < 2:
        return jsonify({'erro': 'Não foi possível analisar pelo menos duas planilhas.', 'resultados': resultados}), 422
    return jsonify({'resultados': resultados})


@app.route('/api/energia', methods=['GET'])
def obter_energia():
    leituras = LEITURAS_ENERGIA[-500:]
    total_kwh = sum(leitura['energia_kwh'] for leitura in leituras)
    total_custo = sum(leitura['custo_brl'] for leitura in leituras)
    por_dispositivo = {}
    for leitura in leituras:
        nome = leitura['dispositivo']
        por_dispositivo.setdefault(nome, {'energia_kwh': 0, 'custo_brl': 0, 'leituras': 0})
        por_dispositivo[nome]['energia_kwh'] += leitura['energia_kwh']
        por_dispositivo[nome]['custo_brl'] += leitura['custo_brl']
        por_dispositivo[nome]['leituras'] += 1
    return jsonify({
        'leituras': leituras,
        'resumo': {
            'energia_kwh': round(total_kwh, 3),
            'custo_brl': round(total_custo, 2),
            'dispositivos': len(por_dispositivo),
            'leituras': len(leituras),
        },
        'por_dispositivo': [
            {'dispositivo': nome, **{chave: round(valor, 3) if chave == 'energia_kwh' else round(valor, 2) if chave == 'custo_brl' else valor for chave, valor in dados.items()}}
            for nome, dados in por_dispositivo.items()
        ],
    })


@app.route('/api/energia/leitura', methods=['POST'])
def receber_leitura_energia():
    dados = request.get_json(silent=True) or {}
    try:
        leitura = {
            'dispositivo': str(dados.get('dispositivo', 'ESP32'))[:80],
            'energia_kwh': round(float(dados['energia_kwh']), 4),
            'custo_brl': round(float(dados.get('custo_brl', 0)), 2),
            'potencia_w': round(float(dados.get('potencia_w', 0)), 2),
            'tensao_v': round(float(dados.get('tensao_v', 0)), 2),
            'corrente_a': round(float(dados.get('corrente_a', 0)), 3),
            'timestamp': dados.get('timestamp') or datetime.now().isoformat(timespec='seconds'),
        }
        if leitura['energia_kwh'] < 0 or leitura['custo_brl'] < 0:
            raise ValueError
    except (KeyError, TypeError, ValueError):
        return jsonify({'erro': 'Envie energia_kwh e valores numéricos válidos.'}), 400
    LEITURAS_ENERGIA.append(leitura)
    del LEITURAS_ENERGIA[:-500]
    return jsonify({'mensagem': 'Leitura recebida.', 'leitura': leitura}), 201


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)