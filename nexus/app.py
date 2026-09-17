# -*- coding: utf-8 -*-
"""
Sistema de Análise de Dados Financeiros para Multinacional de Grãos.
Este arquivo é o backend em Flask, responsável por:
- Leitura de planilhas Excel (pandas)
- Autenticação com bcrypt e sessão única
- Geolocalização para bloquear acessos fora do Brasil
- API para dashboard (maiores gastos)
- API para chat com IA (OpenAI ou fallback offline)
- Rate limiting para evitar ataques de força bruta
"""

import os
import random
import uuid
import pandas as pd
import requests
import bcrypt
from dotenv import load_dotenv
from flask import Flask, request, jsonify, session, render_template, abort
from functools import wraps
from datetime import datetime
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# Carrega as variáveis de ambiente do arquivo .env (seguro)
load_dotenv()

# ----------------------------------------------------------------------
# 1. INICIALIZAÇÃO DO FLASK E CONFIGURAÇÕES
# ----------------------------------------------------------------------
app = Flask(__name__)
# A chave secreta é usada para assinar os cookies de sessão.
# Ela vem do .env, garantindo que não fique exposta no código.
app.secret_key = os.getenv('SECRET_KEY', 'fallback-chave-insegura')

# Configuração do Rate Limiter (limita tentativas de login)
limiter = Limiter(
    app=app,
    key_func=get_remote_address,  # Identifica o cliente pelo IP
    default_limits=["100 per hour", "10 per minute"]  # limite geral
)

# ----------------------------------------------------------------------
# 2. BANCO DE DADOS EM MEMÓRIA (para demonstração)
#    - sessoes_ativas: guarda o token de cada usuário logado
#    - USUARIOS: dicionário com {nome: hash_da_senha}
# ----------------------------------------------------------------------
sessoes_ativas = {}

# Carrega as credenciais do .env. Em produção, use um banco de dados real.
USUARIOS = {}
admin_user = os.getenv('ADMIN_USER')
admin_hash = os.getenv('ADMIN_PASS_HASH')
if admin_user and admin_hash:
    USUARIOS[admin_user] = admin_hash
# Se quiser adicionar mais usuários fixos, pode incluir aqui:
# USUARIOS['analista'] = '$2b$12$...'

# ----------------------------------------------------------------------
# 3. FUNÇÃO PARA CRIAR UMA PLANILHA DE EXEMPLO (SE NÃO EXISTIR)
# ----------------------------------------------------------------------
def criar_planilha_exemplo():
    """
    Verifica se o arquivo dados_graos.xlsx já existe.
    Se não, gera 500 registros aleatórios para testes.
    """
    if os.path.exists('dados_graos.xlsx'):
        return  # Já existe, não faz nada

    categorias = ['Compra de Grãos', 'Transporte', 'Armazenagem', 'Mão de Obra',
                  'Manutenção', 'Seguros', 'Impostos', 'Comissões', 'Energia', 'Embalagens']
    dados = []
    for _ in range(500):
        dados.append({
            'Data': f'2026-{random.randint(1, 8):02d}-{random.randint(1, 28):02d}',
            'Categoria': random.choice(categorias),
            'Valor (R$)': round(random.uniform(1000, 95000), 2),
            'Filial': random.choice(['Brasil', 'Argentina', 'EUA', 'Ucrânia']),
            'Fornecedor': f'Fornecedor_{random.randint(1, 20)}'
        })
    df = pd.DataFrame(dados)
    df.to_excel('dados_graos.xlsx', index=False)
    print("✅ Planilha 'dados_graos.xlsx' criada com dados de exemplo!")

criar_planilha_exemplo()

# ----------------------------------------------------------------------
# 4. FUNÇÃO PARA CARREGAR OS DADOS DA PLANILHA
# ----------------------------------------------------------------------
def carregar_dados():
    """
    Lê a planilha Excel com pandas e retorna um DataFrame.
    Em caso de erro, retorna um DataFrame vazio.
    """
    try:
        df = pd.read_excel('dados_graos.xlsx')
        # Substitui valores nulos na coluna de valor por 0
        df['Valor (R$)'] = df['Valor (R$)'].fillna(0)
        return df
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        # Retorna um DataFrame vazio com as colunas esperadas
        return pd.DataFrame(columns=['Categoria', 'Valor (R$)'])

# ----------------------------------------------------------------------
# 5. GEOLOCALIZAÇÃO POR IP (BLOQUEIO POR PAÍS)
# ----------------------------------------------------------------------
PAISES_PERMITIDOS = ['BR']  # Apenas Brasil

@app.before_request
def verificar_geolocalizacao():
    """
    Executado antes de cada requisição. Verifica se o IP do cliente
    está em um país permitido. Se não estiver, retorna erro 403.
    """
    # Ignora verificação para rotas estáticas e de login (para não travar)
    if request.endpoint in ['static', 'login']:
        return

    # Obtém o IP real do cliente (considerando proxies)
    ip = request.headers.get('X-Forwarded-For', request.remote_addr)
    if ip and ',' in ip:
        ip = ip.split(',')[0].strip()

    # Se for localhost, libera (ambiente de desenvolvimento)
    if ip in ['127.0.0.1', 'localhost']:
        return

    try:
        # Consulta a API gratuita ip-api.com para obter o país
        resposta = requests.get(f'http://ip-api.com/json/{ip}', timeout=3)
        dados = resposta.json()
        pais = dados.get('countryCode', '')

        # Se o país não estiver na lista de permitidos, bloqueia
        if pais not in PAISES_PERMITIDOS:
            print(f"🚫 Acesso bloqueado para IP {ip} - País: {pais}")
            abort(403, description="Acesso negado: apenas usuários do Brasil podem acessar.")
    except Exception as e:
        # Se a API falhar, bloqueia por segurança
        print(f"⚠️ Erro na geolocalização: {e}. Bloqueando por segurança.")
        abort(403, description="Erro na verificação de localização. Acesso negado.")

# ----------------------------------------------------------------------
# 6. DECORATOR PARA EXIGIR LOGIN E SESSÃO ÚNICA
# ----------------------------------------------------------------------
def login_obrigatorio(f):
    """
    Decorator que protege rotas exigindo que o usuário esteja logado
    e que sua sessão seja a única ativa (token válido).
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Verifica se o usuário está na sessão
        if 'usuario' not in session:
            return jsonify({'erro': 'Não autenticado'}), 401

        usuario = session['usuario']
        token_sessao = session.get('token_sessao')

        # Verifica se o token atual corresponde ao último token gerado
        if usuario not in sessoes_ativas or sessoes_ativas[usuario] != token_sessao:
            # Sessão inválida: limpa e força novo login
            session.clear()
            return jsonify({'erro': 'Sessão expirada ou login em outro dispositivo'}), 401

        return f(*args, **kwargs)
    return decorated_function

# ----------------------------------------------------------------------
# 7. ROTAS DA APLICAÇÃO
# ----------------------------------------------------------------------

# Rota principal - serve a página HTML (index.html dentro de templates)
@app.route('/')
def index():
    return render_template('index.html')

# Rota de login - com rate limiting para evitar força bruta
@app.route('/login', methods=['POST'])
@limiter.limit("5 per minute")  # Máximo de 5 tentativas por minuto por IP
def login():
    """
    Recebe usuário e senha, verifica com bcrypt e inicia uma sessão.
    Gera um token único para controle de sessão única.
    """
    dados = request.get_json()
    usuario = dados.get('usuario')
    senha = dados.get('senha')

    if not usuario or not senha:
        return jsonify({'erro': 'Usuário e senha são obrigatórios'}), 400

    # Busca o hash armazenado para este usuário
    hash_armazenado = USUARIOS.get(usuario)
    if not hash_armazenado:
        # Usuário não existe (retorna genérico para não dar pistas)
        return jsonify({'erro': 'Usuário ou senha inválidos'}), 401

    # Verifica a senha usando bcrypt (comparação segura)
    try:
        senha_correta = bcrypt.checkpw(senha.encode('utf-8'), hash_armazenado.encode('utf-8'))
    except ValueError:
        return jsonify({'erro': 'Erro interno na autenticação'}), 500

    if senha_correta:
        # Gera um token único para esta sessão
        novo_token = str(uuid.uuid4())
        # Armazena o token ativo para este usuário (sobrescreve o anterior)
        sessoes_ativas[usuario] = novo_token

        # Guarda informações na sessão do Flask (cookie)
        session['usuario'] = usuario
        session['token_sessao'] = novo_token

        return jsonify({'mensagem': 'Login realizado com sucesso', 'usuario': usuario}), 200
    else:
        return jsonify({'erro': 'Usuário ou senha inválidos'}), 401

# Rota de logout
@app.route('/logout', methods=['POST'])
@login_obrigatorio
def logout():
    """
    Remove o token ativo do usuário e limpa a sessão local.
    """
    usuario = session.get('usuario')
    if usuario in sessoes_ativas:
        del sessoes_ativas[usuario]  # Remove o token, invalidando a sessão
    session.clear()  # Remove os dados do cookie
    return jsonify({'mensagem': 'Desconectado com sucesso'}), 200

# Rota para obter os dados do dashboard (maiores gastos)
@app.route('/api/dados', methods=['GET'])
@login_obrigatorio
def obter_dados():
    """
    Retorna um JSON com:
      - categorias: lista dos nomes das categorias
      - valores: lista dos totais gastos por categoria (top 10)
      - tabela: últimos 20 registros para exibição detalhada
    """
    df = carregar_dados()

    if df.empty:
        return jsonify({'categorias': [], 'valores': [], 'tabela': []})

    # Agrupa por categoria e soma os valores
    gastos_por_categoria = df.groupby('Categoria')['Valor (R$)'].sum().reset_index()
    # Seleciona as 10 categorias com maiores gastos
    top_10 = gastos_por_categoria.nlargest(10, 'Valor (R$)')

    categorias = top_10['Categoria'].tolist()
    valores = top_10['Valor (R$)'].tolist()

    # Pega as últimas 20 transações para a tabela
    tabela = df[['Data', 'Categoria', 'Valor (R$)', 'Filial']].tail(20).to_dict(orient='records')

    return jsonify({
        'categorias': categorias,
        'valores': valores,
        'tabela': tabela
    })

# Rota para o chat com IA
@app.route('/api/chat', methods=['POST'])
@login_obrigatorio
def chat_ia():
    """
    Recebe uma pergunta do usuário, monta um contexto com os dados
    e consulta a OpenAI (se chave disponível) ou usa fallback offline.
    """
    dados = request.get_json()
    pergunta = dados.get('pergunta')

    if not pergunta:
        return jsonify({'erro': 'Pergunta vazia'}), 400

    # Carrega os dados para obter estatísticas
    df = carregar_dados()
    if df.empty:
        contexto = "Não há dados carregados no sistema."
        total_gasto = 0
        maximo = 0
        categorias_unicas = []
        total_registros = 0
    else:
        total_gasto = df['Valor (R$)'].sum()
        media = df['Valor (R$)'].mean()
        maximo = df['Valor (R$)'].max()
        categorias_unicas = df['Categoria'].unique().tolist()
        total_registros = len(df)
        contexto = f"""
        Contexto dos dados financeiros da empresa de grãos:
        - Total de gastos: R$ {total_gasto:,.2f}
        - Gasto médio: R$ {media:,.2f}
        - Maior gasto: R$ {maximo:,.2f}
        - Categorias: {', '.join(categorias_unicas)}
        - Total de registros: {total_registros}
        """

    # Tenta usar a OpenAI se a chave estiver configurada
    openai_api_key = os.getenv('OPENAI_API_KEY')

    if openai_api_key:
        try:
            from openai import OpenAI
            cliente = OpenAI(api_key=openai_api_key)
            resposta_ia = cliente.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": f"Você é um assistente financeiro especializado em análise de dados. {contexto}"},
                    {"role": "user", "content": pergunta}
                ],
                max_tokens=300
            )
            texto_resposta = resposta_ia.choices[0].message.content
        except Exception as e:
            print(f"⚠️ Erro na API OpenAI: {e}. Usando fallback.")
            texto_resposta = gerar_resposta_fallback(pergunta, total_gasto, maximo, categorias_unicas, total_registros)
    else:
        # Modo offline (sem chave)
        texto_resposta = gerar_resposta_fallback(pergunta, total_gasto, maximo, categorias_unicas, total_registros)

    return jsonify({'resposta': texto_resposta})

def gerar_resposta_fallback(pergunta, total_gasto, maximo, categorias, total_registros):
    """
    Função auxiliar que gera respostas simples sem usar API externa.
    Útil para testes ou quando não há chave OpenAI.
    """
    pergunta_lower = pergunta.lower()
    if 'gasto' in pergunta_lower and 'maior' in pergunta_lower:
        return f"O maior gasto registrado é de R$ {maximo:,.2f}. Recomendo revisar essa transação para possíveis reduções."
    elif 'categoria' in pergunta_lower:
        return f"As categorias disponíveis são: {', '.join(categorias)}."
    elif 'total' in pergunta_lower:
        return f"O gasto total acumulado é de R$ {total_gasto:,.2f}."
    else:
        return f"Recebi sua pergunta sobre os dados. Atualmente temos {total_registros} registros e {len(categorias)} categorias. Configure a OPENAI_API_KEY no .env para obter respostas mais avançadas."

# ----------------------------------------------------------------------
# 8. INICIALIZAÇÃO DO SERVIDOR
# ----------------------------------------------------------------------
if __name__ == '__main__':
    # Para rodar com HTTPS (descomente as linhas abaixo e gere os certificados)
    # contexto_ssl = ('cert.pem', 'key.pem')
    # app.run(debug=True, host='0.0.0.0', port=5000, ssl_context=contexto_ssl)

    # Rodando em HTTP (recomendado para desenvolvimento local)
    app.run(debug=True, host='0.0.0.0', port=5000)