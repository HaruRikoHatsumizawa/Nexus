# ============================================================
# app.py - Sistema de Análise de Dados com Flask
# ============================================================

import os
import random
import uuid
import pandas as pd
import requests
import bcrypt
from dotenv import load_dotenv
from flask import Flask, request, jsonify, session, render_template, abort
from functools import wraps
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# Carrega as variáveis do arquivo .env
load_dotenv()

# Cria a aplicação Flask
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'fallback-chave-insegura')

# Configuração do limitador de requisições (segurança)
limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    default_limits=["100 per hour", "10 per minute"]
)

# Dicionário para controlar sessões ativas (uma por usuário)
sessoes_ativas = {}

# Carrega os usuários do .env (apenas admin por enquanto)
USUARIOS = {}
admin_user = os.getenv('ADMIN_USER')
admin_hash = os.getenv('ADMIN_PASS_HASH')
if admin_user and admin_hash:
    USUARIOS[admin_user] = admin_hash

# ============================================================
# FUNÇÃO PARA CRIAR UMA PLANILHA EXEMPLO (se não existir)
# ============================================================
def criar_planilha_exemplo():
    if os.path.exists('dados_graos.xlsx'):
        return
    categorias = ['Compra de Grãos', 'Transporte', 'Armazenagem', 'Mão de Obra',
                  'Manutenção', 'Seguros', 'Impostos', 'Comissões', 'Energia', 'Embalagens']
    dados = []
    for _ in range(500):
        dados.append({
            'Data': f'2026-{random.randint(1,8):02d}-{random.randint(1,28):02d}',
            'Categoria': random.choice(categorias),
            'Valor (R$)': round(random.uniform(1000, 95000), 2),
            'Filial': random.choice(['Brasil', 'Argentina', 'EUA', 'Ucrânia']),
            'Fornecedor': f'Fornecedor_{random.randint(1,20)}'
        })
    pd.DataFrame(dados).to_excel('dados_graos.xlsx', index=False)
    print("✅ Planilha 'dados_graos.xlsx' criada com sucesso!")

criar_planilha_exemplo()

# ============================================================
# FUNÇÃO PARA CARREGAR OS DADOS DO EXCEL
# ============================================================
def carregar_dados():
    try:
        df = pd.read_excel('dados_graos.xlsx')
        df['Valor (R$)'] = df['Valor (R$)'].fillna(0)
        return df
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        return pd.DataFrame(columns=['Categoria', 'Valor (R$)'])

# ============================================================
# GEOLOCALIZAÇÃO - BLOQUEIO POR PAÍS (APENAS BRASIL)
# ============================================================
PAISES_PERMITIDOS = ['BR']

@app.before_request
def verificar_geolocalizacao():
    # Ignora rotas estáticas e de login
    if request.endpoint in ['static', 'login']:
        return

    # Obtém o IP real do usuário
    ip = request.headers.get('X-Forwarded-For', request.remote_addr)
    if ip and ',' in ip:
        ip = ip.split(',')[0].strip()

    # Libera localhost para testes
    if ip in ['127.0.0.1', 'localhost']:
        return

    try:
        resposta = requests.get(f'http://ip-api.com/json/{ip}', timeout=3)
        dados = resposta.json()
        pais = dados.get('countryCode', '')
        if pais not in PAISES_PERMITIDOS:
            abort(403, description="Acesso negado: apenas usuários do Brasil podem acessar.")
    except Exception as e:
        print(f"⚠️ Erro na geolocalização: {e}")
        abort(403, description="Erro na verificação de localização. Acesso negado.")

# ============================================================
# DECORATOR PARA EXIGIR LOGIN E SESSÃO ÚNICA
# ============================================================
def login_obrigatorio(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return jsonify({'erro': 'Não autenticado'}), 401

        usuario = session['usuario']
        token_sessao = session.get('token_sessao')

        if usuario not in sessoes_ativas or sessoes_ativas[usuario] != token_sessao:
            session.clear()
            return jsonify({'erro': 'Sessão expirada ou login em outro dispositivo'}), 401

        return f(*args, **kwargs)
    return decorated_function

# ============================================================
# ROTAS DA APLICAÇÃO
# ============================================================

@app.route('/')
def index():
    """Página principal (frontend)"""
    return render_template('index.html')

@app.route('/login', methods=['POST'])
@limiter.limit("5 per minute")  # Segurança: 5 tentativas por minuto
def login():
    """Autenticação do usuário com verificação de senha (bcrypt)"""
    dados = request.get_json()
    usuario = dados.get('usuario')
    senha = dados.get('senha')

    if not usuario or not senha:
        return jsonify({'erro': 'Usuário e senha são obrigatórios'}), 400

    hash_armazenado = USUARIOS.get(usuario)
    if not hash_armazenado:
        return jsonify({'erro': 'Usuário ou senha inválidos'}), 401

    try:
        if bcrypt.checkpw(senha.encode('utf-8'), hash_armazenado.encode('utf-8')):
            # Gera um token único para esta sessão
            novo_token = str(uuid.uuid4())
            sessoes_ativas[usuario] = novo_token

            session['usuario'] = usuario
            session['token_sessao'] = novo_token

            return jsonify({'mensagem': 'Login realizado com sucesso', 'usuario': usuario}), 200
    except Exception as e:
        print(f"Erro ao verificar senha: {e}")

    return jsonify({'erro': 'Usuário ou senha inválidos'}), 401

@app.route('/logout', methods=['POST'])
@login_obrigatorio
def logout():
    """Encerra a sessão do usuário"""
    usuario = session.get('usuario')
    if usuario in sessoes_ativas:
        del sessoes_ativas[usuario]
    session.clear()
    return jsonify({'mensagem': 'Desconectado com sucesso'}), 200

@app.route('/api/dados', methods=['GET'])
@login_obrigatorio
def obter_dados():
    """Retorna os dados para o dashboard (maiores gastos e tabela)"""
    df = carregar_dados()
    if df.empty:
        return jsonify({'categorias': [], 'valores': [], 'tabela': []})

    gastos_por_categoria = df.groupby('Categoria')['Valor (R$)'].sum().reset_index()
    top_10 = gastos_por_categoria.nlargest(10, 'Valor (R$)')

    categorias = top_10['Categoria'].tolist()
    valores = top_10['Valor (R$)'].tolist()
    tabela = df[['Data', 'Categoria', 'Valor (R$)', 'Filial']].tail(20).to_dict(orient='records')

    return jsonify({
        'categorias': categorias,
        'valores': valores,
        'tabela': tabela
    })

@app.route('/api/chat', methods=['POST'])
@login_obrigatorio
def chat_ia():
    """Recebe uma pergunta e retorna uma resposta (com OpenAI ou fallback)"""
    dados = request.get_json()
    pergunta = dados.get('pergunta')
    if not pergunta:
        return jsonify({'erro': 'Pergunta vazia'}), 400

    df = carregar_dados()
    if df.empty:
        contexto = "Não há dados carregados."
        total_gasto = 0
        maximo = 0
        categorias_unicas = []
    else:
        total_gasto = df['Valor (R$)'].sum()
        maximo = df['Valor (R$)'].max()
        categorias_unicas = df['Categoria'].unique().tolist()
        contexto = f"Total: R$ {total_gasto:,.2f}, Maior: R$ {maximo:,.2f}, Categorias: {', '.join(categorias_unicas)}"

    # Tenta usar a OpenAI se houver chave
    openai_api_key = os.getenv('OPENAI_API_KEY')
    if openai_api_key:
        try:
            from openai import OpenAI
            cliente = OpenAI(api_key=openai_api_key)
            resposta_ia = cliente.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": f"Assistente financeiro. Contexto: {contexto}"},
                    {"role": "user", "content": pergunta}
                ],
                max_tokens=300
            )
            texto_resposta = resposta_ia.choices[0].message.content
        except Exception as e:
            print(f"Erro na OpenAI: {e}")
            texto_resposta = "Erro na API. Usando fallback."
    else:
        # Modo offline (fallback)
        texto_resposta = f"Modo offline. Dados: {len(df)} registros, {len(categorias_unicas)} categorias. Configure a OPENAI_API_KEY no .env para respostas avançadas."

    return jsonify({'resposta': texto_resposta})

# ============================================================
# INICIALIZAÇÃO DO SERVIDOR
# ============================================================
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)