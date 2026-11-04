import streamlit as st
import pandas as pd
import numpy as np
from geopy.geocoders import ArcGIS
from datetime import date, datetime, timedelta
import calendar
import io
import zipfile
import math
from math import radians, cos, sin, asin, sqrt
import urllib.parse
import unicodedata
import re
import requests
import random
import time
import os
import json
import sqlite3
import folium
from streamlit_folium import st_folium
import io

# ==========================================
# 1. CONFIGURAÇÕES GERAIS E CSS
# ==========================================
st.set_page_config(
    page_title="Gestor Inteligente V9.0", 
    page_icon="images.png", 
    layout="wide"
)

# Chave do Google Maps
GOOGLE_MAPS_API_KEY = "AIzaSyA_Jba5fpLURSDv9mpE6Vd8Bzt07jN48BQ"

# --- CAMINHOS DE ARQUIVO ---
CAMINHO_USUARIOS = r"C:\Users\temp\Desktop\TESTE API\PROGRAMAÇÃO\Usuarios_Permitidos.xlsx"
ARQUIVO_CACHE = "memoria_geocoding.csv"
PASTA_HISTORICO = "historico_rotas"
ARQUIVO_SESSAO = "sessao_usuario.json"  

if not os.path.exists(PASTA_HISTORICO):
    os.makedirs(PASTA_HISTORICO)

# --- PALETA DE CORES ---
CORES_HEX = [
    '#e6194b', "#faf61c", "#1a1916", '#3cb44b', '#f58231', 
    "#666968", '#46f0f0', '#f032e6', '#bcf60c', "#0d5224", 
    '#008080', "#580e86", '#9a6324', "#4363d8", '#800000'
]

# --- ESTILO CSS PERSONALIZADO ---
st.markdown("""
<style>
    div.kpi-card {
        padding: 15px;
        border-radius: 10px;
        color: white;
        text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        margin-bottom: 10px;
    }
    div.kpi-title { font-size: 14px; font-weight: bold; margin-bottom: 5px; opacity: 0.9; }
    div.kpi-value { font-size: 28px; font-weight: bold; }
    
    .bg-blue { background-color: #2980b9; }
    .bg-green { background-color: #27ae60; }
    .bg-red { background-color: #c0392b; }
    .bg-orange { background-color: #d35400; }
    
    [data-testid="stDataFrame"] { width: 100%; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. FUNÇÕES DE EMAIL & EXCEL
# ==========================================

def init_db():
    conn = sqlite3.connect('banco_rotas.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS rotas_salvas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome_rota TEXT,
                    tipo_vendedor TEXT,
                    mes_referencia TEXT,
                    qtd_dias INTEGER,
                    usuario_salvou TEXT,
                    data_salvamento TEXT,
                    df_rota_json TEXT,
                    df_sobra_json TEXT
                 )''')
    conn.commit()
    conn.close()

init_db() 

def enviar_email_outlook(destinatario, assunto, corpo_html):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        mail.Send()
        return True, "Email enviado!"
    except Exception as e:
        return False, f"Erro Outlook: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

def carregar_tabela_usuarios():
    try:
        df = pd.read_excel(CAMINHO_USUARIOS, dtype=str)
        df.columns = df.columns.str.lower().str.strip()
        required = ['email', 'senha', 'classe']
        if not all(col in df.columns for col in required):
            st.error(f"O Excel deve ter as colunas: {required}")
            return pd.DataFrame()
        for col in required:
            df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao ler Excel de usuários: {e}")
        return pd.DataFrame()

def salvar_tabela_usuarios(df_novo):
    try:
        with pd.ExcelWriter(CAMINHO_USUARIOS, engine='openpyxl') as writer:
            df_novo.to_excel(writer, index=False)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar Excel (Verifique se ele não está aberto): {e}")
        return False

# ==========================================
# 3. SISTEMA DE LOGIN SIMPLIFICADO
# ==========================================

def tela_login():
    # O st.session_state garante que os dados fiquem salvos apenas na aba de quem abriu
    if 'logged_in' not in st.session_state: 
        st.session_state['logged_in'] = False
    if 'user_data' not in st.session_state: 
        st.session_state['user_data'] = {}

    if st.session_state['logged_in']:
        user_email = st.session_state['user_data'].get('email', 'Usuario')
        user_class = st.session_state['user_data'].get('classe', 'user').lower()
        
        with st.sidebar:
            st.info(f"👤 **{user_email}**\n\nNível: {user_class.upper()}")
            
            if 'adm' in user_class:
                st.divider()
                st.markdown("### 👥 Gestão de Usuários")
                opt_gestao = st.radio("Ação:", ["Nenhuma", "Cadastrar Novo", "Excluir Usuário"])
                
                if opt_gestao == "Cadastrar Novo":
                    with st.form("form_add_user"):
                        new_email = st.text_input("Novo Email")
                        new_pass = st.text_input("Nova Senha", type="password")
                        new_class = st.selectbox("Classe", ["user", "ADM"])
                        if st.form_submit_button("💾 Salvar Usuário"):
                            df = carregar_tabela_usuarios()
                            if new_email in df['email'].values:
                                st.error("Email já cadastrado!")
                            elif new_email and new_pass:
                                new_row = pd.DataFrame({'email': [new_email], 'senha': [new_pass], 'classe': [new_class]})
                                df_final = pd.concat([df, new_row], ignore_index=True)
                                if salvar_tabela_usuarios(df_final):
                                    st.success("Usuário cadastrado!")
                                    time.sleep(1)
                                    st.rerun()
                            else:
                                st.warning("Preencha todos os campos.")
                
                elif opt_gestao == "Excluir Usuário":
                    df = carregar_tabela_usuarios()
                    if not df.empty:
                        list_users = df['email'].unique().tolist()
                        if user_email in list_users: list_users.remove(user_email)
                        to_delete = st.selectbox("Selecione para excluir:", list_users)
                        if st.button("🗑️ Confirmar Exclusão"):
                            df = df[df['email'] != to_delete]
                            if salvar_tabela_usuarios(df):
                                st.success(f"{to_delete} removido!")
                                time.sleep(1)
                                st.rerun()
                    else:
                        st.warning("Nenhum usuário para excluir.")
                st.divider()

            if st.button("🔓 Sair / Logout", use_container_width=True):
                # Limpa a memória apenas deste usuário específico, sem afetar os outros
                st.session_state.clear() 
                st.rerun()
        return True

    st.markdown("<br><br><h1 style='text-align: center;'> Acesso ao Sistema 🔐</h1>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        container = st.container(border=True)
        container.subheader("Credenciais")
        email_input = container.text_input("Email")
        senha_input = container.text_input("Senha", type="password")

        if container.button("Entrar", type="primary", use_container_width=True):
            df_users = carregar_tabela_usuarios()
            if not df_users.empty:
                user_found = df_users[(df_users['email'] == email_input) & (df_users['senha'] == senha_input)]
                if not user_found.empty:
                    st.session_state['user_data'] = user_found.iloc[0].to_dict()
                    st.session_state['logged_in'] = True
                    st.success("Acesso Permitido!")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    container.error("Email ou Senha incorretos.")
            else:
                container.error("Base de usuários não encontrada ou vazia.")
                
    return False

# ==========================================
# 4. FUNÇÕES DE PROCESSAMENTO E AUXILIARES
# ==========================================

def limpar_volume(val):
    if pd.isna(val) or val == '' or val is None: return 0.0
    try:
        if isinstance(val, str):
            v_str = re.sub(r'[^\d.,]', '', val).replace(',', '.')
            return float(v_str) if v_str else 0.0
        return float(val)
    except:
        return 0.0

# =========================================================
# 🧠 BANCO DE DADOS DE GEOCODIFICAÇÃO (SQLITE)
# =========================================================
DB_CACHE = 'memoria_geocoding.db'

# =========================================================
# 🧠 TABELA DEFINITIVA DE CLIENTES CORRIGIDOS
# =========================================================
def iniciar_tabela_correcoes():
    # Conecta ao banco principal de cache geográfico
    conn = sqlite3.connect('memoria_geocoding.db') 
    conn.execute('''
        CREATE TABLE IF NOT EXISTS clientes_corrigidos (
            Codigo TEXT PRIMARY KEY,
            Endereco_Antigo TEXT,
            Endereco_Novo TEXT,
            Bairro_Novo TEXT,
            Latitude REAL,
            Longitude REAL,
            Data_Alteracao TEXT
        )
    ''')
    conn.commit()
    conn.close()

def salvar_correcao_cliente(codigo, end_antigo, end_novo, bairro_novo, lat, lon):
    iniciar_tabela_correcoes()
    conn = sqlite3.connect('memoria_geocoding.db')
    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    conn.execute('''
        INSERT OR REPLACE INTO clientes_corrigidos 
        (Codigo, Endereco_Antigo, Endereco_Novo, Bairro_Novo, Latitude, Longitude, Data_Alteracao)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (str(codigo).strip().upper(), str(end_antigo), str(end_novo), str(bairro_novo), float(lat), float(lon), data_atual))
    
    conn.commit()
    conn.close()

def carregar_correcoes():
    iniciar_tabela_correcoes()
    conn = sqlite3.connect('memoria_geocoding.db')
    try:
        df = pd.read_sql("SELECT * FROM clientes_corrigidos", conn)
    except:
        df = pd.DataFrame()
    conn.close()
    return df

def iniciar_banco_cache():
    conn = sqlite3.connect(DB_CACHE)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS cache_geo (
            Chave TEXT PRIMARY KEY,
            Latitude TEXT,
            Longitude TEXT,
            Score TEXT,
            ID_Cliente TEXT,
            End_Correto TEXT,
            Bairro_Correto TEXT
        )
    ''')
    conn.commit()
    conn.close()

def carregar_memoria():
    iniciar_banco_cache()
    conn = sqlite3.connect(DB_CACHE)
    df = pd.read_sql("SELECT * FROM cache_geo", conn)
    conn.close()
    return df

def salvar_memoria(novo_df_cache):
    iniciar_banco_cache()
    colunas_banco = ['Chave', 'Latitude', 'Longitude', 'Score', 'ID_Cliente', 'End_Correto', 'Bairro_Correto']
    for col in colunas_banco:
        if col not in novo_df_cache.columns:
            novo_df_cache[col] = ""
            
    conn = sqlite3.connect(DB_CACHE)
    novo_df_cache.to_sql('temp_geo', conn, if_exists='replace', index=False)
    conn.execute('''
        INSERT OR REPLACE INTO cache_geo (Chave, Latitude, Longitude, Score, ID_Cliente, End_Correto, Bairro_Correto)
        SELECT Chave, Latitude, Longitude, Score, ID_Cliente, End_Correto, Bairro_Correto FROM temp_geo
    ''')
    conn.commit()
    conn.close()

def gerar_chave_cache(endereco, cidade, cep):
    e = str(endereco).strip().lower()
    c = str(cidade).strip().lower()
    z = str(cep).replace('-', '').replace('.', '').strip()
    return f"{e}|{c}|{z}"

def salvar_no_historico(df_rota, df_sobra, vendedor):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not df_rota.empty:
        df_rota.to_csv(f"{PASTA_HISTORICO}/{timestamp}_{vendedor}_ROTA.csv", index=False)
    if not df_sobra.empty:
        df_sobra.to_csv(f"{PASTA_HISTORICO}/{timestamp}_{vendedor}_SOBRA.csv", index=False)

# ==========================================
# 5. FUNÇÕES VISUAIS & GEOMETRIA
# ==========================================

def corrigir_texto_erp(texto):
    if not isinstance(texto, str): return ""
    texto = str(texto)
    correcoes = {
        'sÃ£o': 'são', 'SÃ£o': 'São', 'Ã£': 'ã', 'Ã©': 'é', 
        'Ã³': 'ó', 'Ã­': 'í', 'Ã§': 'ç', 'Ãª': 'ê', 
        'Ã¢': 'â', 'Ãµ': 'õ', 'Ãº': 'ú', 'Ã¡': 'á',
        'Ã‰': 'É', 'Ã': 'Á', 'Ã“': 'Ó'
    }
    for errado, certo in correcoes.items():
        texto = texto.replace(errado, certo)
    try: 
        texto = texto.encode('cp1252').decode('utf-8')
    except: 
        pass
    return texto.strip()

def limpar_endereco_para_geocoding(endereco):
    endereco = corrigir_texto_erp(endereco)

    if " - " in endereco:
        partes = endereco.split(" - ")
        if any(char.isdigit() for char in partes[0]): 
            endereco = partes[0]
            
    endereco = re.sub(r'[.,\- ]+$', '', endereco)
    if not isinstance(endereco, str): return ""
    
    try:
        endereco = endereco.encode('cp1252').decode('utf-8')
    except:
        pass 

    padrao_inutil = r'(?i)\b(ap|apt|apto|bl|bloco|sl|sala|cj|conjunto|casa|loja|térreo|fundos|km|frente|lado)\b.*'
    endereco = re.sub(padrao_inutil, '', endereco)
    
    nfkd_form = unicodedata.normalize('NFKD', endereco)
    endereco = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    endereco = re.sub(r'[^a-zA-Z0-9\s,-]', '', endereco)
    
    if ',' not in endereco:
        endereco = re.sub(r'([a-zA-Z])\s+(\d+)', r'\1, \2', endereco)
        
    return endereco.strip().strip(",.-")

def is_valid_sp_coord(lat, lon):
    try:
        lat, lon = float(lat), float(lon)
        return (-25.5 <= lat <= -19.0) and (-53.5 <= lon <= -44.0)
    except:
        return False

def remover_acentos(txt):
    if not txt: return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(txt)) if unicodedata.category(c) != 'Mn')

# --- NOVA BUSCA COM ARCGIS ---
def reprocessar_cliente_individual(row_id, novo_end, novo_bairro, cidade, cep):
    geo = ArcGIS(timeout=10)
    
    busca_completa = f"{novo_end}, {novo_bairro}, {cidade}, SP, Brasil"
    busca_sem_bairro = f"{novo_end}, {cidade}, SP, Brasil"
    
    try:
        loc = geo.geocode(busca_completa)
        if loc and is_valid_sp_coord(loc.latitude, loc.longitude):
            return loc.latitude, loc.longitude, "Manual"
            
        loc = geo.geocode(busca_sem_bairro)
        if loc and is_valid_sp_coord(loc.latitude, loc.longitude):
            return loc.latitude, loc.longitude, "Manual"
        
        cep_limpo = str(cep).replace('-', '').replace('.', '').strip()
        loc = geo.geocode(f"{cep_limpo}, {cidade}, SP, Brasil")
        if loc and is_valid_sp_coord(loc.latitude, loc.longitude): 
            return loc.latitude, loc.longitude, "Manual"

        return 0.0, 0.0, "Erro"
    except:
        return 0.0, 0.0, "Erro"

def corrigir_lat_lon(valor):
    try:
        if pd.isna(valor) or valor == 0: return 0.0
        s_val = str(valor).replace(',', '.')
        f_val = float(s_val)
        if f_val == 0: return 0.0
        while abs(f_val) > 180.0: f_val /= 10.0
        return f_val
    except: return 0.0

def get_semana_relativa_str(data_atual, data_inicio):
    if isinstance(data_atual, str): 
        data_atual = pd.to_datetime(data_atual, dayfirst=True).date()
    elif hasattr(data_atual, 'date'):
        data_atual = data_atual.date()
        
    if isinstance(data_inicio, str): 
        data_inicio = pd.to_datetime(data_inicio, dayfirst=True).date()
    elif hasattr(data_inicio, 'date'):
        data_inicio = data_inicio.date()

    dias_sem = ['Segunda-Feira', 'Terça-Feira', 'Quarta-Feira', 'Quinta-Feira', 'Sexta-Feira', 'Sábado', 'Domingo']
    dia_str = dias_sem[data_atual.weekday()]
    delta = (data_atual - data_inicio).days
    num_semana = (delta // 7) + 1
    return f"Semana {num_semana}", dia_str

def get_primeiro_dia_util_proximo_mes():
    hoje = date.today()
    if hoje.month == 12: proximo_mes = date(hoje.year + 1, 1, 1)
    else: proximo_mes = date(hoje.year, hoje.month + 1, 1)
    while proximo_mes.weekday() >= 5: proximo_mes += timedelta(days=1)
    return proximo_mes

def gerar_link_google_maps(subset):
    pontos = []
    for _, r in subset.iterrows():
        lat = r.get('Latitude', 0)
        lon = r.get('Longitude', 0)
        if lat and lon and float(lat) != 0 and float(lon) != 0:
            pontos.append(f"{lat},{lon}")
            
    if len(pontos) < 2: 
        return ""
        
    origin = pontos[0]
    destination = pontos[-1]
    waypoints = "|".join(pontos[1:-1])
    
    url = f"https://www.google.com/maps/dir/?api=1&origin={origin}&destination={destination}"
    if waypoints:
        url += f"&waypoints={waypoints}"
    return url

def haversine_simples(lon1, lat1, lon2, lat2):
    try:
        lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
        a = sin((lat2-lat1)/2)**2 + cos(lat1) * cos(lat2) * sin((lon2-lon1)/2)**2
        return 2 * asin(sqrt(a)) * 6371
    except: return 0.0

@st.cache_data(ttl=3600)
def get_distancia_rodoviaria_real(origem_lat, origem_lon, dest_lat, dest_lon):
    if origem_lat == 0 or dest_lat == 0: return 0.0
    coords = f"{origem_lon},{origem_lat};{dest_lon},{dest_lat}"
    url = f"http://router.project-osrm.org/route/v1/driving/{coords}?overview=false"
    try:
        r = requests.get(url, timeout=2) 
        if r.status_code == 200:
            data = r.json()
            if 'routes' in data and len(data['routes']) > 0:
                return round(data['routes'][0]['distance'] / 1000, 2)
    except: pass
    return round(haversine_simples(origem_lon, origem_lat, dest_lon, dest_lat) * 1.3, 2)

@st.cache_data(ttl=3600)
def get_rota_detalhada_rua_otimizada(lista_pontos):
    if len(lista_pontos) < 2: return []
    caminho_completo = []
    chunk_size = 20
    for i in range(0, len(lista_pontos) - 1, chunk_size - 1):
        chunk = lista_pontos[i : i + chunk_size]
        if len(chunk) < 2: continue
        coords_str = ";".join([f"{p[1]},{p[0]}" for p in chunk])
        url = f"http://router.project-osrm.org/route/v1/driving/{coords_str}?overview=full&geometries=geojson"
        try:
            r = requests.get(url, timeout=5)
            if r.status_code == 200:
                data = r.json()
                if 'routes' in data and len(data['routes']) > 0:
                    geometry = data['routes'][0]['geometry']['coordinates']
                    caminho_completo.extend([[p[1], p[0]] for p in geometry])
                    continue
        except: pass
        caminho_completo.extend(chunk)
    return caminho_completo

def haversine_logistico(lon1, lat1, lon2, lat2, cep1, cep2):
    try:
        dist_reta = haversine_simples(lon1, lat1, lon2, lat2)
        if str(cep1)[:3] != str(cep2)[:3]: return dist_reta * 2.0
        return dist_reta * 1.2
    except: return float('inf')

def otimizar_sequencia_dia(pontos_dia, col_cep):
    if isinstance(pontos_dia, pd.DataFrame):
        if pontos_dia.empty: return []
        pontos_dia = pontos_dia.to_dict('records')
    elif not pontos_dia: return []
    
    pendentes = pontos_dia.copy()
    pendentes.sort(key=lambda x: x.get('Latitude', 0), reverse=True)
    rota = [pendentes.pop(0)]
    rota[0]['Dist_Real_KM'] = 0.0
    
    while pendentes:
        curr = rota[-1]
        best_idx, min_dist = -1, float('inf')
        lat_curr, lon_curr = curr.get('Latitude', 0), curr.get('Longitude', 0)
        cep_curr = curr.get('CEP_Ref', '00000000')
        
        for i, p in enumerate(pendentes):
            lat_p, lon_p = p.get('Latitude', 0), p.get('Longitude', 0)
            cep_p = p.get('CEP_Ref', '00000000')
            
            if lat_curr != 0 and lat_p != 0:
                dist_score = haversine_logistico(lon_curr, lat_curr, lon_p, lat_p, cep_curr, cep_p)
            else:
                dist_score = 0.1 if curr.get('Cidade_Ref', '') == p.get('Cidade_Ref', '') else 100.0
            
            if str(cep_curr)[:3] != str(cep_p)[:3]: dist_score += 2 
            if dist_score < min_dist: min_dist = dist_score; best_idx = i
        
        rota.append(pendentes.pop(best_idx))
    return rota

def recalcular_distancias_dataframe(df_dia, modo_rapido=False):
    if df_dia.empty: return df_dia
    
    df_dia = df_dia.copy()
    distancias = [0.0] 
    
    for i in range(1, len(df_dia)):
        prev = df_dia.iloc[i-1]
        curr = df_dia.iloc[i]
        
        lat1, lon1 = prev.get('Latitude', 0), prev.get('Longitude', 0)
        lat2, lon2 = curr.get('Latitude', 0), curr.get('Longitude', 0)
        
        d = 0.0
        if lat1 != 0 and lat2 != 0:
             d = round(haversine_simples(lat1, lon1, lat2, lon2) * 1.3, 2) if modo_rapido else get_distancia_rodoviaria_real(lat1, lon1, lat2, lon2)
        distancias.append(d)
        
    df_dia['Dist_Real_KM'] = distancias
    return df_dia

def gerar_excel_estilizado(df_rota, df_sobra, df_ignorados):
    output = io.BytesIO()
    df_s_excel = df_sobra.copy() 
    
    if not df_s_excel.empty: 
        if 'Motivo' not in df_s_excel.columns:
            df_s_excel['Motivo'] = 'Sobra / Distância'
        else:
            df_s_excel['Motivo'] = df_s_excel['Motivo'].fillna('Sobra / Distância')
            
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_rota.empty: df_rota.to_excel(writer, sheet_name='Rota_Otimizada', index=False)
        if not df_s_excel.empty: df_s_excel.to_excel(writer, sheet_name='Sobras', index=False)
        
    return output.getvalue()

def render_cal_main(data_ini):
    st.subheader(f"📅 Selecione os Dias de Trabalho")
    ano, mes = data_ini.year, data_ini.month
    dias_selecionados = []
    for offset_mes in range(1):
        mes_atual = mes + offset_mes
        ano_atual = ano
        if mes_atual > 12: mes_atual -= 12; ano_atual += 1
        st.markdown(f"**{calendar.month_name[mes_atual]} / {ano_atual}**")
        cal = calendar.monthcalendar(ano_atual, mes_atual)
        cols = st.columns(7)
        for i, d in enumerate(['Seg','Ter','Qua','Qui','Sex','Sab','Dom']): cols[i].markdown(f"_{d}_")
        for week in cal:
            cols = st.columns(7)
            for i, day in enumerate(week):
                if day != 0:
                    d_obj = date(ano_atual, mes_atual, day)
                    if cols[i].checkbox(f"{day}", value=(i < 5 and d_obj >= data_ini), key=f"chk_{d_obj}"):
                        dias_selecionados.append(d_obj)
        st.markdown("---")
    st.info(f"✅ Dias Selecionados: **{len(dias_selecionados)}**")
    return sorted(dias_selecionados)

# --- LÓGICA ATUALIZADA PARA ORDENAR POR VOLUME CRESCENTE ---
def processar_rota_v101(df, config_visitas_dict, max_dist, datas, col_mun, col_cep, col_end, min_visitas_dia=3, modo_rapido=False, margem_extra_pct=40, maior_volume_fim=True):
    df = df.copy()
    df['Cidade_Ref'] = df[col_mun].fillna('')
    df['CEP_Ref'] = df[col_cep].astype(str).str.replace(r'\D', '', regex=True)
    df['Vol_Cidade'] = df.groupby('Cidade_Ref')['Cidade_Ref'].transform('count')
    df = df.sort_values(by=['Vol_Cidade', 'Cidade_Ref', 'Latitude'], ascending=[False, True, False])
    
    # Adicionamos uma coluna de volume numérico para garantir a soma exata
    df['Vol_Calc'] = df.get('Volume', 0).apply(limpar_volume)
    
    pool = df.to_dict('records')
    dias_gerados = [] 
    if not datas: return pd.DataFrame(), pd.DataFrame()
    datas.sort()
    data_inicio_global = datas[0]
    
    sobras = []
    blocos_temp = []
    fator_margem = 1 + (margem_extra_pct / 100.0)
    limite_absoluto = max_dist * fator_margem

    vendedor_atual = df.iloc[0]['vend_col_temp'] if not df.empty and 'vend_col_temp' in df.columns else 'Padrao'
    max_v = config_visitas_dict.get(vendedor_atual, 15)
    num_datas = len(datas)

    # 1. Agrupamento (Clusters)
    while pool and len(blocos_temp) < num_datas:
        novo_dia = {'visitas': [pool.pop(0)]}
        while len(novo_dia['visitas']) < max_v and pool:
            last = novo_dia['visitas'][-1]
            lat_curr, lon_curr = last.get('Latitude', 0), last.get('Longitude', 0)
            indices_busca = range(min(len(pool), 50))
            best_idx, min_dist = -1, float('inf')
            
            for idx in indices_busca:
                cand = pool[idx]
                lat_p, lon_p = cand.get('Latitude', 0), cand.get('Longitude', 0)
                dist = haversine_simples(lon_curr, lat_curr, lon_p, lat_p) if (lat_curr != 0 and lat_p != 0) else 100.0
                if dist < min_dist: min_dist = dist; best_idx = idx
            
            if best_idx != -1:
                if min_dist <= limite_absoluto:
                    novo_dia['visitas'].append(pool.pop(best_idx))
                else: break 
            else: break
        
        if len(novo_dia['visitas']) < min_visitas_dia:
            sobras.extend(novo_dia['visitas'])
        else:
            blocos_temp.append(novo_dia)
            
    sobras.extend(pool)
    
    # 2. Soma de Volume e Ordenação (O Pulo do Gato)
    for bloco in blocos_temp:
        bloco['vol_total'] = sum(v.get('Vol_Calc', 0) for v in bloco['visitas'])
        
    # Ordena os blocos criados do Menor Volume para o Maior Volume
    blocos_temp.sort(key=lambda x: x['vol_total'], reverse=not maior_volume_fim)
    
    # 3. Distribuição dos blocos ordenados para as datas
    for i, bloco in enumerate(blocos_temp):
        bloco['data'] = datas[i]
        dias_gerados.append(bloco)
    
    # 4. Sequência Lógica
    rota_final = []
    for dia in dias_gerados:
        visitas_otimizadas = otimizar_sequencia_dia(dia['visitas'], col_cep)
        df_dia = pd.DataFrame(visitas_otimizadas)
        df_dia['Roteiro_Data'] = dia['data']
        df_dia['Semana'], df_dia['Dia_Semana'] = get_semana_relativa_str(dia['data'], data_inicio_global)
        df_dia = recalcular_distancias_dataframe(df_dia, modo_rapido)
        df_dia['Roteiro_Ordem'] = range(1, len(df_dia) + 1)
        df_dia['Link_Rota_Google'] = gerar_link_google_maps(df_dia)
        rota_final.append(df_dia)
    
    return (pd.concat(rota_final, ignore_index=True) if rota_final else pd.DataFrame()), pd.DataFrame(sobras)

def mostrar_visualizador_operacional():
    st.title("👁️ Visualizador Operacional de Rotas")
    
    st.markdown("""
        <div style="background-color: #fff3cd; padding: 20px; border-radius: 10px; border-left: 6px solid #ffc107; margin-bottom: 20px;">
            <h4 style="color: #856404; margin-top: 0;">⚠️ Atenção: Regra de Importação</h4>
            <p style="color: #856404; font-size: 15px;">
                Para garantir a precisão do mapa, suba <b>apenas clientes que possuam data de visita programada</b>.
                Linhas com datas vazias ou textos como 'Fechamento' serão ignoradas pelo visualizador.<br>
                                Para garantir o funcionamento da Aplicação é necessario que o arquivo tenha as seguintes colunas na primeira linha da base de dados: .<br>
                <b>Código Cliente - Nome Vendedor - Nome Cliente -  Data da Visita - Rota Matriz - Endereço - Municipio - CEP - Bairro.<b>
            </p>
        </div>
    """, unsafe_allow_html=True)

    upload_view = st.file_uploader("📂 Subir Arquivo de Rota (CSV/Excel)", type=['csv', 'xlsx'], key="up_visualizador")

    if 'arquivo_atual_vis' not in st.session_state or st.session_state.get('arquivo_atual_vis') != (upload_view.name if upload_view else None):
        if 'df_view_pronto' in st.session_state:
            del st.session_state['df_view_pronto']
        st.session_state['arquivo_atual_vis'] = upload_view.name if upload_view else None

    if upload_view:
        if upload_view.name.endswith('.csv'):
            df_view_raw = pd.read_csv(upload_view)
        else:
            df_view_raw = pd.read_excel(upload_view)

        st.write("### 📊 Dados Carregados")
        st.dataframe(df_view_raw.head(), use_container_width=True)

        st.write("### ⚙️ Mapeamento de Colunas")
        colunas = df_view_raw.columns.tolist()
        
        def tentar_achar_coluna(palavras_chave):
            for col in colunas:
                if any(p.lower() in str(col).lower() for p in palavras_chave):
                    return colunas.index(col)
            return 0

        c1, c2, c3 = st.columns(3)
        col_vendedor = c1.selectbox("Vendedor:", colunas, index=tentar_achar_coluna(['vendedor', 'vend']))
        col_data_orig = c2.selectbox("Data da Visita:", colunas, index=tentar_achar_coluna(['data', 'roteiro']))
        col_id_orig = c3.selectbox("Código do Cliente:", colunas, index=tentar_achar_coluna(['cód', 'cod', 'id', 'sap', 'cnpj']))
        
        c4, c5, c6 = st.columns(3)
        col_end_orig = c4.selectbox("Endereço:", colunas, index=tentar_achar_coluna(['endereço', 'endereco', 'rua']))
        col_bairro_orig = c5.selectbox("Bairro:", colunas, index=tentar_achar_coluna(['bairro']))
        col_nome_orig = c6.selectbox("Nome do Cliente:", colunas, index=tentar_achar_coluna(['cliente', 'razão', 'razao', 'nome', 'social']))

        c7, c8, c9 = st.columns(3)
        col_cidade_orig = c7.selectbox("Cidade/Município:", colunas, index=tentar_achar_coluna(['cidade', 'municipio', 'mun']))
        col_cep_orig = c8.selectbox("CEP (Opcional):", colunas, index=tentar_achar_coluna(['cep']))
        
        idx_vol = tentar_achar_coluna(['vol', 'peso', 'qtde', 'quantidade', 'volume', 'litros'])
        col_vol_orig = c9.selectbox("Volume/Peso (Opcional):", ["(Nenhum)"] + colunas, index=idx_vol+1 if idx_vol else 0)

        # --- NOVA SELEÇÃO: ROTA MATRIZ ---
        st.write("### 📅 Organização Semanal (Opcional)")
        col_rota_matriz = st.selectbox("Coluna de Rota/Semana (Ex: 'Rota Matriz' com dados '1ªSem.1-Segunda'):", ["(Nenhuma)"] + colunas, index=tentar_achar_coluna(['rota matriz', 'semana', 'matriz']))

        st.divider()

        if st.button("🚀 Confirmar Mapeamento e Gerar Mapa", type="primary", use_container_width=True):
            with st.spinner("Estruturando dados e buscando coordenadas no satélite..."):
                df_view = df_view_raw.copy()
                
                data_convertida = pd.to_datetime(df_view[col_data_orig], errors='coerce')
                df_view = df_view[data_convertida.notna()].copy()
                df_view['Roteiro_Data'] = data_convertida[data_convertida.notna()].dt.date
                
                df_view['Codigo_Cliente'] = df_view[col_id_orig].astype(str).str.replace(r'\.0$', '', regex=True)
                df_view['Nome_Vendedor'] = df_view[col_vendedor].astype(str)
                df_view['Cliente'] = df_view[col_nome_orig].astype(str)
                df_view['CEP_Ref'] = df_view[col_cep_orig].astype(str)
                
                df_view['Endereço_Limpo'] = df_view[col_end_orig].astype(str).apply(limpar_endereco_para_geocoding)
                df_view['Bairro'] = df_view[col_bairro_orig].astype(str).apply(corrigir_texto_erp)
                df_view['Cidade_Ref'] = df_view[col_cidade_orig].astype(str).apply(corrigir_texto_erp)
                
                df_view['Volume'] = df_view[col_vol_orig] if col_vol_orig != "(Nenhum)" else 0
                df_view['Vol_Calc'] = df_view['Volume'].apply(limpar_volume)

                # --- EXTRAINDO A SEMANA E O DIA DA COLUNA ROTA MATRIZ ---
                def extrair_semana_dia(val):
                    val = str(val).strip()
                    if val == "S/ Rota" or pd.isna(val) or val == 'nan': 
                        return "Semana Indefinida", "Dia Indefinido"
                    try:
                        if "." in val and "-" in val:
                            partes = val.split('.')
                            sem = partes[0].replace('Sem', ' Semana')
                            dia = partes[1].split('-')[-1].upper()
                            return sem, dia
                        else:
                            return "Semana Única", val.upper()
                    except:
                        return "Semana Única", val.upper()

                if col_rota_matriz != "(Nenhuma)":
                    df_view['Rota_Matriz_Raw'] = df_view[col_rota_matriz]
                    df_view['Semana_Visita'] = df_view['Rota_Matriz_Raw'].apply(lambda x: extrair_semana_dia(x)[0])
                    df_view['Dia_Visita'] = df_view['Rota_Matriz_Raw'].apply(lambda x: extrair_semana_dia(x)[1])
                else:
                    df_view['Semana_Visita'] = "Semana Única"
                    df_view['Dia_Visita'] = "Dia " + df_view['Roteiro_Data'].astype(str)


                df_memoria = carregar_memoria()
                
                if 'ID_Cliente' in df_memoria.columns:
                    mem_by_id = df_memoria.drop_duplicates(subset=['ID_Cliente'], keep='last').set_index('ID_Cliente').to_dict('index')
                else:
                    mem_by_id = {}
                
                df_view['Latitude'] = 0.0
                df_view['Longitude'] = 0.0
                
                for idx, row in df_view.iterrows():
                    cod_cli = row['Codigo_Cliente']
                    if cod_cli in mem_by_id:
                        df_view.at[idx, 'Latitude'] = float(mem_by_id[cod_cli].get('Latitude', 0.0))
                        df_view.at[idx, 'Longitude'] = float(mem_by_id[cod_cli].get('Longitude', 0.0))
                        
                mask_sem_coord = (df_view['Latitude'] == 0.0) | (df_view['Latitude'].isna())
                if mask_sem_coord.any():
                    missing_indices = df_view[mask_sem_coord].index
                    novos_dados_memoria = []
                    geo = ArcGIS(timeout=10)
                    
                    st.caption(f"📡 Geocodificando {len(missing_indices)} endereços novos no Satélite...")
                    prog_bar = st.progress(0)
                    
                    for i, idx in enumerate(missing_indices):
                        row = df_view.loc[idx]
                        
                        mun_limpo = row['Cidade_Ref']
                        bairro_busca = row['Bairro']
                        if bairro_busca.lower() == 'nan': bairro_busca = ""
                        end_limpo = row['Endereço_Limpo']
                        cod_cli = row['Codigo_Cliente']
                        cep_ref = row['CEP_Ref']
                        
                        chave = gerar_chave_cache(end_limpo, mun_limpo, cep_ref)
                        
                        try:
                            if "," in end_limpo: apenas_rua = end_limpo.split(',')[0].strip()
                            else: apenas_rua = re.sub(r'\d+$', '', end_limpo).strip()

                            loc = None
                            tipo_score = 'Erro'

                            loc = geo.geocode(f"{end_limpo}, {bairro_busca}, {mun_limpo}, SP, Brasil")
                            if loc: tipo_score = 'Alta'
                            
                            if not loc:
                                loc = geo.geocode(f"{end_limpo}, {mun_limpo}, SP, Brasil")
                                if loc: tipo_score = 'Alta'
                                
                            if not loc and apenas_rua:
                                loc = geo.geocode(f"{apenas_rua}, {mun_limpo}, SP, Brasil")
                                if loc: tipo_score = 'Média'
                                
                            if not loc:
                                cep_limpo = str(cep_ref).replace('-', '').replace('.', '').strip()
                                if cep_limpo != 'nan' and cep_limpo != '':
                                    loc = geo.geocode(f"{cep_limpo}, {mun_limpo}, SP, Brasil")
                                    if loc: tipo_score = 'CEP'

                            if loc and is_valid_sp_coord(loc.latitude, loc.longitude):
                                lat_new, lon_new = loc.latitude, loc.longitude
                                df_view.at[idx, 'Latitude'] = lat_new
                                df_view.at[idx, 'Longitude'] = lon_new
                                df_view.at[idx, 'Geo_Score'] = tipo_score
                                novos_dados_memoria.append({
                                    'Chave': chave, 'Latitude': lat_new, 'Longitude': lon_new, 'Score': tipo_score,
                                    'ID_Cliente': cod_cli, 'End_Correto': end_limpo, 'Bairro_Correto': bairro_busca
                                })
                            else:
                                df_view.at[idx, 'Latitude'] = -1.0
                                df_view.at[idx, 'Longitude'] = -1.0

                        except Exception as e:
                            df_view.at[idx, 'Latitude'] = -1.0
                            df_view.at[idx, 'Longitude'] = -1.0
                            
                        prog_bar.progress((i + 1) / len(missing_indices))
                        
                    if novos_dados_memoria:
                        salvar_memoria(pd.DataFrame(novos_dados_memoria))

                st.session_state['df_view_pronto'] = df_view
                st.success("✅ Mapa processado e congelado na memória!")
                time.sleep(1)
                st.rerun()

    if 'df_view_pronto' in st.session_state:
        df_final = st.session_state['df_view_pronto']

        st.write("---")
        c_f1, c_f2 = st.columns(2)
        
        vendedores = sorted(df_final['Nome_Vendedor'].unique())
        sel_v = c_f1.selectbox("Filtrar Vendedor:", vendedores)
        df_filtro = df_final[df_final['Nome_Vendedor'] == sel_v].copy()
        
        datas_rota = sorted(df_filtro['Roteiro_Data'].unique())
        sel_d = c_f2.selectbox("Filtrar Mapa no Dia:", ["Ver Tudo"] + list(datas_rota))

        if sel_d != "Ver Tudo":
            df_mapa = df_filtro[df_filtro['Roteiro_Data'] == sel_d].copy()
        else:
            df_mapa = df_filtro.copy()

        st.subheader(f"🗺️ Mapa Operacional: {sel_v} - {sel_d}")
        
        df_mapa = df_mapa[(df_mapa['Latitude'] != 0.0) & (df_mapa['Latitude'] != -1.0) & (df_mapa['Latitude'].notna())].copy()

        if not df_mapa.empty:
            features = []
            pts_rota_osrm = []
            dias_unicos_globais = sorted(df_filtro['Roteiro_Data'].unique())
            posicoes_vistas = {}
            distancia_offset = 0.00015  

            for _, r in df_mapa.iterrows():
                lat_orig, lon_orig = float(r['Latitude']), float(r['Longitude'])
                
                chave_pos = f"{round(lat_orig, 5)}_{round(lon_orig, 5)}"
                if chave_pos in posicoes_vistas:
                    mult = posicoes_vistas[chave_pos]
                    lat = lat_orig + (distancia_offset * mult)
                    lon = lon_orig + (distancia_offset * mult)
                    posicoes_vistas[chave_pos] += 1
                else:
                    lat, lon = lat_orig, lon_orig
                    posicoes_vistas[chave_pos] = 1

                pts_rota_osrm.append([lat_orig, lon_orig])
                
                try:
                    cor_idx = dias_unicos_globais.index(r['Roteiro_Data']) % len(CORES_HEX)
                    cor_bg = CORES_HEX[cor_idx]
                except:
                    cor_bg = "#1a73e8"
                    
                id_cliente_display = str(r['Codigo_Cliente'])
                link_maps_oficial = f"https://www.google.com/maps/search/?api=1&query={lat_orig},{lon_orig}"
                
                # Define uma cor padrão caso o cor_dict não exista
                cor_padrao_vendedor = "#1a73e8"
                
                popup_content = f"""
                <div style="font-family: Arial, sans-serif; width: 300px; padding: 5px;">
                    <strong style="font-size: 14px; color: {cor_padrao_vendedor};">{r.get('Cliente', '')}</strong><br>
                    <span style="font-size: 12px; color: #666;">Cód: {r.get('Codigo_Cliente','')}</span>
                    <hr style="border: 0; border-top: 1px solid #eee; margin: 8px 0;">
                    <table width="100%" style="font-size: 12px; border-collapse: collapse; margin-bottom: 8px;">
                        <tr><td width="30%"><strong>Vendedor:</strong></td><td>{r.get('Nome_Vendedor', '')}</td></tr>
                        <tr><td><strong>Município:</strong></td><td>{r.get('Cidade_Ref','')}</td></tr>
                        <tr><td><strong>Bairro:</strong></td><td>{r.get('Bairro','')}</td></tr>
                        <tr><td><strong>CEP:</strong></td><td>{r.get('CEP_Ref','')}</td></tr>
                        <tr valign="top"><td><strong>Endereço:</strong></td><td>{r.get('Endereço_Limpo','')}</td></tr>
                    </table>
                    <div style="border-radius: 4px; overflow: hidden; border: 1px solid #ddd; margin-bottom: 10px;">
                        <img loading="lazy" src="https://maps.googleapis.com/maps/api/streetview?size=300x150&location={lat},{lon}&key={GOOGLE_MAPS_API_KEY}" width="300" height="150">
                    </div>
                    <div style="text-align: center;">
                        <a href="{link_maps_oficial}" target="_blank" style="background-color: {cor_padrao_vendedor}; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; font-size: 12px; font-weight: bold; display: inline-block; width: 90%;">📍 Ver no Google Maps</a>
                    </div>
                </div>
                """
                features.append({
                        "position": {"lat": lat, "lng": lon},
                        "title": str(r['Cliente']),
                        "color": cor_bg,
                        "label": "•",
                        "content": popup_content
                    })

            caminho_osrm_js = "[]"
            cor_linha_osrm = "#1a73e8" 
            
            if len(pts_rota_osrm) > 1 and sel_d != "Ver Tudo":
                try:
                    rota_rua = get_rota_detalhada_rua_otimizada(pts_rota_osrm)
                    coords_finais = rota_rua if rota_rua else pts_rota_osrm
                    lista_dicts = [{"lat": p[0], "lng": p[1]} for p in coords_finais]
                    caminho_osrm_js = json.dumps(lista_dicts)
                    
                    if sel_d in dias_unicos_globais:
                        cor_linha_osrm = CORES_HEX[dias_unicos_globais.index(sel_d) % len(CORES_HEX)]
                except:
                    pass

            todas_lats = [f["position"]["lat"] for f in features]
            todas_lons = [f["position"]["lng"] for f in features]
            center_lat = sum(todas_lats) / len(todas_lats)
            center_lng = sum(todas_lons) / len(todas_lons)

            google_maps_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                    #map {{ height: 100%; }} html, body {{ height: 100%; margin: 0; padding: 0; }}
                    .gm-style-iw-d {{ overflow: auto !important; max-height: 480px !important; }}
                    .custom-control {{ 
                        background-color: #fff; border: 2px solid #fff; border-radius: 2px; 
                        box-shadow: 0 2px 6px rgba(0,0,0,.3); cursor: pointer; 
                        margin-right: 10px; margin-top: 10px; padding: 5px; text-align: center; 
                    }}
                    .custom-control label {{ color: rgb(25,25,25); font-family: Roboto,Arial,sans-serif; font-size: 13px; cursor: pointer; font-weight: bold; }}
                    </style>
                    <script>
                    const data = {json.dumps(features)};
                    const routeData = {caminho_osrm_js}; 
                    
                    let map, infoWindow, marker_list = [];
                    let routePath = null;
                    
                    function initMap() {{
                        const estiloLimpo = [
                            {{ featureType: "poi", elementType: "all", stylers: [{{ visibility: "off" }}] }},
                            {{ featureType: "transit", elementType: "all", stylers: [{{ visibility: "off" }}] }}
                        ];

                        map = new google.maps.Map(document.getElementById("map"), {{
                        zoom: 12, center: {{ lat: {center_lat}, lng: {center_lng} }},
                        gestureHandling: "greedy",
                        mapTypeControl: true, 
                        mapTypeControlOptions: {{ style: google.maps.MapTypeControlStyle.DROPDOWN_MENU, position: google.maps.ControlPosition.TOP_RIGHT }},
                        styles: estiloLimpo
                        }});
                        
                        infoWindow = new google.maps.InfoWindow();
                        
                        data.forEach((point) => {{
                        const marker = new google.maps.Marker({{
                            position: point.position, map, title: point.title, optimize: true,
                            label: {{ text: point.label || '', color: 'white', fontSize: '11px', fontWeight: 'bold' }},
                            icon: {{ path: google.maps.SymbolPath.CIRCLE, fillColor: point.color, fillOpacity: 0.9, strokeWeight: 1, strokeColor: '#fff', scale: 6 }}
                        }});
                        marker.addListener("click", () => {{ infoWindow.setContent(point.content); infoWindow.open(map, marker); }});
                        marker_list.push(marker);
                        }});

                        if (routeData && routeData.length > 1) {{
                            routePath = new google.maps.Polyline({{
                                path: routeData, geodesic: true, strokeColor: "{cor_linha_osrm}", strokeOpacity: 0.8, strokeWeight: 4,
                            }});
                            routePath.setMap(map);
                        }}
                        
                        if (marker_list.length > 0) {{
                            const bounds = new google.maps.LatLngBounds();
                            marker_list.forEach(marker => bounds.extend(marker.getPosition()));
                            if(routeData && routeData.length > 1){{
                                routeData.forEach(p => bounds.extend(p));
                            }}
                            map.fitBounds(bounds);
                        }}

                        map.addListener('zoom_changed', function() {{
                            let currentZoom = map.getZoom();
                            let newScale = 6; 
                            
                            if (currentZoom >= 15) newScale = 10; 
                            else if (currentZoom >= 13) newScale = 8; 
                            else if (currentZoom <= 10) newScale = 4; 

                            marker_list.forEach(marker => {{
                                let icon = marker.getIcon();
                                icon.scale = newScale;
                                marker.setIcon(icon);
                            }});
                        }});
                        
                        createLabelControl(map, estiloLimpo);
                        createMarkerToggleControl(map);
                    }}
                    
                    function createLabelControl(map, estiloLimpo) {{
                        const controlDiv = document.createElement('div');
                        controlDiv.className = 'custom-control';
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox'; checkbox.id = 'toggle-labels'; checkbox.checked = false; 
                        const label = document.createElement('label');
                        label.htmlFor = 'toggle-labels';
                        label.appendChild(checkbox); label.appendChild(document.createTextNode(' 🏪 Ver Comércios (Google)'));
                        controlDiv.appendChild(label);
                        
                        checkbox.addEventListener('change', function() {{
                            if (this.checked) {{ map.setOptions({{ styles: [] }}); }} 
                            else {{ map.setOptions({{ styles: estiloLimpo }}); }} 
                        }});
                        map.controls[google.maps.ControlPosition.TOP_LEFT].push(controlDiv);
                    }}

                    function createMarkerToggleControl(map) {{
                        const controlDiv = document.createElement('div');
                        controlDiv.className = 'custom-control';
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox'; checkbox.id = 'toggle-markers'; checkbox.checked = true; 
                        const label = document.createElement('label');
                        label.htmlFor = 'toggle-markers';
                        label.appendChild(checkbox); label.appendChild(document.createTextNode(' 📍 Meus Clientes & Rota'));
                        controlDiv.appendChild(label);
                        
                        checkbox.addEventListener('change', function() {{
                            const isVisible = this.checked;
                            marker_list.forEach(marker => marker.setVisible(isVisible)); 
                            if (routePath) {{ routePath.setVisible(isVisible); }}
                        }});
                        map.controls[google.maps.ControlPosition.TOP_LEFT].push(controlDiv);
                    }}
                    
                    window.initMap = initMap;
                    </script>
                </head>
                <body>
                    <div id="map"></div>
                    <script src="https://maps.googleapis.com/maps/api/js?key={GOOGLE_MAPS_API_KEY}&callback=initMap&v=weekly" defer></script>
                </body>
                </html>
            """
            st.components.v1.html(google_maps_html, height=750, scrolling=False)
        else:
            st.warning("Nenhum cliente com coordenada válida encontrado para este filtro.")
            
        # =========================================================
        # NOVO: DASHBOARD INFERIOR (LISTAGEM SEMANAL / DIÁRIA)
        # =========================================================
        st.divider()
        st.markdown(f"### 📋 Resumo Estratégico: {sel_v}")
        
        # O Dashboard usa TODOS os dados do Vendedor, ignorando o filtro de "Dia Unico" do Mapa
        tot_clientes = len(df_filtro)
        tot_cidades = df_filtro['Cidade_Ref'].nunique()
        
        # Placas de KPI
        st.markdown(f"""
        <div style="display: flex; gap: 20px; margin-bottom: 20px;">
            <div style="background-color: white; padding: 15px; border-radius: 8px; box-shadow: 1px 1px 5px rgba(0,0,0,0.1); flex: 1; border-top: 4px solid #1a73e8; text-align: center;">
                <h3 style="margin: 0; color: #00144F; font-size: 28px;">{tot_clientes} <span style="font-size: 14px; color: #666; font-weight: normal;">TOTAL CLIENTES</span></h3>
            </div>
            <div style="background-color: white; padding: 15px; border-radius: 8px; box-shadow: 1px 1px 5px rgba(0,0,0,0.1); flex: 1; border-top: 4px solid #1a73e8; text-align: center;">
                <h3 style="margin: 0; color: #00144F; font-size: 28px;">{tot_cidades} <span style="font-size: 14px; color: #666; font-weight: normal;">TOTAL CIDADES</span></h3>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        semanas_unicas = sorted([sem for sem in df_filtro['Semana_Visita'].unique() if pd.notna(sem) and sem != ""])
        
        if semanas_unicas:
            # st.tabs gera a barra de navegação estilo botões
            tabs_semanas = st.tabs(semanas_unicas)
            
            for i, tab in enumerate(tabs_semanas):
                with tab:
                    sem_atual = semanas_unicas[i]
                    df_sem = df_filtro[df_filtro['Semana_Visita'] == sem_atual].copy()
                    
                    # Barra de Pesquisa Interativa do Streamlit
                    busca = st.text_input(f"🔍 Buscar cliente ou cidade...", key=f"busca_sem_{i}")
                    if busca:
                        mask_busca = df_sem['Cliente'].str.contains(busca, case=False, na=False) | df_sem['Cidade_Ref'].str.contains(busca, case=False, na=False)
                        df_sem = df_sem[mask_busca]
                        
                    # Lógica para garantir a ordem correta dos dias da semana
                    ordem_dias = {"SEGUNDA": 1, "TERÇA": 2, "QUARTA": 3, "QUINTA": 4, "SEXTA": 5, "SÁBADO": 6, "SABADO": 6, "DOMINGO": 7}
                    df_sem['Ordem_Dia'] = df_sem['Dia_Visita'].map(ordem_dias).fillna(99)
                    dias_unicos = df_sem.sort_values('Ordem_Dia')['Dia_Visita'].unique()
                    
                    for dia in dias_unicos:
                        df_dia = df_sem[df_sem['Dia_Visita'] == dia].copy()
                        qtd_visitas = len(df_dia)
                        cidades_lista = df_dia['Cidade_Ref'].dropna().unique()
                        qtd_cid = len(cidades_lista)
                        
                        # Pegando algumas cidades para aparecerem no título
                        texto_cidades = ", ".join(cidades_lista[:3]) + ("..." if qtd_cid > 3 else "")
                        
                        # Título dinâmico do Expander (a barra retrátil)
                        titulo_barra = f"🗓️ {dia}  |  {qtd_visitas} visitas  |  📍 {qtd_cid} cid. ({texto_cidades})"
                        
                        # st.expander cria a barra azul escuro que abre e fecha
                        with st.expander(titulo_barra):
                            
                            colunas_desejadas = ['Codigo_Cliente', 'Cliente', 'Endereço_Limpo', 'Bairro', 'Cidade_Ref', 'CEP_Ref', 'Roteiro_Data']
                            colunas_presentes = [c for c in colunas_desejadas if c in df_dia.columns]
                            
                            df_exibir = df_dia[colunas_presentes].rename(columns={
                                'Codigo_Cliente': 'Cód.',
                                'Cliente': 'Razão Social',
                                'Endereço_Limpo': 'Endereço',
                                'Cidade_Ref': 'Município',
                                'CEP_Ref': 'CEP',
                                'Roteiro_Data': 'Data da Visita'
                            })
                            
                            # Tabela do Pandas formatada
                            st.dataframe(df_exibir, hide_index=True, use_container_width=True)
                            
# --- Função Visualizador de Carteiras ---
def mostrar_visualizador_carteiras():
    st.title("🗺️ Visualizador Geral de Carteiras")

    st.markdown("""
        <div style="background-color: #fff3cd; padding: 20px; border-radius: 10px; border-left: 6px solid #ffc107; margin-bottom: 20px;">
            <h4 style="color: #856404; margin-top: 0;">⚠️ Atenção: Regra de Importação</h4>
            <p style="color: #856404; font-size: 15px;">
                Para garantir o funcionamento da Aplicação é necessario que o arquivo tenha as seguintes colunas na primeira linha da base de dados: .<br>
                <b>Código Cliente - Nome Vendedor - Nome Cliente - Endereço - Municipio - CEP - Bairro.<b>
            </p>
        </div>
    """, unsafe_allow_html=True)

    if 'cart_etapa' not in st.session_state: st.session_state['cart_etapa'] = 'config'
    if 'df_cart_raw' not in st.session_state: st.session_state['df_cart_raw'] = None
    if 'df_cart_final' not in st.session_state: st.session_state['df_cart_final'] = None
    if 'cart_reset_count' not in st.session_state: st.session_state['cart_reset_count'] = 0

    with st.sidebar:
        st.divider()
        if st.button("🔄 Reiniciar Visualizador de Carteiras", use_container_width=True):
            st.session_state['cart_etapa'] = 'config'
            st.session_state['df_cart_raw'] = None
            st.session_state['df_cart_final'] = None
            st.session_state['cart_sheet_selected'] = None 
            st.session_state['cart_reset_count'] += 1      
            st.rerun()
            
        st.divider()
        st.markdown("### ⚙️ Atualização Rápida")
        if st.button("🔄 Recarregar Excel (Puxar alterações)", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.session_state['df_cart_raw'] = None
            st.session_state['df_cart_final'] = None
            st.session_state['cart_etapa'] = 'config'
            st.rerun()

    if st.session_state['cart_etapa'] == 'config':
        st.markdown("Importe a base de dados contendo toda a carteira da empresa (Excel ou CSV).")
        modo_leitura = st.sidebar.radio(
            "Como deseja carregar o arquivo?",
            ["Modo Upload (📂)", "Modo Caminho Fixo (Atualiza via botão)"]
        )
        
        uploaded = None

        if modo_leitura == "Modo Caminho Fixo (Atualiza via botão)":
            caminho_input = st.text_input("Cole o caminho completo do arquivo (Ex: C:\\Pasta\\planilha.xlsx):", 
                value=st.session_state.get('caminho_fixo', ""))
            if caminho_input:
                import os
                if os.path.exists(caminho_input):
                    uploaded = caminho_input
                    st.session_state['caminho_fixo'] = caminho_input 
                    st.success("✅ Conectado! Se alterar o Excel, basta clicar em 'Recarregar Dados' ali do lado.")
                else:
                    st.error("⚠️ Arquivo não encontrado. Verifique o caminho digitado.")
        else:
            st.info("Aviso: Se alterar o Excel no seu PC, você precisará fechar no 'X' abaixo e puxar o arquivo novamente.")
            uploaded = st.file_uploader("📂 Importar Tabela da Carteira", type=["xlsx", "csv"], 
                                        key=f"up_cart_{st.session_state['cart_reset_count']}")
        
        if uploaded is not None:
            import os
            nome_arquivo = os.path.basename(uploaded) if isinstance(uploaded, str) else uploaded.name
            
            if nome_arquivo.endswith('.csv'):
                st.session_state['df_cart_raw'] = pd.read_csv(uploaded, dtype=str)
                aba_selecionada = "CSV"
                linha_cabecalho = 1
            else:
                xls = pd.ExcelFile(uploaded)
                sheet_names = xls.sheet_names
                
                if st.session_state.get('cart_sheet_selected') != nome_arquivo:
                    st.divider()
                    st.markdown("### ⚙️ Detalhes do Arquivo Excel")
                    col_aba, col_linha = st.columns(2)
                    aba_selecionada = col_aba.selectbox("Em qual aba estão os dados?", sheet_names, key="cart_aba")
                    linha_cabecalho = col_linha.number_input("Em qual linha começa o cabeçalho?", min_value=1, value=1, step=1, key="cart_linha")
                    
                    if st.button("📥 Confirmar e Carregar Dados", type="primary"):
                        st.session_state['df_cart_raw'] = pd.read_excel(uploaded, sheet_name=aba_selecionada, header=(linha_cabecalho - 1), dtype=str)
                        st.session_state['cart_sheet_selected'] = nome_arquivo
                        st.rerun()
                        
            if st.session_state['df_cart_raw'] is not None:
                df_raw = st.session_state['df_cart_raw']
                cols = df_raw.columns.tolist()
                
                def achar_coluna(palavras_chave):
                    for col in cols:
                        if any(p.lower() in str(col).lower() for p in palavras_chave): return cols.index(col)
                    return 0
                
                st.divider()
                st.markdown("### ⚙️ Mapeamento de Colunas")
                
                cols_opcionais = ["Não mapear"] + list(cols)
                
                def idx_opc(palavras):
                    idx = achar_coluna(palavras)
                    return idx + 1 if idx is not None and idx > 0 else 0

                col1, col2 = st.columns(2)
                
                with col1:
                    c_cod = st.selectbox("CÓDIGO DO CLIENTE", cols, index=achar_coluna(['cód', 'cod', 'id', 'sap']), key="c_cart_cod")
                    c_nome = st.selectbox("NOME DO CLIENTE", cols, index=achar_coluna(['cliente', 'razão', 'nome', 'razao Social']), key="c_cart_nome")
                    c_vendedor_atual = st.selectbox("VENDEDOR", cols, index=achar_coluna(['vendedor', 'vend']), key="c_cart_vend")
                    c_bairro = st.selectbox("BAIRRO", cols, index=achar_coluna(['bairro']), key="c_cart_bairro")
                    c_media = st.selectbox("MÉDIA 3 MESES (Opcional)", cols_opcionais, index=idx_opc(['média', 'media', '3 meses', 'valor']), key="c_cart_media")
                
                with col2:
                    c_cep = st.selectbox("CEP", cols, index=achar_coluna(['cep']), key="c_cart_cep")
                    c_end = st.selectbox("ENDEREÇO", cols, index=achar_coluna(['endereço', 'endereco', 'rua']), key="c_cart_end")
                    c_municipio = st.selectbox("MUNICÍPIO", cols, index=achar_coluna(['cidade', 'município', 'mun']), key="c_cart_mun")
                    c_status = st.selectbox("STATUS (Opcional)", cols_opcionais, index=idx_opc(['status', 'situação', 'situacao']), key="c_cart_status")
                
                st.write("##")
                if st.button("🚀 INICIAR MAPEAMENTO DA CARTEIRA", type="primary", use_container_width=True):
                    st.session_state['cart_cols'] = {
                        'cod': c_cod, 
                        'nome': c_nome, 
                        'cep': c_cep, 
                        'end': c_end, 
                        'bairro': c_bairro, 
                        'mun': c_municipio, 
                        'vend': c_vendedor_atual,
                        'media': c_media if c_media != "Não mapear" else None,
                        'status': c_status if c_status != "Não mapear" else None
                    }
                    st.session_state['cart_etapa'] = 'proc'
                    st.rerun()

    elif st.session_state['cart_etapa'] == 'proc':
        st.title("📡 Cruzando base com a Tabela de Correções e Satélite...")
        df = st.session_state['df_cart_raw'].copy()
        c = st.session_state['cart_cols']
        
        df['LATITUDE'] = 0.0
        df['LONGITUDE'] = 0.0
        df['TIPO_DE_LOCALIZACAO'] = 'Erro'
        
        df_correcoes = carregar_correcoes()
        correcoes_dict = {}
        if not df_correcoes.empty:
            correcoes_dict = df_correcoes.set_index('Codigo').to_dict('index')

        df_memoria = carregar_memoria()
        mem_by_id = {}
        if 'ID_Cliente' in df_memoria.columns:
            df_memoria['ID_Limp'] = df_memoria['ID_Cliente'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
            mem_by_id = df_memoria.drop_duplicates(subset=['ID_Limp'], keep='last').set_index('ID_Limp').to_dict('index')
        
        missing_indices = []
        for idx, row in df.iterrows():
            cod_cli = str(row.get(c['cod'], '')).replace('.0', '').strip().upper()
            end_atual_bruto = row.get(c['end'], '')
            end_atual_limpo = limpar_endereco_para_geocoding(end_atual_bruto) if pd.notna(end_atual_bruto) else ""
            
            if cod_cli in correcoes_dict:
                df.at[idx, 'LATITUDE'] = float(correcoes_dict[cod_cli]['Latitude'])
                df.at[idx, 'LONGITUDE'] = float(correcoes_dict[cod_cli]['Longitude'])
                df.at[idx, 'TIPO_DE_LOCALIZACAO'] = "Manual (Corrigido)"
                df.at[idx, c['end']] = correcoes_dict[cod_cli]['Endereco_Novo']
                df.at[idx, c['bairro']] = correcoes_dict[cod_cli]['Bairro_Novo']
            elif cod_cli in mem_by_id and float(mem_by_id[cod_cli].get('Latitude', 0)) != 0:
                end_memoria = str(mem_by_id[cod_cli].get('End_Correto', '')).strip()
                if end_atual_limpo == end_memoria or end_memoria == "":
                    df.at[idx, 'LATITUDE'] = float(mem_by_id[cod_cli]['Latitude'])
                    df.at[idx, 'LONGITUDE'] = float(mem_by_id[cod_cli]['Longitude'])
                    df.at[idx, 'TIPO_DE_LOCALIZACAO'] = mem_by_id[cod_cli]['Score'] + " (Memória)"
                    if mem_by_id[cod_cli].get('End_Correto'): df.at[idx, c['end']] = mem_by_id[cod_cli]['End_Correto']
                    if mem_by_id[cod_cli].get('Bairro_Correto'): df.at[idx, c['bairro']] = mem_by_id[cod_cli]['Bairro_Correto']
                else:
                    missing_indices.append(idx)
            else:
                missing_indices.append(idx)
                
        if missing_indices:
            import time
            try:
                from geopy.geocoders import ArcGIS
                geo = ArcGIS(timeout=10)
            except Exception:
                pass
            novos_dados_memoria = []
            st.caption(f"Buscando {len(missing_indices)} novos clientes no satélite...")
            prog_bar = st.progress(0)
            
            for i, idx in enumerate(missing_indices):
                row = df.loc[idx]
                cep = str(row[c['cep']]).strip().replace("-","").replace(".","")
                bairro = corrigir_texto_erp(row[c['bairro']]) if pd.notna(row[c['bairro']]) else ""
                municipio = corrigir_texto_erp(row[c['mun']]) if pd.notna(row[c['mun']]) else ""
                end_limpo = limpar_endereco_para_geocoding(row[c['end']])
                cod_cli = str(row.get(c['cod'], '')).replace('.0', '').strip().upper()
                chave = gerar_chave_cache(end_limpo, municipio, cep)
                
                try:
                    busca_completa = f"{end_limpo}, {bairro}, {municipio}, {cep}, SP, Brasil"
                    loc = geo.geocode(busca_completa)
                    
                    if loc and is_valid_sp_coord(loc.latitude, loc.longitude):
                        lat, lon = loc.latitude, loc.longitude
                        df.at[idx, 'LATITUDE'], df.at[idx, 'LONGITUDE'], df.at[idx, 'TIPO_DE_LOCALIZACAO'] = lat, lon, "Exata"
                        novos_dados_memoria.append({'Chave': chave, 'Latitude': lat, 'Longitude': lon, 'Score': "Exata", 'ID_Cliente': cod_cli, 'End_Correto': end_limpo, 'Bairro_Correto': bairro})
                    else:
                        df.at[idx, 'LATITUDE'], df.at[idx, 'LONGITUDE'], df.at[idx, 'TIPO_DE_LOCALIZACAO'] = -1.0, -1.0, 'Erro (Tratamento Manual Necessário)'
                except:
                    df.at[idx, 'LATITUDE'], df.at[idx, 'LONGITUDE'], df.at[idx, 'TIPO_DE_LOCALIZACAO'] = -1.0, -1.0, 'Erro'
                    
                prog_bar.progress((i + 1) / len(missing_indices))
                
            if novos_dados_memoria:
                salvar_memoria(pd.DataFrame(novos_dados_memoria))

        df['VENDEDOR'] = df[c['vend']].fillna("SEM VENDEDOR")
        st.session_state['df_cart_final'] = df
        st.session_state['cart_etapa'] = 'dash'
        st.rerun()

    elif st.session_state['cart_etapa'] == 'dash':
        df = st.session_state['df_cart_final']
        c = st.session_state['cart_cols']

        if 'clientes_selecionados' not in st.session_state:
            st.session_state['clientes_selecionados'] = []

        def callback_adicionar(cod):
            if cod not in st.session_state['clientes_selecionados']:
                st.session_state['clientes_selecionados'].append(cod)
                st.session_state['seletor_clientes_export'] = list(st.session_state['clientes_selecionados'])

        def callback_limpar():
            st.session_state['clientes_selecionados'] = []
            st.session_state['seletor_clientes_export'] = []

        def callback_atualizar_multiselect():
            st.session_state['clientes_selecionados'] = list(st.session_state.get('seletor_clientes_export', []))

        df_mapa = df[(df['LATITUDE'] != 0.0) & (df['LATITUDE'] != -1.0) & (df['LATITUDE'].notna())].copy()

        if df_mapa.empty:
            st.error("Nenhum cliente possui coordenadas válidas para ser exibido.")
            st.stop()

        vendedores_unicos = sorted(df_mapa['VENDEDOR'].dropna().unique())
        col_f1, col_f2 = st.columns([5, 1])
        with col_f1:
            vendedores_selecionados = st.multiselect("Filtrar por Vendedor:", vendedores_unicos, default=vendedores_unicos)
        with col_f2:
            st.markdown(f'<div class="kpi-card bg-blue"><div class="kpi-title">Exibindo</div><div class="kpi-value">{len(df_mapa[df_mapa["VENDEDOR"].isin(vendedores_selecionados)])}</div></div>', unsafe_allow_html=True)

        df_mapa = df_mapa[df_mapa['VENDEDOR'].isin(vendedores_selecionados)]
        vendedores_ativos = df_mapa['VENDEDOR'].unique()
        cor_dict = {vend: CORES_HEX[i % len(CORES_HEX)] for i, vend in enumerate(vendedores_ativos)}

        center_lat = df_mapa['LATITUDE'].mean() if not df_mapa.empty else -23.5505
        center_lng = df_mapa['LONGITUDE'].mean() if not df_mapa.empty else -46.6333

        mapa = folium.Map(
            location=[center_lat, center_lng],
            zoom_start=11,
            tiles='https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}',
            attr='Google Maps'
        )

        for _, row in df_mapa.iterrows():
            lat, lon = row['LATITUDE'], row['LONGITUDE']
            cod_cliente = str(row.get(c['cod'], ''))
            cor_pino = cor_dict.get(row['VENDEDOR'], "#000000")
            link_maps_oficial = f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"
            
            linha_media = ""
            if c.get('media'):
                valor_bruto = row.get(c['media'], '')
                try:
                    valor_num = float(valor_bruto)
                    media_formatada = f"{valor_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except (ValueError, TypeError):
                    media_formatada = str(valor_bruto) if valor_bruto != '' else "-"
                linha_media = f"<tr><td><strong>Média 3 Meses:</strong></td><td>{media_formatada}</td></tr>"

            linha_status = ""
            if c.get('status'):
                status_val = row.get(c['status'], '')
                status_texto = str(status_val) if status_val != '' else "-"
                linha_status = f"<tr><td><strong>Status:</strong></td><td>{status_texto}</td></tr>"

            is_selected = cod_cliente in st.session_state['clientes_selecionados']

            popup_content = f"""
            <div style="font-family: Arial, sans-serif; width: 280px; padding: 5px;">
                <strong style="font-size: 14px; color: {cor_pino};">{row.get(c['nome'],'')}</strong><br>
                <span style="font-size: 12px; color: #666;">Cód: {cod_cliente}</span>
                <hr style="border: 0; border-top: 1px solid #eee; margin: 8px 0;">
                <table width="100%" style="font-size: 12px; border-collapse: collapse; margin-bottom: 8px;">
                    <tr><td width="35%"><strong>Vendedor:</strong></td><td>{row['VENDEDOR']}</td></tr>
                    <tr><td><strong>Município:</strong></td><td>{row.get(c['mun'],'')}</td></tr>
                    <tr><td><strong>Bairro:</strong></td><td>{row.get(c['bairro'],'')}</td></tr>
                    <tr><td><strong>CEP:</strong></td><td>{row.get(c['cep'],'')}</td></tr>
                    <tr valign="top"><td><strong>Endereço:</strong></td><td>{row.get(c['end'],'')}</td></tr>
                    {linha_media}
                    {linha_status}
                </table>
                <div style="border-radius: 4px; overflow: hidden; border: 1px solid #ddd; margin-bottom: 8px;">
                    <img loading="lazy" src="https://maps.googleapis.com/maps/api/streetview?size=280x140&location={lat},{lon}&key={GOOGLE_MAPS_API_KEY}" width="280" height="140">
                </div>
                <div style="text-align: center;">
                    <a href="{link_maps_oficial}" target="_blank" style="background-color: {cor_pino}; color: white; padding: 6px 12px; text-decoration: none; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block; width: 90%;">📍 Ver no Google Maps</a>
                </div>
            </div>
            """

            if is_selected:
                svg_estrela_dourada = """
                <div style="width: 30px; height: 42px; filter: drop-shadow(0px 2px 4px rgba(0,0,0,0.4));">
                    <svg width="30" height="42" viewBox="0 0 30 42" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M15 0C6.71573 0 0 6.71573 0 15C0 26.25 15 42 15 42C15 42 30 26.25 30 15C30 6.71573 23.2843 0 15 0Z" fill="#D3AF37" stroke="#111111" stroke-width="1"/>
                        <polygon points="15,5 18.09,11.26 25,12.27 20,17.14 21.18,24.02 15,20.77 8.82,24.02 10,17.14 5,12.27 11.91,11.26" fill="#FFFFFF"/>
                    </svg>
                </div>
                """
                folium.Marker(
                    location=[lat, lon],
                    tooltip=cod_cliente,
                    popup=folium.Popup(popup_content, max_width=300),
                    icon=folium.DivIcon(
                        html=svg_estrela_dourada,
                        icon_size=(30, 42),
                        icon_anchor=(15, 42)
                    )
                ).add_to(mapa)
            else:
                folium.CircleMarker(
                    location=[lat, lon],
                    radius=6,
                    color='white',
                    weight=1,
                    fill=True,
                    fill_color=cor_pino,
                    fill_opacity=0.9,
                    tooltip=cod_cliente,
                    popup=folium.Popup(popup_content, max_width=300)
                ).add_to(mapa)

        mapa_dados = st_folium(
            mapa, 
            height=750, 
            use_container_width=True, 
            returned_objects=["last_object_clicked_tooltip"],
            key="mapa_carteiras_v8"
        )

        if mapa_dados and mapa_dados.get("last_object_clicked_tooltip"):
            cod_clicado = str(mapa_dados["last_object_clicked_tooltip"])
            df_cli = df_mapa[df_mapa[c['cod']].astype(str) == cod_clicado]
            
            if not df_cli.empty:
                nome_cli = df_cli[c['nome']].values[0]
                vend_cli = df_cli['VENDEDOR'].values[0]
                
                st.markdown("---")
                c_inf, c_btn = st.columns([3, 1])
                
                with c_inf:
                    st.info(f"📍 **Cliente Focado no Mapa:** `{cod_clicado}` — **{nome_cli}** ({vend_cli})")
                    
                with c_btn:
                    if cod_clicado in st.session_state['clientes_selecionados']:
                        st.success("✅ Já está na Lista")
                    else:
                        st.button(
                            "➕ Adicionar à Lista", 
                            type="primary", 
                            use_container_width=True, 
                            key=f"btn_add_{cod_clicado}",
                            on_click=callback_adicionar,
                            args=(cod_clicado,)
                        )

        if st.session_state['clientes_selecionados']:
            st.divider()
            st.markdown("### 📥 Clientes Selecionados para Exportação")
            
            if "seletor_clientes_export" not in st.session_state:
                st.session_state["seletor_clientes_export"] = list(st.session_state['clientes_selecionados'])

            st.multiselect(
                "Gerenciar lista (Clique no 'X' na etiqueta para remover um cliente específico):", 
                options=st.session_state['clientes_selecionados'], 
                key="seletor_clientes_export",
                on_change=callback_atualizar_multiselect
            )
            
            st.button(
                "🗑️ Limpar Toda a Seleção", 
                key="btn_limpar_tudo_export",
                on_click=callback_limpar
            )
            
            df_selecionados = df_mapa[df_mapa[c['cod']].astype(str).isin(st.session_state['clientes_selecionados'])].copy()
            df_selecionados = df_selecionados.loc[:, ~df_selecionados.columns.duplicated()]
            
            raw_cols = [c.get('cod'), c.get('nome'), 'VENDEDOR', c.get('mun'), c.get('bairro'), c.get('cep'), c.get('end')]
            if c.get('media'): raw_cols.append(c['media'])
            if c.get('status'): raw_cols.append(c['status'])
            
            colunas_export = []
            for col in raw_cols:
                if col and col in df_selecionados.columns and col not in colunas_export:
                    colunas_export.append(col)
            
            df_final_export = df_selecionados[colunas_export]
            st.dataframe(df_final_export, use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final_export.to_excel(writer, index=False, sheet_name='Selecionados')
            
            st.download_button(
                label="🚀 Baixar Arquivo Excel",
                data=buffer.getvalue(),
                file_name="clientes_selecionados.xlsx",
                mime="application/vnd.ms-excel",
                type="primary",
                key="btn_baixar_excel_export"
            )            
# ==========================================
# 6. APLICAÇÃO PRINCIPAL
# ==========================================

st.sidebar.title("Gestor Inteligente")
nav = st.sidebar.selectbox("Navegação:", [
    "🚀 Roteirizador", 
    "👁️ Visualizador Operacional",
    "🗺️ Visualizar Carteiras",
    "📂 Histórico & Comparação"
])

# ==========================================================
# NOVO: LIMPANDO A MEMÓRIA AO TROCAR DE ABA
# ==========================================================
if 'aba_atual' not in st.session_state:
    st.session_state['aba_atual'] = nav

if st.session_state['aba_atual'] != nav:
    # O usuário mudou de aba! Vamos limpar os dados da aba antiga
    if st.session_state['aba_atual'] == "🚀 Roteirizador":
        st.session_state['df_main'] = None
        st.session_state['status_processamento'] = False
        st.session_state['processed_data'] = {}
    elif st.session_state['aba_atual'] == "👁️ Visualizador Operacional":
        st.session_state['arquivo_atual_vis'] = None
        if 'df_view_pronto' in st.session_state: del st.session_state['df_view_pronto']
    elif st.session_state['aba_atual'] == "🗺️ Visualizar Carteiras":
        st.session_state['cart_etapa'] = 'config'
        st.session_state['df_cart_raw'] = None
        st.session_state['df_cart_final'] = None
        
    st.session_state['aba_atual'] = nav
    st.rerun() # Reinicia a página limpa na nova aba
# ==========================================================

# ==========================================
# ABA 1: ROTEIRIZADOR INTELIGENTE
# ==========================================
if nav == "🚀 Roteirizador":
    c_title, c_help = st.columns([0.9, 0.1])
    c_title.title("📍🚗 Roteirizador Inteligente")

    st.markdown("""
        <div style="background-color: #fff3cd; padding: 20px; border-radius: 10px; border-left: 6px solid #ffc107; margin-bottom: 20px;">
            <h4 style="color: #856404; margin-top: 0;">⚠️ Atenção: Regra de Importação</h4>
            <p style="color: #856404; font-size: 15px;">
                Para garantir o funcionamento da Aplicação é necessario que o arquivo tenha as seguintes colunas na primeira linha da base de dados: .<br>
                <b>Nome Coordenador - Nome Vendedor - Nome Cliente - Código - Endereço - Municipio - CEP - Bairro - Latitude(Opcional) - Longitude(Opcional) - Volume(Opcional).<b>
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    with c_help:
        with st.popover("💡"):
            st.markdown("### Inteligência V2\n1. **Busca ArcGIS (Alta Precisão)**\n2. **Google Maps + Street View**\n3. **Cálculo Real (OSRM)**\n4. **Balanceamento de Volume Crescente**")

    if 'df_main' not in st.session_state: st.session_state['df_main'] = None
    if 'processed_data' not in st.session_state: st.session_state['processed_data'] = {} 
    if 'status_processamento' not in st.session_state: st.session_state['status_processamento'] = False
    if 'datas_validas' not in st.session_state: st.session_state['datas_validas'] = [] 

    with st.sidebar:
        st.markdown("### ⚙️ Configurações da Rota")
        
        if st.button("🔄 Limpar Tudo / Novo Upload", type="secondary"):
            st.session_state['df_main'] = None
            st.session_state['processed_data'] = {}
            st.session_state['status_processamento'] = False
            st.rerun()
            
        min_v = st.number_input("Mínimo Viável/Dia", 1, 10, 3)
        max_dist = st.number_input("Raio Máx (KM)", value=15)
        margem_extra = st.slider("Margem Extra de Distância (%)", 0, 100, 40)
        
        st.markdown("### 📦 Estratégia de Volume")
        opcao_volume = st.radio(
            "Posicionar maior volume em quais dias?",
            ["📈 Nos últimos dias (Crescente)", "📉 Nos primeiros dias (Decrescente)"]
        )
        st.session_state['maior_volume_fim'] = (opcao_volume == "📈 Nos últimos dias (Crescente)")
        st.divider()

    f = st.file_uploader("📂 Carregar Base do ERP (Excel/CSV)", type=["xlsx","csv"])
    
    # 1. Cria uma variável para lembrar qual foi o último arquivo lido
    if 'nome_arquivo_rot' not in st.session_state:
        st.session_state['nome_arquivo_rot'] = None

    if f:
        try:
            # 2. Atualiza os dados SE a memória estiver vazia OU SE for um arquivo diferente
            if st.session_state['df_main'] is None or st.session_state['nome_arquivo_rot'] != f.name:
                df_raw = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
                st.session_state['df_main'] = df_raw
                st.session_state['status_processamento'] = False 
                st.session_state['nome_arquivo_rot'] = f.name # Memoriza o nome do novo arquivo
        except Exception as e: 
            st.error(f"Erro ao ler o arquivo: {e}")
            st.stop()
            
    if f is None and st.session_state['df_main'] is not None:
        st.info("Usando base carregada na memória.")

    if st.session_state['df_main'] is not None:
        df = st.session_state['df_main'].copy()
        cols = df.columns.tolist()
        
        def idx(lista_chaves): 
            l = [str(x).lower() for x in cols]
            for chave in lista_chaves:
                chave_segura = chave.lower()
                for i, c in enumerate(l):
                    if chave_segura in c: return i
            return 0
        
        with st.expander("⚙️ Mapeamento de Colunas (Automático)", expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            col_coord = c1.selectbox("Coordenação", cols, index=idx(['coord']))
            col_vend = c2.selectbox("Vendedor", cols, index=idx(['vend']))
            col_nome = c3.selectbox("Nome Cliente", cols, index=idx(['cliente', 'razao', 'nome fantasia', 'nome']))
            col_id = c4.selectbox("ID/Código", cols, index=idx(['cód', 'cod', 'código', 'codigo', 'id']))
            c5, c6, c7, c8 = st.columns(4)
            col_end = c5.selectbox("Endereço", cols, index=idx(['Endereço','end','rua','Endereco','Endereço']))
            col_mun = c6.selectbox("Cidade", cols, index=idx(['cid','mun']))
            col_cep = c7.selectbox("CEP", cols, index=idx(['cep']))
            col_bairro = c8.selectbox("Bairro", cols, index=idx(['bair','distrito']))
            
            c9, c10, c11 = st.columns(3)
            col_lat = c9.selectbox("Lat (Opcional)", ["(Nenhuma)"]+cols)
            col_lon = c10.selectbox("Lon (Opcional)", ["(Nenhuma)"]+cols)
            
            # Mapeamento de Volume
            idx_vol = idx(['vol', 'peso', 'qtde', 'quantidade', 'volume', 'litros'])
            col_vol = c11.selectbox("Volume/Peso (Opcional)", ["(Nenhum)"]+cols, index=idx_vol+1 if idx_vol else 0)

        df['Cliente'] = df[col_nome]
        df['Bairro'] = df[col_bairro]
        df['ID_Cliente'] = df[col_id]
        df['Endereço_Limpo'] = df[col_end].apply(limpar_endereco_para_geocoding)
        df['Volume'] = df[col_vol] if col_vol != "(Nenhum)" else 0

        st.write("### 🎯 Metas por Vendedor")
        vendedores_unicos = df[col_vend].dropna().unique()
        config_visitas = {}

        cols_vends = st.columns(3)
        for i, vend in enumerate(vendedores_unicos):
            with cols_vends[i % 3]:
                config_visitas[vend] = st.number_input(f"{vend} (Visitas/Dia)", min_value=1, max_value=60, value=15, key=f"conf_{vend}")

        if not st.session_state['status_processamento']:
            st.divider()
            start_date = st.date_input("Data Início", value=get_primeiro_dia_util_proximo_mes())
            datas_selecionadas = render_cal_main(start_date)
            
            st.session_state['datas_validas'] = datas_selecionadas
            st.write("##")
            if st.button("🚀 PROCESSAR ROTAS (COM BALANCEAMENTO)", type="primary", use_container_width=True):
                if not datas_selecionadas:
                    st.error("Selecione dias válidos no calendário!")
                else:
                    df_memoria = carregar_memoria()
                    
                    mem_by_id = df_memoria[(df_memoria['ID_Cliente'].notna()) & (df_memoria['ID_Cliente'] != '')]
                    mem_by_id = mem_by_id.drop_duplicates(subset=['ID_Cliente'], keep='last').set_index('ID_Cliente').to_dict('index')
                    mem_by_chave = df_memoria.drop_duplicates(subset=['Chave'], keep='last').set_index('Chave').to_dict('index')
                    
                    geo = ArcGIS(timeout=10)
                    
                    if 'Latitude' not in df.columns: df['Latitude'] = np.nan
                    if 'Longitude' not in df.columns: df['Longitude'] = np.nan
                    if 'Geo_Score' not in df.columns: df['Geo_Score'] = 'Manual'
                    
                    if col_lat != "(Nenhuma)":
                        df['Latitude'] = df[col_lat].apply(corrigir_lat_lon)
                        df['Longitude'] = df[col_lon].apply(corrigir_lat_lon)

                    missing_mask = (df['Latitude'].isna()) | (df['Latitude'] == 0)
                    missing_count = missing_mask.sum()
                    novos_dados_memoria = []
                    
                    st.caption("📡 Log de Geocodificação em Tempo Real (ArcGIS)")
                    log_container = st.container(height=250)
                    log_placeholder = log_container.empty()
                    linhas_log = []

                    prog_bar = st.progress(0)
                    status_text = st.empty()
                    
                    if missing_count > 0:
                        indices = df[missing_mask].index
                        for i, row_idx in enumerate(indices):
                            row = df.loc[row_idx]
                            
                            mun_limpo = corrigir_texto_erp(row[col_mun])
                            bairro_busca = corrigir_texto_erp(str(row.get('Bairro', '')))
                            if bairro_busca.lower() == 'nan': bairro_busca = ""
                            
                            end_limpo = limpar_endereco_para_geocoding(row[col_end])
                            chave = gerar_chave_cache(end_limpo, mun_limpo, row[col_cep])
                            cod_cli = str(row.get(col_id, 'S/ID')).replace('.0', '')
                            
                            palavras = end_limpo.lower().split()
                            excecoes = ['de', 'da', 'do', 'das', 'dos', 'e', 'em', 'na', 'no']
                            end_titulo = " ".join([p.capitalize() if p not in excecoes else p for p in palavras])
                            
                            log_status = "Erro"
                            is_memoria = False
                            
                            achou_memoria = False
                            
                            if cod_cli in mem_by_id and float(mem_by_id[cod_cli].get('Latitude', 0)) != 0:
                                df.at[row_idx, 'Latitude'] = float(mem_by_id[cod_cli]['Latitude'])
                                df.at[row_idx, 'Longitude'] = float(mem_by_id[cod_cli]['Longitude'])
                                df.at[row_idx, 'Geo_Score'] = mem_by_id[cod_cli]['Score']
                                
                                if mem_by_id[cod_cli].get('End_Correto'): df.at[row_idx, 'Endereço_Limpo'] = mem_by_id[cod_cli]['End_Correto']
                                if mem_by_id[cod_cli].get('Bairro_Correto'): df.at[row_idx, 'Bairro'] = mem_by_id[cod_cli]['Bairro_Correto']
                                
                                log_status = mem_by_id[cod_cli]['Score']
                                achou_memoria = True
                                is_memoria = True
                                
                            elif chave in mem_by_chave and float(mem_by_chave[chave].get('Latitude', 0)) != 0 and mem_by_chave[chave].get('Score') in ['Alta', 'Média', 'Precisa']:
                                df.at[row_idx, 'Latitude'] = float(mem_by_chave[chave]['Latitude'])
                                df.at[row_idx, 'Longitude'] = float(mem_by_chave[chave]['Longitude'])
                                df.at[row_idx, 'Geo_Score'] = mem_by_chave[chave]['Score']
                                log_status = mem_by_chave[chave]['Score']
                                achou_memoria = True
                                is_memoria = True

                            if not achou_memoria:
                                try:
                                    if "," in end_limpo:
                                        apenas_rua = end_limpo.split(',')[0].strip()
                                    else:
                                        apenas_rua = re.sub(r'\d+$', '', end_limpo).strip()

                                    loc = None
                                    tipo_score = 'Erro'

                                    loc = geo.geocode(f"{end_limpo}, {bairro_busca}, {mun_limpo}, SP, Brasil")
                                    if loc: tipo_score = 'Alta'
                                    
                                    if not loc:
                                        loc = geo.geocode(f"{end_limpo}, {mun_limpo}, SP, Brasil")
                                        if loc: tipo_score = 'Alta'
                                        
                                    if not loc and apenas_rua:
                                        loc = geo.geocode(f"{apenas_rua}, {mun_limpo}, SP, Brasil")
                                        if loc: tipo_score = 'Média'
                                        
                                    if not loc:
                                        cep_limpo = str(row[col_cep]).replace('-', '').replace('.', '').strip()
                                        if cep_limpo != 'nan' and cep_limpo != '':
                                            loc = geo.geocode(f"{cep_limpo}, {mun_limpo}, SP, Brasil")
                                            if loc: tipo_score = 'CEP'

                                    if loc and is_valid_sp_coord(loc.latitude, loc.longitude):
                                        lat, lon = loc.latitude, loc.longitude
                                        df.at[row_idx, 'Latitude'] = lat
                                        df.at[row_idx, 'Longitude'] = lon
                                        df.at[row_idx, 'Geo_Score'] = tipo_score
                                        novos_dados_memoria.append({'Chave': chave, 'Latitude': lat, 'Longitude': lon, 'Score': tipo_score, 'ID_Cliente': cod_cli, 'End_Correto': end_limpo, 'Bairro_Correto': bairro_busca})
                                        log_status = tipo_score
                                    else:
                                        df.at[row_idx, 'Latitude'] = 0.0
                                        df.at[row_idx, 'Longitude'] = 0.0
                                        df.at[row_idx, 'Geo_Score'] = 'Erro'
                                        log_status = 'Erro'

                                except Exception as e:
                                    log_status = 'Erro'

                            if log_status in ['Alta', 'Precisa', 'Manual']:
                                icone = "🟢"; texto_status = "Precisa"
                            elif log_status in ['Baixa', 'Média', 'Aproximada']:
                                icone = "🟠"; texto_status = "Aproximada"
                            else:
                                icone = "🔴"; texto_status = "Erro"
                                
                            tag_mem = " *(Memória)*" if is_memoria else ""
                            mensagem_log = f"{icone} **{cod_cli}** - {end_titulo} - **{texto_status}**{tag_mem}"
                            
                            linhas_log.insert(0, mensagem_log)
                            log_placeholder.markdown("  \n".join(linhas_log[:50]))

                            status_text.text(f"Geocodificando {i+1}/{missing_count}...")
                            prog_bar.progress((i+1)/missing_count)

                    if novos_dados_memoria:
                        salvar_memoria(pd.DataFrame(novos_dados_memoria))
                    
                    st.session_state['df_main'] = df
                    
                    status_text.text("Calculando e Balanceando as rotas...")
                    prog_bar.progress(0)
                    vends = df[col_vend].unique()
                    processed = {}
                    
                    for i, v in enumerate(vends):
                        df_v = df[df[col_vend] == v].copy()
                        df_v['vend_col_temp'] = v 
                        
                        mask_precisa = (
                            (df_v['Latitude'].fillna(0) != 0) & 
                            (df_v['Longitude'].fillna(0) != 0) & 
                            (df_v['Geo_Score'].astype(str).str.strip().str.upper().isin(['ALTA', 'MANUAL', 'PRECISA']))
                        )
                        
                        df_v_rotear = df_v[mask_precisa].copy() 
                        df_v_manual = df_v[~mask_precisa].copy() 
                        
                        r, s = processar_rota_v101(
                            df_v_rotear, config_visitas, max_dist, datas_selecionadas, 
                            col_mun, col_cep, col_end, min_visitas_dia=min_v, 
                            modo_rapido=False, margem_extra_pct=margem_extra,
                            maior_volume_fim=st.session_state.get('maior_volume_fim', True)
                        )
                        
                        if not df_v_manual.empty:
                            df_v_manual['Motivo'] = 'Endereço Aproximado (Revisão)'
                            s = pd.concat([s, df_v_manual], ignore_index=True)
                        
                        salvar_no_historico(r, s, v)
                        processed[v] = {'rota': r, 'sobra': s, 'coord': df_v[col_coord].iloc[0] if col_coord in df_v else 'Geral'}
                        prog_bar.progress((i+1)/len(vends))

                    st.session_state['processed_data'] = processed
                    st.session_state['status_processamento'] = True
                    st.rerun()

        else:
            c_res1, c_res2 = st.columns([3, 1])
            c_res1.success("Processamento concluído!")
            
            if c_res2.button("⚙️ Alterar Configurações / Reprocessar"):
                st.session_state['processed_data'] = {}
                st.session_state['status_processamento'] = False
                st.rerun()

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for v, d in st.session_state['processed_data'].items():
                    if not d['rota'].empty:
                        d['rota']['OBS_Localizacao'] = d['rota']['Geo_Score'].apply(lambda x: 'Localização Precisa' if str(x).lower() in ['alta', 'manual'] else 'Aproximada')
                        xls = gerar_excel_estilizado(d['rota'], d['sobra'], pd.DataFrame())
                        zf.writestr(f"{v}_Rota.xlsx", xls)
            st.download_button("📦 Baixar Tudo (ZIP)", zip_buf.getvalue(), "Rotas.zip", "application/zip", type="primary")
            
            st.divider()
            dados = st.session_state['processed_data']
            
            coords = sorted(list(set([v['coord'] for v in dados.values() if 'coord' in v])))
            if not coords: coords = ['Geral'] 
            
            c1, c2, c3 = st.columns(3)
            sel_c = c1.selectbox("Coordenação", coords)
            
            vends = sorted([k for k, v in dados.items() if v.get('coord', 'Geral') == sel_c])
            sel_v = c2.selectbox("Vendedor", ["Ver Todos"] + vends)
            
            if sel_v == "Ver Todos":
                lista_r, lista_s = [], []
                for v in vends:
                    r_temp = dados[v]['rota'].copy()
                    if not r_temp.empty: r_temp['Nome_Vendedor'] = v
                    lista_r.append(r_temp)
                    
                    s_temp = dados[v]['sobra'].copy()
                    if not s_temp.empty: s_temp['Nome_Vendedor'] = v
                    lista_s.append(s_temp)
                    
                df_r = pd.concat(lista_r, ignore_index=True) if any(not d.empty for d in lista_r) else pd.DataFrame()
                df_s = pd.concat(lista_s, ignore_index=True) if any(not d.empty for d in lista_s) else pd.DataFrame()
            else:
                df_r = dados[sel_v]['rota'].copy()
                if not df_r.empty: df_r['Nome_Vendedor'] = sel_v
                
                df_s = dados[sel_v]['sobra'].copy()
                if not df_s.empty: df_s['Nome_Vendedor'] = sel_v

            if not df_r.empty:
                df_r['Roteiro_Data'] = pd.to_datetime(df_r['Roteiro_Data'], dayfirst=True, format='mixed').dt.date
                
                contagem_dias = df_r['Roteiro_Data'].value_counts().to_dict()
                dias_unicos = sorted(df_r['Roteiro_Data'].unique())
                opcoes_dia = []
                for d in dias_unicos:
                    qtd = contagem_dias.get(d, 0)
                    label = f"{d.strftime('%d/%m/%Y')} ({qtd} Visitas)"
                    opcoes_dia.append((d, label))
            else:
                opcoes_dia = []

            sel_d_tuple = c3.selectbox(
                "Dia (Selecione)", 
                [("Ver Tudo", "Ver Tudo")] + opcoes_dia, 
                format_func=lambda x: x[1]
            )
            sel_d = sel_d_tuple[0]

            st.markdown("### 💾 Salvar Rota no Sistema")
            with st.form("form_sql_save", border=True):
                st.info("Esta ação salvará a tabela de rotas e sobras exibida atualmente para uso futuro no Banco SQL.")
                
                c_f1, c_f2 = st.columns(2)
                nome_rota = c_f1.text_input("Nome da Rota (Ex: Rota SP Zona Sul)")
                
                idx_tipo = 1 if sel_v == "Ver Todos" else 0
                tipo_vend = c_f2.selectbox("É de Um ou Múltiplos Vendedores?", ["Um Vendedor", "Múltiplos Vendedores"], index=idx_tipo)
                
                c_f3, c_f4 = st.columns(2)
                meses_ano = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
                mes_padrao = datetime.now().month - 1
                mes_ref = c_f3.selectbox("Mês Pertencente", meses_ano, index=mes_padrao)
                
                dias_unicos_rota = df_r['Roteiro_Data'].nunique() if not df_r.empty else 0
                c_f4.text_input("Qtde de Dias na Rota (Automático)", value=str(dias_unicos_rota), disabled=True)
                
                if st.form_submit_button("✅ Salvar Rota no Banco de Dados", type="primary", use_container_width=True):
                    if not nome_rota:
                        st.error("Por favor, digite o Nome da Rota antes de salvar!")
                    else:
                        user = st.session_state.get('user_data', {}).get('email', 'Usuario Local')
                        data_hoje = datetime.now().strftime("%d/%m/%Y %H:%M")
                        
                        df_r_str = df_r.to_json(orient='records', date_format='iso') if not df_r.empty else "[]"
                        df_s_str = df_s.to_json(orient='records', date_format='iso') if not df_s.empty else "[]"
                        
                        conn = sqlite3.connect('banco_rotas.db')
                        c = conn.cursor()
                        c.execute("INSERT INTO rotas_salvas (nome_rota, tipo_vendedor, mes_referencia, qtd_dias, usuario_salvou, data_salvamento, df_rota_json, df_sobra_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                  (nome_rota, tipo_vend, mes_ref, dias_unicos_rota, user, data_hoje, df_r_str, df_s_str))
                        conn.commit()
                        conn.close()
                        st.success(f"Rota '{nome_rota}' salva com sucesso no cofre do SQL!")
            
            st.write("##") 
            sel_d = sel_d_tuple[0]
            
            if not df_s.empty: 
                def definir_motivo(row):
                    score = str(row.get('Geo_Score', '')).upper().strip()
                    if score not in ['ALTA', 'MANUAL', 'PRECISA']: return 'Endereço Aproximado / Erro'
                    elif pd.isna(row.get('Motivo')) or str(row.get('Motivo')).strip() == '': return 'Sobra / Distância'
                    else: return row.get('Motivo')
                df_s['Motivo'] = df_s.apply(definir_motivo, axis=1)

            st.markdown("### 📊 Indicadores de Performance")
            
            def contar_problemas(df):
                if df.empty: return 0
                mask = (df['Latitude'] == 0.0) | (df['Longitude'] == 0.0) | (df['Geo_Score'].astype(str).str.upper().isin(['CEP', 'BAIXA', 'MÉDIA', 'ERRO', 'NAN', '']))
                return len(df[mask])

            if sel_d == "Ver Tudo":
                qtd_rot = len(df_r); qtd_sob = len(df_s)
                total_clientes = qtd_rot + qtd_sob
                texto_sobras = str(qtd_sob)
                qtd_aprox = contar_problemas(df_r) + contar_problemas(df_s)
            else:
                sub_df = df_r[df_r['Roteiro_Data'] == sel_d]
                qtd_rot = len(sub_df); qtd_sob = 0; total_clientes = qtd_rot
                texto_sobras = "0"
                qtd_aprox = contar_problemas(sub_df)

            kp1, kp2, kp3, kp4 = st.columns(4)
            kp1.markdown(f'<div class="kpi-card bg-blue"><div class="kpi-title">Total Clientes</div><div class="kpi-value">{total_clientes}</div></div>', unsafe_allow_html=True)
            kp2.markdown(f'<div class="kpi-card bg-green"><div class="kpi-title">Roteirizados</div><div class="kpi-value">{qtd_rot}</div></div>', unsafe_allow_html=True)
            kp3.markdown(f'<div class="kpi-card bg-red"><div class="kpi-title">Sobras</div><div class="kpi-value">{texto_sobras}</div></div>', unsafe_allow_html=True)
            kp4.markdown(f'<div class="kpi-card bg-orange"><div class="kpi-title">Problemas (Aprox/Erro)</div><div class="kpi-value">{qtd_aprox}</div></div>', unsafe_allow_html=True)
            
            st.write("##")

            # ==========================================
            # NOVO: CALENDÁRIO VISUAL DE VOLUME
            # ==========================================
            st.markdown("### 🗓️ Calendário de Ocupação (Clientes x Volume)")
            if not df_r.empty:
                df_r['Vol_Calc'] = df_r['Volume'].apply(limpar_volume)
                resumo_cal = df_r.groupby('Roteiro_Data').agg(
                    Qtd=('Cliente', 'count'),
                    Vol=('Vol_Calc', 'sum')
                ).reset_index()
                
                resumo_cal = resumo_cal.sort_values('Roteiro_Data')
                
                cols_per_row = 6
                for i in range(0, len(resumo_cal), cols_per_row):
                    cols_cal = st.columns(cols_per_row)
                    for j in range(cols_per_row):
                        if i + j < len(resumo_cal):
                            row_cal = resumo_cal.iloc[i + j]
                            dt_obj = row_cal['Roteiro_Data']
                            if isinstance(dt_obj, str): 
                                try: dt_str = pd.to_datetime(dt_obj).strftime('%d/%m/%Y')
                                except: dt_str = dt_obj
                            else:
                                dt_str = dt_obj.strftime('%d/%m/%Y')
                                
                            vol_val = row_cal['Vol']
                            vol_str = f"{vol_val:g}" 
                            
                            cols_cal[j].markdown(f"""
                            <div style="border: 1px solid #ddd; border-radius: 8px; padding: 10px; margin-bottom: 15px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); background-color: #f8f9fa;">
                                <div style="background-color: #00144F; color: white; padding: 6px; border-radius: 5px; font-weight: bold; font-size: 14px;">
                                    {dt_str}
                                </div>
                                <div style="margin-top: 10px; font-size: 14px; color: #333;">
                                    👥 <b>{row_cal['Qtd']}</b> Clientes<br>
                                    📦 <b>{vol_str}</b> Vol
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            
            st.write("##")

            # ==========================================
            # RADAR DE CORREÇÃO
            # ==========================================
            c_tit1, c_tit2 = st.columns([2, 1])
            c_tit1.subheader("🚧 Radar de Endereços (Aproximados)")
            
            if c_tit2.button("🔄 Reprocessar Rotas (Encaixar Corrigidos)", type="primary", use_container_width=True):
                with st.spinner("Recalculando a melhor rota com os endereços corrigidos..."):
                    vendedores_a_reprocessar = [sel_v] if sel_v != "Ver Todos" else list(st.session_state['processed_data'].keys())
                    for vend in vendedores_a_reprocessar:
                        dados_vend = st.session_state['processed_data'].get(vend)
                        if not dados_vend: continue
                        
                        df_r_vend = dados_vend['rota']
                        df_s_vend = dados_vend['sobra']
                        df_full_vend = pd.concat([df_r_vend, df_s_vend], ignore_index=True)
                        if df_full_vend.empty: continue
                        
                        cols_drop = ['Roteiro_Data', 'Roteiro_Ordem', 'Semana', 'Dia_Semana', 'Dist_Real_KM', 'Link_Rota_Google', 'Motivo', 'OBS_Localizacao', 'Status_Local']
                        df_full_vend = df_full_vend.drop(columns=[c for c in cols_drop if c in df_full_vend.columns])
                        
                        mask_precisa = (
                            (df_full_vend['Latitude'].fillna(0) != 0) & 
                            (df_full_vend['Longitude'].fillna(0) != 0) & 
                            (df_full_vend['Geo_Score'].astype(str).str.strip().str.upper().isin(['ALTA', 'MANUAL', 'PRECISA']))
                        )
                        
                        df_v_rotear = df_full_vend[mask_precisa].copy()
                        df_v_manual = df_full_vend[~mask_precisa].copy()
                        df_v_rotear['vend_col_temp'] = vend 

                        r, s = processar_rota_v101(
                            df_v_rotear, config_visitas, max_dist, st.session_state.get('datas_validas', []), 
                            col_mun, col_cep, col_end, min_visitas_dia=min_v, 
                            modo_rapido=False, margem_extra_pct=margem_extra,
                            maior_volume_fim=st.session_state.get('maior_volume_fim', True)
                        )
                        
                        if not df_v_manual.empty:
                            df_v_manual['Motivo'] = 'Endereço Aproximado / Erro'
                            s = pd.concat([s, df_v_manual], ignore_index=True)
                            
                        st.session_state['processed_data'][vend]['rota'] = r
                        st.session_state['processed_data'][vend]['sobra'] = s
                        
                    st.success("Tudo recalculado! Os clientes corrigidos foram encaixados na melhor rota possível.")
                    time.sleep(1.5)
                    st.rerun()

            if not df_r.empty:
                view = df_r if sel_d == "Ver Tudo" else df_r[df_r['Roteiro_Data'] == sel_d].sort_values('Roteiro_Ordem')
                
                if sel_d != "Ver Tudo" and sel_v != "Ver Todos":
                    if st.button(f"🔄 Reotimizar Sequência do Dia {sel_d.strftime('%d/%m')}", type="primary"):
                        df_dia_recalc = df_r[df_r['Roteiro_Data'] == sel_d].copy()
                        rota_nova = otimizar_sequencia_dia(df_dia_recalc, col_cep)
                        df_r = df_r[df_r['Roteiro_Data'] != sel_d]
                        df_dia_novo = pd.DataFrame(rota_nova)
                        df_dia_novo = recalcular_distancias_dataframe(df_dia_novo, modo_rapido=False)
                        df_r = pd.concat([df_r, df_dia_novo], ignore_index=True)
                        st.session_state['processed_data'][sel_v]['rota'] = df_r
                        st.rerun()

                if 'Roteiro_Data' in view.columns:
                    view['Roteiro_Data'] = view['Roteiro_Data'].apply(lambda x: x.strftime('%d/%m/%Y') if not isinstance(x, str) else x)
                
                colunas_desejadas = ['Roteiro_Ordem', 'Nome_Vendedor', 'Roteiro_Data', col_id, 'Cliente', 'Volume', 'Endereço_Limpo', 'Bairro', col_cep, col_mun, 'Dist_Real_KM', 'OBS_Localizacao']
                colunas_finais = [c for c in colunas_desejadas if c in view.columns]
                
                mask_r = ((df_r['Latitude'].fillna(0) == 0) | (df_r['Longitude'].fillna(0) == 0) | (~df_r['Geo_Score'].astype(str).str.strip().str.upper().isin(['ALTA', 'MANUAL', 'PRECISA'])))
                df_al_r = df_r[mask_r].copy()
                if not df_al_r.empty: df_al_r['Status_Local'] = 'Em Rota'

                df_al_s = pd.DataFrame()
                if not df_s.empty:
                    mask_s = ((df_s['Latitude'].fillna(0) == 0) | (df_s['Longitude'].fillna(0) == 0) | (~df_s['Geo_Score'].astype(str).str.strip().str.upper().isin(['ALTA', 'MANUAL', 'PRECISA'])))
                    df_al_s = df_s[mask_s].copy()
                    if not df_al_s.empty: df_al_s['Status_Local'] = 'Nas Sobras'

                df_alertas_full = pd.concat([df_al_r, df_al_s], ignore_index=True)

                if not df_alertas_full.empty:
                    st.error(f"⚠️ **Radar de Correção:** Existem {len(df_alertas_full)} cliente(s) com problema no endereço. Corrija-os abaixo:")
                    
                    cols_alerta = [c for c in colunas_desejadas if c in df_alertas_full.columns]
                    if 'Status_Local' not in cols_alerta: cols_alerta.insert(0, 'Status_Local')
                    
                    estilos_alertas = [{'selector': 'th', 'props': [('background-color', '#c0392b !important'), ('color', '#FFFFFF !important')]}]
                    
                    event_alertas = st.dataframe(
                        df_alertas_full[cols_alerta].style.set_table_styles(estilos_alertas), 
                        on_select="rerun", selection_mode="single-row", use_container_width=True, hide_index=True, key="tb_alertas_global"
                    )
                    
                    if len(event_alertas.selection['rows']) > 0:
                        idx_visual_alerta = event_alertas.selection['rows'][0]

                        if not df_alertas_full.empty and idx_visual_alerta < len(df_alertas_full):
                            row_alerta = df_alertas_full.iloc[idx_visual_alerta]
                            id_cliente = row_alerta[col_id]
                            
                            with st.expander(f"✏️ Corrigir: {row_alerta.get('Cliente', '')}", expanded=True):
                                c_ed1, c_ed2 = st.columns([3, 1])
                                novo_end_al = c_ed1.text_input("Endereço (Rua e Número):", value=row_alerta.get('Endereço_Limpo', ''), key=f"al_end_glob_{id_cliente}")
                                novo_bai_al = c_ed2.text_input("Bairro:", value=row_alerta.get('Bairro', ''), key=f"al_bai_glob_{id_cliente}")
                                c_btn1, c_btn2, c_btn3, c_btn4 = st.columns([2, 1.2, 1.2, 1])
                                
                                if c_btn1.button("📍 Corrigir e Atualizar", type="primary", key=f"al_btn_glob_{id_cliente}"):
                                    with st.spinner("Buscando no satélite..."):
                                        lat_new, lon_new, score_new = reprocessar_cliente_individual(row_alerta[col_id], novo_end_al, novo_bai_al, row_alerta.get('Cidade_Ref', ''), row_alerta.get('CEP_Ref', ''))
                                        if lat_new != 0:
                                            score_new = "Manual"
                                    
                                            # 1. PRIMEIRO nós declaramos quem é o cliente (PUXE ESSAS DUAS LINHAS PARA A ESQUERDA)
                                            vend_alvo = row_alerta.get('Nome_Vendedor', sel_v)
                                            cod_alvo = row_alerta[col_id]

                                            # 2. DEPOIS chamamos a função para salvar na tabela definitiva
                                            end_antigo_log = row_alerta.get('Endereço_Limpo', 'S/ Info')
                                            salvar_correcao_cliente(
                                                codigo=cod_alvo,
                                                # ... resto continua igual ...
                                                end_antigo=end_antigo_log, 
                                                end_novo=novo_end_al, 
                                                bairro_novo=novo_bai_al, 
                                                lat=lat_new, 
                                                lon=lon_new
                                            )

                                            # 3. E aí continua o processo normal atualizando a tela...
                                            df_orig_r = st.session_state['processed_data'][vend_alvo]['rota']
                                            df_orig_s = st.session_state['processed_data'][vend_alvo]['sobra']
                                            
                                            if cod_alvo in df_orig_r[col_id].values:
                                                idx_orig = df_orig_r[df_orig_r[col_id] == cod_alvo].index[0]
                                                df_orig_r.at[idx_orig, 'Latitude'], df_orig_r.at[idx_orig, 'Longitude'], df_orig_r.at[idx_orig, 'Geo_Score'] = lat_new, lon_new, score_new
                                                df_orig_r.at[idx_orig, 'Endereço_Limpo'], df_orig_r.at[idx_orig, 'Bairro'] = novo_end_al, novo_bai_al
                                                st.session_state['processed_data'][vend_alvo]['rota'] = df_orig_r
                                            elif not df_orig_s.empty and cod_alvo in df_orig_s[col_id].values:
                                                idx_orig = df_orig_s[df_orig_s[col_id] == cod_alvo].index[0]
                                                df_orig_s.at[idx_orig, 'Latitude'], df_orig_s.at[idx_orig, 'Longitude'], df_orig_s.at[idx_orig, 'Geo_Score'] = lat_new, lon_new, score_new
                                                df_orig_s.at[idx_orig, 'Endereço_Limpo'], df_orig_s.at[idx_orig, 'Bairro'] = novo_end_al, novo_bai_al
                                                st.session_state['processed_data'][vend_alvo]['sobra'] = df_orig_s
                                                
                                            chave_mem = gerar_chave_cache(novo_end_al, row_alerta.get('Cidade_Ref', ''), row_alerta.get('CEP_Ref', ''))
                                            # Salvar com ID e Endereço Corrigido
                                            salvar_memoria(pd.DataFrame([{
                                                'Chave': chave_mem, 
                                                'Latitude': lat_new, 
                                                'Longitude': lon_new, 
                                                'Score': score_new,
                                                'ID_Cliente': cod_alvo,
                                                'End_Correto': novo_end_al,
                                                'Bairro_Correto': novo_bai_al
                                            }]))
                                            
                                            st.session_state['foco_id'] = cod_alvo 
                                            if 'tb_geral_rota' in st.session_state: del st.session_state['tb_geral_rota']
                                            if 'tb_alertas_global' in st.session_state: del st.session_state['tb_alertas_global']
                                                
                                            st.success("Corrigido com sucesso!")
                                            time.sleep(1)
                                            st.rerun()
                                        else:
                                            st.error("Rua não encontrada ou cidade inválida. Simplifique o texto.")
                                            
                                if c_btn2.button("✅ Liberar OK", key=f"btn_lib_al_{id_cliente}", use_container_width=True):
                                    vend_alvo = row_alerta.get('Nome_Vendedor', sel_v)
                                    cod_alvo = row_alerta[col_id]
                                    for tipo in ['rota', 'sobra']:
                                        df_alvo = st.session_state['processed_data'][vend_alvo][tipo]
                                        if not df_alvo.empty and cod_alvo in df_alvo[col_id].values:
                                            idx = df_alvo[df_alvo[col_id] == cod_alvo].index[0]
                                            df_alvo.at[idx, 'Geo_Score'] = 'Manual'
                                            st.session_state['processed_data'][vend_alvo][tipo] = df_alvo
                                    if 'tb_alertas_global' in st.session_state: del st.session_state['tb_alertas_global']
                                    st.success("Cliente Liberado para Rota!")
                                    time.sleep(0.5)
                                    st.rerun()

                                if c_btn3.button("❌ Jogar p/ Sobras", use_container_width=True, key=f"al_btn_sob_glob_{id_cliente}"):
                                    if row_alerta.get('Status_Local') == 'Em Rota':
                                        vend_alvo = row_alerta.get('Nome_Vendedor', sel_v)
                                        cod_alvo = row_alerta[col_id]
                                        df_orig_r = st.session_state['processed_data'][vend_alvo]['rota']
                                        df_orig_s = st.session_state['processed_data'][vend_alvo]['sobra']
                                        idx_orig = df_orig_r[df_orig_r[col_id] == cod_alvo].index[0]
                                        row_to_move = df_orig_r.loc[[idx_orig]].copy()
                                        df_orig_r = df_orig_r.drop(idx_orig)
                                        df_orig_s = pd.concat([df_orig_s, row_to_move.drop(columns=['Roteiro_Data', 'Roteiro_Ordem', 'Dist_Real_KM', 'Link_Rota_Google', 'Semana', 'Dia_Semana'], errors='ignore')], ignore_index=True)
                                        st.session_state['processed_data'][vend_alvo]['rota'] = df_orig_r
                                        st.session_state['processed_data'][vend_alvo]['sobra'] = df_orig_s
                                        st.rerun()
                                        
                                if c_btn4.button("🗺️ Focar no Mapa", use_container_width=True, key=f"foco_red_{id_cliente}"):
                                    st.session_state['foco_id'] = id_cliente
                                    st.rerun()
                        else:
                            st.info("A lista foi atualizada. Por favor, selecione o cliente novamente na tabela acima.")

                st.divider()

                st.markdown(f"### 📍 Clientes Roterizados")
                
                df_show = view[colunas_finais].copy()
                
                # Formatando as Novas Colunas
                if 'Dist_Real_KM' in df_show.columns:
                    df_show['Distância p/ Anterior'] = df_show['Dist_Real_KM'].apply(lambda x: f"{x:.2f} km")
                    df_show = df_show.drop(columns=['Dist_Real_KM'])
                if 'Roteiro_Ordem' in df_show.columns:
                    df_show = df_show.rename(columns={'Roteiro_Ordem': 'Seq.'})
                
                estilos_tabela = [{'selector': 'th', 'props': [('background-color', '#00144F !important'), ('color', '#FFFFFF !important'), ('text-align', 'left')]}]
                
                event = st.dataframe(
                    df_show.style.set_table_styles(estilos_tabela), 
                    on_select="rerun", selection_mode="multi-row", use_container_width=True, hide_index=True, key="tb_geral_rota"
                )
                
                if len(event.selection['rows']) > 0:
                    indices_selecionados = event.selection['rows']
                    
                    st.write(f"**{len(indices_selecionados)} cliente(s) selecionado(s)**")

                    # Se APENAS 1 cliente estiver selecionado, permite edição manual e focar no mapa
                    if len(indices_selecionados) == 1:
                        idx_visual = indices_selecionados[0]
                        real_idx = view.index[idx_visual]
                        row_atual = df_r.loc[real_idx]
                        
                        with st.expander("✏️ Editor de Endereço (Correção Manual)", expanded=True):
                            c_edit1, c_edit2 = st.columns([3, 1])
                            novo_endereco = c_edit1.text_input("Endereço (Rua e Número):", value=row_atual['Endereço_Limpo'], key=f"rota_end_{row_atual[col_id]}")
                            novo_bairro = c_edit2.text_input("Bairro:", value=row_atual['Bairro'], key=f"rota_bai_{row_atual[col_id]}")
                            
                            c_act1, c_act2 = st.columns([3, 1])
                            
                            if c_act1.button("📍 Buscar Nova Localização", type="primary", key=f"btn_save_m_{real_idx}"):
                                with st.spinner("Buscando..."):
                                    lat_new, lon_new, score_new = reprocessar_cliente_individual(row_atual[col_id], novo_endereco, novo_bairro, row_atual['Cidade_Ref'], row_atual['CEP_Ref'])
                                    if lat_new != 0:
                                        vend_alvo = row_atual.get('Nome_Vendedor', sel_v)
                                        cod_alvo = row_atual[col_id]
                                        df_orig_r = st.session_state['processed_data'][vend_alvo]['rota']
                                        
                                        if cod_alvo in df_orig_r[col_id].values:
                                            idx_orig = df_orig_r[df_orig_r[col_id] == cod_alvo].index[0]
                                            df_orig_r.at[idx_orig, 'Latitude'], df_orig_r.at[idx_orig, 'Longitude'], df_orig_r.at[idx_orig, 'Geo_Score'] = lat_new, lon_new, "Manual"
                                            df_orig_r.at[idx_orig, 'Endereço_Limpo'], df_orig_r.at[idx_orig, 'Bairro'] = novo_endereco, novo_bairro
                                            st.session_state['processed_data'][vend_alvo]['rota'] = df_orig_r
                                            
                                        chave_mem = gerar_chave_cache(novo_endereco, row_atual.get('Cidade_Ref', ''), row_atual.get('CEP_Ref', ''))
                                        
                                        # Salvando na memória com ID!
                                        salvar_memoria(pd.DataFrame([{
                                            'Chave': chave_mem, 
                                            'Latitude': lat_new, 
                                            'Longitude': lon_new, 
                                            'Score': "Manual",
                                            'ID_Cliente': cod_alvo,
                                            'End_Correto': novo_endereco,
                                            'Bairro_Correto': novo_bairro
                                        }]))
                                        
                                        st.success("Localização atualizada com sucesso no banco principal!")
                                        time.sleep(1)
                                        st.rerun()
                                    else:
                                        st.error("Não encontrado. Tente simplificar o nome da rua.")

                            if c_act2.button("🗺️ Focar no Mapa", use_container_width=True, key=f"foco_blue_{row_atual[col_id]}"):
                                st.session_state['foco_id'] = row_atual[col_id]
                                st.rerun()

                    # Ações em Lote para TODOS os clientes selecionados (1 ou mais)
                    st.write("⚙️ **Ações em Lote:**")
                    c_lote1, c_lote2 = st.columns(2)
                    
                    if c_lote1.button("❌ Mover p/ Sobras", use_container_width=True, key="btn_lote_sobra"):
                        real_indices = [view.index[i] for i in indices_selecionados]
                        rows_rem = df_r.loc[real_indices]
                        df_r = df_r.drop(real_indices)
                        cols_drop = ['Roteiro_Data', 'Roteiro_Ordem', 'Dist_Real_KM', 'Link_Rota_Google', 'Semana', 'Dia_Semana']
                        rows_sobra = rows_rem.drop(columns=cols_drop, errors='ignore')
                        df_s = pd.concat([df_s, rows_sobra], ignore_index=True)
                        st.session_state['processed_data'][sel_v]['rota'] = df_r
                        st.session_state['processed_data'][sel_v]['sobra'] = df_s
                        st.rerun()

                    with c_lote2.popover("📅 Mover para Outro Dia", use_container_width=True):
                        todas_datas_possiveis = st.session_state.get('datas_validas', [])
                        if todas_datas_possiveis:
                            target_day_move = st.selectbox("Novo Dia:", todas_datas_possiveis, format_func=lambda x: x.strftime('%d/%m/%Y'), key="pop_lote_date")
                            if st.button("Confirmar Mudança", key="btn_lote_move"):
                                real_indices = [view.index[i] for i in indices_selecionados]
                                rows_to_move = df_r.loc[real_indices].copy()
                                df_r = df_r.drop(real_indices)
                                rows_to_move['Roteiro_Data'] = target_day_move
                                df_r = pd.concat([df_r, rows_to_move], ignore_index=True)
                                st.session_state['processed_data'][sel_v]['rota'] = df_r
                                st.rerun()

            else:
                st.warning("Nenhuma visita agendada.")
            st.divider()

            # ==========================================
            # 3. MAPA OPERACIONAL (GOOGLE MAPS)
            # ==========================================
            st.subheader("🗺️ Mapa de Roteirização (Google Maps)")
            
            if sel_d != "Ver Tudo" and not df_r.empty:
                v_map_btn = df_r[(df_r['Roteiro_Data'] == sel_d) & (df_r['Latitude'] != 0)].sort_values('Roteiro_Ordem')
                if len(v_map_btn) > 1:
                    link_oficial_maps = gerar_link_google_maps(v_map_btn)
                    if link_oficial_maps:
                        st.link_button("📱 Abrir Rota Passo-a-Passo no App Google Maps (Navegação GPS)", link_oficial_maps, type="primary")
                        st.write("##") 
            
            v_map = pd.DataFrame()
            s_map = pd.DataFrame()
            foco_ativo = False
            
            if st.session_state.get('foco_id') is not None:
                foco_id = st.session_state['foco_id']
                if not df_r.empty and foco_id in df_r[col_id].values:
                    v_map = df_r[df_r[col_id] == foco_id].copy()
                    foco_ativo = True
                elif not df_s.empty and foco_id in df_s[col_id].values:
                    s_map = df_s[df_s[col_id] == foco_id].copy()
                    foco_ativo = True
                    
                if foco_ativo:
                    lat_check = float(v_map.iloc[0].get('Latitude', 0.0)) if not v_map.empty else float(s_map.iloc[0].get('Latitude', 0.0))
                    if lat_check == 0.0:
                        st.warning("⚠️ **Impossível Focar:** Este cliente está sem coordenadas válidas. Corrija o endereço primeiro para plotá-lo no mapa!")
                        st.session_state['foco_id'] = None
                        foco_ativo = False
                    else:
                        if st.button("🌍 Limpar Foco e Ver Rota Completa", type="primary"):
                            st.session_state['foco_id'] = None
                            st.rerun()
            else:
                st.session_state['foco_id'] = None 

            if not foco_ativo:
                if not df_r.empty:
                    v_map = df_r if sel_d == "Ver Tudo" else df_r[df_r['Roteiro_Data'] == sel_d]
                if not df_s.empty:
                    s_map = df_s.copy()
            
            if not v_map.empty:
                v_map = v_map[(v_map['Latitude'] != 0) & (v_map['Latitude'].notna())].copy()
            if not s_map.empty:
                s_map = s_map[(s_map['Latitude'] != 0) & (s_map['Latitude'].notna())].copy()

            if not v_map.empty or not s_map.empty:
                features = []
                pts_rota_osrm = []
                
                dias_unicos_globais = sorted(df_r['Roteiro_Data'].unique()) if not df_r.empty else []
                posicoes_vistas = {}
                distancia_offset = 0.00015  
                
                for _, r in v_map.iterrows():
                    lat_orig, lon_orig = r['Latitude'], r['Longitude']
                    
                    chave_pos = f"{round(lat_orig, 5)}_{round(lon_orig, 5)}"
                    if chave_pos in posicoes_vistas:
                        mult = posicoes_vistas[chave_pos]
                        lat = lat_orig + (distancia_offset * mult)
                        lon = lon_orig + (distancia_offset * mult)
                        posicoes_vistas[chave_pos] += 1
                    else:
                        lat, lon = lat_orig, lon_orig
                        posicoes_vistas[chave_pos] = 1

                    pts_rota_osrm.append([lat_orig, lon_orig]) 
                    
                    score_mapa = str(r.get('Geo_Score', '')).lower()
                    if score_mapa in ['cep', 'baixa', 'média', 'aproximada']:
                        cor_bg = "#D3AF37" 
                    else:
                        try:
                            if 'Roteiro_Data' in r and r['Roteiro_Data'] in dias_unicos_globais:
                                cor_idx = dias_unicos_globais.index(r['Roteiro_Data']) % len(CORES_HEX)
                                cor_bg = CORES_HEX[cor_idx]
                            else:
                                cor_bg = "#1a73e8" 
                        except:
                            cor_bg = "#1a73e8"
                        
                    link_maps_oficial = f"https://www.google.com/maps/search/?api=1&query={lat_orig},{lon_orig}"
                    
                    pop_html = f"""
                    <div style="font-family: Arial, sans-serif; width: 300px; padding: 5px;">
                        <strong style="font-size: 14px; color: {cor_bg};">{r.get('Cliente','')}</strong><br>
                        <span style="font-size: 12px; color: #666;">Seq: {r.get('Roteiro_Ordem','-')} | Cód: {r.get(col_id,'')} | Dia: {r.get('Roteiro_Data', '')}</span>
                        <hr style="border: 0; border-top: 1px solid #eee; margin: 8px 0;">
                        <table width="100%" style="font-size: 12px; margin-bottom: 8px;">
                            <tr><td><strong>Endereço:</strong></td><td>{r.get('Endereço_Limpo','')}</td></tr>
                            <tr><td><strong>Bairro:</strong></td><td>{r.get('Bairro','')}</td></tr>
                            <tr><td><strong>Volume:</strong></td><td>{r.get('Volume','0')}</td></tr>
                        </table>
                        <div style="border-radius: 4px; overflow: hidden; border: 1px solid #ddd; margin-bottom: 10px;">
                            <img loading="lazy" src="https://maps.googleapis.com/maps/api/streetview?size=300x150&location={lat_orig},{lon_orig}&key={GOOGLE_MAPS_API_KEY}" width="300" height="150">
                        </div>
                        <div style="text-align: center;">
                            <a href="{link_maps_oficial}" target="_blank" style="background-color: {cor_bg}; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; font-size: 12px; font-weight: bold; display: inline-block; width: 90%;">📍 Ver no Google Maps</a>
                        </div>
                    </div>
                    """
                    features.append({
                        "position": {"lat": lat, "lng": lon}, 
                        "title": str(r.get('Cliente','')),
                        "color": cor_bg,
                        "label": str(r.get('Roteiro_Ordem','')),
                        "content": pop_html
                    })

                for _, s in s_map.iterrows():
                    lat_orig, lon_orig = s['Latitude'], s['Longitude']
                    
                    chave_pos = f"{round(lat_orig, 5)}_{round(lon_orig, 5)}"
                    if chave_pos in posicoes_vistas:
                        mult = posicoes_vistas[chave_pos]
                        lat = lat_orig + (distancia_offset * mult)
                        lon = lon_orig + (distancia_offset * mult)
                        posicoes_vistas[chave_pos] += 1
                    else:
                        lat, lon = lat_orig, lon_orig
                        posicoes_vistas[chave_pos] = 1

                    link_maps_oficial = f"https://www.google.com/maps/search/?api=1&query={lat_orig},{lon_orig}"
                    
                    pop_html_s = f"""
                    <div style="font-family: Arial, sans-serif; width: 300px; padding: 5px;">
                        <strong style="font-size: 14px; color: #c0392b;">{s.get('Cliente','')} (SOBRA)</strong><br>
                        <span style="font-size: 12px; color: #666;">Motivo: {s.get('Motivo','')} | Cód: {s.get(col_id,'')}</span>
                        <hr style="border: 0; border-top: 1px solid #eee; margin: 8px 0;">
                        <table width="100%" style="font-size: 12px; margin-bottom: 8px;">
                            <tr><td><strong>Endereço:</strong></td><td>{s.get('Endereço_Limpo','')}</td></tr>
                        </table>
                        <div style="border-radius: 4px; overflow: hidden; border: 1px solid #ddd; margin-bottom: 10px;">
                            <img loading="lazy" src="https://maps.googleapis.com/maps/api/streetview?size=300x150&location={lat_orig},{lon_orig}&key={GOOGLE_MAPS_API_KEY}" width="300" height="150">
                        </div>
                        <div style="text-align: center;">
                            <a href="{link_maps_oficial}" target="_blank" style="background-color: #c0392b; color: white; padding: 8px 15px; text-decoration: none; border-radius: 4px; font-size: 12px; font-weight: bold; display: inline-block; width: 90%;">📍 Ver no Google Maps</a>
                        </div>
                    </div>
                    """
                    features.append({
                        "position": {"lat": lat, "lng": lon},
                        "title": str(s.get('Cliente','')),
                        "color": "#c0392b",
                        "label": "!",
                        "content": pop_html_s
                    })

                caminho_osrm_js = "[]"
                cor_linha_osrm = "#1a73e8" 
                
                if len(pts_rota_osrm) > 1 and sel_d != "Ver Tudo" and not foco_ativo:
                    with st.spinner("Desenhando traçado das ruas (OSRM)..."):
                        rota_rua = get_rota_detalhada_rua_otimizada(pts_rota_osrm)
                        coords_finais = rota_rua if rota_rua else pts_rota_osrm
                        lista_dicts = [{"lat": p[0], "lng": p[1]} for p in coords_finais]
                        caminho_osrm_js = json.dumps(lista_dicts)
                        
                        if sel_d in dias_unicos_globais:
                            cor_linha_osrm = CORES_HEX[dias_unicos_globais.index(sel_d) % len(CORES_HEX)]

                todas_lats = [f["position"]["lat"] for f in features]
                todas_lons = [f["position"]["lng"] for f in features]
                center_lat = sum(todas_lats) / len(todas_lats) if todas_lats else -23.5505
                center_lng = sum(todas_lons) / len(todas_lons) if todas_lons else -46.6333
                zoom_padrao = 16 if foco_ativo else 12

            google_maps_html = f"""
                <!DOCTYPE html>
                <html>
                  <head>
                    <style>
                      #map {{ height: 100%; }} html, body {{ height: 100%; margin: 0; padding: 0; }}
                      .gm-style-iw-d {{ overflow: auto !important; max-height: 480px !important; }}
                      .custom-control {{ 
                          background-color: #fff; border: 2px solid #fff; border-radius: 2px; 
                          box-shadow: 0 2px 6px rgba(0,0,0,.3); cursor: pointer; 
                          margin-right: 10px; margin-top: 10px; padding: 5px; text-align: center; 
                      }}
                      .custom-control label {{ color: rgb(25,25,25); font-family: Roboto,Arial,sans-serif; font-size: 13px; cursor: pointer; font-weight: bold; }}
                    </style>
                    <script>
                      const data = {json.dumps(features)};
                      const routeData = {caminho_osrm_js}; // <--- VOLTOU A LER O TRAÇADO
                      
                      let map, infoWindow, marker_list = [];
                      let routePath = null;
                      
                      function initMap() {{
                        // Regra para apagar o lixo visual do Google
                        const estiloLimpo = [
                            {{ featureType: "poi", elementType: "all", stylers: [{{ visibility: "off" }}] }},
                            {{ featureType: "transit", elementType: "all", stylers: [{{ visibility: "off" }}] }}
                        ];

                        map = new google.maps.Map(document.getElementById("map"), {{
                          zoom: 12, center: {{ lat: {center_lat}, lng: {center_lng} }},
                          gestureHandling: "greedy",
                          mapTypeControl: true, 
                          mapTypeControlOptions: {{ style: google.maps.MapTypeControlStyle.DROPDOWN_MENU, position: google.maps.ControlPosition.TOP_RIGHT }},
                          styles: estiloLimpo 
                        }});
                        
                        infoWindow = new google.maps.InfoWindow();
                        
                        data.forEach((point) => {{
                          const marker = new google.maps.Marker({{
                            position: point.position, map, title: point.title, optimize: true,
                            
                            // <--- VOLTOU A ORDEM DOS CLIENTES (1, 2, 3...)
                            label: {{ text: point.label, color: 'white', fontSize: '11px', fontWeight: 'bold' }},
                            
                            icon: {{ path: google.maps.SymbolPath.CIRCLE, fillColor: point.color, fillOpacity: 0.9, strokeWeight: 1, strokeColor: '#fff', scale: 6 }}
                          }});
                          marker.addListener("click", () => {{ infoWindow.setContent(point.content); infoWindow.open(map, marker); }});
                          marker_list.push(marker);
                        }});

                        // <--- DESENHA A LINHA AZUL DA ROTA
                        if (routeData && routeData.length > 1) {{
                            routePath = new google.maps.Polyline({{
                                path: routeData, geodesic: true, strokeColor: "{cor_linha_osrm}", strokeOpacity: 0.8, strokeWeight: 4,
                            }});
                            routePath.setMap(map);
                        }}
                        
                        if (marker_list.length > 0) {{
                            const bounds = new google.maps.LatLngBounds();
                            marker_list.forEach(marker => bounds.extend(marker.getPosition()));
                            if(routeData && routeData.length > 1){{
                                routeData.forEach(p => bounds.extend(p));
                            }}
                            map.fitBounds(bounds);
                        }}

                        // ==========================================
                        // MAGIA: TAMANHO DA BOLINHA RESPONSIVA
                        // ==========================================
                        map.addListener('zoom_changed', function() {{
                            let currentZoom = map.getZoom();
                            let newScale = 6; // tamanho padrão
                            
                            if (currentZoom >= 15) newScale = 10; // zoom muito perto (bolinha gigante)
                            else if (currentZoom >= 13) newScale = 8; // zoom médio
                            else if (currentZoom <= 10) newScale = 4; // zoom muito longe (bolinha minúscula)

                            marker_list.forEach(marker => {{
                                let icon = marker.getIcon();
                                icon.scale = newScale;
                                marker.setIcon(icon);
                            }});
                        }});
                        
                        createLabelControl(map, estiloLimpo);
                        createMarkerToggleControl(map);
                      }}
                      
                      function createLabelControl(map, estiloLimpo) {{
                        const controlDiv = document.createElement('div');
                        controlDiv.className = 'custom-control';
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox'; checkbox.id = 'toggle-labels'; checkbox.checked = false; 
                        const label = document.createElement('label');
                        label.htmlFor = 'toggle-labels';
                        label.appendChild(checkbox); label.appendChild(document.createTextNode(' 🏪 Ver Comércios (Google)'));
                        controlDiv.appendChild(label);
                        
                        checkbox.addEventListener('change', function() {{
                            if (this.checked) {{ map.setOptions({{ styles: [] }}); }} 
                            else {{ map.setOptions({{ styles: estiloLimpo }}); }} 
                        }});
                        map.controls[google.maps.ControlPosition.TOP_LEFT].push(controlDiv);
                      }}

                      function createMarkerToggleControl(map) {{
                        const controlDiv = document.createElement('div');
                        controlDiv.className = 'custom-control';
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox'; checkbox.id = 'toggle-markers'; checkbox.checked = true; 
                        const label = document.createElement('label');
                        label.htmlFor = 'toggle-markers';
                        label.appendChild(checkbox); label.appendChild(document.createTextNode(' 📍 Meus Clientes'));
                        controlDiv.appendChild(label);
                        
                        checkbox.addEventListener('change', function() {{
                            const isVisible = this.checked;
                            marker_list.forEach(marker => marker.setVisible(isVisible)); 
                            
                            // Esconde a linha da rota junto com os clientes
                            if (routePath) {{ routePath.setVisible(isVisible); }}
                        }});
                        map.controls[google.maps.ControlPosition.TOP_LEFT].push(controlDiv);
                      }}
                      
                      window.initMap = initMap;
                    </script>
                  </head>
                  <body>
                    <div id="map"></div>
                    <script src="https://maps.googleapis.com/maps/api/js?key={GOOGLE_MAPS_API_KEY}&callback=initMap&v=weekly" defer></script>
                  </body>
                </html>
                """
            st.components.v1.html(google_maps_html, height=750, scrolling=False)

            st.divider()

            # ==========================================
            # 4. GESTÃO DE SOBRAS
            # ==========================================
            if not df_s.empty:
                st.subheader("🚫 Sobras")
                
                cols_padrao_sobra = ['Nome_Vendedor', 'Motivo', col_id, 'Cliente', 'Volume', 'Endereço_Limpo', 'Bairro', col_cep, col_mun, 'OBS_Localizacao']
                cols_view_s = [c for c in cols_padrao_sobra if c in df_s.columns]
                df_show_s = df_s[cols_view_s].copy()

                estilos_tabela_sobras = [{'selector': 'th', 'props': [('background-color', '#00144F !important'), ('color', '#FFFFFF !important'), ('font-weight', 'bold !important'), ('text-align', 'left')]}]
                
                event_s = st.dataframe(
                    df_show_s.style.set_table_styles(estilos_tabela_sobras), 
                    on_select="rerun", selection_mode="multi-row", use_container_width=True, hide_index=True
                )

                if len(event_s.selection['rows']) > 0:
                    indices_selecionados = event_s.selection['rows']
                    
                    if len(indices_selecionados) == 1:
                        idx_visual_s = indices_selecionados[0]
                        real_idx_s = df_s.index[idx_visual_s]
                        row_atual_s = df_s.loc[real_idx_s]

                        with st.expander("✏️ Tentar Corrigir Endereço (Sobra Selecionada)", expanded=True):
                            c_edit1_s, c_edit2_s = st.columns([3, 1])
                            novo_end_s = c_edit1_s.text_input("Endereço:", value=row_atual_s.get('Endereço_Limpo', ''), key=f"ed_s_end_{real_idx_s}")
                            novo_bai_s = c_edit2_s.text_input("Bairro:", value=row_atual_s.get('Bairro', ''), key=f"ed_s_bai_{real_idx_s}")
                            
                            c_btn_s_1, c_btn_s_2 = st.columns(2)
                            
                            if c_btn_s_1.button("📍 Re-Geocodificar", key=f"btn_s_geo_{real_idx_s}", type="primary", use_container_width=True):
                                lat_new, lon_new, score_new = reprocessar_cliente_individual(row_atual_s[col_id], novo_end_s, novo_bai_s, row_atual_s.get('Cidade_Ref', ''), row_atual_s.get('CEP_Ref', ''))
                                if lat_new != 0:
                                    score_new = "Manual"
                                    vend_alvo = row_atual_s.get('Nome_Vendedor', sel_v)
                                    cod_alvo = row_atual_s[col_id]
                                    df_orig_s = st.session_state['processed_data'][vend_alvo]['sobra']
                                    
                                    if cod_alvo in df_orig_s[col_id].values:
                                        idx_orig = df_orig_s[df_orig_s[col_id] == cod_alvo].index[0]
                                        df_orig_s.at[idx_orig, 'Latitude'] = lat_new
                                        df_orig_s.at[idx_orig, 'Longitude'] = lon_new
                                        df_orig_s.at[idx_orig, 'Geo_Score'] = score_new
                                        df_orig_s.at[idx_orig, 'Endereço_Limpo'] = novo_end_s
                                        df_orig_s.at[idx_orig, 'Bairro'] = novo_bai_s
                                        st.session_state['processed_data'][vend_alvo]['sobra'] = df_orig_s
                                        
                                    chave_mem = gerar_chave_cache(novo_end_s, row_atual_s.get('Cidade_Ref', ''), row_atual_s.get('CEP_Ref', ''))
                                    
                                    salvar_memoria(pd.DataFrame([{
                                        'Chave': chave_mem, 
                                        'Latitude': lat_new, 
                                        'Longitude': lon_new, 
                                        'Score': score_new,
                                        'ID_Cliente': cod_alvo,
                                        'End_Correto': novo_end_s,
                                        'Bairro_Correto': novo_bai_s
                                    }]))
                                    
                                    st.success("Endereço corrigido com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.error("Localização não encontrada. Verifique o texto digitado.")
                                    
                            if c_btn_s_2.button("🗺️ Focar no Mapa", key=f"btn_s_foco_{real_idx_s}", use_container_width=True):
                                st.session_state['foco_id'] = row_atual_s[col_id]
                                st.rerun()

                    st.write(f"**{len(indices_selecionados)} sobra(s) selecionada(s)**")
                    c_add1, c_add2 = st.columns([1, 2])
                    
                    with c_add1.popover("➕ Adicionar à Rota"):
                        todas_datas_possiveis = st.session_state.get('datas_validas', [])
                        if todas_datas_possiveis:
                            target_day = st.selectbox("Escolha o Dia:", todas_datas_possiveis, format_func=lambda x: x.strftime('%d/%m/%Y'), key="sb_target_day")
                            
                            if st.button("Confirmar Encaixe", use_container_width=True, type="primary"):
                                real_indices_s = [df_s.index[i] for i in indices_selecionados]
                                rows_add = df_s.loc[real_indices_s].copy()
                                
                                rows_add['Roteiro_Data'] = target_day
                                data_ini_ref = df_r['Roteiro_Data'].min() if not df_r.empty else target_day
                                rows_add['Semana'], rows_add['Dia_Semana'] = get_semana_relativa_str(target_day, data_ini_ref)
                                
                                df_target_day = df_r[df_r['Roteiro_Data'] == target_day]
                                df_others = df_r[df_r['Roteiro_Data'] != target_day]
                                
                                df_combined = pd.concat([df_target_day, rows_add], ignore_index=True)
                                rota_opt = otimizar_sequencia_dia(df_combined, col_cep)
                                
                                df_final_day = pd.DataFrame(rota_opt)
                                df_final_day['Roteiro_Ordem'] = range(1, len(df_final_day) + 1)
                                df_final_day = recalcular_distancias_dataframe(df_final_day, modo_rapido=False)
                                
                                df_r = pd.concat([df_others, df_final_day], ignore_index=True)
                                df_s = df_s.drop(real_indices_s)
                                
                                st.session_state['processed_data'][sel_v]['rota'] = df_r
                                st.session_state['processed_data'][sel_v]['sobra'] = df_s
                                st.success(f"Encaixado no dia {target_day.strftime('%d/%m')}!")
                                time.sleep(1)
                                st.rerun()
                        else:
                            st.warning("Nenhuma data válida encontrada no calendário.")
            else:
                st.success("Não há sobras para este vendedor! 🎉")
            
# ==========================================
# ABA 2: VISUALIZADOR OPERACIONAL
# ==========================================
elif nav == "👁️ Visualizador Operacional":
    mostrar_visualizador_operacional()

# ==========================================
# ABA Nova: VISUALIZAR CARTEIRAS
# ==========================================
elif nav == "🗺️ Visualizar Carteiras":
    mostrar_visualizador_carteiras()

# ==========================================
# ABA 3: HISTÓRICO & COMPARAÇÃO
# ==========================================
elif nav == "📂 Histórico & Comparação":
    st.title("📂 Histórico e Navegação (Banco SQL)")
    
    try:
        conn = sqlite3.connect('banco_rotas.db')
        st.info("Consulte aqui as rotas processadas anteriormente no banco de dados.")
        
        query = """
        SELECT 
            id as 'ID', 
            nome_rota as 'Nome da Rota', 
            tipo_vendedor as 'Tipo de Vendedores', 
            usuario_salvou as 'Usuário', 
            data_salvamento as 'Data Salvo', 
            mes_referencia as 'Mês Pertencente', 
            qtd_dias as 'Qtde Dias' 
        FROM rotas_salvas ORDER BY id DESC
        """
        df_hist_sql = pd.read_sql_query(query, conn)
        
        if not df_hist_sql.empty:
            st.dataframe(df_hist_sql, hide_index=True, use_container_width=True)
            st.divider()
            
            st.markdown("### ⚙️ Ações do Banco de Dados")
            c_hist1, c_hist2 = st.columns([1, 2])
            id_selecionado = c_hist1.selectbox("Selecione o ID da Rota:", df_hist_sql['ID'].tolist())
            
            c_btn1, c_btn2 = st.columns(2)
            if c_btn1.button("📂 Carregar Rota para Roteirizador", type="primary", use_container_width=True):
                c = conn.cursor()
                c.execute("SELECT df_rota_json, df_sobra_json, nome_rota FROM rotas_salvas WHERE id = ?", (id_selecionado,))
                row = c.fetchone()
                if row:
                    import json
                    df_r_loaded = pd.DataFrame(json.loads(row[0]))
                    df_s_loaded = pd.DataFrame(json.loads(row[1]))
                    nome_rota_bd = row[2]
                    
                    if 'Roteiro_Data' in df_r_loaded.columns: 
                        df_r_loaded['Roteiro_Data'] = pd.to_datetime(df_r_loaded['Roteiro_Data'], dayfirst=True, format='mixed').dt.date
                    if 'Roteiro_Data' in df_s_loaded.columns: 
                        df_s_loaded['Roteiro_Data'] = pd.to_datetime(df_s_loaded['Roteiro_Data'], dayfirst=True, format='mixed').dt.date
                    
                    novo_processed_data = {}
                    vendedores_recuperados = []
                    
                    if 'Nome_Vendedor' in df_r_loaded.columns: vendedores_recuperados.extend(df_r_loaded['Nome_Vendedor'].dropna().unique().tolist())
                    if 'Nome_Vendedor' in df_s_loaded.columns: vendedores_recuperados.extend(df_s_loaded['Nome_Vendedor'].dropna().unique().tolist())
                    vendedores_recuperados = list(set(vendedores_recuperados))
                    
                    if not vendedores_recuperados: 
                        vendedores_recuperados = ['Vendedor Recuperado']
                        if not df_r_loaded.empty: df_r_loaded['Nome_Vendedor'] = 'Vendedor Recuperado'
                        if not df_s_loaded.empty: df_s_loaded['Nome_Vendedor'] = 'Vendedor Recuperado'

                    for v in vendedores_recuperados:
                        r_vend = df_r_loaded[df_r_loaded['Nome_Vendedor'] == v] if not df_r_loaded.empty else pd.DataFrame()
                        s_vend = df_s_loaded[df_s_loaded['Nome_Vendedor'] == v] if not df_s_loaded.empty else pd.DataFrame()
                        
                        novo_processed_data[v] = {
                            'rota': r_vend,
                            'sobra': s_vend,
                            'coord': f"💾 SQL: {nome_rota_bd}"
                        }

                    st.session_state['processed_data'] = novo_processed_data
                    st.session_state['status_processamento'] = True
                    st.session_state['df_main'] = pd.concat([df_r_loaded, df_s_loaded], ignore_index=True) 
                    st.success("Rota carregada do SQL com sucesso! Vá para a aba '🚀 Roteirizador'.")

            if c_btn2.button("🗑️ Excluir Rota do Banco", use_container_width=True):
                c = conn.cursor()
                c.execute("DELETE FROM rotas_salvas WHERE id = ?", (id_selecionado,))
                conn.commit()
                st.success("Rota Excluída Permanentemente!")
                time.sleep(1)
                st.rerun()
        else:
            st.info("Nenhuma rota salva no banco de dados ainda.")
        conn.close()
    except Exception as e:
        st.error(f"Erro ao conectar ao banco SQL: {e}")
