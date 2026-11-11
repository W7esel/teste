import re
import sys
import warnings
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, date, time
from functools import lru_cache, wraps
from pathlib import Path
import os

import hdbcli.dbapi as dbapi
import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURAÇÕES VISUAIS STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Painel de Apresentação",
    layout="wide"
)

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0B3D2E;
    }
    h1, h2, h3, h4, h5, h6, p, span, label {
        color: white !important;
    }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
        background-color: white;
        border-radius: 10px;
        padding: 5px;
    }
    /* Deixa a fonte da tabela preta para dar contraste no fundo branco */
    .stDataFrame [data-testid="dataGridCell"] {
        color: black !important; 
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

warnings.filterwarnings("ignore", message=".*SettingWithCopyWarning.*")
warnings.simplefilter(action="ignore", category=UserWarning)

ARQUIVO_CONEXAO = Path(r"C:\Users\B623896\Desktop\CONCT\dados_conexao.xlsx")

# -------------------------------------------------------------
# FUNÇÕES DE CONEXÃO
# -------------------------------------------------------------

def testar_conexao(conn):
    cursor = None
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT CURRENT_DATE FROM DUMMY")
        resultado = cursor.fetchone()
        st.sidebar.success("Conexão SAP OK")
        return True
    except Exception as erro:
        raise ConnectionError(f"Falha ao testar a conexão: {erro}") from erro
    finally:
        if cursor is not None:
            cursor.close()

def carregar_configuracao_conexao():
    if not ARQUIVO_CONEXAO.exists():
        raise FileNotFoundError(f"O arquivo de conexão não foi encontrado:\n{ARQUIVO_CONEXAO}")
    
    df_conexao = pd.read_excel(ARQUIVO_CONEXAO)
    colunas_obrigatorias = {"host", "port", "user", "password", "databasename"}
    colunas_ausentes = colunas_obrigatorias - set(df_conexao.columns)
    
    if colunas_ausentes:
        raise KeyError(f"O arquivo de conexão não possui as colunas: {sorted(colunas_ausentes)}")
    if df_conexao.empty:
        raise ValueError("O arquivo de conexão não possui registros.")
        
    return {
        "host": df_conexao.loc[0, "host"],
        "port": int(df_conexao.loc[0, "port"]),
        "user": df_conexao.loc[0, "user"],
        "password": df_conexao.loc[0, "password"],
        "databasename": df_conexao.loc[0, "databasename"],
    }

def abrir_conexao_hana():
    configuracao = carregar_configuracao_conexao()
    conn = connect_to_hana(
        host=configuracao["host"],
        port=configuracao["port"],
        user=configuracao["user"],
        password=configuracao["password"],
        databasename=configuracao["databasename"]
    )
    testar_conexao(conn)
    return conn

def connect_to_hana(host, port, user, password, databasename):
    try:
        conn = dbapi.connect(
            address=host,
            port=port,
            user=user,
            password=password,
            databasename=databasename,
            sslValidateCertificate=False
        )
        conn.setautocommit(False)
        return conn
    except dbapi.Error as erro:
        raise ConnectionError(f"Erro ao conectar ao SAP HANA: {erro}") from erro

# -------------------------------------------------------------
# EQUIPES FIXAS
# -------------------------------------------------------------

EQUIPES_INFO = [
    {"utd": "FEIRA SUL", "equipe": "VTN2GCFSS01S"},
    {"utd": "SERRINHA", "equipe": "STC2GCSRR014"},
    {"utd": "JUAZEIRO", "equipe": "VTN2GCJZB00C"},
    {"utd": "REMANSO", "equipe": "VTN2GCREM015"},
    {"utd": "REMANSO", "equipe": "VTN2GCREM00D"},
    {"utd": "PAF", "equipe": "VTN2GCPAF007"},
    {"utd": "RIBEIRA DO POMBAL", "equipe": "VTN2GCRPB00G"},
    {"utd": "IRECÊ", "equipe": "VTN2GCIRC00K"},
    {"utd": "ITABERABA", "equipe": "VTN2GCITR007"},
    {"utd": "JEQUIÉ", "equipe": "VTN2GCJEQ008"},
    {"utd": "SEABRA", "equipe": "STC2GCSEA00T"},
    {"utd": "BARREIRAS", "equipe": "VTN2GCBRE00C"},
    {"utd": "LEM", "equipe": "VTN2GCLEM00E"},
    {"utd": "SAJ", "equipe": "VTN2GCSAJ00H"},
    {"utd": "VALENÇA", "equipe": "VTN2GCVAL00B"},
    {"utd": "EUNÁPOLIS", "equipe": "VTN2GCEUN00F"},
    {"utd": "ITABUNA", "equipe": "VTN2GCITN00E"},
    {"utd": "ILHÉUS", "equipe": "STC2GCILH017"},
    {"utd": "PSG", "equipe": "VTN2GCPSG00D"},
    {"utd": "TEF", "equipe": "VTN2GCTEF00B"},
]

# -------------------------------------------------------------
# LÓGICA DE CONSULTA NO BANCO
# -------------------------------------------------------------

def consultar_dados_periodo(conn, data_inicio, data_fim):
    cursor = conn.cursor()
    try:
        lista_equipes = ",".join(f"'{item['equipe']}'" for item in EQUIPES_INFO)
        
        inicio_str = data_inicio.strftime('%Y-%m-%d')
        fim_str = data_fim.strftime('%Y-%m-%d')

        # 1. LOGON DAS EQUIPES (Período Selecionado)
        sql_logon = f"""
        SELECT EQUIPE, INICIO_TURNO
        FROM U619439.TAB_TURNOS_WFM_TEMPO_REAL
        WHERE EQUIPE IN ({lista_equipes})
        AND TO_DATE(INICIO_TURNO) BETWEEN '{inicio_str}' AND '{fim_str}'
        """
        cursor.execute(sql_logon)
        registros_logon = cursor.fetchall()
        colunas_logon = [col[0] for col in cursor.description]
        df_logon = pd.DataFrame(registros_logon, columns=colunas_logon)
        if not df_logon.empty:
            df_logon["INICIO_TURNO"] = pd.to_datetime(df_logon["INICIO_TURNO"], errors="coerce")

        # 2. NOTAS DA GESTÃO CAMPO (Do mês da data inicial)
        mes_alvo = data_inicio.month
        ano_alvo = data_inicio.year
        
        sql_notas = f"""
        SELECT EQUIPE, COUNT(*) AS TOTAL_NOTAS
        FROM U619439.TAB_TAREFA_ENCERRADA_GESTAO_CAMPO
        WHERE EQUIPE IN ({lista_equipes})
          AND (CALL_ID LIKE 'CV48%' OR CALL_ID LIKE '0WOR%')
          AND YEAR(DIA) = {ano_alvo}
          AND MONTH(DIA) = {mes_alvo}
        GROUP BY EQUIPE
        """
        cursor.execute(sql_notas)
        df_notas = pd.DataFrame(cursor.fetchall(), columns=[col[0] for col in cursor.description])
        notas_por_equipe = dict(zip(df_notas["EQUIPE"], df_notas["TOTAL_NOTAS"])) if not df_notas.empty else {}

        # 3. LÓGICA DO DIA ÚTIL ANTERIOR
        hoje = datetime.now().date()
        if hoje.weekday() == 0: 
            dia_anterior = hoje - timedelta(days=3)
        elif hoje.weekday() == 6: 
            dia_anterior = hoje - timedelta(days=2)
        else: 
            dia_anterior = hoje - timedelta(days=1)
            
        sql_produtividade = f"""
        SELECT COUNT(*) AS TOTAL_NOTAS_ONTEM
        FROM U619439.TAB_TAREFA_ENCERRADA_GESTAO_CAMPO
        WHERE EQUIPE IN ({lista_equipes})
          AND (CALL_ID LIKE 'CV48%' OR CALL_ID LIKE '0WOR%')
          AND TO_DATE(DIA) = '{dia_anterior.strftime('%Y-%m-%d')}'
        """
        cursor.execute(sql_produtividade)
        resultado_ontem = cursor.fetchone()
        alvos_dia_anterior = resultado_ontem[0] if resultado_ontem and resultado_ontem[0] else 0

        # 4. ALVOS PROSPECTADOS POR DIA (Para a linha de totalizador)
        sql_diario = f"""
        SELECT TO_DATE(DIA) AS DATA_DIA, COUNT(*) AS TOTAL_NOTAS
        FROM U619439.TAB_TAREFA_ENCERRADA_GESTAO_CAMPO
        WHERE EQUIPE IN ({lista_equipes})
          AND (CALL_ID LIKE 'CV48%' OR CALL_ID LIKE '0WOR%')
          AND TO_DATE(DIA) BETWEEN '{inicio_str}' AND '{fim_str}'
        GROUP BY TO_DATE(DIA)
        """
        cursor.execute(sql_diario)
        df_diario = pd.DataFrame(cursor.fetchall(), columns=[col[0] for col in cursor.description])
        alvos_por_dia = {}
        if not df_diario.empty:
            df_diario["DATA_DIA"] = pd.to_datetime(df_diario["DATA_DIA"]).dt.date
            alvos_por_dia = dict(zip(df_diario["DATA_DIA"], df_diario["TOTAL_NOTAS"]))

        # 5. TABELA ESTRUTURADA
        dias_selecionados = []
        delta = data_fim - data_inicio
        for i in range(delta.days + 1):
            dias_selecionados.append(data_inicio + timedelta(days=i))

        linhas = []
        equipes_logadas_hoje = 0

        for info in EQUIPES_INFO:
            equipe = info["equipe"]
            linha = {
                "UTD": info["utd"],
                "EQUIPE": equipe,
                "ALVOS PROSPECTADOS (MTD)": notas_por_equipe.get(equipe, 0)
            }

            datas_logadas = set()
            if not df_logon.empty:
                dados_equipe = df_logon[df_logon["EQUIPE"] == equipe]
                datas_logadas = set(dados_equipe["INICIO_TURNO"].dt.date.dropna())
            
            if hoje in datas_logadas:
                equipes_logadas_hoje += 1

            for dia in dias_selecionados:
                linha[dia.strftime("%d/%m/%Y")] = "SIM" if dia in datas_logadas else "NÃO"

            linhas.append(linha)

        linha_total = {
            "UTD": "-",
            "EQUIPE": "TOTAL DIÁRIO:",
            "ALVOS PROSPECTADOS (MTD)": sum(notas_por_equipe.values()) 
        }
        for dia in dias_selecionados:
            total_dia = alvos_por_dia.get(dia, 0)
            linha_total[dia.strftime("%d/%m/%Y")] = f"{total_dia} alvos"
            
        linhas.append(linha_total)

        taxa_apresentacao = (equipes_logadas_hoje / len(EQUIPES_INFO)) * 100
        taxa_produtividade = (alvos_dia_anterior / 40.0) * 100 
        
        df_final = pd.DataFrame(linhas)
        return df_final, taxa_apresentacao, taxa_produtividade, hoje, dia_anterior

    finally:
        cursor.close()

# =========================================================
# FUNÇÃO PARA APLICAR CORES NAS CÉLULAS DA TABELA
# =========================================================
def colorir_celulas(val):
    val_str = str(val).upper()
    if val_str == "SIM":
        return 'background-color: #d4edda; color: #155724' 
    elif val_str == "NÃO":
        return 'background-color: #f8d7da; color: #721c24' 
    elif "ALVOS" in val_str or "ALVO" in val_str:
        return 'background-color: #005a2b; color: white; font-weight: bold;' 
    return ''

# =========================================================
# FUNÇÃO PARA DESENHAR O KPI CIRCULAR COM CSS
# =========================================================
def gerar_kpi_circular(titulo, porcentagem, data_ref):
    cor_borda = "#F39C12" 

    html = f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-bottom: 20px; padding: 20px; background-color: #145A32; border-radius: 15px; border: 1px solid #1E8449;">
        <h4 style="color: white; margin-bottom: 20px; text-align: center; font-size: 18px;">{titulo}<br><span style="font-size: 14px; color: #AAB7B8;">({data_ref})</span></h4>
        <div style="width: 140px; height: 140px; border-radius: 50%; border: 12px solid {cor_borda}; display: flex; align-items: center; justify-content: center; background-color: #0B3D2E; box-shadow: 0 0 15px {cor_borda};">
            <span style="font-size: 32px; font-weight: bold; color: white;">{porcentagem:.1f}%</span>
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

# -------------------------------------------------------------
# INTERFACE PRINCIPAL
# -------------------------------------------------------------

st.title("Painel Gerencial de Turmas - WFM")
st.markdown("---")

conn = abrir_conexao_hana()

st.subheader("1. Selecione o Período de Análise")
col_inicio, col_fim, col_btn = st.columns([1, 1, 2])

hoje_padrao = datetime.now().date()
segunda_padrao = hoje_padrao - timedelta(days=hoje_padrao.weekday())

with col_inicio:
    data_inicio = st.date_input("Data Inicial", value=segunda_padrao, format="DD/MM/YYYY")
with col_fim:
    data_fim = st.date_input("Data Final", value=hoje_padrao, format="DD/MM/YYYY")

with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    botao_buscar = st.button("Extrair Dados do SAP HANA", use_container_width=True)

if data_inicio > data_fim:
    st.error("A Data Inicial não pode ser maior que a Data Final.")
    st.stop()

# =========================================================
# LÓGICA DE ESTADO (Para manter tabela editável viva na tela)
# =========================================================

if "dados" not in st.session_state:
    st.session_state.dados = pd.DataFrame()
    st.session_state.taxa_apres = 0.0
    st.session_state.taxa_prod = 0.0
    st.session_state.dt_hoje = ""
    st.session_state.dt_ontem = ""

if botao_buscar:
    with st.spinner('Consultando servidor SAP HANA...'):
        df, t_apres, t_prod, dt_hj, dt_on = consultar_dados_periodo(conn, data_inicio, data_fim)
        st.session_state.dados = df
        st.session_state.taxa_apres = t_apres
        st.session_state.taxa_prod = t_prod
        st.session_state.dt_hoje = dt_hj.strftime("%d/%m/%Y")
        st.session_state.dt_ontem = dt_on.strftime("%d/%m/%Y")

# =========================================================
# RENDERIZAÇÃO DA TABELA E KPIS
# =========================================================

if not st.session_state.dados.empty:
    st.markdown("---")
    st.subheader("2. Quadro de Apresentação (Editável)")
    
    st.info("Dê um duplo-clique nas células de ALVOS ou DATAS para digitar correções e justificativas.")
    
    col_tabela, col_kpis = st.columns([3, 1])
    
    with col_tabela:
        colunas_datas = [col for col in st.session_state.dados.columns if col not in ["UTD", "EQUIPE", "ALVOS PROSPECTADOS (MTD)"]]
        df_estilizado = st.session_state.dados.style.map(colorir_celulas, subset=colunas_datas)

        # Retorna o DataFrame com as edições feitas pelo usuário na tela
        df_editado = st.data_editor(
            df_estilizado,
            disabled=["UTD", "EQUIPE"], 
            use_container_width=True,
            hide_index=True,
            height=860 
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        csv = df_editado.to_csv(index=False, sep=";").encode('utf-8-sig')
        st.download_button(
            label="Baixar Tabela Editada (Excel/CSV)",
            data=csv,
            file_name=f"Painel_Turmas_{data_inicio.strftime('%d-%m')}_a_{data_fim.strftime('%d-%m')}.csv",
            mime="text/csv"
        )
        
    with col_kpis:
        # -----------------------------------------------------
        # CÁLCULO DINÂMICO DOS KPIS BASEADO NAS EDIÇÕES DA TELA
        # -----------------------------------------------------
        taxa_apres_dinamica = st.session_state.taxa_apres
        taxa_prod_dinamica = st.session_state.taxa_prod
        
        col_hoje = st.session_state.dt_hoje
        col_ontem = st.session_state.dt_ontem
        equipes_ausentes_hoje = []
        
        # Recalcula Apresentação e captura ausentes se a coluna de hoje estiver na tabela editada
        if col_hoje in df_editado.columns:
            df_equipes = df_editado[df_editado["EQUIPE"] != "TOTAL DIÁRIO:"]
            
            # Filtra e lista quem está com "NÃO" no dia de hoje
            equipes_ausentes_hoje = df_equipes[df_equipes[col_hoje].astype(str).str.strip().str.upper() == "NÃO"]["EQUIPE"].tolist()
            
            # Conta os "SIM" para a taxa
            qtd_sim = df_equipes[col_hoje].astype(str).str.strip().str.upper().eq("SIM").sum()
            taxa_apres_dinamica = (qtd_sim / len(EQUIPES_INFO)) * 100
            
        # Recalcula Produtividade se a coluna de ontem estiver na tabela editada
        if col_ontem in df_editado.columns:
            linha_total = df_editado[df_editado["EQUIPE"] == "TOTAL DIÁRIO:"]
            if not linha_total.empty:
                valor_celula = str(linha_total.iloc[0][col_ontem])
                numeros = re.findall(r'\d+', valor_celula)
                if numeros:
                    alvos_ontem = int(numeros[0])
                    taxa_prod_dinamica = (alvos_ontem / 40.0) * 100
                else:
                    taxa_prod_dinamica = 0.0

        # Renderiza os KPIs 
        gerar_kpi_circular(
            "Taxa de Apresentação", 
            taxa_apres_dinamica, 
            st.session_state.dt_hoje
        )
        
        gerar_kpi_circular(
            "Taxa de Produtividade", 
            taxa_prod_dinamica, 
            st.session_state.dt_ontem
        )
        
        # -----------------------------------------------------
        # CAIXA DINÂMICA: LISTA DE EQUIPES AUSENTES HOJE
        # -----------------------------------------------------
        if col_hoje in df_editado.columns:
            html_lista = f"""
            <div style="padding: 15px; background-color: #145A32; border-radius: 15px; border: 1px solid #1E8449;">
                <h5 style="color: white; margin-bottom: 10px; text-align: center; font-size: 16px;">
                    Equipes que não logaram no WFM hoje:
                </h5>
                <ul style="color: white; font-size: 14px; list-style-type: none; padding-left: 0; text-align: center;">
            """
            
            if equipes_ausentes_hoje:
                for eq in equipes_ausentes_hoje:
                    html_lista += f"<li style='margin-bottom: 5px; font-weight: bold; color: #F39C12;'>{eq}</li>"
            else:
                html_lista += "<li style='color: #28B463; font-weight: bold;'>Todas as equipes logadas!</li>"
            
            html_lista += "</ul></div>"
            st.markdown(html_lista, unsafe_allow_html=True)
