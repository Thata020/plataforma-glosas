# === app.py ===
# Plataforma de Análise de Glosas - Unificada com script_master e correcao_arquivo

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from unidecode import unidecode
import bcrypt
import logging
import matplotlib.pyplot as plt
import seaborn as sns
from script_master import processar_glosas
from correcao_arquivo import corrigir_caracteres, processar_549

# === CONFIG ===
st.set_page_config(page_title="Glosas Unimed", layout="wide", page_icon="🏥")

# === LOG ===
logging.basicConfig(filename='auditoria_glosas.log', level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s')

def registrar(usu, acao, detalhes=""):
    logging.info(f"USUARIO: {usu} - ACAO: {acao} - DETALHES: {detalhes}")

# === SEGURANÇA ===
def check_password(senha, senha_hash):
    return bcrypt.checkpw(senha.encode('utf-8'), senha_hash.encode('utf-8'))

usuarios = {
    "thalita.moura": {
        "senha": "$2b$12$Kai0k60BAGxa5Sc00N6wy.2TZNXiguFlIUKAJBoeQG/tdCrP3O4f.",
        "perfil": "supervisor"
    },
    "ana.pereira": {
        "senha": "$2b$12$gN54bhAbu7oNNTGq4OC3a.EQt6W1NZ2XAzSItR6MDxBxy.ySBfjFu",
        "perfil": "analista"
    },
    # outros usuários...
}

if 'auth' not in st.session_state:
    st.session_state.auth = False
    st.session_state.user = None

if not st.session_state.auth:
    st.image("logo_unimed.png", width=180)
    st.title("🔐 Login - Plataforma Glosas")
    user = st.selectbox("Usuário:", list(usuarios.keys()))
    pwd = st.text_input("Senha:", type="password")
    if st.button("Entrar"):
        if check_password(pwd, usuarios[user]["senha"]):
            st.session_state.auth = True
            st.session_state.user = user
            registrar(user, "LOGIN_SUCESSO")
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")
            registrar(user, "LOGIN_FALHA")
    st.stop()

# === FUNÇÃO TRATAMENTO ===
def tratar_glosas(df):
    df.columns = [unidecode(str(c)).strip().lower() for c in df.columns]
    df.dropna(how='all', inplace=True)
    df = df[df['motivo da glosa'].notna()]
    df['motivo da glosa'] = df['motivo da glosa'].str.strip().str.upper()
    if 'prestador' in df.columns:
        df = df[~df['prestador'].str.contains("ISENTO", case=False, na=False)]
    df = corrigir_caracteres(df)
    return df

# === INTERFACE PRINCIPAL ===
st.title("🏥 Análise de Glosas - Unimed")
st.sidebar.success(f"Logado como: {st.session_state.user}")

st.header("📤 Envio de Arquivo .xlsx cru")
file = st.file_uploader("Selecione o arquivo de glosas", type="xlsx")

if file:
    registrar(st.session_state.user, "UPLOAD", file.name)
    try:
        st.info("📂 Carregando o arquivo...")
        df = pd.read_excel(file)

        st.info("🛠 Fazendo correções no arquivo...")
        df = tratar_glosas(df)

        st.info("🔍 Verificando se há glosas com regras...")
        df_resultado, df_resumo = processar_glosas(df)

        if 'data' in df_resultado.columns:
            df_resultado['data'] = pd.to_datetime(df_resultado['data'], errors='coerce')
            df_resultado['mes'] = df_resultado['data'].dt.strftime('%Y-%m')
        df_resultado['id'] = range(1, len(df_resultado) + 1)

        st.success("✅ Arquivo processado com sucesso!")
        for i, row in df_resumo.iterrows():
            st.write(f"🔹 Regra {row['Nº da Regra']}: {row['Qtde Glosas']} glosas ({row['Status']})")

        # DOWNLOAD
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_saida = f"resultado_glosas_{timestamp}.xlsx"
        df_resultado.to_excel(nome_saida, index=False)
        with open(nome_saida, "rb") as f:
            st.download_button("📥 Baixar Arquivo Analisado", f, file_name=nome_saida)
        os.remove(nome_saida)

        # DASHBOARD
        st.subheader("📊 Métricas de Análise")
        col1, col2 = st.columns(2)
        col1.metric("Total de Glosas", len(df_resultado))
        col2.metric("Valor Total", f"R$ {df_resultado['valor glosa'].sum():,.2f}")

        if 'mes' in df_resultado.columns:
            st.subheader("📉 Evolução Mensal")
            evolucao = df_resultado.groupby('mes')['valor glosa'].sum().reset_index()
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.lineplot(data=evolucao, x='mes', y='valor glosa', marker='o', ax=ax)
            ax.set_title("Valor de Glosas por Mês")
            st.pyplot(fig)

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
        registrar(st.session_state.user, "ERRO_PROCESSAMENTO", str(e))
        st.stop()

# === EXECUÇÃO COMPLETA ===
st.header("⚙️ Rodar Análise Completa (Arquivo 549_geral.xlsx)")

if st.button("🚀 Executar Robô de Glosas"):
    try:
        st.info("⏳ Corrigindo o arquivo inicial (549_geral.xlsx)...")
        processar_549("549_geral.xlsx", "Atendimentos_Intercambio.xlsx")
        st.success("✅ Arquivo corrigido com sucesso.")

        st.info("🔍 Aplicando regras de glosas...")
        from script_master import main as rodar_regras
        rodar_regras()
        st.success("✅ Regras aplicadas com sucesso.")

        with open("Relatorio_Final_Unimed_Auditoria.xlsx", "rb") as f:
            st.download_button("📥 Baixar Relatório Final", f, file_name="Relatorio_Final_Unimed_Auditoria.xlsx")

        st.success("🏁 Processo completo finalizado!")

    except Exception as e:
        st.error(f"❌ Erro durante a execução: {e}")
        registrar(st.session_state.user, "ERRO_ROBO_COMPLETO", str(e))

# === RODAPÉ ===
st.markdown("---")
st.caption("Desenvolvido por Contas Médicas - Unimed | Versão 3.0")
