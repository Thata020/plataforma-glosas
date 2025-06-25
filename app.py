# === app.py ===
# Plataforma de Análise de Glosas - Unificada com DeepSeek + processamento3.py

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from unidecode import unidecode
import logging
import time
import matplotlib.pyplot as plt
import seaborn as sns

# === CONFIG INICIAL ===
st.set_page_config(page_title="Glosas Unimed", layout="wide", page_icon="🏥")

# === LOG ===
logging.basicConfig(filename='auditoria_glosas.log', level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s')

def registrar(usu, acao, detalhes=""):
    logging.info(f"USUARIO: {usu} - ACAO: {acao} - DETALHES: {detalhes}")

# === SEGURANÇA SIMPLIFICADA ===

usuarios = {
    "thalita.moura": {
        "senha": "1234",
        "perfil": "supervisor"
    },
    "ana.pereira": {
        "senha": "1234",
        "perfil": "analista"
    },
    "ana.santos": {
        "senha": "1234",
        "perfil": "analista"
    },
    "idayane.oliveira": {
        "senha": "1234",
        "perfil": "analista"
    },
    "bruna.silva": {
        "senha": "1234",
        "perfil": "analista"
    },
    "mariana.cunha": {
        "senha": "1234",
        "perfil": "analista"
    },
    "weslane.martins": {
        "senha": "1234",
        "perfil": "analista"
    },
    "riquelme": {
        "senha": "1234",
        "perfil": "analista"
    },
}

def check_password(senha_digitada, senha_salva):
    return senha_digitada == senha_salva


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

# === INTERFACE PRINCIPAL ===
st.title("🏥 Análise de Glosas - Unimed")
st.sidebar.success(f"Logado como: {st.session_state.user}")

# === FUNÇÃO DE TRATAMENTO (processamento3.py embutido) ===
def tratar_glosas(df):
    df.columns = [unidecode(str(c)).strip().lower() for c in df.columns]
    df.dropna(how='all', inplace=True)
    df = df[df['motivo da glosa'].notna()]
    df['motivo da glosa'] = df['motivo da glosa'].str.strip().str.upper()
    df = df[~df['motivo da glosa'].isin(['REAPRESENTACAO', 'CODIGO REMOVIDO'])]
    if 'prestador' in df.columns:
        df = df[~df['prestador'].str.contains("ISENTO", case=False, na=False)]
    return df

# === UPLOAD E PROCESSAMENTO ===
st.header("📤 Envio de Arquivo .xlsx cru")
file = st.file_uploader("Selecione o arquivo 549.xlsx", type="xlsx")

if file:
    registrar(st.session_state.user, "UPLOAD", file.name)
    try:
        df = pd.read_excel(file)
        df = tratar_glosas(df)

        if 'data' in df.columns:
            df['data'] = pd.to_datetime(df['data'], errors='coerce')
            df['mes'] = df['data'].dt.strftime('%Y-%m')
        df['id'] = range(1, len(df) + 1)

        resumo = df['motivo da glosa'].value_counts().reset_index()
        resumo.columns = ['Motivo', 'Qtd']
        st.success("Arquivo processado com sucesso!")

        for i, row in resumo.iterrows():
            st.write(f"🔹 {row['Qtd']} glosas encontradas: {row['Motivo']}")

        # === DOWNLOAD XLS ===
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_saida = f"resultado_glosas_{timestamp}.xlsx"
        df.to_excel(nome_saida, index=False)
        with open(nome_saida, "rb") as f:
            st.download_button("📥 Baixar Arquivo Analisado", f, file_name=nome_saida)
        os.remove(nome_saida)

        # === DASHBOARD ===
        st.subheader("📊 Métricas de Análise")
        col1, col2 = st.columns(2)
        col1.metric("Total de Glosas", len(df))
        col2.metric("Valor Total", f"R$ {df['valor glosa'].sum():,.2f}")

        st.subheader("📉 Evolução Mensal de Glosas")
        if 'mes' in df.columns:
            evolucao = df.groupby('mes')['valor glosa'].sum().reset_index()
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.lineplot(data=evolucao, x='mes', y='valor glosa', marker='o', ax=ax)
            ax.set_title("Valor de Glosas por Mês")
            st.pyplot(fig)

    except Exception as e:
        st.error(f"Erro ao processar: {str(e)}")
        registrar(st.session_state.user, "ERRO_PROCESSAMENTO", str(e))

# === RODAPÉ ===
st.markdown("---")
st.caption("Desenvolvido por Contas Médicas - Unimed | Versão 3.0")
