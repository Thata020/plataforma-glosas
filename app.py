# === app.py ===
# Plataforma de Análise de Glosas - Unificada com DeepSeek + processamento3.py

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from unidecode import unidecode
import bcrypt
import logging
import time
import matplotlib.pyplot as plt
import seaborn as sns
from script_master import processar_glosas
from correcao_arquivo import corrigir_caracteres



# === CONFIG INICIAL ===
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
    "ana.santos": {
        "senha": "$2b$12$JHgqmOF6S7wvy.PDmAsYAeMFSLmyKMAWQ8a.yveHPI2Dnn/RARuNe",
        "perfil": "analista"
    },
    "idayane.oliveira": {
        "senha": "$2b$12$o28V.P8XgGMjZI2zPWpEzuuQYUeOAFocNFJBoEn/aEU1GI21tzO7C",
        "perfil": "analista"
    },
    "bruna.silva": {
        "senha": "$2b$12$P6TyKgFI6DE4iMxX1CaiWeVBdwaoBRzCbYb/jmy0IfpV3l07IWBlS",
        "perfil": "analista"
    },
    "mariana.cunha": {
        "senha": "$2b$12$ETPYnBoSBWy5UHr5D8OH3Ocd4tbsg87xVIRngJJPE/1/gdN4B6ceG",
        "perfil": "analista"
    },
    "weslane.martins": {
        "senha": "$2b$12$wM8EFz7MGTEriaQrK/rCielFw3.kh6gbeKc61/Etr5qKYk93tSfSW",
        "perfil": "analista"
    },
    "riquelme": {
        "senha": "$2b$12$evNWwq8om43/m0Bgf3SmMendfvIvTOLo8o3au0DFxD/Xa9iCPLWf.",
        "perfil": "analista"
    },
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

    # Aplica a correção de caracteres em todas as colunas de texto
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(corrigir_caracteres)

    return df


# === UPLOAD E PROCESSAMENTO ===
st.header("📤 Envio de Arquivo .xlsx cru")
file = st.file_uploader("Selecione o arquivo 549.xlsx", type="xlsx")

if file:
    registrar(st.session_state.user, "UPLOAD", file.name)

    try:
        st.info("📂 Carregando o arquivo...")
        df = pd.read_excel(file)

        st.info("🛠 Fazendo correções no arquivo...")
        df = tratar_glosas(df)

        st.info("🔍 Verificando se há glosas...")
        df = processar_glosas(df)

    except Exception as e:
        st.error("❌ Erro ao processar o arquivo.")
        registrar(st.session_state.user, "ERRO_PROCESSAMENTO", str(e))
        st.stop()

    if 'data' in df.columns:
        df['data'] = pd.to_datetime(df['data'], errors='coerce')
        df['mes'] = df['data'].dt.strftime('%Y-%m')
    df['id'] = range(1, len(df) + 1)

    resumo = df['motivo da glosa'].value_counts().reset_index()
    resumo.columns = ['Motivo', 'Qtd']
    st.success("✅ Arquivo processado com sucesso!")

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
