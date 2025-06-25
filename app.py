import streamlit as st
import pandas as pd
from unidecode import unidecode
import bcrypt
import os
from datetime import datetime

# Configuração básica
st.set_page_config(page_title="Análise de Glosas", layout="wide")

# Dados de usuários (senhas: "unimed123")
usuarios = {
    "thalita.moura": {"senha": "$2b$12$Kai0k60BAGxa5Sc00N6wy.2TZNXiguFlIUKAJBoeQG/tdCrP3O4f.", "perfil": "supervisor"},
    "ana.pereira": {"senha": "$2b$12$gN54bhAbu7oNNTGq4OC3a.EQt6W1NZ2XAzSItR6MDxBxy.ySBfjFu", "perfil": "analista"}
}

# Sistema de login
if 'logado' not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    st.title("🔐 Login")
    usuario = st.selectbox("Usuário:", list(usuarios.keys()))
    senha = st.text_input("Senha:", type="password")
    
    if st.button("Entrar"):
        if bcrypt.checkpw(senha.encode(), usuarios[usuario]["senha"].encode()):
            st.session_state.logado = True
            st.rerun()
        else:
            st.error("Senha incorreta")
    st.stop()

# Página principal (após login)
st.title("🏥 Análise de Glosas")
arquivo = st.file_uploader("Envie o arquivo 549.xlsx", type="xlsx")

if arquivo:
    df = pd.read_excel(arquivo)
    df.columns = [unidecode(str(col)).strip().lower() for col in df.columns]
    
    st.success("Arquivo processado!")
    st.write("Total de registros:", len(df))
    
    if 'valor glosa' in df.columns:
        st.metric("Valor total:", f"R$ {df['valor glosa'].sum():,.2f}")
