import pandas as pd
import numpy as np
import os
import sys
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
from datetime import datetime

# ===================================================
# LISTA COMPLETA DE TODAS AS REGRAS
# ===================================================
TODAS_AS_REGRAS = {
    "R01": "Consulta PA < 24h",
    "R02": "Consulta Retorno < 20 dias", 
    "R03": "US + Doppler",
    "R04": "Procedimentos Duplicados",
    "R05": "SADT acima do permitido",
    "R06": "SADT Seriado",
    "R07": "Fracionamento TC/RM",
    "R08": "Procedimentos Excludentes",
    "R09": "Controle de valor liberado",
    "R10": "HM não pode ser associados",
    "R11": "Pacote fracionamento incorreto",
    "R12": "Fracionamento de Honorário Médico",
    "R13": "Colesterol - Painel não permitido",
    "R14": "Consulta de psicoterapia cobrada mais de 1x no dia",
    "R15": "Atendimento fisiátrico (20103093) max.1x exceto UTI",
    "R16": "Consulta e sessão de psicoterapia cobradas na mesma data",
    "R17": "Intensivista Diarista 10104020: valor e quantidade excessiva",
    "R18": "Intensivista Diarista 10104011: valor e quantidade excessiva",
    "R19": "Probióticos não devem ser pagos no Intercâmbio",
    "R20": "Consulta Eletiva x Puericultura",
    "R21": "Fotodermatoscopia 41301234 máximo 1x por dia",
    "R22": "Fotodermatoscopia 41301234 x Dermatoscopia 41301137",
    "R23": "Dermatosocopia 41301137 máximo 1x por dia",
    "R24": "Acupuntura 31601014 x Estimulação/Infiltração",
    "R25": "Procedimentos não devem ocorrer com consulta",
    "R26": "Infiltração 20103301: máximo 2x por dia",
    "R27": "Facectomia com dobra de apartamento", 
}
# =======================================
# CONFIGURAÇÕES GLOBAIS
# =======================================

# Detecta se está rodando como .exe (empacotado)
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE  = os.path.join(BASE_DIR, "Atendimentos_Intercambio.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "Relatorio_Final_Unimed_Auditoria.xlsx")
GPT_FILE    = os.path.join(BASE_DIR, "Tabelas de base GPT.xlsx")

# Validação de existência do Excel
if not os.path.exists(GPT_FILE):
    print(f"❌ Arquivo não encontrado: {GPT_FILE}")
    input("Pressione Enter para sair...")
    exit()

# Leitura do Excel
df = pd.read_excel(GPT_FILE)
    
# =============================================
# FUNÇÕES AUXILIARES
# =============================================
glosas = pd.DataFrame()

def registrar_glosa(df_glosado, cod_regra, nome_regra, motivo):
    df_glosado = df_glosado.copy()
    if "Competência" not in df_glosado.columns and "Competência" in df.columns:
        df_glosado["Competência"] = df_glosado["Carteirinha"].map(df.set_index("Carteirinha")["Competência"].to_dict())
    df_glosado["Nº da Regra"] = cod_regra
    df_glosado["Nome da Regra"] = nome_regra
    df_glosado["Motivo da Glosa"] = motivo
    return len(df_glosado), df_glosado

def carregar_dados():
    try:
        df_gpt_tc = pd.read_excel(GPT_FILE, sheet_name="TC Intercâmbio")
        df_gpt_rm = pd.read_excel(GPT_FILE, sheet_name="RM Intercâmbio")
    except Exception as e:
        raise ValueError(f"Erro ao carregar arquivo GPT: {str(e)}")

    df = pd.read_excel(INPUT_FILE)
    df.columns = [col.strip().title() for col in df.columns]  # Padroniza nomes

    padrao_colunas = {
        "carteirinha": "Carteirinha",
        "nome beneficiario": "Nome beneficiario",
        "dt procedimento": "Dt Procedimento",
        "hora proc": "Hora Proc",
        "cd procedimento": "Cd Procedimento",
        "descricao": "Descricao",
        "quantidade": "Quantidade",
        "vl unitario": "Vl Unitario",
        "vl liberado": "Vl Liberado",
        "vl calculado": "Vl Calculado",
        "vl anestesista": "Vl Anestesista",
        "vl medico": "Vl Medico",
        "vl custo operacional": "Vl Custo Operacional",
        "vl filme": "Vl Filme",
        "tipo guia": "Tipo Guia",
        "via acesso": "Via Acesso",
        "taxa item": "Taxa Item",
        "grau participantes": "Grau Participantes",
        "tipo receita": "Tipo Receita",
        "executante intercambio": "Executante Intercambio",
        "nr sequencia conta": "Nr Sequencia Conta",
        "status conta": "Status Conta",
        "competencia apresentacao": "Competencia Apresentacao"
    }
    df.columns = [padrao_colunas.get(col.strip().lower(), col.strip()) for col in df.columns]

    if "Competência" not in df.columns:
        if "Competencia Apresentacao" in df.columns:
            df["Competência"] = pd.to_datetime(df["Competencia Apresentacao"], errors="coerce").dt.strftime("%m/%Y")
        else:
            df["Competência"] = ""

    df.fillna({"Hora Proc": "00:00:00"}, inplace=True)
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], dayfirst=True, errors="coerce")
    df["Datahora"] = pd.to_datetime(df["Dt Procedimento"].astype(str) + ' ' + df["Hora Proc"].astype(str), errors='coerce')

    return df, df_gpt_tc, df_gpt_rm

# ================================
# 🔧 Exceções para Rede Master (Alto Custo)
# ================================
EXCECOES_REDE_MASTER = {
    "ASSOC HOSPITALAR MOINHOS DE VENTO"
    # Adicione outros nomes aqui conforme necessário
}

def aplicar_excecao_rede_master(df):
    df = df.copy()
    df["Executante Intercambio"] = df["Executante Intercambio"].astype(str).str.upper().str.strip()
    return df[~df["Executante Intercambio"].isin(EXCECOES_REDE_MASTER)]

def aplicar_excecao(df):
    prestadores_excecao = {
        "HOSPITAL INFANTIL PEQUENO PRINCIPE",
        "IRMandade DA SANTA CASA DE MISERICORDIA E MAT DRACENA",
        "SANTA CASA DE MISERICÓRDIA DE GUARAREMA",
        "SANTA CASA DE MISERICÓRDIA DE CRUZEIRO",
        "HOSPITAL SAO CAMILO",
        "HOSPITAL SAO FRANCISCO",
        "ASSOCIACAO BENEFICENTE HOSPITALAR SAO CAMILO - HOSPITAL PERITIBA",
        "HOSPITAL BOM JESUS",
        "ASSOCIACAO EDUCACIONAL E CARITATIVA - HOSPITAL REGIONAL SAO PAULO",
        "ASS DE CARIDADE S VICENTE DE PAULO - HOSPITAL SAO VICENTE DE PAULO",
        "PIO SOCIADICO DAS DAMAS DE CARIDADE DE CAXIAS DO SUL - HOSPITAL POMPEIA",
        "ORDEM AUXILIADORA DE SENHORAS EVANGELICAS DE NOVA PETROPOLIS - HOSPITAL NOVA PETROPOLIS",
        "HOSPITAL E MATERNIDADE 13 DE MAIO VILA ROMANA S/A",
        "FUNDACAO LUVERDENSE DE SAUDE (CLINICA E HOSPITAL)",
        "SOCIEDADE HOSPITALAR BERTINETTI LTDA",
        "HOSPITAL E MATERNIDADE DOIS PINHEIROS LTDA",
        "HOSPITAL BENEFICENCIA DE JUINA LTDA",
        "FUNDACAO DE SAUDE COMUNITARIA DE SINOP",
        "SOCIEDADE MEDICA SAO LUCAS LTDA (HOSPITAL E CLINICA)",
        "FUNDACAO FILANT E BENEF DE SAUDE ARNALDO GAVAZZA FILHO - HOSPITAL ARNALDO GAVAZZA",
        "IRMANDADE DO HOSPITAL DE NOSSA SENHORA DAS DORES",
        "HOSPITAL MEMORIAL ARCOVERDE LTDA",
        "HOSPITAL SAO PEDRO",
        "INSTITUTO DAS PEQUENAS MISSIONARIAS DE MARIA IMACULADA (HOSPITAL E MATERNIDADE MARIALE KONDER BORNHAUSEN)",
        "HOSPITAL MEMORIAL DE GOIANIA",
        "REDE D'OR SAO LUIZ S.A. - HOSPITAL ESPERANCA OLINDA",
        "CONFERENCIA SAO JOAO DE AVAÍ - HOSPITAL SAO JOSE DO AVAI",
        "HOSPITAL DE CARIDADE DE VARGEM GRANDE DO SUL",
        "SEPACO - HOSPITAL PAINEIRAS",
        "CENTRO NEONATAL TERAPIA INTENSIVA LTDA - MATERNIDADE SAO FRANCISCO",
        "CASA DE SAUDE SANTA RITA DE CASSIA - HOSPITAL DE CLINICAS ALAMEDA",
        "CENTRO DE OLHOS AVENIDA SETE DE SETEMBRO LTDA",
        "CLINICA SAO GONCALO LTDA - HOSPITAL E CLINICA SAO GONCALO",
        "INSTITUTO AVALONCIAGA ICARAI LTDA",
        "CASA DE SAUDE NOSSA SENHORA AUXILIADORA",
        "INSTITUICAO ADVENTISTA SETE BRAS. DE PREV E ASS. A SAUDE - HOSPITAL SILVESTRE",
        "CLINICA SAO GONCALO LTDA - HOSPITAL ICARAI",
        "HOSPITAL OFTALMOLASER SAO GONCALO LTDA - H. OLHOS SAO GONCALO",
        "FILI - OFTALMOCLINICA ICARAI LTDA",
        "FILI - CENTRO DE OLHOS AVENIDA SETE DE SETEMBRO LTDA",
        "HOSPITAL DO CORACAO SAMCORIDIS",
        "CENTRO ORTOPEDICO SAO LUCAS LTDA",
        "Coelho e Bastos/ Inst Fluminense",
        "Clinica Pater",
        "CLÍNICA EGO",
        "Clínica de Hemoterapia",
        "Serum/ GSH",
        "Cdn Clínica De Densitometria Óssea De Niterói Ltda.",
        "Cintila Icaraí Exames Especializados Ltda",
        "Fibroendoscopia Limitada",
        "Orto Trauma Ortopedia E Traumatologia",
        "Gastrotech Ltda",
        "Sugsa - Serviço De Ultrassonografia Alcantara",
        "Serviço Radiológico Gonçalense Ltda",
        "Doppler Serviços Médicos",
        "Perinatal",
        "Oftalmos Reunidos",
        "CLINICAS PIRES DE MELO",
        "INSTITUTO DE UROLOGIA E NEFROLOGIA LTDA",
        "SANTA CASA DE MISERICÓRDIA DA IRMANDADE SENHOR DOS PASSOS DE UBATUBA",
        "INSTITUTO DOUTOR FEITOSA",
        "FUNDAÇÃO ASSISTENCIAL VICOSENSE",
        "SOS CARDIO SERVIÇOS HOSPITALARES LTDA",
        "SANTA CASA DE MISERICÓRDIA DE MACEIÓ",
        "HOSPITAL MEMORIAL ARTHUR RAMOS",
        "HOSPITAL REGIONAL DE PALMITOS",
        "FUNDAÇÃO DE ENSINO SUPERIOR DO VALE DO SAPUCAÍ",
        "HOSPITAL CALIXTO MIDLEJ FILHO",
        "HOSPITAL MANOEL NOVAES",
        "Santa Casa de Misericórdia de Maceió (Filial)"
    }

    df = df.copy()
    df["Executante Intercambio"] = df["Executante Intercambio"].astype(str).str.upper().str.strip()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()

    df["Excluir"] = df.apply(
        lambda row: row["Executante Intercambio"] in prestadores_excecao and 
                    (row["Cd Procedimento"].startswith("7") or 
                     row["Cd Procedimento"].startswith("9") or 
                     len(row["Cd Procedimento"]) == 6), 
        axis=1
    )

    return df[~df["Excluir"]].drop(columns=["Excluir"])


# Regra R01 – Consulta PA < 24h (com lógica ajustada)
def aplicar_regra_r01(df):
    print("Aplicando R01: Consulta PA < 24h...")
    """Regra R01 – Consulta PA < 24h + inclui consulta original como referência"""
    df_pa = df[df["Cd Procedimento"] == 10101039].copy()
    df_pa.sort_values(["Carteirinha", "Executante Intercambio", "Datahora"], inplace=True)
    df_pa["Consulta Anterior"] = df_pa.groupby(["Carteirinha", "Executante Intercambio"])["Datahora"].shift(1)
    df_pa["DifHoras"] = (df_pa["Datahora"] - df_pa["Consulta Anterior"]).dt.total_seconds() / 3600

    registros_glosa = []

    for (carteirinha, prestador), grupo in df_pa.groupby(["Carteirinha", "Executante Intercambio"]):
        grupo = grupo.reset_index(drop=True)
        for i in range(1, len(grupo)):
            consulta_anterior = grupo.loc[i - 1]
            consulta_atual = grupo.loc[i]
            dif_horas = consulta_atual["DifHoras"]

            if pd.notna(dif_horas) and dif_horas < 24:
                linha_anterior = consulta_anterior.copy()
                linha_anterior["Nº da Regra"] = "R01"
                linha_anterior["Nome da Regra"] = "Consulta PA < 24h"
                linha_anterior["Motivo da Glosa"] = f"CONSULTA ORIGINAL (serviu de base para glosa) | Data: {linha_anterior['Dt Procedimento'].date()}"
                registros_glosa.append(linha_anterior)

                linha_atual = consulta_atual.copy()
                linha_atual["Nº da Regra"] = "R01"
                linha_atual["Nome da Regra"] = "Consulta PA < 24h"
                linha_atual["Motivo da Glosa"] = f"CONSULTA REALIZADA COM INTERVALO DE {round(dif_horas, 2)}h — mínimo 24h"
                registros_glosa.append(linha_atual)

    if registros_glosa:
        df_r01 = pd.DataFrame(registros_glosa)
        return len(df_r01), df_r01
    else:
        return 0, pd.DataFrame()

# R02 – Consulta de Retorno < 20 dias
def aplicar_regra_r02(df):
    print("Aplicando R02: Consulta de Retorno...")
    codigos_consulta = ["10101012", "50000560", "50000586", "50001221", "50000055", "10101063", "50000144"]
    intervalo_dias = 20
    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Executante Intercambio"] = df["Executante Intercambio"].astype(str).str.strip().str.upper()
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], errors='coerce', dayfirst=True)
    consultas = df[df["Cd Procedimento"].isin(codigos_consulta)].copy()
    consultas.sort_values(by=["Carteirinha", "Executante Intercambio", "Cd Procedimento", "Dt Procedimento"], inplace=True)
    consultas["Dias desde ultima consulta"] = consultas.groupby(
        ["Carteirinha", "Executante Intercambio", "Cd Procedimento"]
    )["Dt Procedimento"].diff().dt.days
    glosas_final = pd.DataFrame()
    for (carteirinha, prestador, codigo), grupo in consultas.groupby(
        ["Carteirinha", "Executante Intercambio", "Cd Procedimento"]
    ):
        grupo = grupo.sort_values("Dt Procedimento")
        for i in range(1, len(grupo)):
            dias = grupo.iloc[i]["Dias desde ultima consulta"]
            if pd.notna(dias) and dias < intervalo_dias:
                consulta_original = grupo.iloc[i-1].copy()
                motivo_original = (
                    f"CONSULTA ORIGINAL (serviu de base para retorno em {grupo.iloc[i]['Dias desde ultima consulta']} dias) | "
                    f"Data: {consulta_original['Dt Procedimento'].date()}"
                )
                consulta_retorno = grupo.iloc[i].copy()
                motivo_retorno = (
                    f"CONSULTA RETORNO (realizada em {consulta_retorno['Dias desde ultima consulta']} dias, mínimo 20) | "
                    f"Data original: {consulta_original['Dt Procedimento'].date()}"
                )
                _, df_original = registrar_glosa(pd.DataFrame([consulta_original]), "R02", "Consulta Retorno < 20 dias", motivo_original)
                _, df_retorno = registrar_glosa(pd.DataFrame([consulta_retorno]), "R02", "Consulta Retorno < 20 dias", motivo_retorno)
                glosas_final = pd.concat([glosas_final, df_original, df_retorno])
    return len(glosas_final), glosas_final

# R03 – US + Doppler
def aplicar_regra_r03(df):
    print("Aplicando R03: US + Doppler...")
    """
    R03 – US + Doppler
    Glosa quando são cobrados simultaneamente no mesmo prestador:
    - Ultrassom (códigos 40901203, 40901220, 40901211) 
    - Doppler (código 40901386)
    """
    
    # Verificação de colunas obrigatórias
    required_cols = {'Carteirinha', 'Cd Procedimento', 'Executante Intercambio'}
    missing_cols = required_cols - set(df.columns)
    if missing_cols:
        return 0, pd.DataFrame()  # Retorna vazio se faltar coluna crítica

    # Pré-processamento
    df = df.copy()
    df['Cd Procedimento'] = df['Cd Procedimento'].astype(str).str.strip()
    
    # Códigos-alvo
    us_codigos = {"40901203", "40901220", "40901211"}
    doppler_codigo = {"40901386"}
    
    # Filtra os procedimentos relevantes
    df_us = df[df['Cd Procedimento'].isin(us_codigos)].copy()
    df_doppler = df[df['Cd Procedimento'].isin(doppler_codigo)].copy()

    # Identifica pacientes+prestadores com ambos
    pacientes_us = set(zip(df_us['Carteirinha'], df_us['Executante Intercambio']))
    pacientes_doppler = set(zip(df_doppler['Carteirinha'], df_doppler['Executante Intercambio']))
    conflitos = pacientes_us & pacientes_doppler

    # Coleta os registros conflitantes
    registros_glosa = []
    for carteirinha, executante in conflitos:
        # Pega todos os US do paciente+prestador
        registros_glosa.extend(
            df_us[(df_us['Carteirinha'] == carteirinha) & 
                 (df_us['Executante Intercambio'] == executante)]
                 .to_dict('records')
        )
        # Pega todos os Doppler do paciente+prestador
        registros_glosa.extend(
            df_doppler[(df_doppler['Carteirinha'] == carteirinha) & 
                      (df_doppler['Executante Intercambio'] == executante)]
                      .to_dict('records')
        )

    # Gera o DataFrame de glosas
    if registros_glosa:
        df_glosa = pd.DataFrame(registros_glosa)
        motivo = "Cobrança simultânea de US + Doppler no mesmo prestador"
        return registrar_glosa(df_glosa, "R03", "US + Doppler", motivo)
    
    # Caso não encontre glosas
    return 0, pd.DataFrame()  # Retorna DataFrame vazio

# R04 – Procedimentos Duplicados
def aplicar_regra_r04(df):
    print("Aplicando R04: Procedimentos Duplicados...")
    codigos_permitidos = {"30501458": 2, "30501067": 2, "31203060": 2, "30205069": 2, "30602114": 2, "30715016": 3, "40813363": 3, "60033541": 3, "20103301": 6}
    codigos_por_diaria_x2 = ["10104020", "10102019", "60033533", "60033541"]
    codigos_por_diaria_x1 = ["60000368", "60034343"]
    codigos_diarias = ["60000805", "60000589", "60001038"]
    codigos_sem_glosa = ["60000589", "60001038", "60000805", "10101012", "60022965"]
    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df = df[(df["Tipo Receita"].notna()) & (df["Tipo Receita"].str.strip() != "") & (df["Tipo Receita"] != "SADT")].copy()
    df["Diarias"] = df.groupby(["Carteirinha", "Executante Intercambio"])["Cd Procedimento"].transform(
        lambda x: x.str.strip().isin(codigos_diarias).sum()
    ).replace(0, 1).fillna(1)
    cols_chave = ["Carteirinha", "Dt Procedimento", "Cd Procedimento", "Grau Participantes", "Executante Intercambio", "Vl Liberado"]
    df["Quantidade"] = df.groupby(cols_chave)["Cd Procedimento"].transform("count")
    df["Duplicado"] = df["Quantidade"] > 1
    df.loc[(df["Cd Procedimento"].isin(codigos_por_diaria_x2)) & (df["Quantidade"] <= df["Diarias"] * 2), "Duplicado"] = False
    df.loc[(df["Cd Procedimento"].isin(codigos_por_diaria_x1)) & (df["Quantidade"] <= df["Diarias"] * 1), "Duplicado"] = False
    for codigo, limite in codigos_permitidos.items():
        df.loc[(df["Cd Procedimento"] == codigo) & (df["Quantidade"] <= limite), "Duplicado"] = False
    df.loc[df["Cd Procedimento"].isin(codigos_sem_glosa), "Duplicado"] = False
    df_glosa = df[df["Duplicado"]].copy()
    if not df_glosa.empty:
        df_glosa["Motivo_Detalhado"] = (
            "Procedimento registrado " + df_glosa["Quantidade"].astype(str) + "x na mesma data. " +
            "Limite permitido: " +
            df_glosa.apply(lambda x:
                str(codigos_permitidos.get(x["Cd Procedimento"],
                    x["Diarias"]*2 if x["Cd Procedimento"] in codigos_por_diaria_x2 else
                    x["Diarias"]*1 if x["Cd Procedimento"] in codigos_por_diaria_x1 else 1
                )), axis=1)
        )
        return registrar_glosa(df_glosa, "R04", "Procedimentos Duplicados", df_glosa["Motivo_Detalhado"])
    return 0, pd.DataFrame()

# ✅ R05 – SADT acima do permitido
def aplicar_regra_r05(df):
    print("Aplicando R05: SADT acima do permitido...")

    excecoes = {
        '20103069': 3, '20103093': 3, '20103484': 10, '20103506': 10, '20103522': 10,
        '20203012': 4, '20203047': 2, '30502314': 2, '30502322': 2, '40103137': 2,
        '40103315': 2, '40103323': 2, '40301397': 2, '40301400': 2, '40301630': 2,
        '40301648': 1, '40301664': 3, '40302016': 1, '40302040': 3, '40302164': 1,
        '40302318': 3, '40302377': 1, '40302385': 1, '40302423': 3, '40302512': 1,
        '40302520': 2, '40302571': 3, '40302580': 3, '40302615': 2, '40302733': 3,
        '40303110': 3, '40303128': 3, '40304086': 9, '40304361': 2, '40304558': 1,
        '40304590': 1, '40305210': 2, '40306259': 3, '40306798': 2, '40307255': 40,
        '40307263': 40, '40307298': 4, '40307760': 2, '40308383': 1, '40308391': 2,
        '40309070': 3, '40310124': 5, '40310205': 6, '40310248': 3, '40310256': 3,
        '40310264': 3, '40310426': 6, '40310620': 2, '40311180': 1, '40311210': 2,
        '40313190': 4, '40313663': 3, '40316190': 3, '40316270': 1, '40316360': 2,
        '40316378': 4, '40316874': 1, '40601013': 4, '40601110': 4, '40601153': 3,
        '40601196': 2, '40601200': 5, '40601226': 3, '40601250': 4, '40601269': 6,
        '40803040': 2, '40803074': 3, '40803082': 6, '40803112': 6, '40803120': 10,
        '40804011': 3, '40804038': 3, '40804046': 6, '40804054': 2, '40804062': 2,
        '40804070': 2, '40804089': 2, '40804097': 6, '40804100': 6, '40805018': 1,
        '40812049': 3, '40812057': 3, '40813363': 3, '40901122': 1, '40901203': 4,
        '40901211': 4, '40901220': 8, '40901270': 3, '40901289': 2, '40901351': 1,
        '40901386': 6, '40901475': 2, '40901483': 2, '40901530': 2, '41001125': 3,
        '41001133': 6, '41001141': 7, '41101227': 3, '41101286': 2, '41101316': 2,
        '41101323': 2, '41301013': 2, '41301072': 2, '41301080': 2, '41301153': 2,
        '41301170': 2, '41301250': 2, '41301269': 2, '41301307': 2, '41301315': 2,
        '41301323': 1, '41303128': 3, '41401271': 2, '41401360': 4, '41401379': 30,
        '41401387': 2, '41401409': 10, '41401530': 2, '41501012': 2, '41501128': 2,
        '41501144': 2, '50000365': 10, '50000632': 1, '50000810': 10, '50000829': 2,
        '50005103': 2, '50005170': 8, '50005278': 8, '50005286': 8, '50005308': 8,
        '50005340': 8, '50005359': 8, '50005367': 8, '50005383': 8
    }

    df = df.copy()
    df['Cd Procedimento'] = df['Cd Procedimento'].astype(str).str.strip()
    df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0)
    df['Tipo Guia'] = df['Tipo Guia'].astype(str).str.lower().str.strip()
    df['Nr Sequencia Conta'] = df['Nr Sequencia Conta'].astype(str).str.strip()

    # 🔎 Foco em SADT
    df_sadt = df[df['Tipo Receita'].str.upper() == 'SADT'].copy()
    if df_sadt.empty:
        return 0, pd.DataFrame()

    cols_group = ['Carteirinha', 'Executante Intercambio', 'Dt Procedimento', 'Cd Procedimento']
    cols_group = [col for col in cols_group if col in df_sadt.columns]

    df_sadt['Soma Quantidade'] = df_sadt.groupby(cols_group)['Quantidade'].transform('sum')
    df_sadt['Limite'] = df_sadt['Cd Procedimento'].map(excecoes).fillna(1)
    df_sadt['Excedeu'] = df_sadt['Soma Quantidade'] > df_sadt['Limite']

    df_glosa = df_sadt[df_sadt['Excedeu']].copy()

    # 🔴 🔍 NOVO FILTRO – Verificar se na conta existe "Guia de resumo de internação"
    contas_com_resumo = df[df["Tipo Guia"] == "guia de resumo de internação"]["Nr Sequencia Conta"].unique()
    df_glosa = df_glosa[~df_glosa["Nr Sequencia Conta"].isin(contas_com_resumo)]

    if df_glosa.empty:
        return 0, pd.DataFrame()

    df_glosa['Motivo da Glosa'] = (
        "SADT " + df_glosa['Cd Procedimento'] + " realizado " +
        df_glosa['Soma Quantidade'].astype(int).astype(str) + "x (limite: " +
        df_glosa['Limite'].astype(int).astype(str) + "x)"
    )

    return registrar_glosa(df_glosa, "R05", "SADT acima do permitido", df_glosa["Motivo da Glosa"])

# ✅ R06 - SADT Seriado (glosa se quantidade > 1 em uma única linha)
def aplicar_regra_r06(df):
    print("Aplicando R06: SADT Seriado ...")
    """
    R06 - SADT Seriado
    Glosa quando Tipo Receita = SADT, código inicia com 5 e quantidade diferente de 1,
    exceto:
      - Tipo Guia = Guia de resumo de internação (ignorar)
      - Códigos permitidos até 8x
    """
    import pandas as pd

    df_r06 = df.copy()

    # 🔍 Verificação das colunas obrigatórias
    col_obrig = ["Quantidade", "Tipo Receita", "Cd Procedimento"]
    if not all(col in df_r06.columns for col in col_obrig):
        print("⚠️ Coluna obrigatória não encontrada. Regra R06 não será aplicada.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    # 🔹 Exceção: remover guias de internação
    if "Tipo Guia" in df_r06.columns:
        df_r06 = df_r06[~(df_r06["Tipo Guia"].str.strip().str.lower() == "guia de resumo de internação")]

    # 🔧 Preparação
    df_r06["Quantidade"] = pd.to_numeric(df_r06["Quantidade"], errors="coerce").fillna(0)
    df_r06["Cd Procedimento"] = df_r06["Cd Procedimento"].astype(str).str.strip()

    # 🔍 Filtra SADT com código iniciando em 5
    df_r06 = df_r06[
        (df_r06["Tipo Receita"].str.upper() == "SADT") &
        (df_r06["Cd Procedimento"].str.startswith("5"))
    ].copy()

    # 🎯 Lista de códigos permitidos até 8x
    codigos_excecao = ["50005286", "50005170", "50005278", "50005340", "50005308"]
    df_r06["Excecao"] = (
        df_r06["Cd Procedimento"].isin(codigos_excecao) &
        (df_r06["Quantidade"] <= 8)
    )

    # ⚠️ Glosa se quantidade ≠ 1 e não está na exceção
    df_r06 = df_r06[
        (df_r06["Quantidade"] != 1) &
        (~df_r06["Excecao"])
    ].copy()

    if df_r06.empty:
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    # ✅ Registra glosa com padrão geral
    return registrar_glosa(
        df_r06,
        "R06",
        "SADT Seriado",
        "Procedimento SADT iniciado com 5 com quantidade diferente de 1 (exceções aplicadas)"
    )

# Regra R07 – Fracionamento TC/RM
def aplicar_regra_r07(df, df_tc_base, df_rm_base):
    print("Aplicando R07: Fracionamento de Pacotes...")
    """
    R07 - Fracionamento TC/RM
    Glosa exames fracionados (códigos 4100 e 4110) com variação > 5% do valor base esperado.
    """

    def processar_fracionamento(df_relatorio, df_base, prefixo_proc):
        df_relatorio = df_relatorio.copy()
        df_base = df_base.copy()

        df_relatorio.columns = [col.strip().title() for col in df_relatorio.columns]
        df_base.columns = [col.strip().title() for col in df_base.columns]

        df_proc = df_relatorio[df_relatorio["Cd Procedimento"].astype(str).str.startswith(prefixo_proc)].copy()
        df_merge = pd.merge(
            df_proc,
            df_base[["Cd Procedimento", "Vl 1", "Vl 2", "Vl 3", "Vl 4"]],
            on="Cd Procedimento",
            how="left"
        )

        registros_glosa = []

        for (carteirinha, data), grupo in df_merge.groupby(["Carteirinha", "Dt Procedimento"]):
            grupo_sorted = grupo.sort_values(by="Vl Liberado", ascending=False).copy()

            # Nome do beneficiário
            if "Nome beneficiario" not in grupo_sorted.columns and "Nome beneficiario" in df.columns:
                nome_dict = df.set_index("Carteirinha")["Nome beneficiario"].to_dict()
                grupo_sorted["Nome beneficiario"] = grupo_sorted["Carteirinha"].map(nome_dict)

            count = len(grupo_sorted)
            discrepancia = False

            for i, (_, row) in enumerate(grupo_sorted.iterrows()):
                pos = i + 1
                valor_base = row.get(f"Vl {pos}", 0)
                if pd.notna(valor_base):
                    valor_esperado = valor_base * 1.05
                    if round(abs(row["Vl Liberado"] - valor_esperado), 2) > 0.01:
                        discrepancia = True
                        break

            if discrepancia:
                msg = f"{count:02d} exames {prefixo_proc} no mesmo dia. "
                for i in range(count):
                    pos = i + 1
                    vl_base = grupo_sorted.iloc[i].get(f"Vl {pos}")
                    vl_corrigido = vl_base * 1.05 if pd.notna(vl_base) else 0
                    vl_lib = grupo_sorted.iloc[i]["Vl Liberado"]
                    msg += f"{pos}º: esperado R$ {vl_corrigido:.2f}, liberado R$ {vl_lib:.2f}. "

                grupo_sorted["Motivo"] = msg.strip()
                grupo_sorted["Cd Procedimento"] = grupo_sorted["Cd Procedimento"].astype(str)
                grupo_sorted = grupo_sorted[grupo_sorted["Cd Procedimento"] != "41001133"]

                if not grupo_sorted.empty:
                    registros_glosa.append(grupo_sorted)

        if registros_glosa:
            df_resultado = pd.concat(registros_glosa, ignore_index=True)
            return registrar_glosa(
                df_resultado,
                "R07",
                "Fracionamento TC/RM",
                df_resultado["Motivo"]
            )
        else:
            return 0, pd.DataFrame()

    qtd_tc, df_tc = processar_fracionamento(df, df_tc_base, "4100")
    qtd_rm, df_rm = processar_fracionamento(df, df_rm_base, "4110")

    df_r07 = pd.concat([df_tc, df_rm], ignore_index=True) if not df_tc.empty or not df_rm.empty else pd.DataFrame()
    return (len(df_r07), df_r07) if not df_r07.empty else (0, pd.DataFrame())

# Regra R08 – Excludentes
def aplicar_regra_r08(df):
    print("Aplicando R08: Excludentes...")
    """
    R08 – Procedimentos Excludentes (turbo final melhorada)
    Detecta procedimentos excludentes no mesmo atendimento e gera duas linhas de glosa,
    preservando todas as colunas do DataFrame original.
    """
    import pandas as pd

    colunas_necessarias = ['Carteirinha', 'Dt Procedimento', 'Cd Procedimento', 'Executante Intercambio']
    if not all(col in df.columns for col in colunas_necessarias):
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    try:
        df_excludentes = pd.read_excel(GPT_FILE, sheet_name="Excludentes Matriz")
    except Exception as e:
        print(f"Erro ao ler matriz de excludentes: {e}")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str)
    df_excludentes["Cd procedimento 1"] = df_excludentes["Cd procedimento 1"].astype(str)
    df_excludentes["Cd procedimento 2"] = df_excludentes["Cd procedimento 2"].astype(str)

    mapa_excludentes = {}
    for _, row in df_excludentes.iterrows():
        mapa_excludentes.setdefault(row["Cd procedimento 1"], set()).add(row["Cd procedimento 2"])
        mapa_excludentes.setdefault(row["Cd procedimento 2"], set()).add(row["Cd procedimento 1"])

    registros_glosa = []
    df_grupo = df.groupby(["Carteirinha", "Dt Procedimento"])

    for (carteirinha, data), grupo in df_grupo:
        if "Nome beneficiario" in df.columns and "Nome beneficiario" not in grupo.columns:
            grupo["Nome beneficiario"] = grupo["Carteirinha"].map(
                df.set_index("Carteirinha")["Nome beneficiario"].to_dict()
            )

        procedimentos = grupo["Cd Procedimento"].unique()
        procedimentos_set = set(procedimentos)

        for proc in procedimentos:
            excludentes = mapa_excludentes.get(proc, set())
            conflito = procedimentos_set.intersection(excludentes)

            if conflito:
                for codigo_conflito in conflito:
                    linha_proc = grupo[grupo["Cd Procedimento"] == proc]
                    if not linha_proc.empty:
                        linha_proc = linha_proc.copy()
                        linha_proc["Nº da Regra"] = "R08"
                        linha_proc["Nome da Regra"] = "Procedimentos Excludentes"
                        linha_proc["Motivo da Glosa"] = f"Realizado junto com procedimento excludente: {codigo_conflito}"
                        registros_glosa.append(linha_proc)

                    linha_conflito = grupo[grupo["Cd Procedimento"] == codigo_conflito]
                    if not linha_conflito.empty:
                        linha_conflito = linha_conflito.copy()
                        linha_conflito["Nº da Regra"] = "R08"
                        linha_conflito["Nome da Regra"] = "Procedimentos Excludentes"
                        linha_conflito["Motivo da Glosa"] = f"Realizado junto com procedimento excludente: {proc}"
                        registros_glosa.append(linha_conflito)

    if registros_glosa:
        df_glosa_r08 = pd.concat(registros_glosa, ignore_index=True)
    else:
        df_glosa_r08 = pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    return len(df_glosa_r08), df_glosa_r08

# ✅ R09 – US + Doppler (não devem ocorrer juntos no mesmo dia para o mesmo paciente)
def aplicar_regra_r09(df):
    print("Aplicando R09: US + Doppler...")

    colunas_necessarias = {"Carteirinha", "Dt Procedimento", "Cd Procedimento"}
    if not colunas_necessarias.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R09.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df = df.copy()
    df = aplicar_excecao(df)
    df = aplicar_excecao_rede_master(df)  # Aplicação da exceção da Rede Master

    # Códigos relevantes
    codigos_us = {"40310124", "40310132", "40310140"}
    codigos_doppler = {"40310167", "40310175", "40310183"}

    df_r09 = df[df["Cd Procedimento"].isin(codigos_us.union(codigos_doppler))].copy()
    df_r09["Cd Procedimento"] = df_r09["Cd Procedimento"].astype(str).str.strip()
    df_r09["Dt Procedimento"] = pd.to_datetime(df_r09["Dt Procedimento"], errors="coerce")

    grupos = df_r09.groupby(["Carteirinha", "Dt Procedimento"])
    glosas = []

    for (carteirinha, data), grupo in grupos:
        codigos = set(grupo["Cd Procedimento"])
        if codigos_us & codigos and codigos_doppler & codigos:
            for _, linha in grupo.iterrows():
                linha_glosa = linha.copy()
                linha_glosa["Nº da Regra"] = "R09"
                linha_glosa["Nome da Regra"] = "US + Doppler no mesmo dia"
                linha_glosa["Motivo da Glosa"] = "Foram cobrados US e Doppler no mesmo dia para o mesmo paciente."
                glosas.append(linha_glosa)

    if not glosas:
        return 0, pd.DataFrame()

    df_glosa_r09 = pd.DataFrame(glosas)
    return len(df_glosa_r09), df_glosa_r09

# Regra R10 - HM não pode ser associado (Lista Referencial)
def aplicar_regra_r10(df):
    print("Aplicando R10: HM não pode ser associado (Lista Referencial)...")
    """
    R10 – HM não pode ser associados (Lista Referencial)
    Detecta procedimentos não associáveis realizados no mesmo atendimento e gera glosa para ambos,
    exceto se o valor unitário for zero.
    """
    import pandas as pd

    colunas_necessarias = ['Carteirinha', 'Dt Procedimento', 'Cd Procedimento', 'Vl Unitario']
    if not all(col in df.columns for col in colunas_necessarias):
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    try:
        df_nassociaveis = pd.read_excel(GPT_FILE, sheet_name="N associaveis")
    except Exception as e:
        print(f"Erro ao ler a aba 'N associaveis': {e}")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str)
    df["Vl Unitario"] = pd.to_numeric(df["Vl Unitario"], errors="coerce").fillna(0)
    df = df[df["Vl Unitario"] != 0]  # Excluir registros com valor unitário igual a 0

    df_nassociaveis["Cd procedimento 1"] = df_nassociaveis["Cd procedimento 1"].astype(str)
    df_nassociaveis["Cd procedimento 2"] = df_nassociaveis["Cd procedimento 2"].astype(str)

    mapa_nassociaveis = {}
    for _, row in df_nassociaveis.iterrows():
        mapa_nassociaveis.setdefault(row["Cd procedimento 1"], set()).add(row["Cd procedimento 2"])
        mapa_nassociaveis.setdefault(row["Cd procedimento 2"], set()).add(row["Cd procedimento 1"])

    registros_glosa = []
    df_grupo = df.groupby(["Carteirinha", "Dt Procedimento"])

    for (carteirinha, data), grupo in df_grupo:
        grupo = grupo.copy()

        if "Nome beneficiario" in df.columns and "Nome beneficiario" not in grupo.columns:
            grupo["Nome beneficiario"] = grupo["Carteirinha"].map(
                df.set_index("Carteirinha")["Nome beneficiario"].to_dict()
            )

        procedimentos = grupo["Cd Procedimento"].unique()
        procedimentos_set = set(procedimentos)

        for proc in procedimentos:
            nao_associaveis = mapa_nassociaveis.get(proc, set())
            conflito = procedimentos_set.intersection(nao_associaveis)
            if conflito:
                for codigo_conflito in conflito:
                    linha_proc = grupo[grupo["Cd Procedimento"] == proc]
                    if not linha_proc.empty:
                        linha_proc = linha_proc.copy()
                        linha_proc["Nº da Regra"] = "R10"
                        linha_proc["Nome da Regra"] = "HM não pode ser associados (Lista Referencial)"
                        linha_proc["Motivo da Glosa"] = (
                            f"Conforme Lista Referencial, os procedimentos ({proc} e {codigo_conflito}) "
                            "não devem ser cobrados concomitantemente."
                        )
                        registros_glosa.append(linha_proc)

                    linha_conflito = grupo[grupo["Cd Procedimento"] == codigo_conflito]
                    if not linha_conflito.empty:
                        linha_conflito = linha_conflito.copy()
                        linha_conflito["Nº da Regra"] = "R10"
                        linha_conflito["Nome da Regra"] = "HM não pode ser associados (Lista Referencial)"
                        linha_conflito["Motivo da Glosa"] = (
                            f"Conforme Lista Referencial, os procedimentos ({proc} e {codigo_conflito}) "
                            "não devem ser cobrados concomitantemente."
                        )
                        registros_glosa.append(linha_conflito)

    if registros_glosa:
        df_glosa_r10 = pd.concat(registros_glosa, ignore_index=True)
    else:
        df_glosa_r10 = pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    return len(df_glosa_r10), df_glosa_r10

# Regra R11 - Fracionamento de Pacotes 
def aplicar_regra_r11(df):
    print("Aplicando R11: Fracionamento de Pacotes..")
    """
    R11 — Controle de Pacotes 98989898
    - Valida o fracionamento dos dois primeiros pacotes (100% e 50%)
    - Glosa o terceiro pacote em diante
    - Se houver erro de fracionamento nos dois primeiros, mostra as duas linhas com motivo completo
    - Se apenas um pacote, não glosa e não mostra
    """
    import pandas as pd

    df = df.copy()

    colunas_necessarias = ["Carteirinha", "Dt Procedimento", "Cd Procedimento", "Taxa Item", "Vl Liberado"]
    if not all(coluna in df.columns for coluna in colunas_necessarias):
        return 0, pd.DataFrame()

    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Taxa Item"] = pd.to_numeric(df["Taxa Item"], errors="coerce").fillna(0)
    df["Vl Liberado"] = pd.to_numeric(df["Vl Liberado"], errors="coerce").fillna(0)

    df_pacotes = df[df["Cd Procedimento"] == "98989898"].copy()
    registros_glosa = []

    grupos = df_pacotes.groupby(["Carteirinha", "Dt Procedimento"])

    for (carteirinha, data), grupo in grupos:
        grupo_ordenado = grupo.sort_values(by="Vl Liberado", ascending=False).reset_index(drop=True)

        # Garantir Nome beneficiario
        if "Nome beneficiario" in df.columns and "Nome beneficiario" not in grupo_ordenado.columns:
            grupo_ordenado["Nome beneficiario"] = grupo_ordenado["Carteirinha"].map(
                df.set_index("Carteirinha")["Nome beneficiario"].to_dict()
            )

        if len(grupo_ordenado) == 2:
            taxa1 = grupo_ordenado.iloc[0]["Taxa Item"]
            taxa2 = grupo_ordenado.iloc[1]["Taxa Item"]
            if not (abs(taxa1 - 100) <= 2 and abs(taxa2 - 50) <= 2):
                motivo = f"Foram realizados 2 pacotes para este beneficiário. Esperava-se fracionamento de 100% e 50%, porém foi encontrado {taxa1:.0f}% e {taxa2:.0f}%."
                for idx in range(2):
                    linha_erro = grupo_ordenado.iloc[idx].copy()
                    linha_erro["Nº da Regra"] = "R11"
                    linha_erro["Nome da Regra"] = "Pacote fracionamento incorreto"
                    linha_erro["Motivo da Glosa"] = motivo
                    registros_glosa.append(pd.DataFrame([linha_erro]))

        elif len(grupo_ordenado) > 2:
            taxa1 = grupo_ordenado.iloc[0]["Taxa Item"]
            taxa2 = grupo_ordenado.iloc[1]["Taxa Item"]
            if not (abs(taxa1 - 100) <= 2 and abs(taxa2 - 50) <= 2):
                motivo = f"Foram realizados 2 pacotes para este beneficiário. Esperava-se fracionamento de 100% e 50%, porém foi encontrado {taxa1:.0f}% e {taxa2:.0f}%."
                for idx in range(2):
                    linha_erro = grupo_ordenado.iloc[idx].copy()
                    linha_erro["Nº da Regra"] = "R11"
                    linha_erro["Nome da Regra"] = "Pacote fracionamento incorreto"
                    linha_erro["Motivo da Glosa"] = motivo
                    registros_glosa.append(pd.DataFrame([linha_erro]))
            # Glosar a partir do 3º pacote
            glosa_excesso = grupo_ordenado.iloc[2:].copy()
            glosa_excesso["Nº da Regra"] = "R11"
            glosa_excesso["Nome da Regra"] = "Pacote múltiplo excedente"
            glosa_excesso["Motivo da Glosa"] = "Cobrança de mais de 2 pacotes no mesmo atendimento"
            registros_glosa.append(glosa_excesso)

    if registros_glosa:
        df_glosa_r11 = pd.concat(registros_glosa, ignore_index=True)
    else:
        df_glosa_r11 = pd.DataFrame()

    return len(df_glosa_r11), df_glosa_r11

# Regra R12 - Fracionamento de HM conforme Lista Referencial
def aplicar_regra_r12(df):
    print("Aplicando R12: Fracionamento de HM conforme Lista Referencial...")
    codigos_excluidos = {"10101039", "10101012", "10102019", "10104011", "10104020", "10103015", "20103301", "20104022", "20104138", "40401014", "40809161", "41301137", "41301234"}

    df_r12 = df.copy()
    df_r12['Vl Liberado'] = pd.to_numeric(df_r12['Vl Liberado'], errors='coerce')

    # Aplica apenas para Honorários Médicos
    df_r12 = df_r12[df_r12['Tipo Receita'].fillna('').str.strip() == "Honorários"]

    # Verifica coluna do prestador
    possiveis_nomes = ['Executante Intercambio', 'Executante intercambio', 'Executante', 'Prestador', 'Nome prestador']
    nome_coluna_executante = next((col for col in possiveis_nomes if col in df_r12.columns), None)
    executante = df_r12[nome_coluna_executante].fillna('') if nome_coluna_executante else ''

    # Criar chave para agrupamento
    df_r12['chave'] = (
        df_r12['Nome beneficiario'].fillna('') + '|' +
        df_r12['Dt Procedimento'].astype(str) + '|' +
        executante
    )

    glosas_r12 = []

    for chave, grupo in df_r12.groupby('chave'):
        grupo_filtrado = grupo[~grupo['Cd Procedimento'].astype(str).isin(codigos_excluidos)]

        if len(grupo_filtrado) <= 1:
            continue

        graus = grupo_filtrado['Grau Participantes'].dropna().unique()
        if len(graus) > 1:
            continue

        grupo_ordenado = grupo_filtrado.sort_values(by='Vl Liberado', ascending=False).reset_index(drop=True)
        glosas_temp = []
        erro_detectado = False

        for i, row in grupo_ordenado.iterrows():
            percentual = 0
            if i == 0:
                percentual = 100
            elif i == 1:
                via = str(row['Via Acesso']).strip().lower()
                if 'via de acesso diferente' in via:
                    percentual = 70
                else:
                    percentual = 50
            elif i == 2:
                percentual = 40
            elif i == 3:
                percentual = 30
            else:
                percentual = 10

            taxa_item = int(row.get('Taxa Item', 0))
            if percentual != taxa_item:
                erro_detectado = True

            glosa = row.copy()
            glosa['Nº da Regra'] = 'R12'
            glosa['Nome da Regra'] = 'Fracionamento de Honorário Médico'

            if percentual < 100:
                glosa['Motivo da Glosa'] = f'Procedimento remunerado em {percentual}% conforme regra de fracionamento.'
            else:
                glosa['Motivo da Glosa'] = 'Procedimento principal – 100% (sem glosa, base de fracionamento)'

            glosas_temp.append(glosa)

        if erro_detectado:
            glosas_r12.extend(glosas_temp)

    df_glosa_final = pd.DataFrame(glosas_r12)
    return len(df_glosa_final), df_glosa_final

# ✅ R13 - Colesterol (Intercâmbio)
def aplicar_regra_r13(df):
    print("Aplicando R13: Colesterol...")

    # Verifica colunas com nomes corretos (capitalizados)
    colunas_necessarias = {'Carteirinha', 'Dt Procedimento', 'Cd Procedimento'}
    if not colunas_necessarias.issubset(df.columns):
        print("❌ Colunas necessárias não encontradas para R13")
        # Retorna DataFrame vazio mas com a linha de resumo
        df_vazio = pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])
        linha_resumo = pd.DataFrame([{
            "Nº da Regra": "R13",
            "Nome da Regra": "Colesterol - Painel não permitido",
            "Motivo da Glosa": "Colunas necessárias não encontradas"
        }])
        df_vazio = pd.concat([df_vazio, linha_resumo], ignore_index=True)
        return 0, df_vazio

    df = df.copy()
    df['Cd Procedimento'] = df['Cd Procedimento'].astype(str).str.strip()
    df['Carteirinha'] = df['Carteirinha'].astype(str).str.strip()
    df['Dt Procedimento'] = pd.to_datetime(df['Dt Procedimento'], errors='coerce')

    # Códigos relevantes
    codigos_base = {"40301583", "40301605", "40302547"}  # Códigos principais
    codigos_glosaveis = {"40301591", "40302695"}        # Códigos que devem ser glosados
    todos_codigos = codigos_base.union(codigos_glosaveis)

    # Filtra apenas os procedimentos relevantes
    df_analise = df[df['Cd Procedimento'].isin(todos_codigos)].copy()
    
    if df_analise.empty:
        print("Nenhum procedimento relevante encontrado para R13")
        # Retorna DataFrame vazio com linha de resumo
        df_vazio = pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])
        linha_resumo = pd.DataFrame([{
            "Nº da Regra": "R13",
            "Nome da Regra": "Colesterol - Painel não permitido",
            "Motivo da Glosa": "Nenhum procedimento relevante encontrado"
        }])
        df_vazio = pd.concat([df_vazio, linha_resumo], ignore_index=True)
        return 0, df_vazio

    # Agrupa por paciente e data
    grupos = df_analise.groupby(['Carteirinha', 'Dt Procedimento'])
    glosas = []

    for (carteirinha, data), grupo in grupos:
        codigos_presentes = set(grupo['Cd Procedimento'].unique())
        
        # Verifica se tem pelo menos um código base e um glosável
        if codigos_base.intersection(codigos_presentes) and codigos_glosaveis.intersection(codigos_presentes):
            # Adiciona apenas os códigos glosáveis para glosa
            for _, row in grupo[grupo['Cd Procedimento'].isin(codigos_glosaveis)].iterrows():
                motivo = f"Painel de colesterol realizado junto com: {', '.join(codigos_presentes)}"
                qtd, df_glosa = registrar_glosa(
                    pd.DataFrame([row]), 
                    'R13', 
                    "Colesterol - Painel não permitido", 
                    motivo
                )
                glosas.append(df_glosa)

    if glosas:
        df_resultado = pd.concat(glosas, ignore_index=True)
        return len(df_resultado), df_resultado
    
    # Se não encontrou glosas, retorna DataFrame vazio com linha de resumo
    df_vazio = pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])
    linha_resumo = pd.DataFrame([{
        "Nº da Regra": "R13",
        "Nome da Regra": "Colesterol - Painel não permitido",
        "Motivo da Glosa": "Nenhuma ocorrência identificada"
    }])
    df_vazio = pd.concat([df_vazio, linha_resumo], ignore_index=True)
    return 0, df_vazio

# ✅ R14 Consulta de psicoterapia cobrada mais de 1x no dia 
def aplicar_regra_r14(df):
    print("Aplicando R14: Consulta de psicoterapia cobrada mais de 1x no dia...")
    """
    R14 - Consulta de psicoterapia cobrada mais de 1x no dia
    Glosa se:
    1. Um registro tiver Quantidade > 1 (ex: 2, 3...)
    2. Múltiplos registros com Quantidade = 1 na mesma data
    """
    required_cols = {'Carteirinha', 'Dt Procedimento', 'Cd Procedimento', 'Quantidade'}
    if not required_cols.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R14.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df = df.copy()
    codigo = '50000470'
    df['Cd Procedimento'] = df['Cd Procedimento'].astype(str).str.strip()
    df['Dt Procedimento'] = pd.to_datetime(df['Dt Procedimento'], errors='coerce').dt.date
    df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0)

    df_psico = df[df['Cd Procedimento'] == codigo].copy()

    # --- Caso 1: Quantidade > 1 ---
    df_qtd_invalida = df_psico[df_psico['Quantidade'] > 1].copy()
    df_qtd_invalida['Motivo'] = "Quantidade > 1 no mesmo registro (máximo permitido: 1)"

    # --- Caso 2: Múltiplos registros com Quantidade = 1 no mesmo dia ---
    registros_duplicados = []
    for (carteirinha, data), grupo in df_psico.groupby(['Carteirinha', 'Dt Procedimento']):
        if len(grupo) > 1:
            grupo = grupo.copy()
            grupo['Motivo'] = f"Múltiplas sessões na mesma data ({len(grupo)} registros)"
            registros_duplicados.append(grupo)

    df_duplicados = pd.concat(registros_duplicados) if registros_duplicados else pd.DataFrame()

    # --- Consolidar glosas ---
    df_glosa = pd.concat([df_qtd_invalida, df_duplicados]).drop_duplicates()

    if df_glosa.empty:
        return 0, pd.DataFrame()  # ⚠️ Não retorna nada para a aba Glosas se não há glosa

    return registrar_glosa(
        df_glosa,
        "R14",
        "Psicoterapia - Quantidade inválida",
        df_glosa['Motivo']
    )
# ✅ R15 Atendimento fisiátrico (20103093) max.1x exceto UTI"
def aplicar_regra_r15(df):
    print("Aplicando R15: Atendimento fisiátrico (20103093) max.1x exceto UTI...")

    colunas_obrigatorias = {'Cd Procedimento', 'Nome beneficiario', 'Nr Sequencia Conta'}
    if not colunas_obrigatorias.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R15.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df_r15 = df.copy()
    df_r15["Cd Procedimento"] = df_r15["Cd Procedimento"].astype(str).str.strip()

    cod_principal = "20103093"
    cod_excecao = {"60001062", "60001054", "60001038", "60036710"}

    contas_com_20103093 = df_r15[df_r15["Cd Procedimento"] == cod_principal]["Nr Sequencia Conta"].unique()
    df_r15 = df_r15[df_r15["Nr Sequencia Conta"].isin(contas_com_20103093)]

    grupo_conta = df_r15.groupby("Nr Sequencia Conta")
    glosas = []

    for conta, grupo in grupo_conta:
        qtd_proced = (grupo["Cd Procedimento"] == cod_principal).sum()
        tem_excecao = any(grupo["Cd Procedimento"].isin(cod_excecao))

        limite = 2 if tem_excecao else 1

        if qtd_proced > limite:
            qtd_glosar = qtd_proced - limite
            linhas_glosar = grupo[grupo["Cd Procedimento"] == cod_principal].iloc[-qtd_glosar:]

            for _, linha in linhas_glosar.iterrows():
                glosa = linha.copy()
                glosa["Nº da Regra"] = "R15"
                glosa["Nome da Regra"] = "Atendimento fisiátrico (20103093) max.1x exceto UTI"
                if tem_excecao:
                    glosa["Motivo da Glosa"] = (
                        f"Mais de 2 lançamentos do procedimento {cod_principal} com diária de UTI na conta."
                    )
                else:
                    glosa["Motivo da Glosa"] = (
                        f"Atendimento fisiátrico cobrado com quantidade maior que 1 e sem diária de UTI."
                    )
                glosas.append(glosa)

    if not glosas:
        return 0, pd.DataFrame()

    df_glosa_r15 = pd.DataFrame(glosas)
    return len(df_glosa_r15), df_glosa_r15


# ✅ R16 - Consulta e sessão de psicoterapia cobradas na mesma data
def aplicar_regra_r16(df):
    print("Aplicando R16: Consulta e sessão de psicoterapia na mesma data...")

    colunas_necessarias = {"Cd Procedimento", "Dt Procedimento", "Nome beneficiario"}
    if not colunas_necessarias.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R16.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df_r16 = df.copy()
    df_r16["Cd Procedimento"] = df_r16["Cd Procedimento"].astype(str).str.strip()
    df_r16["Dt Procedimento"] = pd.to_datetime(df_r16["Dt Procedimento"], errors="coerce")

    # Procedimentos envolvidos
    cod_psi = "50000470"
    cod_consulta = "50000462"

    # Foco nos atendimentos que têm pelo menos um dos dois códigos
    df_filtrado = df_r16[df_r16["Cd Procedimento"].isin([cod_psi, cod_consulta])].copy()

    # Agrupa por beneficiário e data
    glosas = []
    for (nome, data), grupo in df_filtrado.groupby(["Nome beneficiario", "Dt Procedimento"]):
        codigos = grupo["Cd Procedimento"].unique()
        if cod_psi in codigos and cod_consulta in codigos:
            for _, linha in grupo[grupo["Cd Procedimento"].isin([cod_psi, cod_consulta])].iterrows():
                linha_glosa = linha.copy()
                linha_glosa["Nº da Regra"] = "R16"
                linha_glosa["Nome da Regra"] = "Consulta e sessão de psicoterapia cobradas na mesma data"
                linha_glosa["Motivo da Glosa"] = "Códigos 50000470 e 50000462 cobrados no mesmo dia para o mesmo beneficiário."
                glosas.append(linha_glosa)

    if not glosas:
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df_glosa_r16 = pd.DataFrame(glosas)
    return len(df_glosa_r16), df_glosa_r16

# ✅ R17 - UTI 10104020: valor e quantidade excessiva

def aplicar_regra_r17(df):
    print("Aplicando R17: UTI 10104020 - Valor e quantidade excessiva...")

    colunas_necessarias = {"Cd Procedimento", "Dt Procedimento", "Nome beneficiario", "Vl Liberado", "Quantidade"}
    if not colunas_necessarias.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R17.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df_r17 = df.copy()
    df_r17["Cd Procedimento"] = df_r17["Cd Procedimento"].astype(str).str.strip()
    df_r17["Dt Procedimento"] = pd.to_datetime(df_r17["Dt Procedimento"], errors="coerce")
    df_r17["Vl Liberado"] = pd.to_numeric(df_r17["Vl Liberado"], errors="coerce").fillna(0)
    df_r17["Quantidade"] = pd.to_numeric(df_r17["Quantidade"], errors="coerce").fillna(0)

    cod_proced = "10104020"
    limite_valor = 244.47
    limite_quantidade = 2

    df_foco = df_r17[df_r17["Cd Procedimento"] == cod_proced].copy()

    # 1. Glosa por valor acima do permitido
    df_valor_excedido = df_foco[df_foco["Vl Liberado"] > limite_valor].copy()
    df_valor_excedido["Motivo"] = f"Valor liberado superior a R$ {limite_valor:.2f} - Código 10104020 não inclui acréscimos."

    # 2. Glosa por quantidade > 2 no dia (mesmo que em multiplas linhas)
    glosas_qtd = []
    grupos = df_foco.groupby(["Nome beneficiario", "Dt Procedimento"])
    for (nome, data), grupo in grupos:
        soma_qtd = grupo["Quantidade"].sum()
        if soma_qtd > limite_quantidade:
            grupo_copia = grupo.copy()
            grupo_copia["Motivo"] = f"Soma de quantidades no dia excede o limite de {limite_quantidade}."
            glosas_qtd.append(grupo_copia)

    df_qtd_excedida = pd.concat(glosas_qtd) if glosas_qtd else pd.DataFrame()

    # 3. Consolidar glosas
    df_glosa = pd.concat([df_valor_excedido, df_qtd_excedida]).drop_duplicates()

    if df_glosa.empty:
        return 0, pd.DataFrame([{
            "Nº da Regra": "R17",
            "Nome da Regra": "UTI 10104020 - Valor ou quantidade excessiva",
            "Motivo da Glosa": "Nenhuma irregularidade identificada"
        }])

    return registrar_glosa(
        df_glosa,
        "R17",
        "UTI 10104020 - Valor ou quantidade excessiva",
        df_glosa["Motivo"]
    )

# ✅ R18 - Intensivista Diarista com valor ou quantidade incorreta

def aplicar_regra_r18(df):
    print("Aplicando R18: Intensivista Diarista com valor ou quantidade incorreta...")

    # Verificação de colunas obrigatórias
    obrigatorias = {"Cd Procedimento", "Vl Medico", "Quantidade", "Nome beneficiario", "Dt Procedimento"}
    if not obrigatorias.issubset(df.columns):
        print("⚠️ Colunas obrigatórias ausentes. Pulando R18.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    df_r18 = df.copy()
    df_r18["Cd Procedimento"] = df_r18["Cd Procedimento"].astype(str).str.strip()
    df_r18["Dt Procedimento"] = pd.to_datetime(df_r18["Dt Procedimento"], errors="coerce")
    df_r18["Vl Medico"] = pd.to_numeric(df_r18["Vl Medico"], errors="coerce").fillna(0)
    df_r18["Quantidade"] = pd.to_numeric(df_r18["Quantidade"], errors="coerce").fillna(0)

    cod_alvo = "10104011"
    valor_maximo = 105.00
    glosas = []

    df_cod = df_r18[df_r18["Cd Procedimento"] == cod_alvo].copy()

    # Glosa por valor excedente
    df_valor_excedente = df_cod[df_cod["Vl Medico"] > valor_maximo].copy()
    for _, linha in df_valor_excedente.iterrows():
        linha_glosa = linha.copy()
        linha_glosa["Nº da Regra"] = "R18"
        linha_glosa["Nome da Regra"] = "Intensivista Diarista 10104011: valor e quantidade excessiva"
        linha_glosa["Motivo da Glosa"] = (
            f"Valor superior ao permitido (R$ {linha['Vl Medico']:.2f}). Valor máximo: R$ {valor_maximo:.2f}. "
            "Não se aplica dobra ou acréscimo por acomodação ou urgência/emergência conforme Lista Referencial 2025.02."
        )
        glosas.append(linha_glosa)

    # Glosa por quantidade > 1 no mesmo dia para o mesmo beneficiário
    grupo_qtd = df_cod.groupby(["Nome beneficiario", "Dt Procedimento"])
    for (nome, data), grupo in grupo_qtd:
        if grupo["Quantidade"].sum() > 1:
            for _, linha in grupo.iterrows():
                linha_glosa = linha.copy()
                linha_glosa["Nº da Regra"] = "R18"
                linha_glosa["Nome da Regra"] = "Intensivista Diarista 10104011: valor e quantidade excessiva"
                linha_glosa["Motivo da Glosa"] = (
                    "Cobrança em duplicidade do procedimento 10104011 no mesmo dia. Quantidade permitida: 1 por dia e por paciente. "
                    "Não se aplica dobra ou acréscimo por acomodação ou urgência/emergência conforme Lista Referencial 2025.02."
                )
                glosas.append(linha_glosa)

    if not glosas:
        return 0, pd.DataFrame([{
            "Nº da Regra": "R18",
            "Nome da Regra": "Intensivista Diarista com valor ou quantidade incorreta",
            "Motivo da Glosa": "Nenhuma ocorrência identificada"
        }])

    df_glosa_r18 = pd.DataFrame(glosas)
    return len(df_glosa_r18), df_glosa_r18

# ✅ R19 - Probióticos não devem ser pagos no Intercâmbio

def aplicar_regra_r19(df):
    print("Aplicando R19: Probióticos não devem ser pagos no Intercâmbio...")

    if "Cd Procedimento" not in df.columns:
        print("⚠️ Coluna 'Cd Procedimento' ausente. Pulando R19.")
        return 0, pd.DataFrame(columns=df.columns.tolist() + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"])

    codigos_probioticos = {
        "90490703", "90087771", "90490711", "90528131", "90059182", "90528166", "90528182",
        "90528123", "90528140", "90528158", "90059174", "90059190", "90528115", "90503201",
        "90503228", "90503279", "90503104", "90101391", "90503171", "90503210", "90503236",
        "90503252", "90101413", "90503163", "90503147", "90503090", "90503260", "90503244",
        "90503155", "90503139", "90503198", "90503180", "90101405", "90503120", "90503112",
        "90101383", "90209624", "90209616", "90209446"
    }

    df_r19 = df.copy()
    df_r19["Cd Procedimento"] = df_r19["Cd Procedimento"].astype(str).str.strip()
    df_glosa = df_r19[df_r19["Cd Procedimento"].isin(codigos_probioticos)].copy()

    if df_glosa.empty:
        return 0, pd.DataFrame([{
            "Nº da Regra": "R19",
            "Nome da Regra": "Probióticos não permitidos no Intercâmbio",
            "Motivo da Glosa": "Nenhuma ocorrência identificada"
        }])

    return registrar_glosa(
        df_glosa,
        "R19",
        "Probióticos não permitidos no Intercâmbio",
        "Probióticos não devem ser pagos no Intercâmbio."
    )

# ================================
# R20 – Consulta eletiva x Puericultura
# ================================
def aplicar_regra_r20(df):
    print("Aplicando R20: Consulta eletiva x Puericultura...")

    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Executante Intercambio"] = df["Executante Intercambio"].astype(str).str.strip().str.upper()
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], errors="coerce")

    cod_consulta = "10101012"
    cod_puericultura = "10106146"
    codigos_alvo = {cod_consulta, cod_puericultura}

    df_alvo = df[df["Cd Procedimento"].isin(codigos_alvo)].copy()
    if df_alvo.empty:
        return 0, pd.DataFrame()

    registros_glosa = []

    for (carteirinha, data, prestador), grupo in df_alvo.groupby(["Carteirinha", "Dt Procedimento", "Executante Intercambio"]):
        codigos_presentes = set(grupo["Cd Procedimento"].unique())
        if codigos_alvo.issubset(codigos_presentes):
            for _, linha in grupo.iterrows():
                motivo = "Cobrança de consulta eletiva e puericultura no mesmo dia para o mesmo prestador."
                linha_glosa = linha.copy()
                linha_glosa = pd.DataFrame([linha_glosa])
                _, df_linha_glosa = registrar_glosa(linha_glosa, "R20", "Consulta eletiva x puericultura", motivo)
                registros_glosa.append(df_linha_glosa)

    if registros_glosa:
        df_r20 = pd.concat(registros_glosa, ignore_index=True)
        return len(df_r20), df_r20
    else:
        return 0, pd.DataFrame()

# ================================
# R21 – Fotodermatoscopia 41301234: máximo 1x por dia
# ================================
def aplicar_regra_r21(df):
    print("Aplicando R21: Fotodermatoscopia 41301234 - máximo 1x por dia...")

    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], errors="coerce")

    df_foto = df[df["Cd Procedimento"] == "41301234"].copy()
    if df_foto.empty:
        return 0, pd.DataFrame()

    grupo = df_foto.groupby(["Carteirinha", "Dt Procedimento", "Executante Intercambio"])
    df_foto["Qtd"] = grupo["Cd Procedimento"].transform("count")
    df_glosa = df_foto[df_foto["Qtd"] > 1].copy()

    if not df_glosa.empty:
        motivo = "Fotodermatoscopia 41301234 realizada mais de uma vez no mesmo dia para o mesmo prestador."
        return registrar_glosa(df_glosa, "R21", "Fotodermatoscopia 41301234: máximo 1x por dia", motivo)
    return 0, pd.DataFrame()

# ================================
# R22 – Fotodermatoscopia 41301234 x Dermatoscopia 41301137
# ================================
def aplicar_regra_r22(df):
    print("Aplicando R22: Fotodermatoscopia 41301234 x Dermatoscopia 41301137...")

    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], errors="coerce")
    df["Executante Intercambio"] = df["Executante Intercambio"].astype(str).str.upper().str.strip()

    codigos = {"41301234", "41301137"}
    df_alvo = df[df["Cd Procedimento"].isin(codigos)].copy()
    if df_alvo.empty:
        return 0, pd.DataFrame()

    registros_glosa = []
    for (carteirinha, data, prestador), grupo in df_alvo.groupby(["Carteirinha", "Dt Procedimento", "Executante Intercambio"]):
        codigos_presentes = set(grupo["Cd Procedimento"].unique())
        if codigos == codigos_presentes:
            for _, linha in grupo.iterrows():
                motivo = "Conforme item 1.14.2, não é permitida a cobrança simultânea de Fotodermatoscopia (41301234) e Dermatoscopia (41301137)."
                linha_glosa = linha.copy()
                linha_glosa = pd.DataFrame([linha_glosa])
                _, df_linha_glosa = registrar_glosa(linha_glosa, "R22", "Fotodermatoscopia 41301234 x Dermatoscopia 41301137", motivo)
                registros_glosa.append(df_linha_glosa)

    if registros_glosa:
        df_r22 = pd.concat(registros_glosa, ignore_index=True)
        return len(df_r22), df_r22
    else:
        return 0, pd.DataFrame()

# ================================
# R23 – Dermatoscopia 41301137: máximo 1x por dia
# ================================
def aplicar_regra_r23(df):
    print("Aplicando R23: Dermatoscopia 41301137 - máximo 1x por dia...")

    df = df.copy()
    df["Cd Procedimento"] = df["Cd Procedimento"].astype(str).str.strip()
    df["Dt Procedimento"] = pd.to_datetime(df["Dt Procedimento"], errors="coerce")

    df_dermato = df[df["Cd Procedimento"] == "41301137"].copy()
    if df_dermato.empty:
        return 0, pd.DataFrame()

    grupo = df_dermato.groupby(["Carteirinha", "Dt Procedimento", "Executante Intercambio"])
    df_dermato["Qtd"] = grupo["Cd Procedimento"].transform("count")
    df_glosa = df_dermato[df_dermato["Qtd"] > 1].copy()

    if not df_glosa.empty:
        motivo = "Dermatoscopia 41301137 realizada mais de uma vez no mesmo dia para o mesmo prestador."
        return registrar_glosa(df_glosa, "R23", "Dermatoscopia 41301137: máximo 1x por dia", motivo)
    return 0, pd.DataFrame()

# ================================
# R24 – Acupuntura 31601014 x Estimulação/Infiltração
# ================================
def aplicar_regra_r24(df):
    print("Aplicando R24: Acupuntura x Estimulação/Infiltração...")

    df = df.copy()
    df["Cd procedimento"] = df["Cd procedimento"].astype(str).str.strip()
    df["Dt procedimento"] = pd.to_datetime(df["Dt procedimento"], errors="coerce")
    df["Executante intercambio"] = df["Executante intercambio"].astype(str).str.upper().str.strip()

    cod_acupuntura = "31601014"
    cod_outros = {"31602185", "20103301"}

    df_alvo = df[df["Cd procedimento"].isin({cod_acupuntura}.union(cod_outros))].copy()
    if df_alvo.empty:
        return 0, pd.DataFrame()

    registros_glosa = []
    for (carteirinha, data, prestador), grupo in df_alvo.groupby(["Carteirinha", "Dt procedimento", "Executante intercambio"]):
        codigos_presentes = set(grupo["Cd procedimento"].unique())
        if cod_acupuntura in codigos_presentes and cod_outros & codigos_presentes:
            for _, linha in grupo.iterrows():
                if linha["Cd procedimento"] in cod_outros:
                    motivo = "Procedimentos 31602185 ou 20103301 não devem ser cobrados junto com Acupuntura (31601014) na mesma data."
                    linha_glosa = linha.copy()
                    linha_glosa = pd.DataFrame([linha_glosa])
                    _, df_linha_glosa = registrar_glosa(linha_glosa, "R24", "Acupuntura x Estimulação/Infiltração", motivo)
                    registros_glosa.append(df_linha_glosa)

    if registros_glosa:
        df_r24 = pd.concat(registros_glosa, ignore_index=True)
        return len(df_r24), df_r24
    else:
        return 0, pd.DataFrame()

# ================================
# R25 – Procedimentos que não devem ocorrer com Consulta Eletiva
# ================================
def aplicar_regra_r25(df):
    print("Aplicando R25: Procedimentos não permitidos com Consulta Eletiva...")

    df = df.copy()
    df["Cd procedimento"] = df["Cd procedimento"].astype(str).str.strip()
    df["Dt procedimento"] = pd.to_datetime(df["Dt procedimento"], errors="coerce")
    df["Executante intercambio"] = df["Executante intercambio"].astype(str).str.upper().str.strip()

    cod_consulta = "10101012"
    cod_proibidos = {"20101015", "20101023", "20101074", "20101082", "20101090", "41401514", "40105059"}

    df_alvo = df[df["Cd procedimento"].isin({cod_consulta}.union(cod_proibidos))].copy()
    if df_alvo.empty:
        return 0, pd.DataFrame()

    registros_glosa = []
    for (carteirinha, data, prestador), grupo in df_alvo.groupby(["Carteirinha", "Dt procedimento", "Executante intercambio"]):
        codigos_presentes = set(grupo["Cd procedimento"].unique())
        if cod_consulta in codigos_presentes and cod_proibidos & codigos_presentes:
            for _, linha in grupo.iterrows():
                if linha["Cd procedimento"] in cod_proibidos:
                    motivo = "Procedimento não deve ser realizado junto com consulta eletiva (10101012) na mesma data e prestador."
                    linha_glosa = linha.copy()
                    linha_glosa = pd.DataFrame([linha_glosa])
                    _, df_linha_glosa = registrar_glosa(linha_glosa, "R25", "Consulta eletiva x procedimentos proibidos", motivo)
                    registros_glosa.append(df_linha_glosa)

    if registros_glosa:
        df_r25 = pd.concat(registros_glosa, ignore_index=True)
        return len(df_r25), df_r25
    else:
        return 0, pd.DataFrame()

# ================================
# R26 – Infiltração 20103301: máximo 2x por dia
# ================================
def aplicar_regra_r26(df):
    print("Aplicando R26: Infiltração 20103301 - máximo 2x por dia...")

    df = df.copy()
    df["Cd procedimento"] = df["Cd procedimento"].astype(str).str.strip()
    df["Dt procedimento"] = pd.to_datetime(df["Dt procedimento"], errors="coerce")

    df_inf = df[df["Cd procedimento"] == "20103301"].copy()
    if df_inf.empty:
        return 0, pd.DataFrame()

    grupo = df_inf.groupby(["Carteirinha", "Dt procedimento", "Executante intercambio"])
    df_inf["Qtd"] = grupo["Cd procedimento"].transform("count")
    df_glosa = df_inf[df_inf["Qtd"] > 2].copy()

    if not df_glosa.empty:
        motivo = "Infiltração (20103301) deve ser limitada a no máximo 2 por dia para o mesmo paciente e prestador."
        return registrar_glosa(df_glosa, "R26", "Infiltração 20103301: máximo 2x por dia", motivo)
    return 0, pd.DataFrame()

# ================================
# R27 – Facectomia com dobra de apartamento não pertinente
# ================================
def aplicar_regra_r27(df):
    print("Aplicando R27: Facectomia - possível dobra indevida de apartamento...")

    df = df.copy()
    df["Cd procedimento"] = df["Cd procedimento"].astype(str).str.strip()
    df["Vl liberado"] = pd.to_numeric(df["Vl liberado"], errors="coerce")

    codigos = {"30306027", "30306034"}
    df_alvo = df[df["Cd procedimento"].isin(codigos)].copy()
    df_glosa = df_alvo[df_alvo["Vl liberado"] > 1404].copy()

    if not df_glosa.empty:
        motivo = "Valor do item acima de R$1.404,00 para código de apartamento. Verificar possível cobrança em duplicidade."
        return registrar_glosa(df_glosa, "R27", "Valor elevado - códigos de apartamento", motivo)
    return 0, pd.DataFrame()


# ===================================================
# INICIAR EXECUÇÃO
# ===================================================
if __name__ == "__main__":
    df, df_gpt_tc, df_gpt_rm = carregar_dados()

    # Aplicar todas as regras R01 a R12
    # Exemplo: qtd_glosas_r01, df_r01 = aplicar_regra_r01(df)

    qtd_glosas_r01, df_r01 = aplicar_regra_r01(df)
    qtd_glosas_r02, df_r02 = aplicar_regra_r02(df)
    qtd_glosas_r03, df_r03 = aplicar_regra_r03(df)
    qtd_glosas_r04, df_r04 = aplicar_regra_r04(df)
    qtd_glosas_r05, df_r05 = aplicar_regra_r05(df)
    qtd_glosas_r06, df_r06 = aplicar_regra_r06(df)
    qtd_glosas_r07, df_r07 = aplicar_regra_r07(df, df_gpt_tc, df_gpt_rm)
    qtd_glosas_r08, df_r08 = aplicar_regra_r08(df)
    qtd_glosas_r09, df_r09 = aplicar_regra_r09(df)
    qtd_glosas_r10, df_r10 = aplicar_regra_r10(df)
    qtd_glosas_r11, df_r11 = aplicar_regra_r11(df)
    qtd_glosas_r12, df_r12 = aplicar_regra_r12(df)
    qtd_glosas_r13, df_r13 = aplicar_regra_r13(df)
    qtd_glosas_r14, df_r14 = aplicar_regra_r14(df)
    qtd_glosas_r15, df_r15 = aplicar_regra_r15(df)
    qtd_glosas_r16, df_r16 = aplicar_regra_r16(df)
    qtd_glosas_r17, df_r17 = aplicar_regra_r17(df)
    qtd_glosas_r18, df_r18 = aplicar_regra_r18(df)
    qtd_glosas_r19, df_r19 = aplicar_regra_r19(df)
    qtd_glosas_r20, df_r20 = aplicar_regra_r20(df)
    qtd_glosas_r21, df_r21 = aplicar_regra_r21(df)
    qtd_glosas_r22, df_r22 = aplicar_regra_r22(df)
    qtd_glosas_r23, df_r23 = aplicar_regra_r23(df)
    qtd_glosas_r24, df_r24 = aplicar_regra_r24(df)
    qtd_glosas_r25, df_r25 = aplicar_regra_r25(df)
    qtd_glosas_r26, df_r26 = aplicar_regra_r26(df)
    qtd_glosas_r27, df_r27 = aplicar_regra_r27(df)

# ===================================================
# COMBINAR GLOSAS
# ===================================================
# Agora que as glosas de todas as regras foram registradas, podemos combiná-las
df_glosas_final = pd.concat([df_r01, df_r02, df_r03, df_r04, df_r05, df_r06, df_r07, df_r08, df_r09, df_r10, df_r11,
                             df_r12, df_r13, df_r14, df_r15, df_r16, df_r17, df_r18, df_r19, df_r20, df_r21, df_r22,
                             df_r23, df_r24, df_r25, df_r26, df_r27], ignore_index=True)
# Padronização robusta da Competência para MM/AAAA, extraindo mês e ano da data
if "Competência" in df_glosas_final.columns:
    try:
        df_glosas_final["Competência"] = pd.to_datetime(df_glosas_final["Competência"], errors="coerce")
        df_glosas_final["Competência"] = df_glosas_final["Competência"].apply(lambda x: f"{x.month:02}/{x.year}" if pd.notnull(x) else "")
    except Exception as e:
        print(f"Erro ao formatar a coluna Competência: {e}")
#====================================================
# LIMPEZA FINAL DE COLUNAS TÉCNICAS
# ===================================================
colunas_indesejadas = [
    "Datahora", "DifHoras", "Dias desde ultima consulta",
    "Diarias", "Duplicado", "Motivo_Detalhado", "Soma Quantidade",
    "Limite", "Excedeu", "Vl 1", "Vl 2", "Vl 3", "Vl 4", "Motivo", "Excecao"
]
df_glosas_final.drop(columns=[col for col in colunas_indesejadas if col in df_glosas_final.columns], inplace=True, errors="ignore")

# ===================================================
# GERAR RESUMO (CÓDIGO CORRIGIDO)
# ===================================================
resumo = []

# Primeiro processa as regras que têm glosas
for regra, grupo in df_glosas_final.groupby("Nº da Regra"):
    qtd_glosas = len(grupo)
    resumo.append([regra, grupo["Nome da Regra"].iloc[0], qtd_glosas, "Com Glosas"])

# Depois adiciona as regras que não têm glosas
regras_com_glosas = set(df_glosas_final["Nº da Regra"].unique())
for regra, nome_regra in TODAS_AS_REGRAS.items():
    if regra not in regras_com_glosas:
        resumo.append([regra, nome_regra, 0, "Sem Glosas"])

# Ordena o resumo pela regra
resumo.sort(key=lambda x: x[0])

df_resumo = pd.DataFrame(resumo, columns=["Nº da Regra", "Nome da Regra", "Qtde Glosas", "Status"])

# ===================================================
# GERAR RELATÓRIO EM EXCEL
# ===================================================

# Padronizar a coluna 'Competência' para MM/AAAA, se existir
if "Competência" in df_glosas_final.columns:
    try:
        df_glosas_final["Competência"] = pd.to_datetime(df_glosas_final["Competência"], errors="coerce")
        df_glosas_final["Competência"] = df_glosas_final["Competência"].dt.strftime("%m/%Y")
        df_glosas_final["Competência"] = df_glosas_final["Competência"].fillna("")
    except Exception as e:
        print(f"Erro ao formatar a coluna Competência: {e}")

# Manter TODAS as colunas originais + as 3 novas colunas
colunas_originais = [
    "Competência", "Nr Sequencia Conta", "Status Conta", "Carteirinha", "Nome beneficiario",
    "Dt Procedimento", "Hora Proc", "Cd Procedimento", "Descricao", "Quantidade",
    "Vl Unitario", "Vl Liberado", "Vl Calculado", "Vl Anestesista", "Vl Medico",
    "Vl Custo Operacional", "Vl Filme", "Tipo Guia", "Via Acesso", "Taxa Item",
    "Grau Participantes", "Tipo Receita", "Executante Intercambio"
]

# Colunas de auditoria a serem adicionadas
colunas_glosa = ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"]

# Junta todas
colunas_finais = colunas_originais + colunas_glosa

# Selecionar apenas essas colunas (ignorar se faltar alguma)
df_glosas_final = df_glosas_final[[col for col in colunas_finais if col in df_glosas_final.columns]]

# Exportar Excel
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_glosas_final.to_excel(writer, sheet_name="Glosas", index=False)
    df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

print(f"✅ O arquivo foi salvo em: {OUTPUT_FILE}")
