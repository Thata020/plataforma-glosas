import pandas as pd
from unidecode import unidecode

# Caminho de entrada e saída
INPUT_FILE = "549_geral.xlsx"
OUTPUT_FILE = "Atendimentos_Intercambio.xlsx"

# Lista de colunas desejadas na 549
colunas_549 = [
    "Competencia apresentacao", "Nr sequencia conta", "Status conta", "Carteirinha", "Nome beneficiario", 
    "Dt procedimento", "Hora proc", "Cd procedimento", "Descricao", 
    "Quantidade", "Vl unitario", "Vl liberado", "Vl calculado", "Vl anestesista", "Vl medico", 
    "Vl custo operacional", "Vl filme", "Tipo guia", "Via acesso", "Taxa item", 
    "Grau participantes", "Tipo receita", "Executante intercambio"
]

# Função para corrigir caracteres corrompidos
def corrigir_caracteres(texto):
    if isinstance(texto, str):
        try:
            texto = unidecode(texto)
            substituicoes = {
                "Ã‡": "Ç", "Ã£": "ã", "Ã¡": "á", "Ã©": "é", "Ãª": "ê",
                "Ã³": "ó", "Ã´": "ô", "Ãº": "ú", "Ã­": "í", "Ã¤": "ä",
                "Ã¶": "ö", "Ã¼": "ü", "Ã€": "À", "Ã‰": "É"
            }
            for k, v in substituicoes.items():
                texto = texto.replace(k, v)
        except:
            pass
    return texto

# Função principal
def processar_549(INPUT_FILE, OUTPUT_FILE):
    df = pd.read_excel(INPUT_FILE)

    # Corrige nomes de colunas
    df.columns = [col.strip().lower().replace("  ", " ").replace(" ", "_") for col in df.columns]

    # Mapeamento para renomear colunas
    mapeamento = {
        "competencia_apresentacao": "Competencia apresentacao",
        "nº_sequencia_conta": "Nr sequencia conta",
        "nr_sequencia_conta": "Nr sequencia conta",
        "status_conta": "Status conta",
        "carteirinha": "Carteirinha",
        "nome_beneficiario": "Nome beneficiario",
        "dt_procedimento": "Dt procedimento",
        "hora_proc": "Hora proc",
        "cd_procedimento": "Cd procedimento",
        "descricao": "Descricao",
        "quantidade": "Quantidade",
        "vl_unitario": "Vl unitario",
        "vl_liberado": "Vl liberado",
        "vl_calculado": "Vl calculado",
        "vl_anestesista": "Vl anestesista",
        "vl_medico": "Vl medico",
        "vl_custo_operacional": "Vl custo operacional",
        "vl_filme": "Vl filme",
        "tipo_guia": "Tipo guia",
        "via_acesso": "Via acesso",
        "taxa_item": "Taxa item",
        "grau_participantes": "Grau participantes",
        "tipo_receita": "Tipo receita",
        "executante_intercambio": "Executante intercambio"
    }

    df.rename(columns=mapeamento, inplace=True)

    # Padroniza strings e corrige caracteres
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().apply(corrigir_caracteres)

    # Adiciona colunas de auditoria
    df["Nº da Regra"] = ""
    df["Nome da Regra"] = ""
    df["Motivo da Glosa"] = ""

    # Reorganiza colunas finais
    colunas_finais = colunas_549 + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"]
    df = df[[col for col in colunas_finais if col in df.columns]]

# Função para corrigir caracteres corrompidos
def corrigir_caracteres(texto):
    if isinstance(texto, str):
        try:
            texto = unidecode(texto)  # Remove acentos
            substituicoes = {
    "A!": "á",
    "A(c)": "é",
    "A1": "á",
    "A3": "ó",
    "ASSAPSo": "ção",
    "ClAnico": "Clínico",
    "AnA!lise": "Análise",
    "AplicaASSAPSo": "Aplicação",
    "AtA(c)": "Até",
    "Aª": "ª",
    "Aº": "º",
    "CirurgiAPSo": "Cirurgião",
    "DiA!rias": "Diárias",
    "ExA(c)rese": "Exérese",
    "ExecuASSAPSo": "Execução",
    "FisiA!trica": "Fisiátrica",
    "FisiA!trico": "Fisiátrico",
    "HonorA!rios": "Honorários",
    "LesAues": "Lesões",
    "OperatA3rio": "Operatório",
    "PA3s": "Pós",
    "PrA(c)": "Pré",
    "PrevenASSAPSo": "Prevenção",
    "PunASSão": "Punção",
    "SeqA 1/4elas": "Sequelas",
    "Ã": "A",
    "Ã ": "à",
    "Ã¡": "á",
    "Ã¡vel": "ável",
    "Ã¢": "â",
    "Ã£": "ã",
    "Ã§": "ç",
    "Ãª": "ê",
    "Ãªncia": "ência",
    "Ã­": "í",
    "Ã³": "ó",
    "Ã³gico": "ógico",
    "Ã³rio": "ório",
    "Ã´": "ô",
    "Ã¹": "ù",
    "Ãº": "ú",
    "Ãš": "Ú",
    "Ã“": "Ó",
    "Ã‰": "É",
    "1Âº": "1º",
    "ALMOÃ": "ALMOÇO",
    "ANESTÃ": "ANESTESIA",
    "ANTÃ": "ANTENA",
    "APLICAÃ": "APLICAÇÃO",
    "APRESENTAÃ": "APRESENTAÇÃO",
    "AQUISIÃ": "AQUISIÇÃO",
    "ASSISTÃŠNCIA": "ASSISTÊNCIA",
    "ATÃ": "ATÉ",
    "AblaÃ": "Ablar",
    "AbsorÃ": "Absorção",
    "AcÃºsticas": "Acústicas",
    "AdenÃ³ides": "Adenoides",
    "AdrenocorticotrÃ³fico": "Adrenocorticotrófico",
    "AfÃ": "Afã",
    "AlumÃ": "Alumínio",
    "AlÃ": "Além",
    "AmbulatÃ³rio": "Ambulatório",
    "AminoÃ": "Aminoácido",
    "AmpliaÃ": "Ampliação",
    "AmÃ": "Amã",
    "AnaerÃ³bias": "Anaeróbias",
    "AnaerÃ³bicas": "Anaeróbicas",
    "AnalÃ³gico": "Analógico",
    "AnatÃ": "Anatomia",
    "AnestÃ": "Anestesia",
    "AngÃ": "Angústia",
    "AntebraÃ": "Antebraço",
    "AntiangiogÃªnico": "Antiangiogênico",
    "AntibiÃ³ticos": "Antibióticos",
    "AnticentrÃ": "Anticentróide",
    "AntimÃºsculo": "Antimúsculo",
    "AntineutrÃ³filos": "Antineutrófilos",
    "AntinÃºcleo": "Antinúcleo",
    "AntitireÃ³ide": "Antitireóide",
    "AntÃ": "Antã",
    "AnÃ": "Anã",
    "AplicaÃ": "Aplicação",
    "ApolipoproteÃ": "Apolipoproteína",
    "ApÃ³s": "Após",
    "AraÃºjo": "Araújo",
    "ArritmogÃªnico": "Arritmogênico",
    "ArticulaÃ": "Articulação",
    "ArtÃ": "Artã",
    "ArÃ": "Arã",
    "AscÃ³rbico": "Ascórbico",
    "AspiraÃ": "Aspiração",
    "AssistAancia": "Assistência",
    "AssistÃªncia": "Assistência",
    "AssunÃ": "Assunção",
    "Asnica": "Única",
    "AustrÃ": "Austrália",
    "AvaliaÃ": "Avaliação",
    "AvulsÃµes": "Avulsões",
    "BERÃ": "BERÊ",
    "BactÃ": "Bactéria",
    "BerÃ": "Berê",
    "BiolÃ³gicos": "Biológicos",
    "BioquÃ": "Bioquímica",
    "BiÃ³psia": "Biópsia",
    "BiÃ³psias": "Biópsias",
    "BraÃ": "Braço",
    "BÃ": "Bã",
    "CAPTAÃ": "CAPTAÇÃO",
    "CIRÃšRGICA": "CIRÚRGICA",
    "CirAorgica": "Cirúrgica",
    "CIRÃšRGICO": "CIRÚRGICO",
    "CORONAVÃ": "CORONAVÍRUS",
    "CrAC/nio": "Crânio",
    "DEFICIÃ": "DEFICIÊNCIA",
    "DEGENERAÃ": "DEGENERAÇÃO",
    "DIARRÃ‰IA": "DIARREIA",
    "DILATAÃ": "DILATAÇÃO",
    "DIREÃ": "DIREÇÃO",
    "DISFUNÃ": "DISFUNÇÃO",
    "DISPLASIAÃ": "DISPLASIA",
    "DOENÃ": "DOENÇA",
    "DOSAGEMÃ": "DOSAGEM",
    "DURAÃ": "DURAÇÃO",
    "DÃ": "Dã",
    "DÃ³i": "Dói",
    "EducaÃ": "Educação",
    "ElevaÃ": "Elevação",
    "EmergÃªncia": "Emergência",
    "EmÃ": "Emã",
    "EncaminhaÃ": "Encaminhamento",
    "EnfisemaÃ": "Enfisema",
    "EnxaquecaÃ": "Enxaqueca",
    "EscleroseÃ": "Esclerose",
    "EsforÃ": "Esforço",
    "EstenoseÃ": "Estenose",
    "EstenÃ³tico": "Estenótico",
    "EstudosÃ": "Estudos",
    "ExameÃ": "Exame",
    "ExcreÃ": "Excreção",
    "ExposiÃ": "Exposição",
    "ExpressÃ£o": "Expressão",
    "ExtensÃ£o": "Extensão",
    "FÃ": "Fã",
    "FÃ¡rmaco": "Fármaco",
    "FÃ­gado": "Fígado",
    "FÃ­sico": "Físico",
    "FÃ³rmula": "Fórmula",
    "GÃ¡strico": "Gástrico",
    "GÃªnero": "Gênero",
    "HemorrÃ¡gico": "Hemorrágico",
    "HidrataÃ§Ã£o": "Hidratação",
    "HipertensÃ£o": "Hipertensão",
    "HistÃ³rico": "Histórico",
    "IncisAPSo": "Incisão",
    "InfecÃ§Ã£o": "Infecção",
    "InformaÃ§Ã£o": "Informação",
    "InteraÃ§Ã£o": "Interação",
    "IntercAC/mbio": "Intercâmbio",
    "LocalizaÃ§Ã£o": "Localização",
    "LÃ¡bio": "Lábio",
    "LesAPSo": "Lesão",
    "MÃ£e": "Mãe",
    "MÃ©dico": "Médico",
    "MicrocirAorgico": "Microcirúrgico",
    "MÃºsculo": "Músculo",
    "NAPSo": "Não",
    "NecrosÃ£o": "Necrosão",
    "NutriÃ§Ã£o": "Nutrição",
    "ObservaÃ§Ã£o": "Observação",
    "OrientaÃ§Ã£o": "Orientação",
    "PatolÃ³gico": "Patológico",
    "PÃ¡": "Pá",
    "ProteAna": "Proteína",
    "PÃºblico": "Público",
    "RecuperaÃ§Ã£o": "Recuperação",
    "RegiÃ£o": "Região",
    "ResoluÃ§Ã£o": "Resolução",
    "SangraÃ§Ã£o": "Sangração",
    "SessAPSo": "Sessão",
    "SituaÃ§Ã£o": "Situação",
    "TransiÃ§Ã£o": "Transição",
    "TÃ©cnico": "Técnico",
    "TAXA DE SALA PARA APLICAA++AfO DE MEDICAA++AfO": "Taxa de sala para aplicação de medicação",
    "UlcerÃ¡vel": "Ulcerável",
    "VacinaÃ§Ã£o": "Vacinação",
    "ZÃ³ster": "Zóster",
}
            for padrao, substituicao in substituicoes.items():
                texto = texto.replace(padrao, substituicao)
        except Exception:
            return texto  # Retorna o original se houver erro
    return texto

# Função para padronizar os nomes das colunas
def padronizar_nomes_colunas(df):
    df.columns = df.columns.str.strip().str.lower().str.replace("_", " ").str.capitalize()
    df.columns = [unidecode(col) for col in df.columns]  # Remove acentos
    return df

def processar_549(arquivo_entrada, arquivo_saida):
    try:
        df = pd.read_excel(arquivo_entrada)

        # Padroniza os nomes das colunas
        df.columns = [str(c).strip().upper() for c in df.columns]
        df.dropna(how="all", inplace=True)

        # Ajuste: garantir que "Dt procedimento" existe antes de filtrar
        if "DT PROCEDIMENTO" in df.columns:
            df = df[df["DT PROCEDIMENTO"].notna()]

        # Padroniza nomes das colunas
        df.columns = df.columns.str.strip().str.lower().str.replace("_", " ").str.capitalize()
        df.columns = [unidecode(col) for col in df.columns]

        # Garante que todas as colunas necessárias existem
        for col in colunas_549:
            if col not in df.columns:
                df[col] = None
        df = df[colunas_549]

        # Corrige caracteres corrompidos
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].apply(corrigir_caracteres)

        # Formata datas
        if "Dt procedimento" in df.columns:
            df["Dt procedimento"] = pd.to_datetime(df["Dt procedimento"], errors="coerce").dt.strftime("%d/%m/%Y")

        # Remove quantidade inválida
        if "Quantidade" in df.columns:
            df = df[df["Quantidade"] > 0]
            
        # Garante que colunas de auditoria existam
        df["Nº da Regra"] = ""
        df["Nome da Regra"] = ""
        df["Motivo da Glosa"] = ""

        # Ajusta as colunas existentes
        colunas_existentes = [col for col in colunas_549 if col in df.columns]
        df = df[colunas_existentes + ["Nº da Regra", "Nome da Regra", "Motivo da Glosa"]]

        # Salvar resultado
        df.to_excel(arquivo_saida, index=False)
        print(f"✅ Arquivo corrigido salvo como '{arquivo_saida}'")
    except Exception as e:
        print(f"❌ Erro ao processar o arquivo: {e}")

# Executa somente se rodar como script principal
if __name__ == "__main__":
    processar_549(INPUT_FILE, OUTPUT_FILE)
