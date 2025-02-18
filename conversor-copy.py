import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st


def remove_accents(text):
    if pd.isna(text):
        return text
    text = str(text)
    return ''.join(c for c in unicodedata.normalize('NFD', text) 
                  if unicodedata.category(c) != 'Mn')

def find_header_row(df_temp):
    """
    Encontra a linha do cabeçalho no DataFrame.
    Retorna o número da linha onde o cabeçalho está localizado.
    """
    for i, row in df_temp.iterrows():
        lista = list(map(lambda val: str(val).replace(' ', '') if isinstance(val, str) else val, row.values))
        if "Nome" in lista and "Tipo" in lista:
            return i
    return None  # Retorna None se o cabeçalho não for encontrado

def process_sheet(df):
    """
    Processa uma única planilha do Excel.
    Retorna um DataFrame no formato XLSForm.
    """
    # Encontrar a linha do cabeçalho
    header_row = find_header_row(df)
    
    if header_row is None:
        st.warning("Cabeçalho não encontrado nesta planilha. Pulando...")
        return None
    
    # Ler o DataFrame a partir da linha do cabeçalho
    df = df.iloc[header_row:].reset_index(drop=True)
    df.columns = df.iloc[0]  # Define a primeira linha como cabeçalho
    df = df.drop(0).reset_index(drop=True)  # Remove a linha do cabeçalho duplicado
    
    # Verificar se as colunas esperadas estão presentes
    expected_columns = ["Nome", "Tipo", "Rótulo (Label)", "Valores", "Anexo"]
    lista = list(map(lambda val: str(val).strip() if isinstance(val, str) else val, df.columns))
    missing_columns = [col for col in expected_columns if col not in lista]
    
    if missing_columns:
        st.warning(f"As seguintes colunas não foram encontradas nesta planilha: {missing_columns}. Pulando...")
        return None
    
    # Remover linhas e colunas completamente vazias
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    
    # Mapear colunas para o formato XLSForm
    column_mappings = {
        "Nome": "name",
        "Tipo": "type",
        "Rótulo (Label)": "label::Portugues (pt)",
        "Valores": "choices",
        "Domínio": "hint::Portugues (pt)",
        "Anexo": "media"
    }
    
    # Renomear colunas
    df.columns = df.columns.str.strip()  # Remove espaços extras dos nomes das colunas
    df = df.rename(columns=column_mappings)
    
      # Remover acentos da coluna name
    df["name"] = df["name"].apply(remove_accents)  # <--- NOVO TRECHO
    
    # Mapear tipos para XLSForm
    type_mapping = {
        "Numérico": "integer",
        "Númerico": "integer",
        "Texto": "text",
        "data": "date",
        "Sequência de caracteres": "text",
        "Sequência de Caracteres": "text",
        "Seleção": "select_one",
        "Múltipla Escolha": "select_multiple"
    }
    
    
    # Normalizar as chaves do dicionário para minúsculas
    type_mapping_normalized = {k.lower(): v for k, v in  type_mapping.items() }
    # Normalizar os valores da coluna "type" para minúsculas e aplicar o mapeamento
    df["type"] = df["type"].str.lower().str.lstrip().str.rstrip().replace(type_mapping_normalized)
    
    # Criar estrutura padrão do XLSForm
    survey_columns = [
        "type", "name", "label::Portugues (pt)", "hint::Portugues (pt)",
        "appearance", "required", "constraint", "guidance_hint", "relevant"
    ]
    
    # Adicionar colunas ausentes ao DataFrame df
    for col in survey_columns:
        if col not in df.columns:
            df[col] = ""
    
    return df[survey_columns]

def convert_to_xlsform(file):
    # Ler todas as planilhas do arquivo Excel
    xls = pd.ExcelFile(file)
    all_surveys = []
    
    for sheet_name in xls.sheet_names:
        st.write(f"Processando planilha: {sheet_name}")
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        processed_sheet = process_sheet(df)
        
        if processed_sheet is not None:
            all_surveys.append(processed_sheet)
    
    if not all_surveys:
        st.error("Nenhuma planilha válida encontrada no arquivo.")
        return None
    
    # Concatenar todas as planilhas processadas
    survey = pd.concat(all_surveys, ignore_index=True)
    
        # Adicionar linhas padrão ao survey
    standard_rows = [
        ["start", "start", "", "", "", "", "", "", ""],
        ["end", "end", "", "", "", "", "", "", ""],
        ["start-geopoint", "start-geopoint", "", "", "", "", "", "", ""],
        ["today", "today", "", "", "", "", "", "", ""],
        ["username", "username", "", "", "", "", "", "", ""],
        ["deviceid", "deviceid", "", "", "", "", "", "", ""],
        ["phonenumber", "phonenumber", "", "", "", "", "", "", ""],
        ["audit", "audit", "", "", "", "", "", "", ""]
    ]
    
    # Criar DataFrame com as linhas padrão
    standard_df = pd.DataFrame(standard_rows, columns=survey.columns)
    
    # Concatenar as linhas padrão com os dados processados
    survey = pd.concat([standard_df, survey], ignore_index=True)
    
    
    
    # Criar aba 'choices' com lista de opções únicas
    if "choices" in survey.columns:
        choices = survey[["choices"]].dropna().drop_duplicates()
        choices = choices.assign(list_name=choices["choices"], name=choices["choices"], label=choices["choices"])
    else:
        choices = pd.DataFrame(columns=["list_name", "name", "label::Portugues (pt)"])
    
    # Criar aba 'settings'
    settings = pd.DataFrame({
        "form_title": ["Meu XLSForm"],
        "form_id": ["meu_xlsform"]
    })
    
    # Salvar em memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        survey.to_excel(writer, sheet_name="survey", index=False)
        choices.to_excel(writer, sheet_name="choices", index=False)
        settings.to_excel(writer, sheet_name="settings", index=False)
    
    output.seek(0)
    return output

# Interface com Streamlit
st.title("Conversor de Excel para XLSForm (Ochili)")
file = st.file_uploader("Envie seu arquivo Excel", type=["xlsx"])

if file:
    st.success("Arquivo carregado com sucesso!")
    converted_file = convert_to_xlsform(file)
    if converted_file:
        st.download_button(
            label="Baixar XLSForm",
            data=converted_file,
            file_name="xlsform_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )