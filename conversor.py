import os
import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st


def remover_grupos_vazios(df):
    """
    Remove grupos que estão vazios (apenas com begin_group e end_group consecutivos).
    
    Parâmetros:
        df (pd.DataFrame): DataFrame com os dados do formulário
    
    Retorna:
        pd.DataFrame: DataFrame sem os grupos vazios
    """
    # Identificar índices dos grupos vazios
    indices_para_remover = []
    grupo_aberto = None

    for idx, row in df.iterrows():
        if row['type'] == 'begin_group':
            # Registrar início do grupo
            grupo_aberto = {
                'start': idx,
                'name': row['name']
            }
        elif row['type'] == 'end_group' and grupo_aberto:
            # Verificar se é o fechamento do mesmo grupo
            if grupo_aberto['name'] == row.get('name', ''):
                end_idx = idx
                # Verificar se o grupo está vazio (sem outras linhas entre begin e end)
                conteudo_grupo = df.iloc[grupo_aberto['start']+1:end_idx]
                if conteudo_grupo.empty:
                    indices_para_remover.extend([grupo_aberto['start'], end_idx])
                grupo_aberto = None

    # Remover grupos vazios e resetar índice
    df_limpo = df.drop(indices_para_remover).reset_index(drop=True)
    
    return df_limpo
 
 
def remove_line_breaks(df):
    if 'name' in df.columns:
        df['name'] = df['name'].astype(str).str.replace(r'[\n\r]', '', regex=True)
        df = setar_obrigatoriedade(df,'hint::Portugues (pt)')
    return df

def setar_obrigatoriedade(df, hint_col='hint::Portugues (pt)'):
    """
    Define a obrigatoriedade com base na presença de (*) no hint
    Retorna o DataFrame modificado
    """
    required_col = 'required'
    
    # Verificar se a coluna de hint existe
    if hint_col not in df.columns:
        raise ValueError(f"Coluna {hint_col} não encontrada no DataFrame")
    
    # Criar coluna required se não existir
    if required_col not in df.columns:
        df = df.copy()
        df[required_col] = False
    
    # Procurar por asterisco no hint
    mask = df[hint_col].str.contains(r'\*', case=False, na=False)
    
    # Atualizar coluna required
    df.loc[mask, required_col] = "True"
    df.loc[~mask, required_col] = "False"
    
    return df



def remove_accents(text):
    if pd.isna(text):
        return text
    text = str(text)
    return ''.join(c for c in unicodedata.normalize('NFD', text) 
                  if unicodedata.category(c) != 'Mn')

def is_valid_variable_name(name):
    """Verifica se o nome da variável está no padrão aceitável"""
    if pd.isna(name):
        return False
    
    # Remover acentos e caracteres especiais
    normalized = unicodedata.normalize('NFKD', str(name))
    ascii_name = normalized.encode('ASCII', 'ignore').decode('ASCII')
    
    # Verificar caracteres válidos (letras, números e underscores)
    return re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', ascii_name) is not None




def check_variable_names(df):
    """Verifica nomes de variáveis e retorna os inválidos destacando o erro dentro da string"""

    st.write("Verificando nomes das variáveis...")

    if 'name' not in df.columns:
        st.error("Coluna 'name' não encontrada no dataset!")
        return False

    #st.write(df["name"].to_string(index=False))  # Mostra os nomes antes da verificação

    invalid_vars = []

    for idx, row in df.iterrows():
        var_name = str(row['name']).strip()
        
        if pd.isna(row['name']) or var_name == "":
            invalid_vars.append({
                'linha': idx + 2,
                'nome_original': var_name,
                'nome_formatado': "(ERRO->Vazio)",
                'erro': 'Nome da variável está vazio ou nulo'
            })
            continue

        erro_detectado = False
        nome_formatado = var_name

        # Verificar se começa com número
        if re.match(r'^\d', var_name):
            nome_formatado = f"(ERRO->{var_name[0]})" + var_name[1:]
            invalid_vars.append({
                'linha': idx + 2,
                'nome_original': var_name,
                'nome_formatado': nome_formatado,
                'erro': f"Nome não pode começar com número ('{var_name[0]}')"
            })
            erro_detectado = True

        # Verificar caracteres inválidos
        for char in re.finditer(r'[^a-zA-Z0-9_]', var_name):
            pos = char.start()
            nome_formatado = var_name[:pos] + f"(ERRO->{char.group()})" + var_name[pos+1:]
            invalid_vars.append({
                'linha': idx + 2,
                'nome_original': var_name,
                'nome_formatado': nome_formatado,
                'erro': f"Caractere inválido '{char.group()}'"
            })
            erro_detectado = True

        # Verificar espaços
        if ' ' in var_name:
            pos = var_name.find(" ")
            nome_formatado = var_name[:pos] + "(ERRO-> )" + var_name[pos+1:]
            invalid_vars.append({
                'linha': idx + 2,
                'nome_original': var_name,
                'nome_formatado': nome_formatado,
                'erro': "Nome contém espaços"
            })
            erro_detectado = True

        if erro_detectado:
            continue

    # Exibir erros formatados
    if invalid_vars:
        error_msg = "ERRO: Variáveis com nomes inválidos encontradas:\n\n"
        for var in invalid_vars:
            error_msg += f"Linha {var['linha']}: {var['nome_original']}\n"
            error_msg += f"  → Erro: {var['erro']}\n"
            error_msg += f"  → Nome ajustado: {var['nome_formatado']}\n\n"

        error_msg += "\nRegras para nomes válidos:\n"
        error_msg += "- Sem espaços, acentos ou caracteres especiais\n"
        error_msg += "- Somente letras, números e underscores\n"
        error_msg += "- Não pode começar com número\n"

        st.error(error_msg)
        return False

    return True




def find_header_row(df_temp):
    for i, row in df_temp.iterrows():
        lista = list(map(lambda val: str(val).replace(' ', '') if isinstance(val, str) else val, row.values))
        if "Nome" in lista and "Tipo" in lista:
            return i
    return None

def process_sheet(df):
    header_row = find_header_row(df)
    if header_row is None:
        return None
    
    df = df.iloc[header_row:].reset_index(drop=True)
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)
    
    expected_columns = ["Nome", "Tipo", "Rótulo (Label)", "Valores", "Anexo"]
    lista = list(map(lambda val: str(val).strip() if isinstance(val, str) else val, df.columns))
    missing_columns = [col for col in expected_columns if col not in lista]
    
    if missing_columns:
        st.write(f"As seguintes colunas não foram encontradas nesta planilha: {missing_columns}. Pulando...")
        return None
    
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    column_mappings = {
        "Nome": "name",
        "Tipo": "type",
        "Rótulo (Label)": "label::Portugues (pt)",
        "Valores": "choices",
        "Domínio": "hint::Portugues (pt)",
        "Anexo": "media"
    }
    
    df.columns = df.columns.str.strip()
    df = df.rename(columns=column_mappings)
    
    df["name"] = df["name"].apply(remove_accents)
    
    type_mapping = {
        "númerico": "integer",
        "numérico": "integer",
        "texto": "text",
        "data": "date",
        "sequência de caracteres": "text",
        "Sequência de caracteres": "text",
        "seleção": "select_one",
        "múltipla escolha": "select_multiple"
    }
   
   
       # Normalizar as chaves do dicionário para minúsculas
    type_mapping_normalized = {k.lower(): v for k, v in  type_mapping.items() }
    # Normalizar os valores da coluna "type" para minúsculas e aplicar o mapeamento
    df["type"] = df["type"].str.lower().str.lstrip().str.rstrip().replace(type_mapping_normalized)
    
    survey_columns = [
    "type", 
    "name", 
    "label::Portugues (pt)", 
    "hint::Portugues (pt)",
    "required",
    "appearance", 
    "constraint",
    "calculation",
    "constraint_message",
    "relevant",
    "choice_filter"
]
    
    for col in survey_columns:
        if col not in df.columns:
            df[col] = ""
    
    return df[survey_columns]

def add_groups(survey_df, groups_df):
    groups_df = groups_df.dropna(subset=['inicio', 'fim'])
    existing_groups = set()
    st.write("======================= ADICIONANDO GUPOS ==========================")

    for _ in range(2):  # Duas verificações
        for _, group in groups_df.iterrows():
            group_name = group['name']
            #st.write(f"======================================================================================")
            #st.write(f"Processando grupo: {group_name}")
            
            if group_name in existing_groups:
                #st.write(f"Grupo {group_name} já foi adicionado. Pulando...")
                continue
                
            start_field = remove_accents(group['inicio']).strip()
            end_field = remove_accents(group['fim']).strip()
            #st.write(f"Campos de início/fim: '{start_field}' / '{end_field}'")
            #st.write(f"Campo start_field está em metadados?: { start_field in survey_df['name'].str.strip().tolist()}")
            #st.write(f"Campo end_field está em metadados?: { end_field in survey_df['name'].str.strip().tolist()}")
             
               
            start_mask = survey_df['name'].str.strip() == start_field.strip()
            end_mask = survey_df['name'].str.strip() == end_field.strip()
            
            if not start_mask.any() or not end_mask.any():
                #st.write(f"Campos de início/fim não encontrados. Pulando... {end_mask.any()} {start_mask.any()}")
                continue
            
            start_idx = survey_df[start_mask].index[0]
            end_idx = survey_df[end_mask].index[0]
            
            
            # Verificar se o grupo já foi adicionado
            #group_exists = survey_df.iloc[start_idx-1:end_idx+1]['type'].isin(['begin_group', 'end_group']).any()
            
            group_exists = ((survey_df.iloc[start_idx-1:end_idx+1]['type'].isin(['begin_group', 'end_group'])) & 
                (survey_df.iloc[start_idx-1:end_idx+1]['name'] == group_name)).any()


            if group_exists:
                #st.write(f"Grupo {group_name} já foi adicionado. Pulando...")
                existing_groups.add(group_name)
                continue
                
            # Adicionar begin_group
            new_row = {'type': 'begin_group', 'name': group_name, 
                      'label::Portugues (pt)': group['label'].upper(),'appearance': 'field-list'}
            
            survey_df.loc[start_idx - 0.5] = new_row
            
            # Adicionar end_group
            new_row = {'type': 'end_group'}
            survey_df.loc[end_idx + 0.5] = new_row
            
            #st.write(f"Grupo {group_name} adicionado com sucesso.")
            #st.write(f"=============================================================")
             
            existing_groups.add(group_name)
            
            # Reordenar e resetar índices
            survey_df = survey_df.sort_index().reset_index(drop=True)
    
    return survey_df
#=========================================================================

# Função para adicionar cálculos automáticos baseados em padrões de um Excel
def adicionar_calculos_automaticos(df, excel_path):
    """
    Adiciona cálculos automáticos baseados em padrões de um Excel
    
    Parâmetros:
        df (pd.DataFrame): DataFrame principal do formulário
        excel_path (str): Caminho para o Excel com os padrões
    
    Retorna:
        pd.DataFrame: DataFrame com os cálculos adicionados
    """
    # Ler o arquivo de padrões
    try:
        padroes_df = pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"Erro ao ler arquivo de padrões: {str(e)}")
        return df

    # Verificar colunas necessárias
    if not all(col in padroes_df.columns for col in ['name', 'pergunta', 'padrao']):
        st.error("Arquivo de padrões deve conter as colunas: name, pergunta, padrao")
        return df

    for _, row in padroes_df.iterrows():
        target_var = row['name']
        pergunta = str(row['pergunta'])
        padroes = [p.strip().lower() for p in str(row['padrao']).split(',')]

        # Verificar se a variável alvo existe
        if target_var not in df['name'].values:
            st.warning(f"Variável alvo '{target_var}' não encontrada no formulário")
            continue

        # Filtrar variáveis da mesma pergunta
        pergunta_filter = df['name'].str.contains(f'_{pergunta}_', case=False, na=False)
        vars_pergunta = df[pergunta_filter]['name'].tolist()

        # Filtrar variáveis que correspondem aos padrões
        vars_somar = []
        for var in vars_pergunta:
            var_clean = var.lower()
            if any(padrao in var_clean for padrao in padroes):
                vars_somar.append(var)

        if not vars_somar:
            st.warning(f"Nenhuma variável encontrada para {target_var} com padrões: {', '.join(padroes)}")
            continue

        # Montar expressão de cálculo
        calculo = ' + '.join([f'${{{var}}}' for var in vars_somar])

        # Atualizar o cálculo na variável alvo
        df.loc[df['name'] == target_var, 'calculation'] = calculo
        df.loc[df['name'] == target_var, 'type'] = 'calculate'

    return df


# Função para converter os dados do Excel para XLSForm
def convert_to_xlsform(data_file, groups_file, padroes_file):
    # Processar dados principais
    xls = pd.ExcelFile(data_file)
    all_surveys = []
    
    for sheet_name in xls.sheet_names:
        #st.write(f"Processando planilha: {sheet_name}")
        df = pd.read_excel(data_file, sheet_name=sheet_name, header=None)
        processed = process_sheet(df)
        if processed is not None:
            #st.write(f"Planilha {sheet_name} processada com sucesso.")
            all_surveys.append(processed)
    
    if not all_surveys:
        return None
    
    survey = pd.concat(all_surveys, ignore_index=True)
    survey=remove_line_breaks(survey)
    # Validação dos nomes das variáveis
    if not check_variable_names(survey):
        st.stop()  # Interrompe a execução
    
    
    st.write("Planilhas processadas com sucesso.")
    # Processar grupos
    groups_df = pd.read_excel(groups_file)
    groups_df['name'] = groups_df['name'].apply(remove_accents)
    groups_df['inicio'] = groups_df['inicio'].apply(remove_accents)
    groups_df['fim'] = groups_df['fim'].apply(remove_accents)
    
    
    survey = add_groups(survey, groups_df)
    survey=remover_grupos_vazios(survey)
    survey = adicionar_calculos_automaticos(survey, padroes_file)
    
    
    
    # Adicionar linhas padrão
    standard_rows = [
    ["start", "start", "", "", "", "", "", "", "", "", ""],
    ["end", "end", "", "", "", "", "", "", "", "", ""],
    ["start-geopoint", "start-geopoint", "", "", "", "", "", "", "", "", ""],
    ["today", "today", "", "", "", "", "", "", "", "", ""],
    ["username", "username", "", "", "", "", "", "", "", "", ""],
    ["deviceid", "deviceid", "", "", "", "", "", "", "", "", ""],
    ["phonenumber", "phonenumber", "", "", "", "", "", "", "", "", ""],
    ["audit", "audit", "", "", "", "", "", "", "", "", ""]
   ]
    
    survey = pd.concat([pd.DataFrame(standard_rows, columns=survey.columns), survey], ignore_index=True)
    
    # Criar abas adicionais
    choices = pd.DataFrame(columns=["list_name", "name", "label::Portugues (pt)"])
    if "choices" in survey.columns:
        choices = survey[["choices"]].dropna().drop_duplicates()
        choices = choices.assign(list_name=choices["choices"], name=choices["choices"], label=choices["choices"])
    
    settings = pd.DataFrame({"form_title": ["Formulário PAT"], "form_id": ["form_pat"]})
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        survey.to_excel(writer, sheet_name='survey', index=False)
        choices.to_excel(writer, sheet_name='choices', index=False)
        settings.to_excel(writer, sheet_name='settings', index=False)
    
    output.seek(0)
    return output


os.system("cls")
# Interface Streamlit
st.title("Conversor de Excel para XLSForm")
data_file = st.file_uploader("Arquivo principal com os dados", type=["xlsx"])
groups_file = st.file_uploader("Arquivo com a definição dos grupos", type=["xlsx"])
padroes_file = st.file_uploader("Arquivo com a definição dos somatorios", type=["xlsx"])

if data_file and groups_file and padroes_file:
    converted = convert_to_xlsform(data_file, groups_file, padroes_file)
    if converted:
        st.download_button(
            label="Baixar XLSForm",
            data=converted,
            file_name="formulario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )