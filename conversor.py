import os
import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st

# Criar um espaço vazio para "limpar" a tela
placeholder = st.empty()

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
 

 

def aplicar_regex(df):
    arquivo_validacoes="regex.xlsx"
    # Carregar a tabela de validações
    validacoes = pd.read_excel(arquivo_validacoes)

    # Garantir que as colunas necessárias estão presentes
    colunas_necessarias = {"padrao", "excepto", "constraint", "constraint_message"}
    if not colunas_necessarias.issubset(validacoes.columns):
        raise ValueError(f"O arquivo {arquivo_validacoes} deve conter as colunas: {colunas_necessarias}")

    # Adiciona colunas de validação ao DataFrame original
    df["constraint"] = None
    df["constraint_message"] = None

    for index, row in validacoes.iterrows():
        padrao = row["padrao"].lower().strip()
        excepto = str(row["excepto"]).lower().strip() if pd.notna(row["excepto"]) else None
        constraint = row["constraint"]
        constraint_message = row["constraint_message"]

        # Aplicar a validação: se o padrão estiver na variável e excepto não estiver
        mask = df["name"].str.contains(padrao, case=False, na=False)
        if excepto:
            mask &= ~df["name"].str.contains(excepto, case=False, na=False)

        # Aplicar as constraints às linhas que atendem ao critério
        df.loc[mask, "constraint"] = constraint
        df.loc[mask, "constraint_message"] = constraint_message
        #df.loc[mask, "appearance"] = "w10"

    return df



 
def atualizar_df_com_selects(df, caminho_selects):
    """
    Atualiza o DataFrame com os campos relevant, choice_filter e type
    com base no arquivo selects.xlsx.
    """
    try:
        # Ler o arquivo de selects
        selects_df = pd.read_excel(caminho_selects)

        # Normalizar os nomes das colunas (remover espaços e converter para minúsculas)
        selects_df.columns = selects_df.columns.str.strip().str.lower()

        # Verificar se as colunas necessárias estão presentes
        colunas_necessarias = {"type", "variavel", "choice_filter"}
        colunas_arquivo = set(selects_df.columns)
        
        if not colunas_necessarias.issubset(colunas_arquivo):
            colunas_faltantes = colunas_necessarias - colunas_arquivo
            raise ValueError(f"O arquivo {caminho_selects} deve conter as colunas: {colunas_faltantes}")

        # Criar uma cópia do DataFrame para evitar modificações inplace
        novo_df = df.copy()

        # Iterar sobre as linhas do arquivo de selects
        for _, select_row in selects_df.iterrows():
            variavel = select_row["variavel"]
            tipo = select_row["type"]
            choice_filter = select_row.get("choice_filter", "")  # Usar get para evitar KeyError

            # Criar a máscara correta para encontrar as variáveis que terminam com 'variavel'
            mask = novo_df["name"].str.endswith(variavel, na=False)
                        # Pegar o primeiro índice onde a condição é verdadeira
  
            # Verificar se a máscara encontrou algo antes de atualizar
            if mask.any():
                # Atualizar os valores apenas nas linhas filtradas
                novo_df.loc[mask, "type"] = tipo
                if pd.notna(choice_filter):
                    first_index = mask.idxmax()
                    variavel_original = novo_df.loc[first_index, "name"]
                    ListChoices=choice_filter.split("=")
                    choice_1=ListChoices[0]
                    choice_2=ListChoices[1].replace('${','').replace('}','')
                    prefixo=variavel_original.split('_')[0]
                    choice_final=f"{choice_1}=${{{prefixo}_{choice_2}}}"
                    
                    novo_df.loc[mask, "choice_filter"] = f"{choice_final}"  
            
        return novo_df

    except Exception as e:
        raise ValueError(f"Erro ao processar o arquivo {caminho_selects}: {str(e)}")


 
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
    #st.write(f"Variaveis existentes no config.xls: {survey_df['name'].tolist()}")
    for _ in range(2):  # Duas verificações
        for _, group in groups_df.iterrows():
            group_name = group['name']
            #st.write(f"======================================================================================")
            #st.write(f"Processando grupo: {group_name}")
              # Limpar a tela
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
            st.write(f"Grupos adicionados com sucesso. {group_name}")
    return survey_df
 
#=========================================================================
# 📌 Dicionário de regras sem os prefixos (mapeia somente o sufixo real da variável)

# 📌 Dicionário de regras sem os prefixos
REGRAS = {
    "DGE_SQE_B0_P0_id_questionario": {
        "constraint": "regex(., '^[0-9]{1,10}$')",
        "constraint_msg": "Deve conter somente dígitos e ter no máximo 10 caracteres.",
        "calculation": "substr(uuid(), 0, 8)",
        "label_varavel": "ID do questionário"
    },
    "DGE_SQE_B0_P1_codigo_escola": {
        "constraint": "regex(., '^[0-9]{1,11}$')",
        "constraint_msg": "Deve conter somente dígitos e ter no máximo 11 caracteres.",
        "calculation": "substr(uuid(), 0, 8)",
        "label_varavel": "Código da escola"
    },
    "DGE_SQE_B0_P3_fim_ano_lectivo": {
        "calculation": lambda var_name: f"${{{var_name.replace('fim', 'inicio').replace('P3', 'P2')}}} + 1",
        "constraint": "",
        "constraint_msg": "",
        "label_varavel": "fim do ano letivo"
    }
}

def gerar_campos_automaticos(df, variaveis):
    """
    Modifica variáveis existentes para 'calculate' e cria 'notes' correspondentes.
    Agora funciona para qualquer variável automática sem depender dos prefixos do questionário.
    """
    df = df.copy()

    # 🔹 Remove valores NaN na coluna "name"
    df = df.dropna(subset=["name"])

    for var_sufixo in reversed(variaveis):
        # 🔍 Encontra qualquer variável que termine exatamente com o nome esperado
        match_indices = df.index[df['name'].str.endswith(var_sufixo, na=False)].tolist()

        if not match_indices:
            st.warning(f"Variável terminando com '{var_sufixo}' não encontrada. Pulando...")
            continue

        # Pega o primeiro índice correspondente
        idx = match_indices[0]
        var_name = df.at[idx, 'name']

        # Aplica regras, se existirem
        if var_sufixo in REGRAS:
            regra = REGRAS[var_sufixo]
            calculation = regra["calculation"] if isinstance(regra["calculation"], str) else regra["calculation"](var_name)
            constraint = regra["constraint"]
            constraint_msg = regra["constraint_msg"]
            label_varavel=regra["label_varavel"]
        else:
            st.warning(f"Regras para '{var_sufixo}' não encontradas. Pulando...")
            continue

        # Modifica a linha existente (calculate)
        df.at[idx, 'type'] = 'calculate'
        df.at[idx, 'calculation'] = calculation
        df.at[idx, 'constraint'] = constraint
        df.at[idx, 'constraint_message'] = constraint_msg
        df.at[idx, 'label::Portugues (pt)'] = f'Valor gerado automaticamente para {var_sufixo.replace("_", " ")}'

        # Criar uma linha "note" dinâmica abaixo
        note_row = {
            'type': 'note',
            'name': f'show_aux_{var_sufixo.split("_")[-1]}',
            'label::Portugues (pt)': f'{label_varavel} : ${{{var_name}}}',
            'hint::Portugues (pt)': '',
            'required': 'false',
            'appearance': '',
            'constraint': '',
            'calculation': '',
            'constraint_message': '',
            'relevant': '',
            'choice_filter': ''
        }

        # Inserir a linha note logo abaixo
        df.loc[idx + 0.5] = note_row

    # Reordenar e resetar índices
    return df.sort_index().reset_index(drop=True)



#=========================================================================
 
# Função para adicionar cálculos automáticos baseados em padrões de um Excel
def adicionar_calculos_automaticos(df, excel_path):
    """
    Adiciona cálculos automáticos baseados em padrões de um Excel, evitando ciclos.

    Parâmetros:
        df (pd.DataFrame): DataFrame principal do formulário
        excel_path (str): Caminho para o Excel com os padrões

    Retorna:
        pd.DataFrame: DataFrame atualizado com os cálculos adicionados.
    """
    try:
        padroes_df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Erro ao ler arquivo de padrões: {str(e)}")
        return df

    if not all(col in padroes_df.columns for col in ['name', 'pergunta', 'padrao', 'excepto']):
        print("Arquivo de padrões deve conter as colunas: name, pergunta, padrao, excepto")
        return df

    existing_calculations = df.set_index('name')['calculation'].dropna().to_dict()

    def has_cycle(var, visited):
        """ Verifica se há um ciclo nos cálculos antes de adicionar. """
        if var in visited:
            print(f"⚠️ Ciclo detectado: {var}")
            return True  # Ciclo detectado
        if var not in existing_calculations:
            return False  # Variável não tem cálculo ainda

        visited.add(var)
        for ref_var in existing_calculations[var].split('+'):
            ref_var = ref_var.strip('${}')
            if has_cycle(ref_var, visited):
                return True
        visited.remove(var)
        return False

    for _, row in padroes_df.iterrows():
        target_var = row['name']
        pergunta = str(row['pergunta']).strip()
        padroes = [p.strip().lower() for p in str(row['padrao']).split(',')]
        excepto = [e.strip().lower() for e in str(row['excepto']).split(',') if e.strip()]
        #st.write(f"Variável alvo: {target_var}")
        #st.write(f"Padroes: {padroes}")
        #st.write(f"Excepto: {excepto}")
        #st.write(f"Pergunta: {pergunta}")
        #st.write(f"=============================================================")

        if target_var not in df['name'].values:
            print(f"⚠️ Variável alvo '{target_var}' não encontrada no formulário.")
            continue

        # Filtrar variáveis da mesma pergunta
        pergunta_filter = df['name'].str.contains(f'_{pergunta}_', case=False, na=False)
        vars_pergunta = df[pergunta_filter]['name'].tolist()

        # Filtrar variáveis que devem ser somadas, excluindo as do "excepto"
        vars_somar = []
        for var in vars_pergunta:
            var_clean = var.lower()
            if var_clean == target_var.lower():
                continue
                #st.write(f"Variavel DO RESULTADO DA SOMA : {target_var}")
                #st.write(f"Variável somando: {var_clean}")
                #st.write(f"Padrão está em var: {any(padrao in var_clean for padrao in padroes)}")
                #st.write(f"Exceto está em var: {any(exc in var_clean for exc in excepto)}")
                #st.write(f"LISTA DE padoes:{padroes}")
                
            if any(padrao in var_clean for padrao in padroes)==True and not any(exc in var_clean for exc in excepto):
                vars_somar.append(var)
         
        if not vars_somar:
            #print(f"⚠️ Nenhuma variável encontrada para {target_var} com padrões: {', '.join(padroes)} (exceto: {', '.join(excepto)})")
            continue

        new_calculation = '+'.join([f'${{{var}}}' for var in vars_somar])

        if has_cycle(target_var, set()):
            print(f"❌ Cálculo ignorado para {target_var} para evitar ciclo.")
            continue

        # Atualizar o cálculo na variável alvo
        df.loc[df['name'] == target_var, 'calculation'] = new_calculation
        df.loc[df['name'] == target_var, 'type'] = 'calculate'
    st.write("Cálculos automáticos adicionados com sucesso.")
    st.write("==CONCLUÍDO==")
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
    
    
    
    #survey=remover_grupos_vazios(survey)
    survey = adicionar_calculos_automaticos(survey, padroes_file)
    # Lista de variáveis para automação
    survey = gerar_campos_automaticos(survey, ['DGE_SQE_B0_P0_id_questionario', 'DGE_SQE_B0_P1_codigo_escola','DGE_SQE_B0_P3_fim_ano_lectivo'])
    survey=aplicar_regex(survey)
    #survey=adicionar_validacao_tempo_real(survey)
    survey=atualizar_df_com_selects(survey, "selects.xlsx")
    survey = add_groups(survey, groups_df)
    
    
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
    
    
   # survey = pd.concat([pd.DataFrame(standard_rows, columns=survey.columns), survey], ignore_index=True)
    
        # Criar um DataFrame vazio com as mesmas colunas de survey
    standard_rows_df = pd.DataFrame(columns=survey.columns)

    # Adicionar os dados corretamente
    for row in standard_rows:
        row_dict = dict(zip(survey.columns, row + [""] * (len(survey.columns) - len(row))))
        standard_rows_df = pd.concat([standard_rows_df, pd.DataFrame([row_dict])], ignore_index=True)

    # Concatenar com survey
    survey = pd.concat([standard_rows_df, survey], ignore_index=True)
    
    
    # Criar abas adicionais
    #choices = pd.DataFrame(columns=["list_name", "name", "label::Portugues (pt)"])
    #if "choices" in survey.columns:
    #    choices = survey[["choices"]].dropna().drop_duplicates()
    #    choices = choices.assign(list_name=choices["choices"], name=choices["choices"], label=choices["choices"])
        
    # Carregar o arquivo choiceGood.xlsx
    caminho_choices = "formWithChoiceGood.xlsx"
    # Ler a aba "choices" do arquivo
    choices = pd.read_excel(caminho_choices, sheet_name="choices")
    # Verificar se o arquivo foi carregado corretamente
    if choices.empty:
        raise ValueError("A aba 'choices' do arquivo formWithChoiceGood.xlsx está vazia!")
    
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