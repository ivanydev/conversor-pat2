import re

import pandas as pd


def limpar_label(label):
    """
    Remove partes incorretas de um label mantendo os valores dentro de ${}.
    """
    padrao_corrupto = r'Total de (Qea|Qee|Qepe) .*? (?=\d{1,2} Num|Total)'  # Remover tudo até números ou palavras-chave úteis
    
    # Separar texto e variáveis (${...})
    partes = re.split(r'(\${.*?})', label)
    
    # Limpar apenas as partes de texto
    partes_corrigidas = [re.sub(padrao_corrupto, 'Total ', p) if not p.startswith('${') else p for p in partes]
    
    return ''.join(partes_corrigidas)

def corrigir_xlsform(caminho_arquivo):
    """
    Lê um arquivo XLSForm, corrige a coluna label::Portugues (pt) e salva um novo arquivo.
    """
    df = pd.read_excel(caminho_arquivo,  header=None)
    
    # Verificar se a coluna existe
    coluna_label = 'label::Portugues (pt)'
    if coluna_label not in df.columns:
        raise ValueError(f"A coluna '{coluna_label}' não foi encontrada no arquivo.")
    
    # Aplicar a correção
    df[coluna_label] = df[coluna_label].apply(limpar_label)
    
    # Salvar o arquivo corrigido
    novo_caminho = caminho_arquivo.replace('.xls', '_corrigido.xls')
    df.to_excel(novo_caminho, index=False)
    
    return novo_caminho

corrigir_xlsform("c:\\Users\\a\\Downloads\\FORMULÁRIO GERAL.xlsx")