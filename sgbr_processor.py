import re
import pandas as pd
from io import BytesIO

def process_sgbr_pdf(texto_completo):
    """
    Processa um PDF do modelo SGBr, extrai a tabela, ordena por Número em ordem crescente,
    adiciona números ausentes e retorna o DataFrame e o Excel em memória.
    """
    dados = []
    
    # Padrões regex para identificação das colunas
    numero_pattern = r'^\d{5}$'  # Exatamente 5 dígitos para Número (ajustado conforme dados)
    modelo_pattern = r'^\d{2}$'  # Exatamente 2 dígitos para Modelo
    serie_pattern = r'^\d$'      # Exatamente 1 dígito para Série
    data_pattern = r'^\d{2}/\d{2}/\d{4}$'  # DD/MM/YYYY
    total_pattern = r'^\d{1,3}(?:[,\.]\d{2})$'  # e.g., 47,00, 293,80
    
    # Filtrar linhas indesejadas
    linhas = texto_completo.split('\n')
    linhas_filtradas = [
        linha for linha in linhas
        if not (
            re.match(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', linha.strip()) or
            linha.strip().startswith('Relatório') or
            linha.strip().lower().startswith('https') or
            linha.strip().lower().startswith('valor total') or
            "Número ModeloSérie Natureza de operação CPF/CNPJ Chave de acesso Protocolo aut. Data emis. Total nota" in linha or
            "Número Modelo Série Natureza de operação CPF/CNPJ Chave de acesso Protocolo aut. Data emis. Total nota" in linha or
            linha.strip().startswith('SGBr Sistemas') or 
            linha.strip().startswith('Status NFC-e:') or 
            linha.strip().startswith('Número ModeloSérie') or
            linha.strip().startswith('Totais') or
            linha.strip().startswith('Canceladas') or
            linha.strip().startswith('Rejeitadas') or
            linha.strip().startswith('Contingência') or
            linha.strip().startswith('Não enviadas') or
            linha.strip().startswith('Total líquido')
        )
    ]
    
    # Processar linhas para o DataFrame
    for linha in linhas_filtradas:
        partes = linha.strip().split()
        if len(partes) >= 6:  # Mínimo de partes para uma linha válida
            numero = partes[0] if partes[0] and re.match(numero_pattern, partes[0]) else None
            modelo = partes[1] if partes[1] and re.match(modelo_pattern, partes[1]) else None
            serie = partes[2] if partes[2] and re.match(serie_pattern, partes[2]) else None
            data = partes[-2] if partes[-2] and re.match(data_pattern, partes[-2]) else None
            total = partes[-1] if partes[-1] and re.match(total_pattern, partes[-1]) else None
            descricao = ' '.join(partes[3:-2]).strip() if len(partes) > 6 else ''
            
            # Adicionar linha apenas se todos os campos obrigatórios forem válidos
            if numero and modelo and serie and data and total:
                dados.append([numero, modelo, serie, descricao, data, total])
    
    # Criar DataFrame
    colunas = ['Número', 'Modelo', 'Série', 'Descrição', 'Data emis.', 'Total nota']
    df = pd.DataFrame(dados, columns=colunas)
    
    # Converter 'Número' para inteiro para ordenação e preenchimento
    df['Número'] = df['Número'].astype(int)
    
    # Identificar números ausentes na sequência
    if not df.empty:
        min_numero = df['Número'].min()
        max_numero = df['Número'].max()
        todos_numeros = set(range(min_numero, max_numero + 1))
        numeros_presentes = set(df['Número'])
        numeros_ausentes = todos_numeros - numeros_presentes
        
        # Adicionar linhas para números ausentes
        for numero in numeros_ausentes:
            dados.append([f'{numero:05d}', 'NaN', 'NaN', 'NaN', 'NaN', 'NaN'])

    # Recriar DataFrame com números ausentes
    df = pd.DataFrame(dados, columns=colunas)
    
    # Ordenar por 'Número' em ordem crescente
    df['Número'] = df['Número'].astype(int)
    df = df.sort_values(by='Número', ascending=True)
    df['Número'] = df['Número'].apply(lambda x: f'{x:05d}')
    
    # Gerar Excel em memória
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    return output
