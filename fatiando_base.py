import pandas as pd
import os

def separar_municipios_simples():
    """Versão simples apenas para separar os arquivos"""
    
    # Ler arquivo principal
    df = pd.read_excel("dados.xlsx")
    
    # Listar municípios únicos
    municipios = df['Municipio'].astype(str).str.strip().unique()
    
    print(f"🏙️ Separando {len(municipios)} municípios...")
    
    for municipio in municipios:
        # Filtrar dados do município
        df_municipio = df[df['Municipio'].astype(str).str.strip() == municipio]
        
        # Criar nome do arquivo
        nome_arquivo = municipio.lower().replace(' ', '_') + ".xlsx"
        
        # Salvar arquivo separado
        df_municipio.to_excel(nome_arquivo, index=False)
        print(f"✅ {municipio}: {len(df_municipio)} registros -> {nome_arquivo}")

# Executar a separação
separar_municipios_simples()