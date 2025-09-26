'''naturezas_interesse = [
            'ROUBO', 'FURTO', 'HOMICIDIO','HOMICÍDIO', 'FEMINICIDIO', 'FEMINICÍDIO', 
            'HOMICÍDIO DOLOSO', 'FEMINICÍDIO DOLOSO', 'HOMICIDIO DOLOSO', 'FEMINICIDIO DOLOSO'
        ]
        
        # Converter para uppercase para comparação case-insensitive
        df['Natureza_Upper'] = df['Natureza Ocorrencia'].astype(str).str.upper().str.strip()
        
        # Mapear variações para nomes padronizados
        mapeamento_naturezas = {
            'HOMICÍDIO': 'HOMICIDIO',
            'FEMINICÍDIO': 'FEMINICIDIO',
            'HOMICÍDIO DOLOSO': 'HOMICIDIO',
            'HOMICIDIO DOLOSO': 'HOMICIDIO',
            'FEMINICIDIO DOLOSO': 'FEMINICIDIO',
            'FEMINICÍDIO DOLOSO': 'FEMINICIDIO'
        }'''

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import numpy as np
import os
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import warnings
from PIL import Image
warnings.filterwarnings('ignore')



#___________________________________________trocar as variáveis_____________________________________________________


tempo_periodo='dia 01 à 24 do Mês de Setembro de 2025' #####trocar aqui o período analisado

#________________________________________________________________________________________________
# Configuração dos arquivos e unidades
CONFIG_ARQUIVOS = {
    'rondonopolis': {
        'arquivo': 'rondonopolis.xlsx',
        'unidade': 'Sede',
       'cidade': 'Rondonópolis'
    },
    'pedra_preta': {
        'arquivo': 'pedra_preta.xlsx',
        'unidade': '1º Pelotão de Pedra Preta',
        'cidade': 'Pedra Preta'
    },
    'guiratinga': {
        'arquivo': 'guiratinga.xlsx',
        'unidade': '2º Pelotão de Guiratinga',
        'cidade': 'Guiratinga'
    },
     
    'itiquira': {
        'arquivo': 'itiquira.xlsx',
        'unidade': '3º Pelotão de Itiquira',
        'cidade': 'Itiquira'
    },
    'sao_jose_do_povo': {
        'arquivo': 'sao_jose_do_povo.xlsx',
        'unidade': 'NPM de São José do Povo',
        'cidade': 'São José do Povo'
    },
    'tesouro': {
        'arquivo': 'tesouro.xlsx',
        'unidade': 'NPM de Tesouro',
        'cidade': 'Tesouro'
    },
    

    
}


def adicionar_cabecalho(ax,unidade):
    """Adiciona o cabeçalho com brasões e textos"""
    try:
        # Carregar imagens
        plt.rcParams['figure.dpi'] = 300
        img_pm = np.array(Image.open("pmmt.png"))
        img_bpm = np.array(Image.open("bpm.png"))

        brasao_pm = OffsetImage(img_pm, zoom=0.15)  # Zoom aumentado
        brasao_bpm = OffsetImage(img_bpm, zoom=0.15)

      
        
        
        # Texto da instituição
        texto_instituicao = f"Polícia Militar do Estado de Mato Grosso\n4º Comando Regional\n5º Batalhão de Polícia Militar\n{unidade}"

        # Adicionar brasão da PMMT à esquerda
        if brasao_pm:
            ab_pm = AnnotationBbox(brasao_pm, (0.1, 0.85), 
                                  xycoords='axes fraction', 
                                  frameon=False, boxcoords="axes fraction")
            ax.add_artist(ab_pm)
        
        # Adicionar texto centralizado
        ax.text(0.5, 0.85, texto_instituicao, 
               ha='center', va='center', transform=ax.transAxes,
               fontsize=12, fontweight='bold', linespacing=1.5)
        
        # Adicionar brasão do 5º BPM à direita
        if brasao_bpm:
            ab_bpm = AnnotationBbox(brasao_bpm, (0.9, 0.85), 
                                   xycoords='axes fraction', 
                                   frameon=False, boxcoords="axes fraction")
            ax.add_artist(ab_bpm)
        
        # Adicionar linha separadora
        #ax.axhline(y=0.85, color='black', linewidth=1)
        
    except Exception as e:
        print(f"Erro ao adicionar cabeçalho: {e}")





# Configuração do estilo dos gráficos
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")

def ler_dados_excel(arquivo_excel):
    """
    Lê o arquivo dados.xlsx da mesma pasta do código
    """
    try:
        # Caminho do arquivo na mesma pasta
        #arquivo_excel = "dados.xlsx"
        
        # Verificar se o arquivo existe
        if not os.path.exists(arquivo_excel):
            raise FileNotFoundError(f"Arquivo {arquivo_excel} não encontrado na pasta do código")
        
        print("Lendo arquivo Excel...")
        # Ler o arquivo Excel
        df = pd.read_excel(arquivo_excel)
        print(f"Arquivo lido com sucesso. {len(df)} linhas encontradas.")
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Bairro', 'Dia Semana Fato', 'Desc Faixa 6Hora Fato', 'Natureza Ocorrencia']
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        
        if colunas_faltantes:
            print(f"Colunas faltantes: {colunas_faltantes}")
            print("Colunas disponíveis:", df.columns.tolist())
        
        print("Processando dados...")
        
        # Filtrar apenas as naturezas de interesse (incluindo variações possíveis)
        naturezas_interesse = [
            'ROUBO', 'FURTO', 'HOMICIDIO','HOMICÍDIO', 'FEMINICIDIO', 'FEMINICÍDIO', 
            'HOMICÍDIO DOLOSO', 'FEMINICÍDIO DOLOSO', 'HOMICIDIO DOLOSO', 'FEMINICIDIO DOLOSO'
        ]
        
        # Converter para uppercase para comparação case-insensitive
        df['Natureza_Upper'] = df['Natureza Ocorrencia'].astype(str).str.upper().str.strip()
        
        # Mapear variações para nomes padronizados
        mapeamento_naturezas = {
            'HOMICÍDIO': 'HOMICIDIO',
            'FEMINICÍDIO': 'FEMINICIDIO',
            'HOMICÍDIO DOLOSO': 'HOMICIDIO',
            'HOMICIDIO DOLOSO': 'HOMICIDIO',
            'FEMINICIDIO DOLOSO': 'FEMINICIDIO',
            'FEMINICÍDIO DOLOSO': 'FEMINICIDIO'
        }
        
        df['Natureza_Padronizada'] = df['Natureza_Upper'].replace(mapeamento_naturezas)
        df_filtrado = df[df['Natureza_Upper'].isin(naturezas_interesse)].copy()
        
        print(f"Linhas após filtragem por natureza: {len(df_filtrado)}")
        
        if len(df_filtrado) == 0:
            print("AVISO: Nenhum dado encontrado para as naturezas especificadas")
            print("Naturezas encontradas no arquivo:", df['Natureza Ocorrencia'].unique())
            df_filtrado = df.copy()
        
        return df_filtrado
    
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

def criar_grafico_bairros(df, titulo):
    """Cria gráfico de bairros"""
    fig, ax = plt.subplots(figsize=(8.27, 5.85))  # A4 em paisagem
    
    try:
        df_bairros = df.copy()
        df_bairros['Bairro'] = df_bairros['Bairro'].astype(str).str.strip()
        df_bairros['Bairro'] = df_bairros['Bairro'].replace(['nan', 'NaN', 'None', ''], 'NÃO INFORMADO')
        
        top_bairros = df_bairros['Bairro'].value_counts().head(5)
        
        bars = ax.bar(top_bairros.index, top_bairros.values, 
                     color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFD166'])
        
        ax.set_title(titulo, fontsize=14, fontweight='bold', pad=15)
        ax.set_xlabel('Bairro', fontsize=10)
        ax.set_ylabel('Número de Ocorrências', fontsize=10)
        ax.tick_params(axis='x', rotation=45, labelsize=9)
        ax.tick_params(axis='y', labelsize=9)
        ax.grid(True, alpha=0.3, axis='y')
        
        # Adicionar valores nas barras
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                   f'{int(height)}', ha='center', va='bottom', 
                   fontweight='bold', fontsize=10)
        
        plt.tight_layout()
        return fig
        
    except Exception as e:
        print(f"Erro no gráfico de bairros: {e}")
        ax.text(0.5, 0.5, 'Erro ao processar dados', 
               ha='center', va='center', transform=ax.transAxes, fontsize=12)
        return fig

def criar_grafico_dias_semana(df, titulo):
    """Cria gráfico de dias da semana"""
    fig, ax = plt.subplots(figsize=(8.27, 5.85))  # A4 em paisagem
    
    try:
        if 'Dia Semana Fato' not in df.columns:
            ax.text(0.5, 0.5, 'Dados não disponíveis', 
                   ha='center', va='center', transform=ax.transAxes, fontsize=12)
            return fig
        
        df_dias = df.copy()
        df_dias['Dia Semana Fato'] = df_dias['Dia Semana Fato'].astype(str).str.strip().str.upper()
        
        ordem_dias = ['SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 
                     'SEXTA-FEIRA', 'SÁBADO', 'DOMINGO']
        
        contagem_dias = df_dias['Dia Semana Fato'].value_counts()
        contagem_dias = contagem_dias.reindex(ordem_dias, fill_value=0)
        
        bars = ax.bar(contagem_dias.index, contagem_dias.values, 
                     color=['#FF9AA2', '#FFB7B2', '#FFDAC1', '#E2F0CB', '#B5EAD7', '#C7CEEA', '#F8B195'])
        
        ax.set_title(titulo, fontsize=14, fontweight='bold', pad=15)
        ax.set_xlabel('Dia da Semana', fontsize=10)
        ax.set_ylabel('Número de Ocorrências', fontsize=10)
        ax.tick_params(axis='x', rotation=45, labelsize=9)
        ax.tick_params(axis='y', labelsize=9)
        ax.grid(True, alpha=0.3, axis='y')
        
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                       f'{int(height)}', ha='center', va='bottom', 
                       fontweight='bold', fontsize=10)
        
        plt.tight_layout()
        return fig
        
    except Exception as e:
        print(f"Erro no gráfico de dias: {e}")
        ax.text(0.5, 0.5, 'Erro ao processar dados', 
               ha='center', va='center', transform=ax.transAxes, fontsize=12)
        return fig

def criar_grafico_horarios(df, titulo):
    """Cria gráfico de faixas horárias"""
    fig, ax = plt.subplots(figsize=(8.27, 5.85))  # A4 em paisagem
    
    try:
        if 'Desc Faixa 6Hora Fato' not in df.columns:
            ax.text(0.5, 0.5, 'Dados não disponíveis', 
                   ha='center', va='center', transform=ax.transAxes, fontsize=12)
            return fig
        
        df_horarios = df.copy()
        df_horarios['Desc Faixa 6Hora Fato'] = df_horarios['Desc Faixa 6Hora Fato'].astype(str).str.strip()
        
        ordem_faixas = ['00:00 AS 05:59', '06:00 AS 11:59', '12:00 AS 17:59', '18:00 AS 23:59']
        
        contagem_faixas = df_horarios['Desc Faixa 6Hora Fato'].value_counts()
        contagem_faixas = contagem_faixas.reindex(ordem_faixas, fill_value=0)
        
        bars = ax.bar(contagem_faixas.index, contagem_faixas.values, 
                     color=['#264653', '#2A9D8F', '#E9C46A', '#F4A261'])
        
        ax.set_title(titulo, fontsize=14, fontweight='bold', pad=15)
        ax.set_xlabel('Faixa Horária', fontsize=10)
        ax.set_ylabel('Número de Ocorrências', fontsize=10)
        ax.tick_params(axis='x', rotation=45, labelsize=9)
        ax.tick_params(axis='y', labelsize=9)
        ax.grid(True, alpha=0.3, axis='y')
        
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                       f'{int(height)}', ha='center', va='bottom', 
                       fontweight='bold', fontsize=10)
        
        plt.tight_layout()
        return fig
        
    except Exception as e:
        print(f"Erro no gráfico de horários: {e}")
        ax.text(0.5, 0.5, 'Erro ao processar dados', 
               ha='center', va='center', transform=ax.transAxes, fontsize=12)
        return fig

def criar_pagina_capa(total_ocorrencias,unidade,cidade):
    """Cria página de capa"""
    #fig = plt.figure(figsize=(8.27, 11.69))  # A4 em retrato
    #plt.axis('off')

    fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 em retrato
    ax.axis('off')

    adicionar_cabecalho(ax,unidade)
    
    texto_capa = f"""
    RELATÓRIO DE ANÁLISE CRIMINAL
    
    Análise Detalhada por Tipo de Crime
    Município: {cidade} 
    
    Total de Ocorrências: {total_ocorrencias:,}
    
    Relatório Inclui:
    • Análise Geral (Todos os Crimes)
    • Análise Específica por Tipo de Crime:
      - Roubo
      - Furto
      - Homicídio
      - Feminicídio
    
    Para cada tipo de crime são apresentados:
    • Top 5 bairros
    • Distribuição por dia da semana
    • Ocorrências por faixa horária

    Windsney de Oliveira Bandeira -MAJ PM
    Seção de Planejamento e Estatística do 5º BPM-{datetime.now().strftime('%d/%m/%Y')}
    """

    plt.text(0.5, 0.7, f'RELATÓRIO DE ANÁLISE CRIMINAL\n{tempo_periodo}', 
            ha='center', va='center', fontsize=16, fontweight='bold', 
            transform=plt.gca().transAxes)
    plt.text(0.5, 0.4, texto_capa, ha='center', va='center', 
            fontsize=11, transform=plt.gca().transAxes,
            bbox=dict(boxstyle="round,pad=1", facecolor="lightgray", alpha=0.8))
    
    plt.tight_layout()
    return fig

def criar_relatorio_pdf(df, nome_arquivo,unidade,cidade):
    
    """
    Cria relatório em PDF com gráficos individuais para cada natureza
    """
    print("Criando relatório PDF...")
    
    try:
        with PdfPages(nome_arquivo) as pdf:
            total_ocorrencias = len(df) if df is not None else 0
            
            # Página 1: Capa
            print("Criando capa...")
            fig_capa = criar_pagina_capa(total_ocorrencias,unidade,cidade)
            pdf.savefig(fig_capa, bbox_inches='tight')
            plt.close()
            
            if df is not None and not df.empty:
                # ANÁLISE GERAL (TODOS OS CRIMES)
                print("Criando análise geral...")
                
                # Gráfico 1: Bairros (Geral)
                fig = criar_grafico_bairros(df, "TOP 5 BAIRROS - TODOS OS CRIMES")
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
                
                # Gráfico 2: Dias da semana (Geral)
                fig = criar_grafico_dias_semana(df, "DISTRIBUIÇÃO POR DIA DA SEMANA - TODOS OS CRIMES")
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
                
                # Gráfico 3: Horários (Geral)
                fig = criar_grafico_horarios(df, "DISTRIBUIÇÃO POR FAIXA HORÁRIA - TODOS OS CRIMES")
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()


                # Gráfico 4: Heatmap por dia e natureza
                print("Criando heatmap...")
                fig, ax = plt.subplots(figsize=(12, 8))
                
                # PREPARAR DADOS PARA HEATMAP - CORREÇÃO AQUI
                if 'Dia Fato' in df.columns and 'Natureza_Padronizada' in df.columns:
                    # Agrupar por dia e natureza
                    dados_dia = df.groupby(['Dia Fato', 'Natureza_Padronizada']).size().unstack(fill_value=0)
                    
                    # Completar todos os dias de 1 a 31
                    todos_dias = range(1, 32)
                    dados_dia = dados_dia.reindex(todos_dias, fill_value=0)
                    
                    # Criar heatmap
                    sns.heatmap(dados_dia.T, annot=True, fmt='d', cmap='YlOrRd', 
                               linewidths=0.5, ax=ax, cbar_kws={'label': 'Número de Ocorrências'})
                    
                    ax.set_title('HEATMAP - DISTRIBUIÇÃO DIÁRIA POR TIPO DE CRIME', 
                               fontsize=16, fontweight='bold')
                    ax.set_xlabel('Dia do Mês')
                    ax.set_ylabel('Tipo de Crime')
                    
                else:
                    ax.text(0.5, 0.5, 'Dados insuficientes para heatmap\n(necessária coluna "Dia Fato" e "Natureza_Padronizada")', 
                           ha='center', va='center', transform=ax.transAxes, fontsize=12)
                    ax.set_title('HEATMAP - DADOS INSUFICIENTES', fontsize=16, fontweight='bold')
                
                plt.tight_layout()
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()




                
                # ANÁLISE POR TIPO DE CRIME
                if 'Natureza_Padronizada' in df.columns:
                    naturezas = df['Natureza_Padronizada'].unique()
                    
                    for natureza in naturezas:
                        print(f"Criando análise para {natureza}...")
                        df_natureza = df[df['Natureza_Padronizada'] == natureza]
                        
                        if len(df_natureza) > 0:
                            # Página de separação para cada natureza
                            fig_sep = plt.figure(figsize=(8.27, 11.69))
                            plt.axis('off')
                            plt.text(0.5, 0.5, f'ANÁLISE: {natureza}\n\n{len(df_natureza)} ocorrências', 
                                    ha='center', va='center', fontsize=16, fontweight='bold',
                                    transform=plt.gca().transAxes)
                            pdf.savefig(fig_sep, bbox_inches='tight')
                            plt.close()
                            
                            # Gráfico 1: Bairros para a natureza
                            fig = criar_grafico_bairros(df_natureza, f"TOP 5 BAIRROS - {natureza}")
                            pdf.savefig(fig, bbox_inches='tight')
                            plt.close()
                            
                            # Gráfico 2: Dias da semana para a natureza
                            fig = criar_grafico_dias_semana(df_natureza, f"DISTRIBUIÇÃO POR DIA DA SEMANA - {natureza}")
                            pdf.savefig(fig, bbox_inches='tight')
                            plt.close()
                            
                            # Gráfico 3: Horários para a natureza
                            fig = criar_grafico_horarios(df_natureza, f"DISTRIBUIÇÃO POR FAIXA HORÁRIA - {natureza}")
                            pdf.savefig(fig, bbox_inches='tight')
                            plt.close()
                
                # Página final com resumo
                fig_resumo = plt.figure(figsize=(8.27, 11.69))
                plt.axis('off')
                
                texto_resumo = "RESUMO ESTATÍSTICO\n\n"
                texto_resumo += "="*50 + "\n\n"
                texto_resumo += f"Total de Ocorrências: {len(df):,}\n\n"
                
                if 'Natureza_Padronizada' in df.columns:
                    texto_resumo += "Distribuição por Tipo de Crime:\n"
                    distribuicao = df['Natureza_Padronizada'].value_counts()
                    for nat, total in distribuicao.items():
                        perc = (total / len(df)) * 100
                        texto_resumo += f"• {nat}: {total} ({perc:.1f}%)\n"
                
                texto_resumo += f"\nRelatório de {cidade} Produzido em: {datetime.now().strftime('%d/%m/%Y')}\nSeção de Planejamento e Estatística do 5º BPM"
                
                plt.text(0.1, 0.9, texto_resumo, fontsize=12, 
                        fontfamily='monospace', transform=plt.gca().transAxes,
                        verticalalignment='top', linespacing=1.5)
                pdf.savefig(fig_resumo, bbox_inches='tight')
                plt.close()
                
            else:
                # Página alternativa se não houver dados
                fig = plt.figure(figsize=(8.27, 11.69))
                plt.axis('off')
                plt.text(0.5, 0.5, 'Nenhum dado disponível para análise', 
                        ha='center', va='center', fontsize=16, 
                        transform=plt.gca().transAxes)
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
        
        print(f"✓ Relatório PDF gerado: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"✗ Erro ao criar PDF: {e}")
        return False

def processar_unidade(chave_unidade, config):
    """Processa uma unidade específica"""
    print(f"\n{'='*60}")
    print(f"PROCESSANDO: {config['cidade']}")
    print(f"{'='*60}")
    
    arquivo = config['arquivo']
    unidade = config['unidade']
    
    cidade = config['cidade']
    
    if not os.path.exists(arquivo):
        print(f"✗ Arquivo {arquivo} não encontrado. Pulando...")
        return False
    
    # Ler dados
    df = ler_dados_excel(arquivo)
    
    if df is not None and not df.empty:
        print(f"✓ Dados carregados: {len(df)} ocorrências")
        
        # Nome do arquivo PDF de saída
        nome_pdf = f"relatorio_{chave_unidade}.pdf"
        
        # Gerar relatório
        sucesso = criar_relatorio_pdf(df, nome_pdf,unidade,cidade)
        
        if sucesso:
            print(f"✓ Relatório gerado com sucesso: {nome_pdf}")
            return True
        else:
            print(f"✗ Erro ao gerar relatório para")
            return False
    else:
        print(f"✗ Não foi possível processar os dados de")
        return False

def main():
    """Função principal que processa todas as unidades"""
    print("PROCESSADOR DE RELATÓRIOS CRIMINAIS")
    print("="*60)
    
    # Verificar se as imagens existem
    if not os.path.exists("pmmt.png"):
        print("AVISO: Arquivo 'pmmt.png' não encontrado")
    if not os.path.exists("bpm.png"):
        print("AVISO: Arquivo 'bpm.png' não encontrado")
    
    # Processar cada unidade
    resultados = {}
    
    for chave_unidade, config in CONFIG_ARQUIVOS.items():
        resultados[chave_unidade] = processar_unidade(chave_unidade, config)
    
    # Resumo final
    print(f"\n{'='*60}")
    print("RESUMO DA EXECUÇÃO")
    print(f"{'='*60}")
    
    sucessos = sum(resultados.values())
    total = len(resultados)
    
    print(f"Unidades processadas com sucesso: {sucessos}/{total}")
    
    for chave_unidade, sucesso in resultados.items():
        status = "✓ SUCESSO" if sucesso else "✗ FALHA"
        print(f"{status} - {CONFIG_ARQUIVOS[chave_unidade]['cidade']}")
    
    if sucessos == total:
        print("\n🎉 TODOS OS RELATÓRIOS FORAM GERADOS COM SUCESSO!")
    else:
        print(f"\n⚠️  {total - sucessos} relatório(s) não foram gerados.")

if __name__ == "__main__":
    os.environ['MPLBACKEND'] = 'Agg'
    main()