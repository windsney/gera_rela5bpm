import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import numpy as np
import os
from matplotlib.backends.backend_pdf import PdfPages
import warnings
warnings.filterwarnings('ignore')

# Configuração do estilo dos gráficos
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")

#___________________________________________trocar as variáveis_____________________________________________________

unidade='5º BPM-Sede'#####trocar aqui o nome da unidade
tempo_periodo='Período: Dias 1 a 31' #####trocar aqui o período analisado

#________________________________________________________________________________________________


def ler_dados_excel():
    """
    Lê o arquivo dados.xlsx da mesma pasta do código
    """
    try:
        # Caminho do arquivo na mesma pasta
        arquivo_excel = "dados.xlsx"
        
        # Verificar se o arquivo existe
        if not os.path.exists(arquivo_excel):
            raise FileNotFoundError(f"Arquivo {arquivo_excel} não encontrado na pasta do código")
        
        print("Lendo arquivo Excel...")
        # Ler o arquivo Excel
        df = pd.read_excel(arquivo_excel)
        print(f"Arquivo lido com sucesso. {len(df)} linhas encontradas.")
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Natureza Ocorrencia', 'Ano Fato', 'Mes Fato', 'Dia Fato']
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        
        if colunas_faltantes:
            raise ValueError(f"Colunas faltantes: {colunas_faltantes}")
        
        print("Processando datas...")
        # Criar coluna de data completa
        df['Data_Fato'] = pd.to_datetime({
            'year': df['Ano Fato'],
            'month': df['Mes Fato'],
            'day': df['Dia Fato']
        }, errors='coerce')
        
        # Remover linhas com datas inválidas
        linhas_antes = len(df)
        df = df.dropna(subset=['Data_Fato'])
        linhas_depois = len(df)
        print(f"Linhas com datas válidas: {linhas_depois}/{linhas_antes}")
        
        if linhas_depois == 0:
            raise ValueError("Nenhuma data válida encontrada após processamento")
        
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
        
        # Aplicar padronização
        df_filtrado.loc[:, 'Natureza_Padronizada'] = df_filtrado['Natureza_Upper'].replace(mapeamento_naturezas)
        
        print(f"Linhas após filtragem por natureza: {len(df_filtrado)}")
        print("Naturezas encontradas:", df_filtrado['Natureza_Padronizada'].unique())
        
        if len(df_filtrado) == 0:
            print("AVISO: Nenhum dado encontrado para as naturezas especificadas")
            print("Naturezas encontradas no arquivo:", df['Natureza Ocorrencia'].unique())
        
        return df_filtrado
    
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

def criar_relatorio_pdf(df, nome_arquivo=f"relatorio_{unidade}.pdf"):
    """
    Cria relatório em PDF com análise do período completo (1-31 dias)
    """
    print("Criando relatório PDF...")
    
    try:
        with PdfPages(nome_arquivo) as pdf:
            # Página 1: Capa
            fig = plt.figure(figsize=(11, 8.5))
            plt.axis('off')
            
            info_periodo = "Período não disponível"
            total_ocorrencias = 0
            mes_analisado = "Não especificado"
            
            if df is not None and not df.empty:
                info_periodo = f"{df['Data_Fato'].min().strftime('%d/%m/%Y')} a {df['Data_Fato'].max().strftime('%d/%m/%Y')}"
                total_ocorrencias = len(df)
                mes_analisado = df['Mes Fato'].iloc[0] if 'Mes Fato' in df.columns else "Não especificado"
            
            texto_capa = f"""
            RELATÓRIO COMPLETO DE ANÁLISE CRIMINAL - {unidade}
            
            Análise do Mês {mes_analisado}
            Período: {tempo_periodo}
            
            Período Analisado: {info_periodo}
            Total de Ocorrências: {total_ocorrencias}
            
            Naturezas Analisadas:
            • Roubo
            • Furto  
            • Homicídio
            • Feminicídio
            
            Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}
            """
            
            plt.text(0.5, 0.7, 'RELATÓRIO COMPLETO DE CRIMES', 
                    ha='center', va='center', fontsize=18, fontweight='bold', 
                    transform=plt.gca().transAxes)
            plt.text(0.5, 0.4, texto_capa, ha='center', va='center', 
                    fontsize=11, transform=plt.gca().transAxes,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray"))
            
            pdf.savefig(fig)
            plt.close()
            
            if df is not None and not df.empty:
                # Página 2: Gráfico de barras - Distribuição geral
                fig, ax = plt.subplots(figsize=(12, 8))
                
                # Contagem por natureza
                contagem_naturezas = df['Natureza_Padronizada'].value_counts()
                
                # Gráfico de barras
                bars = ax.bar(contagem_naturezas.index, contagem_naturezas.values, 
                             color=['#ff6b6b', '#4ecdc4', '#45b7d1', '#96ceb4'])
                
                ax.set_title('DISTRIBUIÇÃO GERAL DE CRIMES - PERÍODO 1-31 DIAS', 
                           fontsize=16, fontweight='bold')
                ax.set_xlabel('Tipo de Crime', fontsize=12)
                ax.set_ylabel('Número de Ocorrências', fontsize=12)
                ax.grid(True, alpha=0.3, axis='y')
                
                # Adicionar valores nas barras
                for bar in bars:
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                           f'{int(height)}', ha='center', va='bottom', fontweight='bold')
                
                plt.tight_layout()
                pdf.savefig(fig)
                plt.close()
                
                # Página 3: Gráfico de linhas - Evolução diária
                fig, ax = plt.subplots(figsize=(14, 8))
                
                # Agrupar por dia e natureza
                dados_dia = df.groupby(['Dia Fato', 'Natureza_Padronizada']).size().unstack(fill_value=0)
                
                # Completar todos os dias de 1 a 31
                todos_dias = range(1, 32)
                dados_dia = dados_dia.reindex(todos_dias, fill_value=0)
                
                # Plotar gráfico de linhas
                for natureza in dados_dia.columns:
                    ax.plot(dados_dia.index, dados_dia[natureza], 
                           marker='o', linewidth=2.5, markersize=6, label=natureza)
                
                ax.set_title('EVOLUÇÃO DIÁRIA DOS CRIMES - DIAS 1 A 31', 
                           fontsize=16, fontweight='bold')
                ax.set_xlabel('Dia do Mês', fontsize=12)
                ax.set_ylabel('Número de Ocorrências', fontsize=12)
                ax.legend(title='Tipo de Crime', fontsize=10)
                ax.grid(True, alpha=0.3)
                ax.set_xlim(1, 31)
                ax.set_xticks(range(1, 32, 2))  # Mostrar dias ímpares para melhor visualização
                
                plt.tight_layout()
                pdf.savefig(fig)
                plt.close()
                
                # Página 4: Gráfico de pizza - Distribuição percentual
                fig, ax = plt.subplots(figsize=(10, 8))
                
                contagem_naturezas = df['Natureza_Padronizada'].value_counts()
                cores = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99']
                
                wedges, texts, autotexts = ax.pie(contagem_naturezas.values, 
                                                 labels=contagem_naturezas.index, 
                                                 autopct='%1.1f%%',
                                                 startangle=90,
                                                 colors=cores)
                
                # Melhorar aparência dos textos
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
                
                ax.set_title('DISTRIBUIÇÃO PERCENTUAL DOS CRIMES', 
                           fontsize=16, fontweight='bold')
                
                plt.tight_layout()
                pdf.savefig(fig)
                plt.close()
                
                # Página 5: Gráfico de área - Acumulado diário
                fig, ax = plt.subplots(figsize=(14, 8))
                
                # Calcular acumulado por dia
                dados_dia_acumulado = dados_dia.cumsum()
                
                # Plotar gráfico de área
                ax.stackplot(dados_dia_acumulado.index, 
                           dados_dia_acumulado.T, 
                           labels=dados_dia_acumulado.columns,
                           alpha=0.7)
                
                ax.set_title('ACUMULADO DIÁRIO DE OCORRÊNCIAS', 
                           fontsize=16, fontweight='bold')
                ax.set_xlabel('Dia do Mês', fontsize=12)
                ax.set_ylabel('Ocorrências Acumuladas', fontsize=12)
                ax.legend(title='Tipo de Crime', loc='upper left')
                ax.grid(True, alpha=0.3)
                ax.set_xlim(1, 31)
                ax.set_xticks(range(1, 32, 2))
                
                plt.tight_layout()
                pdf.savefig(fig)
                plt.close()
                
                # Página 6: Heatmap por dia e natureza
                fig, ax = plt.subplots(figsize=(12, 8))
                
                # Preparar dados para heatmap
                heatmap_data = dados_dia.T
                
                sns.heatmap(heatmap_data, annot=True, fmt='d', cmap='YlOrRd', 
                           linewidths=0.5, ax=ax, cbar_kws={'label': 'Número de Ocorrências'})
                
                ax.set_title('HEATMAP - DISTRIBUIÇÃO DIÁRIA POR TIPO DE CRIME', 
                           fontsize=16, fontweight='bold')
                ax.set_xlabel('Dia do Mês')
                ax.set_ylabel('Tipo de Crime')
                
                plt.tight_layout()
                pdf.savefig(fig)
                plt.close()
                
                # Página 7: Análise estatística detalhada
                fig = plt.figure(figsize=(11, 8.5))
                plt.axis('off')
                
                # Cálculos estatísticos
                total_geral = len(df)
                contagem_naturezas = df['Natureza_Padronizada'].value_counts()
                dia_mais_ocorrencias = df['Dia Fato'].value_counts().idxmax()
                max_ocorrencias_dia = df['Dia Fato'].value_counts().max()
                
                texto_analise = "RELATÓRIO ESTATÍSTICO - ANÁLISE COMPLETA\n\n"
                texto_analise += "="*50 + "\n"
                texto_analise += f"PERÍODO ANALISADO: Dias {tempo_periodo} do mês {mes_analisado}\n"
                texto_analise += f"TOTAL DE OCORRÊNCIAS: {total_geral:,}\n\n"
                
                texto_analise += "DISTRIBUIÇÃO POR TIPO DE CRIME:\n"
                texto_analise += "-" * 30 + "\n"
                for natureza, total in contagem_naturezas.items():
                    percentual = (total / total_geral) * 100
                    texto_analise += f"• {natureza}: {total} ocorrências ({percentual:.1f}%)\n"
                
                texto_analise += f"\nDIA COM MAIOR NÚMERO DE OCORRÊNCIAS:\n"
                texto_analise += "-" * 30 + "\n"
                texto_analise += f"• Dia {dia_mais_ocorrencias}: {max_ocorrencias_dia} ocorrências\n"
                
                # Médias e estatísticas
                texto_analise += f"\nESTATÍSTICAS DIÁRIAS:\n"
                texto_analise += "-" * 30 + "\n"
                texto_analise += f"• Média diária: {total_geral/31:.1f} ocorrências/dia\n"
                
                # Dias sem ocorrências
                dias_com_ocorrencias = df['Dia Fato'].unique()
                dias_sem_ocorrencias = [dia for dia in range(1, 32) if dia not in dias_com_ocorrencias]
                texto_analise += f"• Dias sem ocorrências: {len(dias_sem_ocorrencias)}\n"
                
                if dias_sem_ocorrencias:
                    texto_analise += f"  {dias_sem_ocorrencias}\n"
                
                plt.text(0.1, 0.95, texto_analise, fontsize=11, 
                        fontfamily='monospace', transform=plt.gca().transAxes,
                        verticalalignment='top', linespacing=1.3)
                
                pdf.savefig(fig)
                plt.close()
            
            else:
                # Página alternativa se não houver dados
                fig = plt.figure(figsize=(11, 8.5))
                plt.axis('off')
                plt.text(0.5, 0.5, 'Nenhum dado disponível para análise', 
                        ha='center', va='center', fontsize=16, 
                        transform=plt.gca().transAxes)
                pdf.savefig(fig)
                plt.close()
        
        print(f"✓ Relatório PDF gerado: {nome_arquivo}")
        return True
        
    except Exception as e:
        print(f"✗ Erro ao criar PDF: {e}")
        return False

def main():
    """
    Função principal do programa
    """
    print("=" * 60)
    print("ANÁLISE COMPLETA DE DADOS CRIMINAIS (1-31 DIAS)")
    print("=" * 60)
    
    # Verificar se arquivo existe
    if not os.path.exists("dados.xlsx"):
        print("ERRO: Arquivo 'dados.xlsx' não encontrado na pasta atual")
        print("Por favor, coloque o arquivo na mesma pasta do script")
        return
    
    # Ler dados
    df = ler_dados_excel()
    
    if df is not None and not df.empty:
        print(f"\n✓ Dados carregados com sucesso!")
        print(f"✓ Período: {df['Data_Fato'].min().strftime('%d/%m/%Y')} a {df['Data_Fato'].max().strftime('%d/%m/%Y')}")
        print(f"✓ Total de ocorrências: {len(df):,}")
        print(f"✓ Tipos de crime encontrados: {', '.join(df['Natureza_Padronizada'].unique())}")
        print(f"✓ Período analisado: Dias 1 a 31")
        
        # Gerar relatório PDF
        sucesso = criar_relatorio_pdf(df)
        
        if sucesso:
            print("\n" + "=" * 60)
            print("PROCESSO CONCLUÍDO COM SUCESSO!")
            print("=" * 60)
            print("Relatório gerado: 'relatorio_crimes_completo.pdf'")
            print("\nConteúdo do relatório:")
            print("• Capa com informações gerais")
            print("• Gráfico de barras - Distribuição geral")
            print("• Gráfico de linhas - Evolução diária")
            print("• Gráfico de pizza - Distribuição percentual")
            print("• Gráfico de área - Acumulado diário")
            print("• Heatmap - Distribuição detalhada")
            print("• Análise estatística completa")
        else:
            print("\nErro ao gerar relatório PDF")
    
    else:
        print("\nNão foi possível processar os dados. Verifique o arquivo Excel.")

if __name__ == "__main__":
    # Configurar para evitar problemas de display
    import os
    os.environ['MPLBACKEND'] = 'Agg'
    
    main()