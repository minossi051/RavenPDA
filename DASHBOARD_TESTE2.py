#funcionalidades: 
#gráfico de pizza da percentagem de PoPs mapeados X
#infográfico em barras:
                    #Número de incidentes por POP durante 2024 X
                    #Data exata do incidente ao clicar em algum deles 
#planilha de serviços executados desde 23/09/2024
#Operação Verão (Relação dos splits)
#Operação Lítio (Relação dos UPS)


import streamlit as st
import pandas as pd
import plotly.express as px
# Caminho do arquivo Excel
base_dados = r"C:\Users\vicenzo-minossi\OneDrive - Governo do Estado do Rio Grande do Sul\POPS\procergs-diop-dif-pir\DIF-PIR Execução POPS.xlsx"

st.title('POPS - Status de Serviços')

# Nome da aba específica a ser carregada
aba_alvo = 'Panorama POPS RS'

# Carregar os dados da aba especificada
try:
    df = pd.read_excel(base_dados, sheet_name=aba_alvo, engine='openpyxl')  # Especifica a aba e usa openpyxl
except Exception as e:
    st.error(f"Ocorreu um erro ao carregar os dados da aba '{aba_alvo}': {e}")
    st.stop()

# Verificar se os dados foram carregados corretamente
st.write(f"Informações '{aba_alvo}'")
st.dataframe(df)

# Verificar e corrigir tipos de dados
if 'Última vistoria' in df.columns:
    try:
        # Converter a coluna de datas, tratando valores inválidos
        df['Última vistoria'] = pd.to_datetime(df['Última vistoria'], errors='coerce')
        st.write("Coluna 'Última vistoria' convertida para formato de data.")
    except Exception as e:
        st.error(f"Erro ao converter 'Última vistoria' para data: {e}")

# Tratar células vazias
df.fillna("Não informado", inplace=True)

# Filtrar e contar serviços em andamento
coluna_status = 'Serviços em andamento (Inserir SOL, se houver)'
coluna_nome = 'Nome POP'

if coluna_status in df.columns and coluna_nome in df.columns:
    try:
        # Identificar serviços em andamento
        servico_andamento = df[
            (df[coluna_status] != "Não informado") & 
            (~df[coluna_status].str.contains('Sem serviços em andamento', case=False, na=False))
        ]
        
        # Selecionar apenas as colunas desejadas
        resultado_filtrado = servico_andamento[[coluna_nome, coluna_status]]
        total_andamento = len(resultado_filtrado)
        
        st.subheader(f"Total de POPs com serviços em andamento: {total_andamento}")
        st.dataframe(resultado_filtrado)
    except Exception as e:
        st.error(f"Erro ao filtrar os serviços em andamento: {e}")
else:
    st.error(f"Uma ou ambas as colunas '{coluna_status}' e '{coluna_nome}' não foram encontradas na aba '{aba_alvo}'.")

coluna_rack = 'Rack mapeado?'

if coluna_rack in df.columns:
    try:
        # Contar os valores 'Sim' e 'Não'
        contagem = df[coluna_rack].value_counts()

        # Garantir que valores ausentes ('Sim' ou 'Não') não causem erro
        sim = contagem.get('Sim', 0)
        nao = contagem.get('Não', 0)

        total = sim + nao
        if total > 0:
            porcentagem_sim = (sim / total) * 100
            porcentagem_nao = (nao / total) * 100

            st.subheader("Porcentagem de Racks Mapeados")
            st.write(f"Total de POPs: {total}")
            st.write(f"Racks Mapeados: {porcentagem_sim:.2f}% ({sim})")
            st.write(f"Racks Não Mapeados: {porcentagem_nao:.2f}% ({nao})")

            # Gráfico de pizza
            st.write("Distribuição de Racks Mapeados:")
            st.pyplot(pd.DataFrame(
                {'Status': ['Sim', 'Não'], 'Quantidade': [sim, nao]}
            ).set_index('Status').plot.pie(
                y='Quantidade', autopct='%1.1f%%', figsize=(6, 6), legend=False
            ).figure)
        else:
            st.warning("Não há informações suficientes para calcular os racks mapeados.")
    except Exception as e:
        st.error(f"Erro ao calcular racks mapeados: {e}")
else:
    st.error(f"A coluna '{coluna_rack}' não foi encontrada na aba '{aba_alvo}'.")



# Nome da aba específica a ser carregada
aba_incidentes = 'Lista de incidentes'

# Carregar os dados da aba especificada
try:
    df_incidentes = pd.read_excel(base_dados, sheet_name=aba_incidentes, engine='openpyxl')
except Exception as e:
    st.error(f"Ocorreu um erro ao carregar os dados da aba '{aba_incidentes}': {e}")
    st.stop()

# Verificar se os dados foram carregados corretamente
st.write(f"Visualizando os dados da aba '{aba_incidentes}'")
st.dataframe(df_incidentes)

# Garantir que a coluna 'Incidente' seja tratada como string
df_incidentes['Incidente'] = df_incidentes['Incidente'].astype(str)
Resumo = df_incidentes['Resumo'].astype(str)

# Garantir que a coluna 'Data de abertura' esteja no formato datetime
if 'Data de abertura' in df_incidentes.columns:
    try:
        df_incidentes['Data de abertura'] = pd.to_datetime(df_incidentes['Data de abertura'], errors='coerce')
        st.write("Coluna 'Data de abertura' convertida para formato de data.")
    except Exception as e:
        st.error(f"Erro ao converter 'Data de abertura' para data: {e}")

# Filtrar os incidentes entre 2022 e 2024
df_incidentes_filtrados = df_incidentes[(df_incidentes['Data de abertura'] >= '2024-01-01') &
                                         (df_incidentes['Data de abertura'] <= '2024-12-31')]
                                           

# Verificar se há a coluna 'POP' e 'Serviço afetado' no dataframe
if 'Serviço afetado' in df_incidentes.columns:
    # Contar os incidentes por POP
    incidentes_por_pop = df_incidentes_filtrados.groupby('Serviço afetado').size().reset_index(name='Quantidade de Incidentes')
    
    # Exibir gráfico de barras
    fig = px.bar(incidentes_por_pop, x='Serviço afetado', y='Quantidade de Incidentes',
                 title='Quantidade de Incidentes por POP (Janeiro - Dezembro (2024))', labels={'Serviço afetado': 'POP', 'Quantidade de Incidentes': 'Número de Incidentes'})
    
    st.plotly_chart(fig)
else:
    st.error("A coluna 'Serviço afetado' não foi encontrada nos dados.")