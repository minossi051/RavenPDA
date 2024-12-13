#programmed tasks 1.2.1

#tarefas para este expeiente (13/12):
#incluir relação de modelos de  UPS
#incluir relação de tempo de uso de ups (antigos)
#incluir infográfico relação de tempo de ausência de UPS nos PoPs
#incluir relação de autonomia de PoPs (antigos e novos)

#running issues -
#nome 'S' na planilha dos splits precisa ser trocado para STATUS, assim como aqui no código




base_dados = r"C:\Users\vicenzo-minossi\OneDrive - Governo do Estado do Rio Grande do Sul\POPS\procergs-diop-dif-pir\DIF-PIR Execução POPS.xlsx"
import pandas as pd
import folium
import streamlit as st
from folium.plugins import MarkerCluster
from streamlit_folium import folium_static
from PIL import Image

st.title('POPS - Dados de serviço')
# Dashboard destinado à apresentação no dia 16/12
# Nome da aba específica a ser carregada

aba_splits = 'Instalações Splits NOV.DEZ2024'
aba_coordenadas = 'Panorama POPS RS'

# Carrega os dados da aba específica
try:
    df_splits = pd.read_excel(base_dados, sheet_name=aba_splits, engine='openpyxl')
except Exception as e:
    st.error(f"Ocorreu um erro ao carregar os dados da aba '{aba_splits}': {e}")
    st.stop()

# Progresso de instalação dos PoPs
#mostrar apenas as colunas que contém nenhum valor vazio
#mostrar somente os splits instalados
#mostra PoP + Situação + Data da instalação
if 'Situação' in df_splits.columns:
    try:
        # Filtra os splits com situação 'OK' ou 'Verificar dreno'
        coluna_situacao = 'Situação' 
        valores_objetivo = ['OK','Verificar dreno']

        df_splits_filtrado = df_splits[df_splits['Situação'].isin(['OK', 'Verificar dreno'])]
        df_splits_filtrado = df_splits[df_splits[coluna_situacao].isin(valores_objetivo)]
        total_splits_filtrado = len(df_splits_filtrado)

        total_splits = len(df_splits)

        # Calcula a porcentagem de splits instalados
        porcentagem_splits = (total_splits_filtrado / total_splits) * 100    if total_splits > 0 else 0

        # Exibe as métricas
        st.subheader("Resumo de Splits Instalados")
        st.metric(f"Total de splits instalados", f"{total_splits_filtrado} de {total_splits} ({porcentagem_splits:.2f}%)")

        # Exibe a barra de progresso
        st.progress(porcentagem_splits / 100)  # A barra de progresso é uma porcentagem (de 0 a 1)

    except Exception as e:
        st.error(f"Erro ao filtrar e contar splits instalados: {e}")

    # Exibe o dataframe com o progresso de instalação
    colunas_progresso = ['POP', 'Situação','Data da instalação','Potência (BTUs)']
    df_progresso = df_splits_filtrado[colunas_progresso]
    df_progresso = df_splits_filtrado.dropna(axis=1, how='any')
    st.dataframe(df_progresso)

else:

    st.warning("A coluna 'Situação' não está presente nos dados de splits.")


#datas das conclusões das instalaões
coluna_data = ['POP', 'Data da instalação']

# Verifica se a coluna 'Supervisão' está presente
if 'Supervisão' in df_splits.columns:

    # Filtra os dados onde 'POP' e 'Data da instalação' não são nulos
    df_data = df_splits[df_splits[coluna_data].notna().all(axis=1)]  # Apenas as linhas em que todas as colunas da lista têm dados

    # Adiciona a coluna 'Supervisão' se ela existir
    coluna_data.append('Supervisão')

    # Exibe o subtítulo para o dashboard
    st.subheader("Datas de Conclusão das Instalações")

    # Remove as colunas com valores nulos
    df_data = df_data[coluna_data].dropna(axis=1, how='all')

    # Exibe o dataframe sem as colunas nulas
    st.dataframe(df_data)

else:
    st.warning(f"A coluna 'Supervisão' não está presente nos dados.")

    st.subheader("Próximas instalações agendadas")


#relção de POP - Situação do split - origem
st.subheader('Situação dos Splits')

# Colunas de interesse
colunas_status = ['POP', 'Motivo da instalação', 'Origem','Situação']

# Verifica se todas as colunas de interesse estão presentes no dataframe
colunas_existentes = [col for col in colunas_status if col in df_splits.columns]

if colunas_existentes:
    # Filtra as linhas onde as colunas selecionadas não possuem valores vazios
    df_status_filtrado = df_splits[colunas_existentes].dropna(how='any')

    # Exibe o dataframe filtrado no dashboard
    st.dataframe(df_status_filtrado)
else:
    st.warning(f"As colunas {colunas_status} não estão presentes no dataframe.")


# Cronograma de instalações
st.subheader("Próximas instalações agendadas")
if 'Situação' in df_splits.columns:
    # Filtra os POPs com 'Instalação agendada' na coluna 'Situação'
    df_agendados = df_splits[df_splits['Situação'].str.contains('Instalação agendada', na=False)]
    
    # Exibe a tabela filtrada com a data extraída
    colunas_agendadas = ['POP', 'Situação']
    st.dataframe(df_agendados[colunas_agendadas])

#mapa foi retirado

caminho_imagem = r"C:\\Users\\vicenzo-minossi\\Desktop\\16.12\\horas_de_viagem.png"

st.subheader('Itinerário planejado')
# Carrega e exibe a imagem
try:
    imagem = Image.open(caminho_imagem)
    st.image(imagem, caption="Itinerário planejado pela chefia", use_container_width=True)
except Exception as e:
    st.error(f"Erro ao carregar a imagem: {e}")




####################################################################################################################
# Segunda parte do dashboard: Operação Lítio - Dados sobre os UPS dos PoPs: Status NOBREAK, Modelo UPS, tempo de USO, incidentes.



aba_ups = 'Levantamento UPS'
#carregando dados
try:
    df_ups = pd.read_excel(base_dados, sheet_name=aba_ups, engine='openpyxl')
except Exception as e:
    st.error(f'ERRO:{e}')
    st.stop


#tratamento de dados
# classificar PoPs QUE não possuem UPS
if 'S' in df_ups.columns:
    try:
        coluna_status = 'S'
        linhas_nobreak = ['SEM NOBREAK']
        linhas_sem_dados = ['Sem dados']

        df_ups_filtrado = df_ups[df_ups['S'].isin(['SEM NOBREAK'])]
        df_ups_filtrado = df_ups[df_ups[coluna_status].isin(linhas_nobreak)]
        total_pops_semups = len(df_ups_filtrado)

        colunas_status = ['POP', 'S']
        df_ups = df_ups_filtrado[colunas_status]
        df_ups = df_ups_filtrado.dropna(axis=1, how='any')
        st.dataframe(df_ups)
        st.metric("Total de PoPs Sem Nobreak", total_pops_semups)
   
    except Exception as e:
        st.error(f'Erro: {e}')

else:

    st.warning("A coluna 'S' não está presente nos dados dos UPS.")


# esse código ta me testando a paciência 
#refiz toda a lógica pra funcionar
try:
 df_ups_sem_dados = pd.read_excel(base_dados, sheet_name=aba_ups,engine='openpyxl')
except Exception as e:
    st.error(f'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA {e}')
if 'S' in df_ups_sem_dados.columns:
    try:
        # Exibir valores únicos na coluna 'S' para verificar possíveis problemas
        # Normalizar os valores da coluna 'S' (remover espços ea tornar maiúsculos)
        df_ups_sem_dados['S'] = df_ups_sem_dados['S'].str.strip().str.upper() 
        linhas_sem_dados = ['SEM DADOS']  #Certifique-se de que os valores correspondem ao esperado

        # Diagnóstico: Testar se algum valor esperado existe
        if any(valor in df_ups_sem_dados['S'].values for valor in linhas_sem_dados):
            # Filtrar os PoPs com "Sem dados"
            df_ups_semdados = df_ups_sem_dados[df_ups_sem_dados['S'].isin(linhas_sem_dados)]
            total_pops_semdados = len(df_ups_semdados)

            # Exibir os resultados filtrados
            colunas_semdados = ['POP', 'S']
            st.subheader("PoPs com 'Sem dados'")
            st.dataframe(df_ups_semdados[colunas_semdados])
            st.metric("Total de PoPs Sem Dados", total_pops_semdados)
        else:
            st.warning("Nenhum valor 'Sem dados' foi encontrado na coluna 'S'.")

    except Exception as e:
        st.error(f'Erro: {e}')
else:
    st.warning('A coluna "S" não está presente nos dados dos UPS.')







