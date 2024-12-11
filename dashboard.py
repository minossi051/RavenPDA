#mudanças dashboard 1.2
#aprimoramento do filtro de dados da planilha principal
#adição do mapa interativo







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
coluna_data = 'Data da instalação'
if coluna_data in df_splits.columns:
    df_data = df_splits[df_splits[coluna_data].notna()]
    
    st.subheader("Datas de Conclusão das Instalações")
    st.dataframe(df_data[['POP', coluna_data]]) 
else:
    st.warning(f"A coluna '{coluna_data}' não está presente nos dados.")


#relção de POP - Situação do split - origem
st.subheader('Situação dos Splits')

# Colunas de interesse
colunas_status = ['POP', 'Motivo da instalação', 'Origem']

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
# Mapa com roteiro das instalações
st.subheader("Mapa Roteiro das Instalações")
roteiro_inst = ['VIAMAO', 'NOVO HAMBURGO', 'TAQUARA', 'ALEGRETE', 'CARAZINHO',
                'ERECHIM', 'PALMEIRA DAS MISSOES', 'SANTA ROSA']

try:
    df_coord = pd.read_excel(
        base_dados,
        sheet_name=aba_coordenadas,
        usecols=['Município', 'Latitude', 'Longitude'],
        engine='openpyxl'
    )

    # Filtrar somente as cidades do roteiro
    df_coord_filtrado = df_coord[df_coord['Município'].str.upper().isin(roteiro_inst)]

    # Converter Latitude e Longitude para valores numéricos
    df_coord_filtrado['Latitude'] = pd.to_numeric(df_coord_filtrado['Latitude'], errors='coerce')
    df_coord_filtrado['Longitude'] = pd.to_numeric(df_coord_filtrado['Longitude'], errors='coerce')

    # Verificar se há coordenadas válidas
    if df_coord_filtrado[['Latitude', 'Longitude']].isnull().any().any():
        st.error("Algumas coordenadas possuem valores inválidos. Verifique os dados.")
    else:
        # Criar o mapa
        mapa = folium.Map(location=[df_coord_filtrado['Latitude'].mean(),
                                    df_coord_filtrado['Longitude'].mean()],
                          zoom_start=6)

        # Adicionar marcadores no mapa
        marker_cluster = MarkerCluster().add_to(mapa)
        for _, row in df_coord_filtrado.iterrows():
            folium.Marker(
                location=[row['Latitude'], row['Longitude']],
                popup=row['Município'],
                icon=folium.Icon(color='blue')
            ).add_to(marker_cluster)

        # Traçar o caminho entre as cidades na ordem específica
        roteiro_ordenado = df_coord_filtrado.set_index('Município').loc[roteiro_inst]
        caminho = [(row['Latitude'], row['Longitude']) for _, row in roteiro_ordenado.iterrows()]
        folium.PolyLine(caminho, color='red', weight=2.5, opacity=1).add_to(mapa)

        # Exibir o mapa no Streamlit
        st.write("Mapa dos pontos de instalação pendentes")
        folium_static(mapa)

except Exception as e:
    st.error(f"Erro ao carregar ou processar as coordenadas: {e}")

caminho_imagem = r"C:\\Users\\vicenzo-minossi\\Desktop\\16.12\\horas_de_viagem.png"

st.subheader('Itinerário planejado')
# Carrega e exibe a imagem
try:
    imagem = Image.open(caminho_imagem)
    st.image(imagem, caption="Itinerário planejado pela chefia", use_container_width=True)
except Exception as e:
    st.error(f"Erro ao carregar a imagem: {e}")
