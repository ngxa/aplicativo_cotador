import streamlit as st
import pandas as pd
import io
import datetime as dt 
from PIL import Image
from openpyxl import Workbook
from io import BytesIO


# Store the initial value of widgets in session state
if "visibility" not in st.session_state:
    st.session_state.visibility = "visible"
    st.session_state.disabled = False


# Página Inicial
def home_page():
    # Carregar e exibir uma imagem local
    image = Image.open('Ativo-1Logo-1.png')
    st.image(image)
    st.write("Cotador de precipitação de nível municipal")

# Página 2
def page_2():
    st.title("Cotador")
    file = 'MUNICIPAL_AON - calculator.csv'
    df = pd.read_csv(file, sep=';')
    
    dia_atual = dt.datetime.today()

    df[['inicio1', 'fim1']] = df['Risk_period 1'].str.split(' / ', n=1, expand=True)
    df[['inicio2', 'fim2']] = df['Risk_period 2'].str.split(' / ', n=1, expand=True)
    ano_atual = 2023

    df['inicio1'] = pd.to_datetime(df['inicio1'] + f"-{ano_atual}", format="%m-%d-%Y")
    df['fim1'] = pd.to_datetime(df['fim1'] + f"-{ano_atual}", format="%m-%d-%Y")
    df['Risk_period 1'] = df['inicio1'].dt.strftime('%d/%m') + ' a ' + df['fim1'].dt.strftime('%d/%m')
    df['inicio2'] = pd.to_datetime(df['inicio2'] + f"-{ano_atual}", format="%m-%d-%Y")
    df['fim2'] = pd.to_datetime(df['fim2'] + f"-{ano_atual}", format="%m-%d-%Y")
    df['Risk_period 2'] = df['inicio2'].dt.strftime('%d/%m') + ' a ' + df['fim2'].dt.strftime('%d/%m')


    cidades = df['Município'].unique().tolist()

    # Selecionar até 10 cidades usando o st.multiselect
    cidades_selecionadas = st.multiselect("Selecione as cidades:", cidades, [], key="cidades", placeholder = 'Escolha os Municípios')
    cont= 1
    periodo1 = []
    periodo2 = []
    area = []
    valor = []
    # Verificar se alguma cidade foi selecionada
    if cidades_selecionadas:
        for cidade in cidades_selecionadas:
            df_filtrado = df[df['Município'] == cidade]
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write('Escolha os Períodos:')

                p1 = st.checkbox(df_filtrado['Risk_period 1'].tolist()[0], key="checkbox1" + str(cont))
                p2 = st.checkbox(df_filtrado['Risk_period 2'].tolist()[0], key="checkbox2" + str(cont))
                cont += 1
                periodo1.append(p1)
                periodo2.append(p2)
            with col2:
                a = st.text_input(f'Area para {cidade} em Hectares (ha):')
                area.append(a)
            with col3:
                p = st.text_input(f'Preço para {cidade} em Reais (R$):')
                valor.append(p)


    if st.button("Mostrar Cotações"):
        area = [float(a.replace(",", ".")) for a in area]
        valor = [float(v.replace(",", ".")) for v in valor]


        CIDADE = []
        TAXA = []
        PERIODO = []
        AREA = []
        VALOR = []
        PRECO = []
        MAX_IND=[]
        MAX_STRIKE = []
        PAID_MM = []
        GATILHO = []
        SAIDA = []
        for c, p1, p2, a, v in zip(cidades_selecionadas, periodo1, periodo2, area, valor):
            result = df[df['Município'] == c]
             
            taxa = result.iloc[:, 4].values[0]
            lmi = a * v
            preco = taxa * lmi
            if p1 == True:
                saida1 = result['Exit 1'].values[0]
                strike1 = result.iloc[:, 8].values[0]
                ind_max1 = lmi/2
                paid_mm1 = ind_max1/strike1
                CIDADE.append(c)
                TAXA.append(taxa * 100)
                PERIODO.append(result['Risk_period 1'].values[0])
                AREA.append(a)
                VALOR.append(v)
                PRECO.append(preco)
                MAX_IND.append(lmi)
                MAX_STRIKE .append(ind_max1)
                PAID_MM.append(round(paid_mm1, 2))
                GATILHO.append(round(strike1))
                SAIDA.append(saida1)
                
                
            if p2 == True:
                saida2 = result['Exit 2'].values[0]
                strike2 = result.iloc[:, 12].values[0]
                ind_max2 = lmi/2
                paid_mm2 = ind_max2/strike2
                CIDADE.append(c)
                TAXA.append(taxa * 100)
                PERIODO.append(result['Risk_period 2'].values[0])
                AREA.append(a)
                VALOR.append(v)
                PRECO.append(preco)
                MAX_IND.append(lmi)
                MAX_STRIKE .append(ind_max2)
                PAID_MM.append(round(paid_mm2, 2))
                GATILHO.append(round(strike2))
                SAIDA.append(saida2)


        dic = {'Município': CIDADE,
                'Período': PERIODO,
                'Taxa Final (%)': TAXA,
                'Gatilho (mm)': GATILHO,
                'Saída (mm)': SAIDA,
                'R$ por hectare': VALOR,
                'Área (ha)': AREA,
                'LMI': MAX_IND,
                'Tick (R$/mm)': PAID_MM,
                }
           
        resultado = pd.DataFrame(dic)
        st.dataframe(resultado)

        
        # Função para salvar o DataFrame em um arquivo Excel
        def save_dataframe_to_excel(df, file_name):
            workbook = Workbook()
            sheet = workbook.active

            # Adicione os dados do DataFrame ao arquivo Excel
            for row in df.iterrows():
                sheet.append(row[0].tolist())

            # Salve o arquivo Excel em memória
            excel_buffer = BytesIO()
            workbook.save(excel_buffer)
            excel_buffer.seek(0)

            # Crie um link de download para o arquivo Excel
            st.download_button(
                "Download Excel",
                data=excel_buffer,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=file_name
            )

        # Use a função para salvar o DataFrame
        save_dataframe_to_excel(resultado, "cotador_" + dia_atual.strftime('%Y%m%d')+".xlsx")

st.set_page_config(page_title = "Cotador Municipal")
# Configurar o estado da sessão
if 'page' not in st.session_state:
    st.session_state.page = "Página Inicial"

# Criar um menu de navegação
menu = ["Página Inicial", "Kovr Cotador"]
choice = st.sidebar.selectbox("Navegação", menu)

# Roteamento com base na escolha do usuário
if choice == "Página Inicial":
    home_page()
elif choice == "Kovr Cotador":
    page_2()






