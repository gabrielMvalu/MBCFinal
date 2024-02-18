# mbc_docs.py
# mbc_docs.py
from openai import OpenAI
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.api import ExponentialSmoothing


st.set_page_config(layout="wide")

st.header(':blue[Pagina Principală]', divider='rainbow')
st.write(':violet[Bine ați venit la aplicația pentru completarea - Planului de Afaceri! -]')


col1, col2 = st.columns(2)

with col1:
    
    # Definirea datelor de vânzări
    sales_mireasma = [0, 0, 0, 0, 0, 120, 150, 100, 230, 240, 270, 265]
    sales_proaspat = [80, 100, 120, 140, 160, 180, 200, 220, 200, 180, 160, 140]
    sales_parfumat = [50, 70, 90, 110, 130, 120, 110, 100, 90, 80, 70, 60]
    sales_zero = [20, 100, 60, 40, 120, 100, 110, 60, 40, 80, 10, 70]
    
    # Setarea titlului și a layout-ului aplicației Streamlit
    st.title('Analiza și Predicția Vânzărilor pentru Produse')
    st.sidebar.header('Setări')
    
    # Selecția produsului pentru analiză
    produs_selectat = st.sidebar.selectbox('Alege un produs:', ('Mireasmă', 'Proaspăt', 'Parfumat', 'Zero'))
    
    # Maparea selecției la setul de date corespunzător
    sales_data = {
        'Mireasmă': sales_mireasma,
        'Proaspăt': sales_proaspat,
        'Parfumat': sales_parfumat,
        'Zero': sales_zero
    }
    
    # Preluarea datelor pentru produsul selectat
    sales_selected = sales_data[produs_selectat]
    
    # Crearea modelului de netezire exponențială și predicția vânzărilor viitoare
    model = ExponentialSmoothing(sales_selected, trend='add', seasonal='add', seasonal_periods=6).fit()
    pred_sales = model.forecast(3)
    
    # Crearea unei figuri pentru grafic
    fig, ax = plt.subplots()
    
    # Adăugarea datelor istorice și a predicțiilor pe grafic
    ax.plot(sales_selected, label='Vânzări Istorice', marker='o')
    ax.plot(np.arange(len(sales_selected), len(sales_selected) + 3), pred_sales, label='Predicții Vânzări', marker='o', linestyle='--')
    
    # Setarea titlului și a legendei
    ax.set_title(f'Analiza și Predicția Vânzărilor pentru {produs_selectat}')
    ax.set_xlabel('Perioada')
    ax.set_ylabel('Vânzări')
    ax.legend()
    
    # Afișarea graficului în Streamlit
    st.pyplot(fig)



with col2:

    with st.expander(" ℹ️ Mesaj Informativ ℹ️  "):
        st.write("""
            Vă informăm că aceast bot se află într-o fază incipientă de dezvoltare. 
            În acest moment, funcționalitatea este limitată la furnizarea de răspunsuri generale.
        """)
    
    with st.sidebar:
        openai_api_key = st.text_input("OpenAI API Key", key="chatbot_api_key", type="password")
        
    if not openai_api_key:
        st.info("Vă rugăm să introduceți cheia API OpenAI în bara laterală.")
    else:
        # Inițializarea clientului OpenAI cu cheia API introdusă
        client = OpenAI(api_key=openai_api_key)
    
        # Inițializarea stării sesiunii pentru model și mesaje
        if "openai_model" not in st.session_state:
            st.session_state["openai_model"] = "gpt-4-1106-preview"
    
        if "messages" not in st.session_state:
            st.session_state.messages = []
    
        # Afișarea mesajelor anterioare
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
    
        # Input pentru mesaj nou de la utilizator
        if prompt := st.chat_input("Adaugati mesajul aici."):
            st.session_state.messages.append({"role": "user", "content": f"{prompt}" })
            with st.chat_message("user"):
                st.markdown(prompt)
    
            # Generarea răspunsului asistentului și afișarea acestuia
            with st.chat_message("assistant"):
                stream = client.chat.completions.create(
                    model=st.session_state["openai_model"],
                    messages=[
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages
                    ],
                    stream=True,
                )
                response = st.write_stream(stream)
            st.session_state.messages.append({"role": "assistant", "content": response})

