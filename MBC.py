# mbc_docs.py
from openai import OpenAI
import streamlit as st

st.set_page_config(layout="wide")



with st.sidebar:
    openai_api_key = st.text_input("OpenAI API Key", key="chatbot_api_key", type="password")
    

st.header(':blue[Pagina Principală]', divider='rainbow')
st.write(':violet[Bine ați venit la aplicația pentru completarea - Planului de Afaceri! -]')

# Setarea modelului OpenAI
if "openai_model" not in st.session_state:
    st.session_state["openai_model"] = "gpt-3.5-turbo"

# Inițializarea istoricului de mesaje
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "Cu ce te pot ajuta?"}]

# Afișarea mesajelor din istoric
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Acceptarea inputului utilizatorului
if prompt := st.chat_input("Ce doriți să întrebați?"):
    # Adăugarea mesajului utilizatorului în istoric
    st.session_state.messages.append({"role": "user", "content": prompt})
    
    # Afișarea mesajului utilizatorului
    with st.chat_message("user"):
        st.markdown(prompt)

    # Obținerea răspunsurilor stream-uite de la OpenAI și afișarea lor
    with st.chat_message("assistant"):
        stream = client.chat.completions.create(
            model=st.session_state["openai_model"],
            messages=[
                {"role": m["role"], "content": m["content"]}
                for m in st.session_state.messages
            ],
            stream=True
        )
        
        # Afișarea fiecărui fragment din stream-ul de răspunsuri
        for response in stream:
            st.markdown(response.choices[0].message["content"])
            st.session_state.messages.append({"role": "assistant", "content": response.choices[0].message["content"]})
