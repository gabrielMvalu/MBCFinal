# mbc_docs.py
from openai import OpenAI
import streamlit as st

# Setarea configurației paginii
st.set_page_config(layout="wide")

# Inițializarea API Key în sidebar
with st.sidebar:
    openai_api_key = st.text_input("OpenAI API Key", key="chatbot_api_key", type="password")

# Verificarea dacă OpenAI API Key este furnizată
if not openai_api_key:
    st.error("Vă rugăm să furnizați OpenAI API Key în sidebar pentru a continua.")
    st.stop()

# Inițializarea clientului OpenAI
client = OpenAI(api_key=openai_api_key)

# Setarea titlului și a mesajului de întâmpinare
st.header('Pagina Principală')
st.write('Bine ați venit la aplicația pentru completarea Planului de Afaceri!')

# Inițializarea modelului și a istoricului de mesaje dacă nu există deja
if "openai_model" not in st.session_state:
    st.session_state["openai_model"] = "gpt-3.5-turbo"

if "messages" not in st.session_state:
    st.session_state["messages"] = [{"role": "assistant", "content": "Cu ce te pot ajuta?"}]

# Afișarea istoricului de mesaje
for message in st.session_state["messages"]:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Acceptarea și procesarea inputului de la utilizator
if prompt := st.chat_input("Ce doriți să întrebați?"):
    st.session_state["messages"].append({"role": "user", "content": prompt})

    response = client.chat.completions.create(
        model=st.session_state["openai_model"],
        messages=st.session_state["messages"]
    )

    # Adăugarea răspunsului asistentului la istoricul de mesaje și afișarea acestuia
    assistant_message = response.choices[0].message["content"]
    st.session_state["messages"].append({"role": "assistant", "content": assistant_message})
    with st.chat_message("assistant"):
        st.markdown(assistant_message)
