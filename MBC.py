# mbc_docs.py
# mbc_docs.py
from openai import OpenAI
import streamlit as st

st.set_page_config(layout="wide")

st.header(':blue[Pagina Principală]', divider='rainbow')
st.write(':violet[Bine ați venit la aplicația pentru completarea - Planului de Afaceri! -]')


equipment_list = [
        "autogreder", "autogreder cu sistem 3d", "autogudronator", "automacara",
        "automaturatoare cu apa", "autopompa cu malaxor", "autopompa de beton",
        "buldoexcavator", "buldoexcavator cu roti egale", "buldoexcavator cu roti inegale",
        "buldozer", "cilindru compactor terasament", "ciur mobil rotativ", "compactor de sol",
        "compactor/cilindru vibrator compactor", "concasor cu falci si presortare",
        "excavator pe pneuri", "excavator pe senile", "finisor beton", "finisor de asfalt pe pneuri",
        "finisor de asfalt pe senile", "finisorul de asfalt", "foreza", "freza", "freza de asfalt",
        "greder", "incarcator frontal", "incarcator multifunctional cu brat telescopic",
        "incarcator pe pneuri", "macara mobila cu brat telescopic",
        "manipulator telescopic/incarcator multifunctional rotativ", "masina badijonat",
        "masina de colmatat rosturi", "midiexcavatorele", "minibuldoexcavator", "miniexcavator",
        "miniincarcator", "motostivuitor", "panouri fotovoltaice mobile", "pompa de beton",
        "reciclator pentru beton/asfalt", "repartizator mixturi asfaltice", "sistem de ghidare automata",
        "stabilizator terasament", "statia de betoane", "statie de asfalt", "statie mobila de concasare",
        "statie mobila de sortare cu spalare", "telehandler", "tocatorul pentru resturi vegetale",
        "vibroprese fabricare pavele, boltari, borduri", "macara", "masina de repartizat emulsie (autogudronator)",
        "grupul electrogen", "microexcavator", "miniexcavator"
    ]



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
        st.session_state.messages.append({"role": "user", "content": f"Vreau sa ma asisti si sa mi raspunzi care din elementele astea:{prompt} le poti identifica in: {equipment_list} ***Vei raspunde enumerand doar utilajele pe care le ai identificat!!!! Nimic altceva!!!! ")
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

