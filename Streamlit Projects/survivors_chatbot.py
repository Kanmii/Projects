# London AI Strategy Copilot
# Installation: pip install streamlit langchain-community langchain-core langchain-chroma langchain-ollama chromadb plotly pathlib-ng
# Ollama models: ollama pull gemma3:1b && ollama pull nomic-embed-text
# Run: streamlit run app.py

import streamlit as st
import plotly.express as px
from pathlib import Path
from typing import List
import re

from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_ollama import OllamaEmbeddings
from langchain_chroma import Chroma
from langchain_community.chat_models import ChatOllama
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnablePassthrough
from langchain_community.chat_message_histories import StreamlitChatMessageHistory

# Configuration
PDF_FOLDER = Path(r"C:\Users\TECHTALENT-009\Desktop\Tyy\Python Programming\The survivors group\pdfs")
CHROMA_DIR = Path("data/chroma_1.db")
EMBEDDING_MODEL = "nomic-embed-text"
OLLAMA_MODEL = "gemma3:1b"
TEMPERATURE = 0.1
CHUNK_SIZE = 1000
CHUNK_OVERLAP = 100

def load_pdfs(folder: Path) -> List:
    """Read all PDFs in folder and return list of Document objects."""
    all_docs = []
    for pdf_path in folder.glob("*.pdf"):
        loader = PyPDFLoader(str(pdf_path))
        pages = loader.load()
        
        fallback_title = pdf_path.stem
        for page in pages:
            page.metadata.setdefault("title", fallback_title)
            page.metadata.setdefault("author", "Unknown Author")
        all_docs.extend(pages)
    return all_docs

def chunk_documents(documents: List) -> List:
    """Split documents into smaller chunks for vector search."""
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n", "\n", ".", " ", ""]
    )
    return splitter.split_documents(documents)

@st.cache_resource(show_spinner="🔄 Indexing and embedding PDFs…")
def get_vectorstore():
    """Load existing Chroma DB or build new one."""
    embedding_fn = OllamaEmbeddings(model=EMBEDDING_MODEL)

    if CHROMA_DIR.exists():
        return Chroma(persist_directory=str(CHROMA_DIR), embedding_function=embedding_fn)

    documents = load_pdfs(PDF_FOLDER)
    chunked_docs = chunk_documents(documents)
    return Chroma.from_documents(
        documents=chunked_docs,
        embedding=embedding_fn,
        persist_directory=str(CHROMA_DIR)
    )

# Safety filters
BANNED_KEYWORDS = [
    "kill", "bomb", "die", "hate", "rape", "suicide", "explode", "abuse",
    "stab", "swear", "threat", "attack", "shoot", "violence", "murder"
]

def is_safe_input(user_input: str) -> bool:
    """Check if input contains banned words."""
    return not any(re.search(rf"\b{kw}\b", user_input.lower()) for kw in BANNED_KEYWORDS)

def is_london_question(user_input: str) -> bool:
    """Check if question is about London."""
    return bool(re.search(r"\blondon\b", user_input.lower()))

def build_rag_chain(vectorstore):
    """Create RAG chain: retrieve → prompt → LLM → parse."""
    
    retriever = vectorstore.as_retriever(
        search_type="mmr",
        search_kwargs={"k": 6, "fetch_k": 20, "score_threshold": 0.7}
    )

    instructions = """
You are an intelligent, polite, London-focused assistant.
If the user's question is not about London, politely say:
"I'm sorry, I can only answer questions related to London."

Guidelines:
• Be polite, friendly, and brief
• Use only the provided context, never guess
• Cite PDF title or section when helpful
• Use bullet points when appropriate
• If context is empty, admit you don't know
• No medical/legal/financial advice
• No profanity, hate, or violent content
• For safety concerns, direct to emergency services
"""

    prompt = ChatPromptTemplate.from_template(
        f"""{instructions}

CONTEXT:
{{context}}

QUESTION:
{{question}}

RESPONSE:
""")

    llm = ChatOllama(model=OLLAMA_MODEL, temperature=TEMPERATURE)

    return (
        {"context": retriever, "question": RunnablePassthrough()}
        | prompt
        | llm
        | StrOutputParser()
    )

def main():
    """Main Streamlit application."""
    
    st.set_page_config(page_title="AI Strategy Copilot (London)", layout="wide")
    st.title("🧠 Your London AI Copilot")
    st.markdown("Ask anything about the London PDF documents.")

    # Debug chat history
    with st.expander("📝 Chat History (Debug)"):
        st.json(st.session_state.get("langchain_messages", {}))

    # Initialize chat history
    msgs = StreamlitChatMessageHistory(key="langchain_messages")
    if not msgs.messages:
        msgs.add_ai_message("Hi! I know only what's inside the London PDFs. How can I help?")

    # Build vector store and RAG chain
    vectorstore = get_vectorstore()
    rag_chain = build_rag_chain(vectorstore)

    # Display chat history
    for msg in msgs.messages:
        with st.chat_message(msg.type):
            st.write(msg.content)

    # Chat input
    if question := st.chat_input("Enter your London-related question here…"):
        st.chat_message("user").write(question)
        msgs.add_user_message(question)

        # Safety checks
        if not is_safe_input(question):
            warning = "⚠️ Please keep the conversation respectful and safe."
            st.chat_message("ai").write(warning)
            msgs.add_ai_message(warning)

        elif not is_london_question(question):
            msg = "🛑 I'm sorry, I can only answer questions related to London."
            st.chat_message("ai").write(msg)
            msgs.add_ai_message(msg)

        else:
            with st.spinner("Thinking…"):
                response = rag_chain.invoke(question)
            st.chat_message("ai").write(response)
            msgs.add_ai_message(response)

if __name__ == "__main__":
    main()