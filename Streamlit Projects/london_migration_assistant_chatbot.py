
# pip install streamlit langchain-community langchain-core langchain-chroma langchain-ollama chromadb plotly pandas pypdf

# ollama pull gemma3:1b -> to pull and install ollama gemma exe 
# ollama pull nomic-embed-text -> to pull and install ollama embedding exe 

# streamlit run london_migration_assistant_chatbot.py -> to run  the app


# app.py — London Migration Assistant (Enhanced RAG Application)
# -----------------------------------------------------------------------------
# A comprehensive AI assistant for people planning to move to London.
# Covers housing, visas, jobs, healthcare, transport, education, and more.
# Built with RAG (Retrieval-Augmented Generation) using local PDFs.

# ───────────────────────────── imports ─────────────────────────────
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from typing import List, Dict, Optional
import re
import datetime as dt
import pandas as pd
import json

# LangChain imports
from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
# from langchain_community.embeddings import OllamaEmbeddings
from langchain_ollama import OllamaEmbeddings 
# from langchain_community.vectorstores import Chroma
from langchain_chroma import Chroma
from langchain_community.chat_models import ChatOllama
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnablePassthrough, RunnableParallel
from langchain_community.chat_message_histories import StreamlitChatMessageHistory

# ───────────────────────── configuration ─────────────────────────
# Enhanced configuration for London migration assistant

PDF_FOLDER = Path(r"C:\Users\TECHTALENT-009\Desktop\Tyy\Python Programming\The survivors group\pdfs")
CHROMA_DIR = Path("data/london_migration_db")
EMBEDDING_MODEL = "nomic-embed-text"
OLLAMA_MODEL = "gemma3:1b"
TEMPERATURE = 0.3  # Slightly higher for more helpful responses
CHUNK_SIZE = 1200  # Larger chunks for better context
CHUNK_OVERLAP = 200  # More overlap for continuity
MAX_TOKENS = 2048

# Migration categories for better organization
MIGRATION_CATEGORIES = {
    "🏠 Housing & Accommodation": ["housing", "rent", "buy", "property", "accommodation", "landlord", "deposit", "council tax"],
    "💼 Work & Employment": ["job", "work", "employment", "visa", "permit", "salary", "tax", "career", "profession"],
    "🏥 Healthcare & Services": ["nhs", "healthcare", "doctor", "gp", "hospital", "medical", "insurance", "prescription"],
    "🚇 Transport & Travel": ["transport", "tube", "bus", "train", "oyster", "travel", "car", "driving", "license"],
    "🏛️ Legal & Immigration": ["visa", "immigration", "permit", "citizen", "passport", "legal", "law", "rights"],
    "🎓 Education & Schools": ["school", "university", "education", "student", "course", "qualification", "degree"],
    "💰 Finance & Banking": ["bank", "account", "money", "finance", "credit", "loan", "mortgage", "budget", "cost"],
    "🌆 Lifestyle & Culture": ["culture", "food", "entertainment", "social", "community", "weather", "lifestyle"],
    "📋 Admin & Bureaucracy": ["council", "register", "tax", "ni", "insurance", "benefits", "pension", "utilities"]
}

# ──────────────────── enhanced PDF processing ────────────────────

def load_and_categorize_pdfs(folder: Path) -> Dict[str, List]:
    """Load PDFs and categorize them by content type for better retrieval."""
    categorized_docs = {category: [] for category in MIGRATION_CATEGORIES.keys()}
    all_docs = []
    
    if not folder.exists():
        st.error(f"PDF folder not found: {folder}")
        return {"uncategorized": []}
    
    pdf_files = list(folder.glob("*.pdf"))
    if not pdf_files:
        st.warning(f"No PDF files found in {folder}")
        return {"uncategorized": []}
    
    progress_bar = st.progress(0)
    for i, pdf_path in enumerate(pdf_files):
        try:
            loader = PyPDFLoader(str(pdf_path))
            pages = loader.load()
            
            # Enhanced metadata
            fallback_title = pdf_path.stem.replace("_", " ").title()
            category = categorize_document(pdf_path.name, pages)
            
            for page in pages:
                page.metadata.update({
                    "title": fallback_title,
                    "source_file": pdf_path.name,
                    "category": category,
                    "processed_date": dt.datetime.now().isoformat(),
                    "page_number": page.metadata.get("page", 0)
                })
                
            categorized_docs[category].extend(pages)
            all_docs.extend(pages)
            
        except Exception as e:
            st.error(f"Error loading {pdf_path.name}: {str(e)}")
        
        progress_bar.progress((i + 1) / len(pdf_files))
    
    progress_bar.empty()
    return categorized_docs, all_docs

def categorize_document(filename: str, pages: List) -> str:
    """Automatically categorize documents based on filename and content."""
    filename_lower = filename.lower()
    
    # Sample content from first few pages
    content_sample = " ".join([page.page_content[:500] for page in pages[:3]]).lower()
    
    for category, keywords in MIGRATION_CATEGORIES.items():
        if any(keyword in filename_lower or keyword in content_sample for keyword in keywords):
            return category
    
    return "📋 General Information"

def enhanced_chunk_documents(documents: List) -> List:
    """Enhanced document chunking with better separators and metadata preservation."""
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n\n", "\n\n", "\n", ".", "!", "?", ";", " ", ""],
        keep_separator=True,
        add_start_index=True
    )
    
    chunks = splitter.split_documents(documents)
    
    # Add chunk-specific metadata
    for i, chunk in enumerate(chunks):
        chunk.metadata["chunk_id"] = i
        chunk.metadata["chunk_length"] = len(chunk.page_content)
        
    return chunks

# ──────────────────── enhanced vector store ────────────────────

@st.cache_resource(show_spinner="🔄 Building your London knowledge base...")
def get_enhanced_vectorstore():
    """Enhanced vector store with better indexing and metadata."""
    embedding_fn = OllamaEmbeddings(model=EMBEDDING_MODEL)
    
    if CHROMA_DIR.exists() and any(CHROMA_DIR.iterdir()):
        try:
            return Chroma(persist_directory=str(CHROMA_DIR), embedding_function=embedding_fn)
        except Exception as e:
            st.warning(f"Rebuilding vector store due to error: {e}")
            import shutil
            shutil.rmtree(CHROMA_DIR)
    
    # Build new vector store
    categorized_docs, all_docs = load_and_categorize_pdfs(PDF_FOLDER)
    
    if not all_docs:
        st.error("No documents found to process!")
        return None
    
    chunked_docs = enhanced_chunk_documents(all_docs)
    
    # Create vector store with collection metadata
    vectorstore = Chroma.from_documents(
        documents=chunked_docs,
        embedding=embedding_fn,
        persist_directory=str(CHROMA_DIR),
        collection_metadata={"description": "London Migration Assistant Knowledge Base"}
    )
    
    return vectorstore

# ──────────────────── enhanced safety and validation ────────────────────

def enhanced_safety_filter(user_input: str) -> tuple[bool, str]:
    """Enhanced safety filter with specific feedback."""
    
    # Inappropriate content
    banned_patterns = [
        r'\b(kill|murder|bomb|explode|terrorist|attack)\b',
        r'\b(hate|racist|discrimination)\b',
        r'\b(suicide|self.harm|hurt.myself)\b',
        r'\b(illegal|scam|fraud|cheat)\b'
    ]
    
    for pattern in banned_patterns:
        if re.search(pattern, user_input.lower()):
            return False, "⚠️ Please keep our conversation respectful and focused on helpful information about moving to London."
    
    return True, ""

def is_london_migration_related(user_input: str) -> tuple[bool, str]:
    """Check if question is related to London migration with helpful guidance."""
    
    london_keywords = [
        "london", "uk", "britain", "british", "england", "english",
        "move", "relocate", "migrate", "immigration", "expat"
    ]
    
    migration_keywords = [
        "visa", "housing", "job", "work", "transport", "healthcare", "school",
        "cost", "life", "live", "living", "move", "relocate", "settle"
    ]
    
    input_lower = user_input.lower()
    
    has_london = any(keyword in input_lower for keyword in london_keywords)
    has_migration = any(keyword in input_lower for keyword in migration_keywords)
    
    if has_london or has_migration:
        return True, ""
    
    suggestion = """
🎯 **I specialize in helping people move to London!**

I can help you with:

**🏠 Housing & Living**
• Finding accommodation and understanding rental markets
• Cost of living and budgeting advice

**💼 Work & Legal**
• Visa and immigration requirements
• Job market insights and employment guidance

**🏥 Essential Services**
• Healthcare (NHS) registration and medical services
• Banking, utilities, and administrative setup

**🚇 Daily Life**
• Transport systems and getting around London
• Education options and school admissions
• Cultural adaptation and lifestyle tips

---

**💡 Try asking something like:**
• *"How much does housing cost in London?"*
• *"What visa do I need to work in London?"*
• *"How do I register with the NHS?"*
• *"What's the best area to live in London?"*
"""
    
    return False, suggestion

# ──────────────────── enhanced RAG chain ────────────────────

def build_enhanced_rag_chain(vectorstore):
    """Enhanced RAG chain with better prompting and retrieval."""
    
    # Enhanced retriever with multiple search strategies
    base_retriever = vectorstore.as_retriever(
        search_type="mmr",
        search_kwargs={
            "k": 8,
            "fetch_k": 25,
            "lambda_mult": 0.7,  # Balance between relevance and diversity
        }
    )
    
    # Enhanced system prompt
    system_prompt = """
You are the **London Migration Assistant** - a specialized AI helper for people planning to move to London.

🎯 **Your Mission**: Provide accurate, practical, and empathetic guidance on ALL aspects of relocating to London.

📚 **Knowledge Scope**: You have access to comprehensive PDFs covering:
• Housing markets, rental processes, and property buying
• Visa requirements and immigration procedures
• Job markets, employment law, and career opportunities
• NHS healthcare system and medical services
• Transport networks (Tube, buses, trains)
• Education system and school admissions
• Banking, finance, and cost of living
• Cultural adaptation and social integration
• Legal requirements and bureaucratic processes

🗣️ **Communication Style**:
• Be warm, encouraging, and understanding - moving countries is stressful!
• Provide specific, actionable advice with concrete examples
• Structure complex information with clear headings and bullet points
• Always cite the PDF source when providing specific facts or figures
• If information might be outdated, mention this and suggest verification

⚡ **Response Strategy**:
1. **Acknowledge** the person's situation empathetically
2. **Answer** their specific question using the provided context
3. **Expand** with related helpful information they might not have considered
4. **Guide** them to next steps or additional resources when relevant

🚨 **Important Rules**:
• ONLY use information from the provided PDF context
• If context is insufficient, clearly state "I don't have specific information about this in my current documents"
• Never guess about visa requirements, legal processes, or official procedures
• For urgent matters (medical, legal, safety), direct them to official authorities
• When discussing costs, always mention these may have changed and suggest checking current rates

Remember: You're not just answering questions - you're helping someone navigate one of life's biggest transitions!
"""

    prompt = ChatPromptTemplate.from_template(f"""
{system_prompt}

**CONTEXT FROM LONDON DOCUMENTS:**
{{context}}

**PERSON'S QUESTION:**
{{question}}

**YOUR HELPFUL RESPONSE:**
""")

    llm = ChatOllama(
        model=OLLAMA_MODEL, 
        temperature=TEMPERATURE,
        num_ctx=MAX_TOKENS
    )

    # Enhanced chain with parallel processing
    chain = (
        RunnableParallel({
            "context": base_retriever,
            "question": RunnablePassthrough()
        })
        | prompt
        | llm
        | StrOutputParser()
    )
    
    return chain

# ──────────────────── enhanced UI components ────────────────────

def render_sidebar():
    """Enhanced sidebar with migration categories and quick tips."""
    
    with st.sidebar:
        st.markdown("## 🎯 Quick Navigation")
        
        # Category buttons
        st.markdown("### 📋 Ask about:")
        for category, keywords in MIGRATION_CATEGORIES.items():
            if st.button(category, key=f"cat_{category}"):
                example_questions = get_example_questions(category)
                st.session_state['suggested_question'] = example_questions[0]
        
        st.markdown("---")
        
        # Quick stats
        if 'vectorstore' in st.session_state:
            st.markdown("### 📊 Knowledge Base")
            try:
                collection = st.session_state.vectorstore._collection
                doc_count = collection.count()
                st.metric("Documents processed", doc_count)
            except:
                st.info("Knowledge base loaded ✅")
        
        st.markdown("---")
        
        # Migration timeline
        st.markdown("### 📅 Typical Migration Timeline")
        timeline_data = {
            "Phase": ["Research", "Visa Prep", "Applications", "Moving", "Settling"],
            "Months": [1, 2, 3, 1, 6]
        }
        fig = px.bar(timeline_data, x="Phase", y="Months", title="Planning Timeline")
        st.plotly_chart(fig, use_container_width=True)

def get_example_questions(category: str) -> List[str]:
    """Generate example questions for each category."""
    examples = {
        "🏠 Housing & Accommodation": [
            "What's the average rent for a 1-bedroom flat in London?",
            "How do I find reliable accommodation before moving?",
            "What documents do I need to rent in London?"
        ],
        "💼 Work & Employment": [
            "What visa do I need to work in London?",
            "How competitive is the job market in London?",
            "What's the average salary for my profession?"
        ],
        "🏥 Healthcare & Services": [
            "How do I register with the NHS?",
            "Do I need private health insurance in London?",
            "How do I find a local GP?"
        ],
        "🚇 Transport & Travel": [
            "How much does an Oyster card cost?",
            "What's the best way to get around London?",
            "Can I use my foreign driving license?"
        ]
    }
    return examples.get(category, ["Tell me about " + category.lower()])

def render_chat_interface(rag_chain, msgs):
    """Enhanced chat interface with better formatting."""
    
    # Display chat history with enhanced formatting
    for msg in msgs.messages:
        with st.chat_message(msg.type):
            if msg.type == "ai":
                # Enhanced AI message formatting
                st.markdown(msg.content)
            else:
                st.write(msg.content)
    
    # Handle suggested questions
    if 'suggested_question' in st.session_state:
        st.info(f"💡 Try asking: {st.session_state.suggested_question}")
        if st.button("Ask this question"):
            st.session_state['auto_question'] = st.session_state.suggested_question
            del st.session_state.suggested_question
            st.rerun()
    
    # Auto-submit suggested question
    if 'auto_question' in st.session_state:
        question = st.session_state.auto_question
        del st.session_state.auto_question
        process_question(question, rag_chain, msgs)
        st.rerun()
    
    # Main chat input - moved to after message display
    if question := st.chat_input("Ask me anything about moving to London... 🇬🇧"):
        process_question(question, rag_chain, msgs)
        st.rerun()

def process_question(question: str, rag_chain, msgs):
    """Process user question with enhanced error handling."""
    
    # Add user message to history first
    msgs.add_user_message(question)
    
    # Safety checks
    is_safe, safety_msg = enhanced_safety_filter(question)
    if not is_safe:
        msgs.add_ai_message(safety_msg)
        return
    
    is_relevant, relevance_msg = is_london_migration_related(question)
    if not is_relevant:
        msgs.add_ai_message(relevance_msg)
        return
    
    # Process with RAG
    try:
        response = rag_chain.invoke(question)
        msgs.add_ai_message(response)
        
        # Add follow-up suggestions
        follow_ups = generate_follow_up_questions(question)
        if follow_ups:
            follow_up_text = "\n\n---\n**💡 You might also want to ask:**\n"
            for follow_up in follow_ups[:2]:
                follow_up_text += f"• {follow_up}\n"
            msgs.add_ai_message(msgs.messages[-1].content + follow_up_text)
            
    except Exception as e:
        error_msg = f"🔧 I encountered a technical issue: {str(e)}\n\nPlease try rephrasing your question or contact support if this persists."
        msgs.add_ai_message(error_msg)

def generate_follow_up_questions(original_question: str) -> List[str]:
    """Generate contextual follow-up questions."""
    question_lower = original_question.lower()
    
    if "housing" in question_lower or "rent" in question_lower:
        return [
            "What areas of London are most affordable?",
            "What are the typical rental contract terms?",
            "How much should I budget for utilities?"
        ]
    elif "visa" in question_lower or "work" in question_lower:
        return [
            "How long does visa processing typically take?",
            "What documents do I need for my visa application?",
            "Can my family come with me on my work visa?"
        ]
    elif "transport" in question_lower:
        return [
            "Which London transport zones should I consider for housing?",
            "Are there discounts available for transport cards?",
            "How reliable is London public transport?"
        ]
    
    return []

# ──────────────────── main application ────────────────────

def main():
    """Enhanced main application with better structure."""
    
    # Page configuration
    st.set_page_config(
        page_title="London Migration Assistant",
        page_icon="🇬🇧",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Header
    st.markdown("""
    # 🇬🇧 London Migration Assistant
    ### Your AI-powered guide to moving to London
    
    I'm here to help you navigate every aspect of relocating to London - from finding housing and understanding visas 
    to getting around the city and settling into British life. Ask me anything!
    """)
    
    # Initialize session state
    if 'vectorstore' not in st.session_state:
        with st.spinner("🔄 Setting up your London knowledge base..."):
            st.session_state.vectorstore = get_enhanced_vectorstore()
    
    vectorstore = st.session_state.vectorstore
    if not vectorstore:
        st.error("❌ Could not load the knowledge base. Please check your PDF folder.")
        return
    
    # Build RAG chain
    if 'rag_chain' not in st.session_state:
        st.session_state.rag_chain = build_enhanced_rag_chain(vectorstore)
    
    # Initialize chat history
    msgs = StreamlitChatMessageHistory(key="london_migration_chat")
    if not msgs.messages:
        welcome_msg = """
👋 **Welcome to your London Migration Assistant!** 

I'm here to help you with everything about moving to London. I have comprehensive information about:

🏠 **Housing**: Renting, buying, areas, costs  
💼 **Work**: Visas, jobs, salaries, taxes  
🏥 **Healthcare**: NHS, doctors, insurance  
🚇 **Transport**: Tubes, buses, travel cards  
📋 **Admin**: Banking, council tax, utilities  
🎓 **Education**: Schools, universities  
🌆 **Lifestyle**: Culture, food, social life  

**What would you like to know about moving to London?**
        """
        msgs.add_ai_message(welcome_msg)
    
    # Render UI
    col1, col2 = st.columns([3, 1])
    
    with col1:
        render_chat_interface(st.session_state.rag_chain, msgs)
    
    with col2:
        render_sidebar()

if __name__ == "__main__":
    main()