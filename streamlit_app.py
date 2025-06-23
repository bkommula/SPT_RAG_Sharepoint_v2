### REFRESHES CHROMADB ON EVERY USER QUERY AND LOGS IT ###

__import__('pysqlite3')
import sys
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')

import streamlit as st
import os
from dotenv import load_dotenv
from langchain_chroma import Chroma
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from openai import OpenAI
from langchain.prompts import ChatPromptTemplate
from langchain.schema import HumanMessage
from datetime import datetime

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "openai")
CHROMA_PATH = "chroma_db"

PROMPT_TEMPLATE = """
Use the following context to answer the user's question.

Context:
{context}

---

User: {question}
Assistant:"""

INFO_KEYWORDS = [
    "what", "how", "explain", "give me", "tell me", "who", "where", "when",
    "document", "file", "details", "report", "show", "find", "source"
]

def is_information_query_heuristic(user_input):
    result = any(kw in user_input.lower() for kw in INFO_KEYWORDS)
    print(f"[Heuristic] '{user_input}' → Match: {result}")
    return result

def is_information_query_llm(user_input):
    prompt = f"""
Decide if the following user input is asking for information from documents (e.g., question, file search) or just casual/small talk.

Input: "{user_input}"
Answer only with "yes" or "no".
"""
    classifier_llm = ChatOpenAI(model="gpt-4o-mini-2024-07-18", openai_api_key=OPENAI_API_KEY, temperature=0)
    result = classifier_llm.invoke([HumanMessage(content=prompt)])
    decision = result.content.strip().lower()
    print(f"[LLM Intent Classifier] Decision: {decision}")
    return decision == "yes"

def is_information_query(user_input):
    return is_information_query_heuristic(user_input) or is_information_query_llm(user_input)

def generate_casual_response(user_input):
    casual_prompt = f"""
You're a friendly AI assistant. The user said: "{user_input}"
Respond naturally and appropriately to their tone, without referring to documents or sources.
Keep it short, friendly, and human-like.
"""
    casual_llm = ChatOpenAI(model="gpt-4o-mini-2024-07-18", openai_api_key=OPENAI_API_KEY, temperature=0.6)
    print(f"[LLM Casual Reply Triggered] User Input: {user_input}")
    response = casual_llm.invoke([HumanMessage(content=casual_prompt)])
    return response.content.strip()

def get_fresh_vectorstore():
    print("[Vectorstore] Reloading ChromaDB client fresh from disk...")
    embedding = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)
    db = Chroma(persist_directory=CHROMA_PATH, embedding_function=embedding)

    try:
        total_chunks = db._collection.count()
        all_docs = db._collection.get()
        print(f"[Vectorstore] Total chunks in Chroma DB: {total_chunks}")
        print(f"[Vectorstore] Total documents retrieved: {len(all_docs['documents'])}")
        
        if any(doc is None for doc in all_docs["documents"]):
            print("[⚠️ Document Health] Found NoneType page_content — potential broken chunk(s)")
        else:
            print("[✅ Document Health] All documents have valid content")

    except Exception as e:
        print(f"[Vectorstore] Could not fetch chunk count: {e}")
    
    # For UI confirmation
    st.session_state["last_reload_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[🕒 Timestamp] Reloaded at: {st.session_state['last_reload_time']}")

    print("[Vectorstore] Reloaded.")
    return db

# Initialize session state
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

# Page config and title
st.set_page_config(page_title="AI Chat (Docs)", layout="centered")
st.title("🧠 AI Chat Assistant (SharePoint Docs)")

# Display chat history
for msg, role in st.session_state.chat_history:
    with st.chat_message(role):
        st.markdown(msg)

# User input
user_input = st.chat_input("Ask your question...")
if user_input:
    print(f"[User Input] {user_input}")
    st.chat_message("user").markdown(user_input)
    st.session_state.chat_history.append((user_input, "user"))

    if not is_information_query(user_input):
        print("[Intent] Detected casual conversation.")
        bot_response = generate_casual_response(user_input)
    else:
        print("[Intent] Information-seeking query")
        db = get_fresh_vectorstore()

        # Show last reload time in UI
        st.info(f"🔄 Vector store reloaded at `{st.session_state['last_reload_time']}`")

        with st.status("🔍 Searching the vector store...", expanded=False):
            try:
                raw_results = db.similarity_search_with_relevance_scores(user_input, k=2)
                results = []
                for doc, score in raw_results:
                    if not doc or not isinstance(doc.page_content, str):
                        print(f"[⚠️ Skipping invalid doc] {doc}")
                        continue
                    results.append((doc, score))

                print("[Vectorstore Results]")
                for doc, score in results:
                    excerpt = doc.page_content[:80].replace("\n", " ") if doc.page_content else "<Empty>"
                    print(f"  - Score: {score:.2f}, Excerpt: {excerpt}...")
            except Exception as e:
                print(f"[❌ Vectorstore Error] {e}")
                results = []

        if not results or results[0][1] < 0.65:
            print("[Relevance] No sufficiently relevant results.")
            bot_response = "❌ I couldn't find relevant information for that."
        else:
            context = "\n\n---\n\n".join([doc.page_content for doc, _ in results])
            sources = list({doc.metadata.get("source", "unknown") for doc, _ in results})
            prompt = ChatPromptTemplate.from_template(PROMPT_TEMPLATE).format(
                context=context, question=user_input
            )

            messages = [{"role": role, "content": msg} for msg, role in st.session_state.chat_history]
            messages.append({"role": "user", "content": prompt})

            with st.status("🤖 Generating response...", expanded=False) as status:
                if LLM_PROVIDER == "openrouter":
                    print("[LLM] Using OpenRouter")
                    client = OpenAI(api_key=OPENROUTER_API_KEY, base_url="https://openrouter.ai/api/v1")
                    response = client.chat.completions.create(
                        model="openai/gpt-4o-mini",
                        messages=messages
                    )
                    bot_response = response.choices[0].message.content
                else:
                    print("[LLM] Using OpenAI")
                    llm = ChatOpenAI(openai_api_key=OPENAI_API_KEY, model="gpt-4o-mini-2024-07-18", temperature=0)
                    bot_response = llm.invoke(messages).content

                if sources:
                    formatted_sources = "\n\n📎 **Sources:**\n" + "\n".join([f"- [{src}]({src})" for src in sources])
                    bot_response += formatted_sources

                status.update(label="✅ Done", state="complete")

    with st.chat_message("assistant"):
        st.markdown(bot_response)
    st.session_state.chat_history.append((bot_response, "assistant"))
