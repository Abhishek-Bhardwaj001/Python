from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings
from dotenv import load_dotenv
from langchain_groq import ChatGroq
import os
from langchain_core.messages import SystemMessage, HumanMessage
from Gen_AI.config import db_directory,docs_path
from Gen_AI.rag_systems.utils.data_ingestion_helpers import load_text_documents,split_documents, create_vector_store
from Gen_AI.llm_prompts import basic_rag_prompt

#=============== Configurables ====================
print(f"Database DIR:{db_directory}")
new_db_path = str(db_directory / "text_data_vectors")

embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

load_dotenv()
api_secret = os.getenv('GROQ_API')
llm_chat = ChatGroq(model = 'llama-3.1-8b-instant',api_key=api_secret)

# ============ Data Ingestion =================================
def data_ingestion():
    documents = load_text_documents(docs_path=docs_path,verbose=True)
    chunks = split_documents(documents)
    vector_store = create_vector_store(chunks,new_db_path,verbose=True)
    return vector_store

# =============== Creating Vector DB Retriever =========================
db = Chroma(persist_directory=new_db_path,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space":"cosine"})
if db._collection.count()==0:
    print(f"[WARN] The vector DB for Documents is not found. Creating New from {docs_path}")
    data_ingestion()
retriever = db.as_retriever(search_kwargs = {"k":5})

# ====================== User Query ===================================
query_input =  input("Type your query:\n")
relevant_docs = retriever.invoke(query_input)

print(f"user Query:\n{query_input}")
# Query Example: When did microsoft bought git?

combined_input = "".join(doc.page_content for doc in relevant_docs)

# ============ LLm Response generation ====================
message = basic_rag_prompt(query_input,combined_input)
response = llm_chat.invoke(message)

print(f"Chatbot Response:\n{response.content}")