import json
import os
from typing import Any, List, Callable, Union
from unstructured.partition.pdf import partition_pdf
from unstructured.chunking.title import chunk_by_title
from langchain_core.documents import Document
from langchain_chroma import Chroma
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from langchain_groq import ChatGroq
from anthropic import Anthropic
import base64
from bs4 import BeautifulSoup
from langchain_huggingface import HuggingFaceEmbeddings

from dotenv import load_dotenv
from Gen_AI.rag_systems.services.document_vectorization_wth_unstructured import partition_document,process_elements
from Gen_AI.rag_systems.services.data_vectorization import DataVectorization
from Gen_AI.rag_systems.utils.data_ingestion_helpers import create_vector_store
from Gen_AI.rag_systems.utils.data_transform_helpers import convert_langchain_doc,merge_documents,assign_metadata_to_chunks
from Gen_AI.rag_systems.bots.task_agents import generate_ai_summary,format_query_bot
from Gen_AI.rag_systems.bots.ragbots import history_aware_ragbot
from Gen_AI.llm_prompts import chat_history_rag_prompt
from Gen_AI.rag_systems.utils.text_formatter import format_history
from Gen_AI.config import pdf_path,db_directory
load_dotenv()

CLAUDE_API = os.getenv("CLAUDE_API")
LLM_SECRETS = os.getenv("GROQ_API")

llm_chatbot = ChatGroq(model='llama-3.1-8b-instant',api_key = LLM_SECRETS)
client = Anthropic(api_key=CLAUDE_API) # Used for AI Image Summary
pdf_vector = str(db_directory / "attention_is_all_you_need_pdf_vector")

def data_ingestion():
    elements = partition_document(file_path)
    processed_documents = process_elements(elements,generate_ai_summary,client)
    full_text, metadata_tracker = merge_documents(processed_documents)
    vectorize = DataVectorization()
    chunks = vectorize.text_splitter(full_text, embedding_model_token_limit = 512,chunk_overlap_percent=0.25,verbose=False)
    chunked_documents = assign_metadata_to_chunks(chunks, metadata_tracker)
    documents = convert_langchain_doc(chunked_documents)
    db = create_vector_store(documents,db_directory=pdf_vector)

embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

db = Chroma(persist_directory=pdf_vector,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space":"cosine"})
retriever = db.as_retriever(search_kwargs={"k": 3})
chat_history: List[Union[HumanMessage, AIMessage]] = []

def main() -> None:    
    while True:
        user_query = str(input("Please ask your query here:"))
        #query Example: "What are the two main components of the Transformer architecture? "
        if user_query in ['quit','exit','bye']:
            print("See you, take care!")
            break
        else:
            ai_answer = history_aware_ragbot(user_query, chat_history,format_history, retriever, llm_chatbot, format_query_bot, chat_history_rag_prompt)
            print(f"chatbot Response:\n{ai_answer}")
            chat_history.append(HumanMessage(content=user_query))
            chat_history.append(AIMessage(content=ai_answer))
            chat_history[:] = chat_history[-10:]

if __name__=="__main__":
      main() 
