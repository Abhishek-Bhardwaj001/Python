from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings
from dotenv import load_dotenv
import os
from typing import Any, List, Callable, Union
from langchain_groq import ChatGroq
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from utils.text_formatter import format_query, format_chat_history_message, format_history

load_dotenv()
LLM_SECRETS = os.getenv("GROQ_API")

db_directory = "db/chroma_db"
embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

db = Chroma(persist_directory=db_directory,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space":"cosine"})


retriever = db.as_retriever(search_kwargs = {"k":5})

llm_chatbot = ChatGroq(model='llama-3.1-8b-instant',api_key = LLM_SECRETS) 
print(f"user Query Example:How much did microsoft pay to acquire github?")
chat_history: List[Union[HumanMessage, AIMessage]] = []

def ragbot(
    user_query: str,
    chat_history: List[Union[HumanMessage, AIMessage]],
    retriever: Any,
    chatbot: Any,
    format_query: Callable[[str, str, Any], str],
    format_chat_history_message: Callable[[str, str, str], List[Union[SystemMessage, HumanMessage]]],
) -> None:
    history_text = format_history(chat_history)
    search_query = format_query(user_query, history_text, chatbot)
    relevant_docs = retriever.invoke(search_query)
    combined_context = "\n\n".join(doc.page_content for doc in relevant_docs)
    message = format_chat_history_message(search_query, history_text, combined_context)
    response = chatbot.invoke(message)
    return response.content.strip()

def main() -> None:    
    while True:
        user_query = str(input("Please ask your query here:"))
        if user_query in ['quit','exit','bye']:
            print("See you, take care!")
            break
        else:
            ai_answer = ragbot(user_query, chat_history, retriever, llm_chatbot, format_query, format_chat_history_message)
            print(f"chatbot Response:\n{ai_answer}")
            chat_history.append(HumanMessage(content=user_query))
            chat_history.append(AIMessage(content=ai_answer))
            chat_history[:] = chat_history[-30:]
if __name__=="__main__":
      main()       