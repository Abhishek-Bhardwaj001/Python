from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings
from dotenv import load_dotenv
import os
from typing import Any, List, Callable, Union
from langchain_groq import ChatGroq
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
from Gen_AI.rag_systems.utils.text_formatter import format_history
from Gen_AI.llm_prompts import chat_history_rag_prompt
from Gen_AI.config import db_directory
from Gen_AI.rag_systems.bots.task_agents import format_query_bot
load_dotenv()
LLM_SECRETS = os.getenv("GROQ_API")

print(f"Database DIR:{db_directory}")
new_db_path = str(db_directory / "text_data_vectors")

embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

db = Chroma(persist_directory=new_db_path,
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
    print(f"Reformatted Search Query: {search_query}")
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
            ai_answer = ragbot(user_query, chat_history, retriever, llm_chatbot, format_query_bot, chat_history_rag_prompt)
            print(f"chatbot Response:\n{ai_answer}")
            chat_history.append(HumanMessage(content=user_query))
            chat_history.append(AIMessage(content=ai_answer))
            chat_history[:] = chat_history[-30:]
if __name__=="__main__":
      main()       