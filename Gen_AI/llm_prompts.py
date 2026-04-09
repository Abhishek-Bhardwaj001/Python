from langchain_core.messages import SystemMessage, HumanMessage
from typing import List, Any


def basic_rag_prompt(query_input,document_context:str):
    message_prompt = [SystemMessage(content = "You are an AI Assistant trained on documents"),
            HumanMessage(content = f"""Question:
                            {query_input}
                            Context: {document_context}
                            respond to users query with response curated only from documents provided in context""")
                            ]
    return message_prompt

def chat_history_rag_prompt(
    user_query: str,
    chat_history: str,
    combined_context: str,
) -> List[Any]:
    if chat_history:
        print("Chat history found")

        message = [SystemMessage(content="""Role: You are an RAG Agent specifically designed to respond to user question based on context and chat history provided to you" \
                    Strict Rules:
                    - Answer ONLY using provided context
                    - DO NOT mention chat history
                    - DO NOT say 'Based on context' or 'Based on chat history'
                    - DO NOT make assumptions
                    - Respond with ONLY the final answer

                    If answer not found, respond with:
                    'I don't know'"""),
                HumanMessage(content = f""" Use the context and chat history below and respond to user question
                                            Question: {user_query}
                                            Context: {combined_context}
                                            Chat History: {chat_history}
                                            """)]
    else:
        print("No Chat history found starting a new session")
        message = [SystemMessage(content="""Role: You are an RAG Agent specifically designed to respond to user question based on context provided to you.
                                 Response Format: Respond only with Answer"""),
        HumanMessage(content = f""" Use the context below and respond to user question
                                    Question: {user_query}
                                    Context: {combined_context}
                                    """)]
    return message