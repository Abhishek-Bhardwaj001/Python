from typing import List, Union, Any, Callable
from langchain_core.messages import AIMessage, SystemMessage, HumanMessage

def format_query(
    user_query: str,
    chat_history: str,
    chatbot: Any,
) -> str:
    if chat_history:
        print("Chat history found")
        message = [SystemMessage(content="""
                    Your task is to decide whether the user's question depends on previous conversation context.

                    Rules:
                    - If the question is a follow-up that depends on chat history, rewrite it into a standalone question.
                    - If the question is already standalone, return it EXACTLY as written.
                    - Do not use outside knowledge.
                    - Do not make guess.

                    Do NOT paraphrase standalone questions.

                    Return ONLY the final question text."""),
                HumanMessage(content = f""" Use the chat history below and rewrite user question.
                                            Question: {user_query}
                                            Chat History: {chat_history}
                                            """)]
        reformat_result = chatbot.invoke(message)
        search_query = reformat_result.content.strip()
    else:
        print("No Chat history found starting a new session")
        search_query = user_query
    return search_query

def format_message(
    user_query: str,
    chat_history: str,
    combined_context: str,
) -> List[Any]:
    if chat_history:
        print("Chat history found")

        message = [SystemMessage(content="Role: You are an RAG Agent specifically designed to respond to user question based on context and chat history provided to you"),
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

def format_history(chat_history: List[Union[HumanMessage, AIMessage]]) -> str:
    return "\n".join(
        f"{'User' if isinstance(msg, HumanMessage) else 'Assistant'}: {msg.content}"
        for msg in chat_history
    )