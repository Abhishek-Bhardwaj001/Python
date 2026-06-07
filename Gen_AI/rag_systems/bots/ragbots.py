from typing import Any, List, Callable, Union
from langchain_core.messages import SystemMessage, HumanMessage, AIMessage
def history_aware_ragbot(
    user_query: str,
    chat_history: List[Union[HumanMessage, AIMessage]],
    format_history: Callable,
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