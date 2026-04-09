from typing import List, Union, Any
from langchain_core.messages import AIMessage, SystemMessage, HumanMessage


def format_history(chat_history: List[Union[HumanMessage, AIMessage]]) -> str:
    return "\n".join(
        f"{'User' if isinstance(msg, HumanMessage) else 'Assistant'}: {msg.content}"
        for msg in chat_history
    )
