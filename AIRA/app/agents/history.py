from langchain_core.messages import AIMessage, HumanMessage


def to_lc_messages(history):
    """Convert UI history dicts into LangChain messages.

    The UI stores turns as {"role": ..., "content": ..., "audio": ...}. Only role
    and content matter to the LLM, so audio blobs are dropped. Blank turns are
    skipped so a failed STT attempt never reaches the model.
    """
    messages = []

    for turn in history or []:
        content = (turn.get("content") or "").strip()
        if not content:
            continue

        if turn.get("role") == "user":
            messages.append(HumanMessage(content=content))
        else:
            messages.append(AIMessage(content=content))

    return messages
