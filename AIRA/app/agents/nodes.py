from app.agents.orchestrator_agent import AIRA
from langchain_core.messages import AIMessage

llm = AIRA()


def llm_call(state):
    response = llm.invoke(state.messages)

    # `messages` uses the add_messages reducer, so return only the new turn.
    return {
        "messages": [AIMessage(content=response)]
    }
