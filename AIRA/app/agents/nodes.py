from app.agents.orchestrator_agent import AIRA
from langchain_core.messages import AIMessage

llm = AIRA()

def llm_call(state):
    response = llm.invoke(state.messages)
    return {"messages": [AIMessage(content=response)]}
