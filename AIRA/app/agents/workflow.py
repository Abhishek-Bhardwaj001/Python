from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import StateGraph,add_messages,START,END
from app.agents.nodes import llm_call
from app.models.orchestrator_state import State
from langchain_core.messages import BaseMessage, HumanMessage, SystemMessage,AIMessage

class Workflow():
    def __init__(self):
        self.app = self._compile()
    
    def run(self, query, thread_id="default"):
        session = {
            "configurable":{
                'thread_id': thread_id
            }
        }
        response = self.app.invoke({
            'messages': [HumanMessage(content=query)]},
            config=session
        )
        return response['messages'][-1].content

    def _compile(self):
        memory = MemorySaver()
 
        # ========================= Add Node ========================
        workflow = StateGraph(State)
        workflow.add_node("llm_call",llm_call)

        # =========================== Add Edge ==================
        workflow.add_edge(START,"llm_call")
        workflow.add_edge("llm_call",END)
        return workflow.compile(checkpointer=memory)