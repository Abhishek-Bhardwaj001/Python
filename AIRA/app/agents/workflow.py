from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import StateGraph, START, END
from langchain_core.messages import HumanMessage, RemoveMessage

from app.agents.history import to_lc_messages
from app.agents.nodes import llm_call
from app.models.orchestrator_state import State
from config.settings import default_thread_id


class Workflow():
    def __init__(self):
        self.app = self._compile()

    def run(self, query, thread_id=default_thread_id, history=None):
        """Answer `query` on the conversation identified by `thread_id`.

        The checkpointer owns the conversation. `history` is only used to
        rehydrate a thread that the checkpointer has never seen (a fresh
        Streamlit session, or the app restarting with UI state still in the
        browser); once the thread exists it is ignored, so nothing is inserted
        twice.
        """
        session = self._session(thread_id)

        messages = []
        if history and not self._has_history(session):
            messages.extend(to_lc_messages(history))
        messages.append(HumanMessage(content=query))

        response = self.app.invoke({'messages': messages}, config=session)
        return response['messages'][-1].content

    def get_history(self, thread_id=default_thread_id):
        """Return the messages the checkpointer currently holds for a thread."""
        return self._messages(self._session(thread_id))

    def reset(self, thread_id=default_thread_id):
        """Drop a thread's conversation so the next run starts clean."""
        self.app.update_state(
            self._session(thread_id),
            {'messages': [RemoveMessage(id=m.id) for m in self.get_history(thread_id)]}
        )

    # ========================= Internals ========================
    def _session(self, thread_id):
        return {
            "configurable": {
                'thread_id': thread_id
            }
        }

    def _messages(self, session):
        state = self.app.get_state(session)
        values = getattr(state, 'values', None) or {}

        if isinstance(values, dict):
            return values.get('messages') or []
        return getattr(values, 'messages', None) or []

    def _has_history(self, session):
        return bool(self._messages(session))

    def _compile(self):
        memory = MemorySaver()

        # ========================= Add Node ========================
        workflow = StateGraph(State)
        workflow.add_node("llm_call", llm_call)

        # =========================== Add Edge ==================
        workflow.add_edge(START, "llm_call")
        workflow.add_edge("llm_call", END)
        return workflow.compile(checkpointer=memory)
