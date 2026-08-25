import os
from dotenv import load_dotenv
from langchain_groq import ChatGroq
from config.settings import orchestrator_prompt
from langchain_core.messages import SystemMessage, HumanMessage
class AIRA():
    def __init__(self,ai_engine="groq"):
        self.ai_engine = ai_engine
        load_dotenv()
        self.orchestrator_llm = self._llm()
    
    def invoke(self, messages):
        print("Entered Invoke function for Orchestrator A.I.R.A thinking...")
        full_messages = [SystemMessage(content=orchestrator_prompt)] + list(messages)
        response = self.orchestrator_llm.invoke(full_messages)
        return response.content

    def _llm(self):
        if self.ai_engine != "groq":
            raise ValueError(
                f"Unsupported orchestrator engine '{self.ai_engine}'. Only 'groq' is supported right now."
            )

        my_secret = os.getenv("GROQ_API")
        if not my_secret:
            raise RuntimeError("GROQ_API is not set. Add it to your .env before starting AIRA.")

        return ChatGroq(model='openai/gpt-oss-20b', api_key=my_secret, max_tokens=250)
