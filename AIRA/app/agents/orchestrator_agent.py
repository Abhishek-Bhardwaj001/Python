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
        try:
            if self.ai_engine == "groq":
                my_secret = os.getenv("GROQ_API")
                return ChatGroq(model = 'openai/gpt-oss-20b',api_key = my_secret)
            else:
                print("The Only Supported Orchestrator Model right now is GROQ")
        except Exception as e:
            print(f"API Error: {e}")
