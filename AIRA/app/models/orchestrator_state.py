from pydantic import BaseModel,Field
from typing import Annotated, List,Optional
from langgraph.graph import add_messages
from langchain_core.messages import BaseMessage, HumanMessage, SystemMessage,AIMessage

class State(BaseModel):
    messages :Annotated[List[BaseMessage],add_messages] =Field(description="LLM Messages")