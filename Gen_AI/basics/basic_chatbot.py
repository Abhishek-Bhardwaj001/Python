import os

from dotenv import load_dotenv
from langchain_groq import ChatGroq

load_dotenv()  # This loads the variables from .env

my_secret = os.getenv("GROQ_API_KEY")

llm_chat = ChatGroq(model = 'llama-3.1-8b-instant')

message = [
    ("system", "You are an expert on Mithical Creatures"),
    ("human", "List out all the Mithical creatures?")
]

llm_response = llm_chat.invoke(message)

print(llm_response.content)