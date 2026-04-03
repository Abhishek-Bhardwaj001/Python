import json
import os
from typing import List
from unstructured.partition.pdf import partition_pdf
from unstructured.chunking.title import chunk_by_title


from langchain_core.documents import Document
from langchain_chroma import Chroma
from langchain_core.messages import HumanMessage,SystemMessage
from langchain_groq import ChatGroq
from dotenv import load_dotenv

load_dotenv()
LLM_SECRETS = os.getenv("GROQ_API")

groq = ChatGroq(model='llama-3.1-8b-instant',api_key = LLM_SECRETS)

def format_ai_input():
    message = [SystemMessage(content="Role:You are an AI agent trained on ")]
response = groq.invoke()
print(response.content)
