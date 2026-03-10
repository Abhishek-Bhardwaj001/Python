from langchain_chroma import Chroma
from langchain_huggingface import HuggingFaceEmbeddings
from dotenv import load_dotenv
from langchain_groq import ChatGroq
import os
from langchain_core.messages import SystemMessage, HumanMessage



db_directory = "db/chroma_db"

embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

db = Chroma(persist_directory=db_directory,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space":"cosine"})
# if len(db)==0:
#     raise FileNotFoundError("The vector DB for Documents is not found")
query = "How much did microsoft pay to acquire github?"

retriever = db.as_retriever(search_kwargs = {"k":5})

relevant_docs = retriever.invoke(query)

print(f"user Query:{query}")

print("-----Context-----")

# for i,doc in enumerate(relevant_docs,1):
#     print(f"document at {i}:\n {doc.page_content}\n")

combined_input = "".join(doc.page_content for doc in relevant_docs)

# print(f"\n--------------Combined Input-----------------\n{combined_input}")
load_dotenv()
api_secret = os.getenv('GROQ_API')
llm_chat = ChatGroq(model = 'llama-3.1-8b-instant',api_key=api_secret)

query_input = input("Type your query:\n")
message = [SystemMessage(content = "You are an AI Assistant trained on documents"),
           HumanMessage(content = f"""Question:
                        {query_input}
                        Context: {combined_input}
                        respond to users query with response curated only from documents provided in context""")
                        ]
response = llm_chat.invoke(message)

print(f"Chatbot Response:{response.content}")

# Checkpoint: Build History Aware Chatbot