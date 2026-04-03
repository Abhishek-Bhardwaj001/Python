from langchain_core.messages import SystemMessage, HumanMessage

def basic_rag_prompt(query_input,document_context:str):
    message_prompt = [SystemMessage(content = "You are an AI Assistant trained on documents"),
            HumanMessage(content = f"""Question:
                            {query_input}
                            Context: {document_context}
                            respond to users query with response curated only from documents provided in context""")
                            ]
    return message_prompt