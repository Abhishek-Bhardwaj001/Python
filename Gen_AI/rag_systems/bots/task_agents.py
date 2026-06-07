from typing import List, Union, Any
from langchain_core.messages import AIMessage, SystemMessage, HumanMessage

def format_query_bot(
    user_query: str,
    chat_history: str,
    chatbot: Any,
) -> str:
    if chat_history:
        message = [SystemMessage(content="""
                    Your task is to decide whether the user's question depends on previous conversation context.

                    Rules:
                    - If the question is a follow-up that depends on chat history, rewrite it into a standalone question.
                    - If the question is already standalone, return it EXACTLY as written.
                    - Do not use outside knowledge.
                    - Do not make guess.

                    Do NOT paraphrase standalone questions.

                    Return ONLY the final question text."""),
                HumanMessage(content = f""" Use the chat history below and rewrite user question.
                                            Question: {user_query}
                                            Chat History: {chat_history}
                                            """)]
        reformat_result = chatbot.invoke(message)
        search_query = reformat_result.content.strip()
    else:
        search_query = user_query
    return search_query


def generate_ai_summary(image_base64,claude_bot,model="claude-sonnet-4-5"):

    response = claude_bot.messages.create(
        model=model,
        max_tokens=500,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": image_base64
                        }
                    },
                    {
                        "type": "text",
                        "text": "Describe this image for document retrieval"
                    }
                ]
            }
        ]
    )

    return response.content[0].text