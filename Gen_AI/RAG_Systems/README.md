# RAG_Systems

This directory contains a small retrieval-augmented generation (RAG) system implemented using LangChain and the Chroma vector database. It includes utilities for ingesting text documents, building a vector store, and running both a simple and history-aware chatbot.

## 📁 Directory Structure

```
RAG_Systems/
├── db/
│   └── chroma_db/          # Persisted Chroma vector store data
├── docs/                   # Source text documents used for ingestion
│   ├── google.txt
│   └── microsoft.txt
├── src/                    # Main example scripts
│   ├── basic_rag.py        # Basic RAG chatbot demonstration
│   ├── history_aware_rag.py# History-aware RAG chatbot with query reformulation
│   └── ingestion_pipeline.py # Data ingestion & vector store creation
└── utils/
    └── text_formatter.py   # Helper for formatting queries & messages
```

## 🛠️ Setup Instructions

1. **Install dependencies**

   ```bash
   pip install -r requirements.txt  # ensure you have langchain, chroma, dotenv, etc.
   ```

2. **Add environment variables**

   Create a `.env` file at the project root (or inside `Gen_AI` if using relative imports) containing your API key(s):

   ```dotenv
   GROQ_API=your_groq_api_key_here
   ```

3. **Prepare documents**

   Place any `.txt` files you want the system to index inside the `docs/` folder. The provided examples already include `google.txt` and `microsoft.txt`.

4. **Run the ingestion pipeline**

   ```bash
   python src/ingestion_pipeline.py
   ```

   This will load documents, split them into chunks, and create (or load) a Chroma vector store at `db/chroma_db`.

5. **Start a chatbot session**

   - Basic version:
     ```bash
     python src/basic_rag.py
     ```

   - History-aware version:
     ```bash
     python src/history_aware_rag.py
     ```

   Follow the on-screen prompts to enter queries. The history-aware script will attempt to rewrite follow-up questions and maintain conversation context.

## 🧩 Component Details

- **`ingestion_pipeline.py`**: Handles document loading, splitting, and vector store creation or loading.
- **`basic_rag.py`**: Demonstrates a simple query workflow without tracking conversation history.
- **`history_aware_rag.py`**: Adds query reformulation and history management using utilities from `utils/text_formatter.py`.
- **`text_formatter.py`**: Contains `format_query` and `format_message` helpers that wrap LLM logic for cleaning user input and building prompt messages.

## 🚀 Extending the System

- Add more documents to `docs/` and rerun the ingestion pipeline.
- Customize the embedding or LLM models in the scripts.
- Adapt `text_formatter` to implement additional prompt engineering or use a different rewriting strategy.

---

Feel free to explore and modify the code—this is a lightweight starting point for building your own RAG-based assistants!
