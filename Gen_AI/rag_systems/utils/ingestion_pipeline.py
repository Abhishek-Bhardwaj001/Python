import os
from typing import List, Callable
from langchain_community.document_loaders import TextLoader, DirectoryLoader
from langchain_text_splitters import CharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
from langchain_core.documents import Document
from data_ingestion_helpers import load_documents,split_documents, create_vector_store


def ingest_documents(load_documents:Callable[[str],dict], #Callable[[arguments],return_type]
         split_documents:Callable[[object],dict],
         create_vector_store:Callable[[list,str],Chroma],
         docs_path:str="docs",
         database_name:str="chroma_db",
         verbose:bool=False) -> Chroma:
    "Main ingestion pipeline"
    if verbose:
        print("Data Load initiated")
    db_directory = f"db/{database_name}"

    if os.path.exists(db_directory):
        if verbose:
            print("Vector store already exists, loading...")

        embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
        vector_store = Chroma(
            persist_directory=db_directory,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space": "cosine"}
        )
        if verbose:
            print(f"Loaded existing vector store with {vector_store._collection.count()} chunks")
        return vector_store
    if verbose:
        print("DB directory not found, creating new...")
    documents = load_documents(docs_path)
    chunks = split_documents(documents)
    vector_store = create_vector_store(chunks, db_directory)  # ✅ Added missing call
    return chunks



if __name__ == "__main__":
    docs_path="docs"
    documents = load_documents(docs_path)
    chunks = split_documents(documents)
    vector_store = create_vector_store(chunks,db_directory)
    for i in chunks[:1]:
        print(f"Document :\n {i}")

