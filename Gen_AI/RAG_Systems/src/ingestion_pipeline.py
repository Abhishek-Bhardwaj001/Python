import os
from typing import List
from langchain_community.document_loaders import TextLoader, DirectoryLoader
from langchain_text_splitters import CharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
from langchain_core.documents import Document

def load_documents(docs_path: str = "docs") -> List[Document]:
    "Function for loading data files"
    print(f"Loading documents from {docs_path}")

    if not os.path.exists(docs_path):
        raise FileNotFoundError(f"The directory {docs_path} does not exist.")

    loader = DirectoryLoader(
        path=docs_path,
        glob="*.txt",
        loader_cls=TextLoader,
        loader_kwargs={"autodetect_encoding": True}
    )

    documents = loader.load()

    if len(documents) == 0:
        raise FileNotFoundError(f"No documents found in {docs_path}.")

    for i, doc in enumerate(documents[:2]):
        print(f"\nDocument {i+1}:")
        print(f"  Source: {doc.metadata['source']}")
        print(f"  Content length: {len(doc.page_content)}")

    return documents


def split_documents(
    documents: List[Document],
    chunk_size: int = 1000,
    chunk_overlap: int = 200,
) -> List[Document]:
    "Split documents into chunks"
    text_splitter = CharacterTextSplitter(chunk_size=chunk_size, chunk_overlap=chunk_overlap)

    chunks = text_splitter.split_documents(documents)

    if chunks:
        for i, chunk in enumerate(chunks[:5]):
            print(f"Chunk source: {chunk.metadata['source']}")
        print(f"Total chunks: {len(chunks)}")

    return chunks


def create_vector_store(chunks: List[Document], db_directory: str = "db/chroma_db") -> Chroma:
    "Create vector store"
    embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")

    vector_store = Chroma.from_documents(
        documents=chunks,
        embedding=embedding_model,
        persist_directory=db_directory,
        collection_metadata={"hnsw:space": "cosine"}
    )

    print(f"Vector store created at {db_directory}")
    return vector_store


def main() -> Chroma:
    "Main ingestion pipeline"
    docs_path = "docs"
    db_directory = "db/chroma_db"

    if os.path.exists(db_directory):
        print("Vector store already exists, loading...")

        embedding_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
        vector_store = Chroma(
            persist_directory=db_directory,
            embedding_function=embedding_model,
            collection_metadata={"hnsw:space": "cosine"}
        )

        print(f"Loaded existing vector store with {vector_store._collection.count()} chunks")
        return vector_store

    print("DB directory not found, creating new...")
    documents = load_documents(docs_path)
    chunks = split_documents(documents)
    vector_store = create_vector_store(chunks, db_directory)  # ✅ Added missing call
    return vector_store


if __name__ == "__main__":
    main()