from __future__ import annotations

from dataclasses import dataclass
from typing import List, Literal, Optional

from langchain_text_splitters import RecursiveCharacterTextSplitter


SplitterMode = Literal["prose", "code", "markdown", "pdf"]


@dataclass
class SplitterConfig:
    mode: SplitterMode = "prose"
    model_name: str = "text-embedding-3-small"
    chunk_size: int = 500
    chunk_overlap: int = 50
    keep_separator: bool = True
    is_separator_regex: bool = False


class ImprovedTextSplitter:
    def __init__(self, default_model_name: str = "text-embedding-3-small"):
        self.default_model_name = default_model_name

    def split_text(
        self,
        text: str,
        mode: SplitterMode = "prose",
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        model_name: Optional[str] = None,
        keep_separator: bool = True,
    ) -> List[str]:
        if not text or not text.strip():
            return []

        model_name = model_name or self.default_model_name
        config = SplitterConfig(
            mode=mode,
            model_name=model_name,
            chunk_size=chunk_size,
            chunk_overlap=chunk_overlap,
            keep_separator=keep_separator,
        )
        splitter = self._build_splitter(config)
        return splitter.split_text(text)

    def split_documents(
        self,
        documents: List[str],
        mode: SplitterMode = "prose",
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        model_name: Optional[str] = None,
        keep_separator: bool = True,
    ) -> List[List[str]]:
        return [
            self.split_text(
                text=doc,
                mode=mode,
                chunk_size=chunk_size,
                chunk_overlap=chunk_overlap,
                model_name=model_name,
                keep_separator=keep_separator,
            )
            for doc in documents
        ]

    def _build_splitter(self, config: SplitterConfig) -> RecursiveCharacterTextSplitter:
        return RecursiveCharacterTextSplitter.from_tiktoken_encoder(
            model_name=config.model_name,
            chunk_size=config.chunk_size,
            chunk_overlap=config.chunk_overlap,
            separators=self._get_separators(config.mode),
            keep_separator=config.keep_separator,
            is_separator_regex=config.is_separator_regex,
        )

    def _get_separators(self, mode: SplitterMode) -> List[str]:
        separators = {
            "prose": [
                "\n\n",
                "\n",
                ". ",
                "? ",
                "! ",
                "; ",
                ", ",
                " ",
                "",
            ],
            "code": [
                "\nclass ",
                "\ndef ",
                "\nasync def ",
                "\n@",
                "\nif __name__ == \"__main__\":",
                "\n\n",
                "\n",
                " ",
                "",
            ],
            "markdown": [
                "\n# ",
                "\n## ",
                "\n### ",
                "\n#### ",
                "\n##### ",
                "\n###### ",
                "\n```",
                "\n---\n",
                "\n\n",
                "\n",
                ". ",
                " ",
                "",
            ],
            "pdf": [
                "\n\n",
                "\n",
                ". ",
                ": ",
                "; ",
                " ",
                "",
            ],
        }
        return separators[mode]


if __name__ == "__main__":
    sample_prose = (
        "Retrieval-augmented generation combines search with language models. "
        "A good chunking strategy improves retrieval quality. "
        "Different document types require different separators and chunking logic.\n\n"
        "For prose, sentence boundaries usually matter more than strict line boundaries."
    )

    sample_code = (
        "import math\n\n"
        "class VectorStore:\n"
        "    def __init__(self, client):\n"
        "        self.client = client\n\n"
        "    def search(self, query):\n"
        "        return self.client.search(query)\n\n"
        "def rerank(results):\n"
        "    return sorted(results)\n"
    )

    sample_markdown = (
        "# RAG Overview\n\n"
        "RAG has two core pipelines.\n\n"
        "## Indexing\n\n"
        "Chunking, embedding, and storage happen here.\n\n"
        "## Retrieval\n\n"
        "Query rewriting, retrieval, reranking, and generation happen here.\n"
    )

    splitter = ImprovedTextSplitter()

    print("--- Prose ---")
    prose_chunks = splitter.split_text(sample_prose, mode="prose", chunk_size=40, chunk_overlap=8)
    for i, chunk in enumerate(prose_chunks):
        print(f"[{i}] {chunk}\n")

    print("--- Code ---")
    code_chunks = splitter.split_text(sample_code, mode="code", chunk_size=50, chunk_overlap=10)
    for i, chunk in enumerate(code_chunks):
        print(f"[{i}] {chunk}\n")

    print("--- Markdown ---")
    markdown_chunks = splitter.split_text(sample_markdown, mode="markdown", chunk_size=50, chunk_overlap=10)
    for i, chunk in enumerate(markdown_chunks):
        print(f"[{i}] {chunk}\n")
