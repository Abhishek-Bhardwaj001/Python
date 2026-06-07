from __future__ import annotations

import re
import math
from dataclasses import dataclass
from typing import Any, Callable, Dict, Iterable, List, Literal, Optional, Sequence, Tuple

from langchain_text_splitters import RecursiveCharacterTextSplitter


Chunk = str
SimilarityFn = Callable[[str, str], float]
EmbeddingFn = Callable[[List[str]], List[List[float]]]


@dataclass
class ChunkingResult:
    """
    Container for the result returned by a chunking strategy.

    Attributes:
        strategy: Name of the chunking strategy used.
        chunks: Final list of text chunks produced by the strategy.
        metadata: Extra information about how the strategy behaved,
            such as detected pages, thresholds, or chunk counts.
    """

    strategy: str
    chunks: List[Chunk]
    metadata: Dict[str, Any]


class ChunkingStrategies:
    """
    Collection of practical chunking strategies for RAG ingestion.

    This class provides multiple chunking approaches that you can apply
    depending on document structure and retrieval requirements.

    Implemented strategies:
    - fixed_size_chunking
    - sliding_window_chunking
    - sentence_based_chunking
    - paragraph_based_chunking
    - recursive_token_chunking
    - markdown_header_chunking
    - page_based_chunking
    - code_aware_chunking
    - semantic_chunking
    - adaptive_chunking
    - proposition_like_chunking
    - late_chunking_proxy
    """

    def __init__(self, default_model_name: str = "text-embedding-3-small"):
        """
        Initialize the chunking strategy helper.

        Args:
            default_model_name: Tokenizer / model name used by token-aware
                splitters created with LangChain's tiktoken-backed helpers.

        Returns:
            None.
        """
        self.default_model_name = default_model_name

    def fixed_size_chunking(
        self,
        text: str,
        chunk_size: int = 500,
        chunk_overlap: int = 50,
    ) -> ChunkingResult:
        """
        Split text into fixed-size character windows.

        This is the simplest chunking method. It ignores document structure
        and slices the text into chunks based only on a fixed character budget.

        Args:
            text: Raw input text to split.
            chunk_size: Maximum number of characters in each chunk.
            chunk_overlap: Number of overlapping characters between
                adjacent chunks.

        Returns:
            ChunkingResult containing the generated chunks and strategy metadata.
        """
        chunks = self._character_window_split(text, chunk_size, chunk_overlap)
        return ChunkingResult(
            strategy="fixed_size_chunking",
            chunks=chunks,
            metadata={
                "chunk_size": chunk_size,
                "chunk_overlap": chunk_overlap,
                "unit": "characters",
            },
        )

    def sliding_window_chunking(
        self,
        text: str,
        window_size: int = 500,
        step_size: int = 400,
    ) -> ChunkingResult:
        """
        Split text using a sliding window over characters.

        Unlike fixed-size chunking expressed in chunk-overlap terms,
        this method uses an explicit window and step size.

        Args:
            text: Raw input text to split.
            window_size: Number of characters per window.
            step_size: Number of characters to move the window forward
                after each chunk. Smaller values create more overlap.

        Returns:
            ChunkingResult with chunks and metadata about the chosen window.
        """
        if not text.strip():
            return ChunkingResult("sliding_window_chunking", [], {"window_size": window_size, "step_size": step_size})

        chunks: List[str] = []
        start = 0

        # Keep moving the window over the text until the full document is covered.
        while start < len(text):
            end = start + window_size
            chunks.append(text[start:end])
            start += step_size

        return ChunkingResult(
            strategy="sliding_window_chunking",
            chunks=chunks,
            metadata={"window_size": window_size, "step_size": step_size, "unit": "characters"},
        )

    def sentence_based_chunking(
        self,
        text: str,
        max_sentences_per_chunk: int = 5,
        sentence_overlap: int = 1,
    ) -> ChunkingResult:
        """
        Group text by sentences instead of raw characters.

        This method first segments the text into sentences and then bundles
        a fixed number of sentences into each chunk.

        Args:
            text: Raw input text to split.
            max_sentences_per_chunk: Maximum number of sentences in each chunk.
            sentence_overlap: Number of trailing sentences from the previous
                chunk to repeat in the next chunk.

        Returns:
            ChunkingResult with sentence-aligned chunks.
        """
        sentences = self._split_sentences(text)
        chunks = self._group_items_with_overlap(sentences, max_sentences_per_chunk, sentence_overlap, join_with=" ")
        return ChunkingResult(
            strategy="sentence_based_chunking",
            chunks=chunks,
            metadata={
                "max_sentences_per_chunk": max_sentences_per_chunk,
                "sentence_overlap": sentence_overlap,
                "sentence_count": len(sentences),
            },
        )

    def paragraph_based_chunking(
        self,
        text: str,
        max_paragraphs_per_chunk: int = 3,
        paragraph_overlap: int = 1,
    ) -> ChunkingResult:
        """
        Group text by paragraph boundaries.

        This strategy works well for prose-heavy documents where paragraphs
        often represent coherent units of meaning.

        Args:
            text: Raw input text to split.
            max_paragraphs_per_chunk: Maximum number of paragraphs per chunk.
            paragraph_overlap: Number of trailing paragraphs from the previous
                chunk to repeat in the next chunk.

        Returns:
            ChunkingResult with paragraph-aligned chunks.
        """
        paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
        chunks = self._group_items_with_overlap(paragraphs, max_paragraphs_per_chunk, paragraph_overlap, join_with="\n\n")
        return ChunkingResult(
            strategy="paragraph_based_chunking",
            chunks=chunks,
            metadata={
                "max_paragraphs_per_chunk": max_paragraphs_per_chunk,
                "paragraph_overlap": paragraph_overlap,
                "paragraph_count": len(paragraphs),
            },
        )

    def recursive_token_chunking(
        self,
        text: str,
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        separators: Optional[List[str]] = None,
        model_name: Optional[str] = None,
    ) -> ChunkingResult:
        """
        Split text recursively using token-aware measurement.

        This is a strong default for many RAG pipelines because it tries to
        preserve larger structures first and only falls back to smaller units
        when necessary.

        Args:
            text: Raw input text to split.
            chunk_size: Maximum tokens per chunk.
            chunk_overlap: Overlap in tokens between adjacent chunks.
            separators: Ordered separator priority list. If not provided,
                a generic prose-friendly default is used.
            model_name: Tokenizer-backed model name used by the splitter.

        Returns:
            ChunkingResult with token-aware recursive chunks.
        """
        splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
            model_name=model_name or self.default_model_name,
            chunk_size=chunk_size,
            chunk_overlap=chunk_overlap,
            separators=separators or ["\n\n", "\n", ". ", "? ", "! ", "; ", " ", ""],
            keep_separator=True,
        )
        chunks = splitter.split_text(text)
        return ChunkingResult(
            strategy="recursive_token_chunking",
            chunks=chunks,
            metadata={
                "chunk_size": chunk_size,
                "chunk_overlap": chunk_overlap,
                "model_name": model_name or self.default_model_name,
                "separator_count": len(separators or ["\n\n", "\n", ". ", "? ", "! ", "; ", " ", ""]),
            },
        )

    def markdown_header_chunking(
        self,
        text: str,
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        model_name: Optional[str] = None,
    ) -> ChunkingResult:
        """
        Split Markdown-like text using heading-aware separators.

        This strategy is useful when the input contains Markdown headings,
        horizontal rules, or fenced code blocks.

        Args:
            text: Markdown or Markdown-like text.
            chunk_size: Maximum tokens per chunk.
            chunk_overlap: Overlap in tokens between chunks.
            model_name: Tokenizer-backed model name used by the splitter.

        Returns:
            ChunkingResult with heading-aware Markdown chunks.
        """
        separators = [
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
        ]
        return self.recursive_token_chunking(text, chunk_size, chunk_overlap, separators, model_name)._replace_strategy("markdown_header_chunking")

    def page_based_chunking(
        self,
        text: str,
        page_break_pattern: str = r"(?:\f|\n\s*---PAGE\s+\d+---\s*\n|\n\s*PAGE\s+\d+\s*\n)",
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        model_name: Optional[str] = None,
    ) -> ChunkingResult:
        """
        Split text by page boundaries first, then recursively chunk each page.

        This strategy is useful for PDFs, slide exports, and page-structured
        reports where each page is a meaningful unit.

        Args:
            text: Extracted text containing explicit page markers.
            page_break_pattern: Regex pattern used to detect page boundaries.
            chunk_size: Maximum tokens per page-level subchunk.
            chunk_overlap: Token overlap between adjacent chunks on the same page.
            model_name: Tokenizer-backed model name used by the splitter.

        Returns:
            ChunkingResult with page-tagged chunks.
        """
        pages = [p.strip() for p in re.split(page_break_pattern, text) if p.strip()]
        chunks: List[str] = []

        # Chunk each page independently so page-local context is preserved.
        for i, page in enumerate(pages, start=1):
            page_result = self.recursive_token_chunking(
                page,
                chunk_size=chunk_size,
                chunk_overlap=chunk_overlap,
                separators=["\n\n", "\n", ". ", ": ", "; ", " ", ""],
                model_name=model_name,
            )
            chunks.extend([f"[PAGE {i}]\n{chunk}" for chunk in page_result.chunks])

        return ChunkingResult(
            strategy="page_based_chunking",
            chunks=chunks,
            metadata={
                "page_count": len(pages),
                "chunk_size": chunk_size,
                "chunk_overlap": chunk_overlap,
            },
        )

    def code_aware_chunking(
        self,
        text: str,
        chunk_size: int = 500,
        chunk_overlap: int = 50,
        model_name: Optional[str] = None,
    ) -> ChunkingResult:
        """
        Split source code with code-aware separators.

        This strategy prioritizes class, function, decorator, and file-level
        boundaries before falling back to more generic splitting.

        Args:
            text: Raw source code text.
            chunk_size: Maximum tokens per chunk.
            chunk_overlap: Overlap in tokens between adjacent chunks.
            model_name: Tokenizer-backed model name used by the splitter.

        Returns:
            ChunkingResult with code-aware chunks.
        """
        separators = [
            "\nclass ",
            "\ndef ",
            "\nasync def ",
            "\n@",
            "\nif __name__ == \"__main__\":",
            "\n\n",
            "\n",
            " ",
            "",
        ]
        return self.recursive_token_chunking(text, chunk_size, chunk_overlap, separators, model_name)._replace_strategy("code_aware_chunking")

    def semantic_chunking(
        self,
        text: str,
        similarity_fn: SimilarityFn,
        max_sentences_per_chunk: int = 8,
        similarity_threshold: float = 0.45,
    ) -> ChunkingResult:
        """
        Split text where semantic similarity drops between adjacent sentences.

        This is a lightweight semantic chunking pattern. It relies on an
        external similarity function so you can plug in your embedding model.

        Args:
            text: Raw input text to split.
            similarity_fn: Function that accepts two sentence strings and
                returns a similarity score where larger means more similar.
            max_sentences_per_chunk: Hard cap on sentences per chunk to avoid
                overly large semantic groups.
            similarity_threshold: If similarity between adjacent sentences falls
                below this threshold, a new chunk is started.

        Returns:
            ChunkingResult with semantically grouped sentence chunks.
        """
        sentences = self._split_sentences(text)
        if not sentences:
            return ChunkingResult("semantic_chunking", [], {"reason": "empty_input"})

        chunks: List[str] = []
        current_chunk: List[str] = [sentences[0]]
        boundaries = 0

        # Start a new chunk whenever topic continuity appears to drop.
        for prev_sentence, next_sentence in zip(sentences, sentences[1:]):
            sim = similarity_fn(prev_sentence, next_sentence)
            should_split = sim < similarity_threshold or len(current_chunk) >= max_sentences_per_chunk
            if should_split:
                chunks.append(" ".join(current_chunk).strip())
                current_chunk = [next_sentence]
                boundaries += 1
            else:
                current_chunk.append(next_sentence)

        if current_chunk:
            chunks.append(" ".join(current_chunk).strip())

        return ChunkingResult(
            strategy="semantic_chunking",
            chunks=chunks,
            metadata={
                "sentence_count": len(sentences),
                "semantic_boundaries": boundaries,
                "max_sentences_per_chunk": max_sentences_per_chunk,
                "similarity_threshold": similarity_threshold,
            },
        )

    def adaptive_chunking(
        self,
        text: str,
        low_complexity_chunk_size: int = 700,
        high_complexity_chunk_size: int = 300,
        complexity_threshold: float = 18.0,
        model_name: Optional[str] = None,
    ) -> ChunkingResult:
        """
        Split text with chunk size adapted to estimated local complexity.

        This strategy uses a simple readability-inspired proxy: average sentence
        length in words inside each paragraph. Dense paragraphs get smaller
        chunk sizes; simpler paragraphs get larger chunk sizes.

        Args:
            text: Raw input text to split.
            low_complexity_chunk_size: Token budget used for simpler regions.
            high_complexity_chunk_size: Token budget used for denser regions.
            complexity_threshold: Average words per sentence above which a
                paragraph is treated as complex.
            model_name: Tokenizer-backed model name used by the splitter.

        Returns:
            ChunkingResult with chunks created from dynamically chosen sizes.
        """
        paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
        chunks: List[str] = []
        complexity_log: List[Dict[str, Any]] = []

        for idx, paragraph in enumerate(paragraphs):
            complexity = self._estimate_complexity(paragraph)
            chunk_size = high_complexity_chunk_size if complexity >= complexity_threshold else low_complexity_chunk_size
            result = self.recursive_token_chunking(
                paragraph,
                chunk_size=chunk_size,
                chunk_overlap=max(1, chunk_size // 10),
                separators=["\n", ". ", "; ", ", ", " ", ""],
                model_name=model_name,
            )
            chunks.extend(result.chunks)
            complexity_log.append({
                "paragraph_index": idx,
                "complexity_score": complexity,
                "chosen_chunk_size": chunk_size,
            })

        return ChunkingResult(
            strategy="adaptive_chunking",
            chunks=chunks,
            metadata={
                "paragraph_count": len(paragraphs),
                "complexity_threshold": complexity_threshold,
                "low_complexity_chunk_size": low_complexity_chunk_size,
                "high_complexity_chunk_size": high_complexity_chunk_size,
                "complexity_log": complexity_log,
            },
        )

    def proposition_like_chunking(
        self,
        text: str,
        max_propositions_per_chunk: int = 6,
        proposition_overlap: int = 1,
    ) -> ChunkingResult:
        """
        Split text into short proposition-like clauses and then group them.

        This is a lightweight approximation of proposition chunking. Instead of
        full information-extraction, it uses punctuation and conjunction cues to
        create smaller fact-like text units.

        Args:
            text: Raw input text to split.
            max_propositions_per_chunk: Maximum proposition-like units in a chunk.
            proposition_overlap: Number of trailing proposition-like units to
                repeat in the next chunk.

        Returns:
            ChunkingResult with grouped proposition-like chunks.
        """
        propositions = self._extract_proposition_like_units(text)
        chunks = self._group_items_with_overlap(
            propositions,
            group_size=max_propositions_per_chunk,
            overlap=proposition_overlap,
            join_with=" ",
        )
        return ChunkingResult(
            strategy="proposition_like_chunking",
            chunks=chunks,
            metadata={
                "proposition_count": len(propositions),
                "max_propositions_per_chunk": max_propositions_per_chunk,
                "proposition_overlap": proposition_overlap,
            },
        )

    def late_chunking_proxy(
        self,
        text: str,
        embedding_fn: EmbeddingFn,
        sentence_group_size: int = 6,
    ) -> ChunkingResult:
        """
        Build sentence groups first, then attach group-level embeddings metadata.

        True late chunking usually embeds the full document with a long-context
        model and derives chunk embeddings afterward. This implementation is a
        practical proxy: it forms sentence groups and computes embeddings for
        those groups after the grouping step.

        Args:
            text: Raw input text to split.
            embedding_fn: Function that accepts a list of strings and returns
                a list of numeric embeddings.
            sentence_group_size: Number of sentences per final group.

        Returns:
            ChunkingResult whose metadata includes group embeddings.
        """
        sentences = self._split_sentences(text)
        groups = self._group_items_with_overlap(sentences, sentence_group_size, overlap=0, join_with=" ")
        embeddings = embedding_fn(groups) if groups else []
        return ChunkingResult(
            strategy="late_chunking_proxy",
            chunks=groups,
            metadata={
                "group_count": len(groups),
                "sentence_group_size": sentence_group_size,
                "embeddings": embeddings,
            },
        )

    def _split_sentences(self, text: str) -> List[str]:
        """
        Split raw text into sentence-like segments.

        Args:
            text: Raw input text.

        Returns:
            List of sentence-like strings.
        """
        if not text.strip():
            return []
        parts = re.split(r"(?<=[.!?])\s+", text.strip())
        return [p.strip() for p in parts if p.strip()]

    def _group_items_with_overlap(
        self,
        items: Sequence[str],
        group_size: int,
        overlap: int,
        join_with: str,
    ) -> List[str]:
        """
        Group a list of items into overlapping windows.

        Args:
            items: Sequence of text units such as sentences or paragraphs.
            group_size: Maximum number of items in each group.
            overlap: Number of trailing items to repeat in the next group.
            join_with: Separator string used to join grouped items.

        Returns:
            List of grouped text chunks.
        """
        if not items:
            return []
        if group_size <= 0:
            raise ValueError("group_size must be greater than 0")
        if overlap >= group_size:
            raise ValueError("overlap must be smaller than group_size")

        chunks: List[str] = []
        step = group_size - overlap

        for start in range(0, len(items), step):
            group = items[start : start + group_size]
            if group:
                chunks.append(join_with.join(group).strip())
        return chunks

    def _character_window_split(self, text: str, chunk_size: int, chunk_overlap: int) -> List[str]:
        """
        Split text into overlapping character windows.

        Args:
            text: Raw input text.
            chunk_size: Maximum characters per chunk.
            chunk_overlap: Number of overlapping characters.

        Returns:
            List of character-window chunks.
        """
        if not text.strip():
            return []
        if chunk_overlap >= chunk_size:
            raise ValueError("chunk_overlap must be smaller than chunk_size")

        chunks: List[str] = []
        step = chunk_size - chunk_overlap

        for start in range(0, len(text), step):
            chunk = text[start : start + chunk_size]
            if chunk:
                chunks.append(chunk)
        return chunks

    def _estimate_complexity(self, paragraph: str) -> float:
        """
        Estimate local paragraph complexity using average words per sentence.

        Args:
            paragraph: Paragraph text.

        Returns:
            Floating-point complexity score. Larger means denser or more complex.
        """
        sentences = self._split_sentences(paragraph)
        if not sentences:
            return 0.0
        word_counts = [len(re.findall(r"\b\w+\b", sentence)) for sentence in sentences]
        return sum(word_counts) / len(word_counts)

    def _extract_proposition_like_units(self, text: str) -> List[str]:
        """
        Convert text into short proposition-like units.

        The method uses punctuation and simple conjunction-based boundaries
        as a lightweight approximation of fact-level segmentation.

        Args:
            text: Raw input text.

        Returns:
            List of short proposition-like units.
        """
        if not text.strip():
            return []

        # First split into sentences, then break long sentences further
        # using conjunctions and punctuation that often separate facts.
        sentences = self._split_sentences(text)
        units: List[str] = []

        for sentence in sentences:
            parts = re.split(r"\s+(?:and|but|because|while|whereas|which|that)\s+|[;:]", sentence)
            units.extend([part.strip(" ,") for part in parts if part.strip(" ,")])

        return units


def lexical_jaccard_similarity(a: str, b: str) -> float:
    """
    Compute a simple lexical Jaccard similarity between two strings.

    This helper is useful as a placeholder similarity function for
    semantic_chunking when an embedding-backed similarity function is not
    yet available.

    Args:
        a: First text string.
        b: Second text string.

    Returns:
        Similarity score in the range [0, 1].
    """
    tokens_a = set(re.findall(r"\b\w+\b", a.lower()))
    tokens_b = set(re.findall(r"\b\w+\b", b.lower()))
    if not tokens_a and not tokens_b:
        return 1.0
    if not tokens_a or not tokens_b:
        return 0.0
    return len(tokens_a & tokens_b) / len(tokens_a | tokens_b)


def dummy_embedding_fn(texts: List[str]) -> List[List[float]]:
    """
    Return simple deterministic numeric vectors for demonstration purposes.

    This is only a placeholder so the module can be executed without an
    external embedding API. Replace it with your real embedding function in
    production.

    Args:
        texts: List of strings to embed.

    Returns:
        List of small numeric vectors, one per input string.
    """
    vectors: List[List[float]] = []
    for text in texts:
        length = len(text)
        word_count = len(re.findall(r"\b\w+\b", text))
        sentence_count = max(1, len(re.split(r"(?<=[.!?])\s+", text.strip())))
        vectors.append([float(length), float(word_count), float(sentence_count)])
    return vectors


def _patch_replace_strategy() -> None:
    """
    Add a small helper method to ChunkingResult for internal reuse.

    This helper avoids duplicating metadata copying logic when one strategy
    builds on top of another and only needs to rename the strategy field.

    Args:
        None.

    Returns:
        None.
    """

    def _replace_strategy(self: ChunkingResult, strategy_name: str) -> ChunkingResult:
        return ChunkingResult(strategy=strategy_name, chunks=self.chunks, metadata=self.metadata)

    setattr(ChunkingResult, "_replace_strategy", _replace_strategy)


_patch_replace_strategy()


if __name__ == "__main__":
    sample_text = (
        "# RAG Overview\n\n"
        "Retrieval-augmented generation combines retrieval and generation. "
        "It improves grounding by bringing relevant external context into the prompt. "
        "Chunking is a core ingestion step.\n\n"
        "## Chunking\n\n"
        "Fixed-size chunking is simple but can split ideas badly. "
        "Semantic chunking groups sentences based on meaning. "
        "Page-based chunking works well for PDFs and slide decks.\n\n"
        "## Code Example\n\n"
        "def search(query):\n    return vector_db.search(query)\n"
    )

    strategies = ChunkingStrategies()

    demo_results = [
        strategies.fixed_size_chunking(sample_text, chunk_size=120, chunk_overlap=20),
        strategies.sentence_based_chunking(sample_text, max_sentences_per_chunk=2, sentence_overlap=1),
        strategies.recursive_token_chunking(sample_text, chunk_size=80, chunk_overlap=10),
        strategies.markdown_header_chunking(sample_text, chunk_size=80, chunk_overlap=10),
        strategies.semantic_chunking(sample_text, similarity_fn=lexical_jaccard_similarity, max_sentences_per_chunk=2, similarity_threshold=0.10),
        strategies.adaptive_chunking(sample_text, low_complexity_chunk_size=90, high_complexity_chunk_size=50, complexity_threshold=10.0),
        strategies.proposition_like_chunking(sample_text, max_propositions_per_chunk=3, proposition_overlap=1),
        strategies.late_chunking_proxy(sample_text, embedding_fn=dummy_embedding_fn, sentence_group_size=2),
    ]

    for result in demo_results:
        print(f"\n=== {result.strategy} ===")
        print(f"chunk_count: {len(result.chunks)}")
        print(f"metadata: {result.metadata}")
        for i, chunk in enumerate(result.chunks[:3]):
            print(f"[{i}] {chunk}\n")
