import re
import json
import math
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Tuple
from collections import Counter


@dataclass
class ChunkMetrics:
    chunk_id: int
    char_count: int
    word_count: int
    sentence_count: int
    starts_mid_sentence: bool
    ends_mid_sentence: bool
    header_present: bool
    code_block_count: int
    table_like_line_count: int
    intra_chunk_coherence: float


@dataclass
class CorpusMetrics:
    chunk_count: int
    total_chars: int
    total_words: int
    mean_chunk_chars: float
    median_chunk_chars: float
    min_chunk_chars: int
    max_chunk_chars: int
    std_chunk_chars: float
    mean_chunk_words: float
    sentence_boundary_violation_rate: float
    header_attachment_rate: float
    code_block_fragmentation_rate: float
    table_fragmentation_rate: float
    mean_intra_chunk_coherence: float
    mean_adjacent_chunk_similarity: float
    overlap_redundancy_ratio: float
    duplicate_chunk_ratio: float
    lexical_coverage_ratio: float


class ChunkIngestionEvaluator:
    def __init__(self):
        self.sentence_splitter = re.compile(r'(?<=[.!?])\s+')
        self.word_pattern = re.compile(r"\b\w+\b", re.UNICODE)
        self.header_pattern = re.compile(r'^(#{1,6}\s+.+|[A-Z][A-Za-z0-9\s\-/()]{2,80}:?|\d+(\.\d+)*\s+.+)$', re.MULTILINE)
        self.code_fence_pattern = re.compile(r'```.*?```', re.DOTALL)

    def split_sentences(self, text: str) -> List[str]:
        text = text.strip()
        if not text:
            return []
        return [s.strip() for s in self.sentence_splitter.split(text) if s.strip()]

    def tokenize(self, text: str) -> List[str]:
        return self.word_pattern.findall(text.lower())

    def starts_mid_sentence(self, chunk: str) -> bool:
        chunk = chunk.strip()
        if not chunk:
            return False
        if chunk.startswith(("#", "-", "*", "```", "1.", "2.", "3.")):
            return False
        return bool(re.match(r'^[a-z,(\[]', chunk))

    def ends_mid_sentence(self, chunk: str) -> bool:
        chunk = chunk.strip()
        if not chunk:
            return False
        if chunk.endswith((".", "!", "?", '"', "'", "```", ":")):
            return False
        if chunk.endswith((")", "]")):
            return False
        return True

    def has_header(self, chunk: str) -> bool:
        return bool(self.header_pattern.search(chunk[:300]))

    def count_code_blocks(self, chunk: str) -> int:
        return len(self.code_fence_pattern.findall(chunk))

    def count_table_like_lines(self, chunk: str) -> int:
        lines = [line.strip() for line in chunk.splitlines() if line.strip()]
        table_lines = 0
        for line in lines:
            if line.count("|") >= 2:
                table_lines += 1
            elif re.search(r'\s{2,}', line) and len(line.split()) >= 3:
                table_lines += 1
        return table_lines

    def tf_vector(self, text: str) -> Counter:
        return Counter(self.tokenize(text))

    def cosine_similarity(self, a: Counter, b: Counter) -> float:
        if not a or not b:
            return 0.0
        intersection = set(a) & set(b)
        dot = sum(a[t] * b[t] for t in intersection)
        norm_a = math.sqrt(sum(v * v for v in a.values()))
        norm_b = math.sqrt(sum(v * v for v in b.values()))
        if norm_a == 0 or norm_b == 0:
            return 0.0
        return dot / (norm_a * norm_b)

    def intra_chunk_coherence(self, chunk: str) -> float:
        sentences = self.split_sentences(chunk)
        if len(sentences) <= 1:
            return 1.0 if sentences else 0.0
        vectors = [self.tf_vector(s) for s in sentences]
        sims = []
        for i in range(len(vectors) - 1):
            sims.append(self.cosine_similarity(vectors[i], vectors[i + 1]))
        return sum(sims) / len(sims) if sims else 0.0

    def adjacent_chunk_similarity(self, chunks: List[str]) -> float:
        if len(chunks) <= 1:
            return 0.0
        vectors = [self.tf_vector(c) for c in chunks]
        sims = [self.cosine_similarity(vectors[i], vectors[i + 1]) for i in range(len(vectors) - 1)]
        return sum(sims) / len(sims) if sims else 0.0

    def overlap_redundancy_ratio(self, chunks: List[str]) -> float:
        if len(chunks) <= 1:
            return 0.0
        total_overlap_words = 0
        total_words = sum(len(self.tokenize(c)) for c in chunks)
        for i in range(len(chunks) - 1):
            a = self.tokenize(chunks[i])
            b = self.tokenize(chunks[i + 1])
            suffix_limit = min(50, len(a), len(b))
            max_overlap = 0
            for k in range(1, suffix_limit + 1):
                if a[-k:] == b[:k]:
                    max_overlap = k
            total_overlap_words += max_overlap
        return total_overlap_words / total_words if total_words else 0.0

    def duplicate_chunk_ratio(self, chunks: List[str]) -> float:
        normalized = [re.sub(r'\s+', ' ', c.strip().lower()) for c in chunks if c.strip()]
        if not normalized:
            return 0.0
        counts = Counter(normalized)
        duplicates = sum(count - 1 for count in counts.values() if count > 1)
        return duplicates / len(normalized)

    def lexical_coverage_ratio(self, full_document: str, chunks: List[str]) -> float:
        doc_tokens = set(self.tokenize(full_document))
        chunk_tokens = set()
        for chunk in chunks:
            chunk_tokens.update(self.tokenize(chunk))
        if not doc_tokens:
            return 0.0
        return len(doc_tokens & chunk_tokens) / len(doc_tokens)

    def estimate_fragmentation_rate(self, chunks: List[str], pattern_type: str = "code") -> float:
        if not chunks:
            return 0.0
        fragmented = 0
        relevant = 0
        for i, chunk in enumerate(chunks):
            if pattern_type == "code":
                has_signal = ("def " in chunk or "class " in chunk or "```" in chunk)
                boundary_risk = chunk.strip().endswith(":") or chunk.strip().startswith(("return", "elif", "else", "except"))
            else:
                has_signal = self.count_table_like_lines(chunk) >= 2
                boundary_risk = bool(re.search(r'\|\s*$', chunk.strip())) or bool(re.search(r'^\|', chunk.strip()))
            if has_signal:
                relevant += 1
                if boundary_risk:
                    fragmented += 1
                elif i < len(chunks) - 1:
                    curr = self.tokenize(chunk)
                    nxt = self.tokenize(chunks[i + 1])
                    if curr and nxt and len(set(curr[-10:]) & set(nxt[:10])) > 4:
                        fragmented += 1
        return fragmented / relevant if relevant else 0.0

    def evaluate_chunks(self, full_document: str, chunks: List[str]) -> Dict[str, Any]:
        chunk_metrics: List[ChunkMetrics] = []
        char_counts = []
        word_counts = []

        for idx, chunk in enumerate(chunks):
            sentences = self.split_sentences(chunk)
            cm = ChunkMetrics(
                chunk_id=idx,
                char_count=len(chunk),
                word_count=len(self.tokenize(chunk)),
                sentence_count=len(sentences),
                starts_mid_sentence=self.starts_mid_sentence(chunk),
                ends_mid_sentence=self.ends_mid_sentence(chunk),
                header_present=self.has_header(chunk),
                code_block_count=self.count_code_blocks(chunk),
                table_like_line_count=self.count_table_like_lines(chunk),
                intra_chunk_coherence=self.intra_chunk_coherence(chunk),
            )
            chunk_metrics.append(cm)
            char_counts.append(cm.char_count)
            word_counts.append(cm.word_count)

        chunk_count = len(chunk_metrics)
        total_chars = sum(char_counts)
        total_words = sum(word_counts)
        boundary_violations = sum(1 for c in chunk_metrics if c.starts_mid_sentence or c.ends_mid_sentence)
        header_hits = sum(1 for c in chunk_metrics if c.header_present)
        coherence_values = [c.intra_chunk_coherence for c in chunk_metrics]

        corpus = CorpusMetrics(
            chunk_count=chunk_count,
            total_chars=total_chars,
            total_words=total_words,
            mean_chunk_chars=(sum(char_counts) / chunk_count) if chunk_count else 0.0,
            median_chunk_chars=sorted(char_counts)[chunk_count // 2] if chunk_count else 0.0,
            min_chunk_chars=min(char_counts) if char_counts else 0,
            max_chunk_chars=max(char_counts) if char_counts else 0,
            std_chunk_chars=(self._stddev(char_counts) if char_counts else 0.0),
            mean_chunk_words=(sum(word_counts) / chunk_count) if chunk_count else 0.0,
            sentence_boundary_violation_rate=(boundary_violations / chunk_count) if chunk_count else 0.0,
            header_attachment_rate=(header_hits / chunk_count) if chunk_count else 0.0,
            code_block_fragmentation_rate=self.estimate_fragmentation_rate(chunks, "code"),
            table_fragmentation_rate=self.estimate_fragmentation_rate(chunks, "table"),
            mean_intra_chunk_coherence=(sum(coherence_values) / len(coherence_values)) if coherence_values else 0.0,
            mean_adjacent_chunk_similarity=self.adjacent_chunk_similarity(chunks),
            overlap_redundancy_ratio=self.overlap_redundancy_ratio(chunks),
            duplicate_chunk_ratio=self.duplicate_chunk_ratio(chunks),
            lexical_coverage_ratio=self.lexical_coverage_ratio(full_document, chunks),
        )

        return {
            "corpus_metrics": asdict(corpus),
            "chunk_metrics": [asdict(c) for c in chunk_metrics],
        }

    def _stddev(self, values: List[int]) -> float:
        if len(values) <= 1:
            return 0.0
        mean = sum(values) / len(values)
        variance = sum((x - mean) ** 2 for x in values) / len(values)
        return math.sqrt(variance)


def print_summary(results: Dict[str, Any]) -> None:
    m = results["corpus_metrics"]
    print("=== Corpus-level ingestion metrics ===")
    for k, v in m.items():
        if isinstance(v, float):
            print(f"{k}: {v:.4f}")
        else:
            print(f"{k}: {v}")


def example_usage() -> None:
    full_document = """
    # Authentication
    Authentication is required for all requests. Use an API key in the Authorization header.
    
    # Rate Limits
    The API allows 100 requests per minute. Exceeding the limit returns HTTP 429.
    
    # Python Example
    ```python
    def fetch_users(client):
        response = client.get('/users')
        return response.json()
    ```
    
    # Errors
    Common errors include HTTP 400, 401, 403, and 500.
    """.strip()

    chunks = [
        "# Authentication\nAuthentication is required for all requests. Use an API key in the Authorization header.",
        "# Rate Limits\nThe API allows 100 requests per minute. Exceeding the limit returns HTTP 429.",
        "# Python Example\n```python\ndef fetch_users(client):\n    response = client.get('/users')\n    return response.json()\n```",
        "# Errors\nCommon errors include HTTP 400, 401, 403, and 500.",
    ]

    evaluator = ChunkIngestionEvaluator()
    results = evaluator.evaluate_chunks(full_document, chunks)
    print_summary(results)
    print("\n=== First chunk example ===")
    print(json.dumps(results["chunk_metrics"][0], indent=2))


if __name__ == "__main__":
    example_usage()
