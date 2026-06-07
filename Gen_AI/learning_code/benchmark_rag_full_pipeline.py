import os
import csv
import json
import math
import time
import shutil
from statistics import mean

from langchain_core.documents import Document
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_chroma import Chroma
from langchain_qdrant import QdrantVectorStore
from qdrant_client import QdrantClient
from langchain_text_splitters import RecursiveCharacterTextSplitter, CharacterTextSplitter, MarkdownTextSplitter

try:
    from langchain_experimental.text_splitter import SemanticChunker
    SEMANTIC_CHUNKER_AVAILABLE = True
except Exception:
    SEMANTIC_CHUNKER_AVAILABLE = False


# --------------------------------------------------
# 1. SOURCE DOCUMENTS
# Replace this with your actual page-level docs or raw text.
# --------------------------------------------------
SOURCE_DOCS = [
    Document(
        page_content=(
            "Figure 1: The Transformer model architecture. The Transformer uses stacked self-attention "
            "and point-wise fully connected layers for both encoder and decoder. "
            "The encoder is composed of a stack of N = 6 identical layers. "
            "Each layer has two sub-layers. The first is a multi-head self-attention mechanism. "
            "The second is a position-wise fully connected feed-forward network. "
            "Residual connections are employed around each sub-layer, followed by layer normalization."
        ),
        metadata={"source": "attention_is_all_you_need.pdf", "page_number": 3, "doc_id": "p3"},
    ),
    Document(
        page_content=(
            "The decoder is also composed of a stack of N = 6 identical layers. "
            "In addition to the two sub-layers in each encoder layer, the decoder inserts a third sub-layer "
            "that performs multi-head attention over the output of the encoder stack. "
            "To preserve the auto-regressive property, masking is applied in the decoder self-attention so that "
            "positions cannot attend to subsequent positions."
        ),
        metadata={"source": "attention_is_all_you_need.pdf", "page_number": 3, "doc_id": "p3b"},
    ),
]


# --------------------------------------------------
# 2. EVAL QUERIES
# Note: relevant_match_terms are substrings expected in at least one relevant chunk.
# This avoids relabeling chunk IDs for every chunking strategy.
# --------------------------------------------------
EVAL_QUERIES = [
    {
        "query": "How many layers are in the encoder?",
        "relevant_match_terms": ["stack of N = 6 identical layers"],
    },
    {
        "query": "What is the first sub-layer in the encoder?",
        "relevant_match_terms": ["first is a multi-head self-attention mechanism"],
    },
    {
        "query": "What is the second sub-layer in the encoder?",
        "relevant_match_terms": ["position-wise fully connected feed-forward network"],
    },
    {
        "query": "What extra sub-layer is added in the decoder?",
        "relevant_match_terms": ["decoder inserts a third sub-layer", "multi-head attention over the output of the encoder stack"],
    },
    {
        "query": "Why is masking used in decoder self-attention?",
        "relevant_match_terms": ["positions cannot attend to subsequent positions", "masking is applied in the decoder self-attention"],
    },
]


# --------------------------------------------------
# 3. EMBEDDING MODELS
# --------------------------------------------------
EMBEDDING_CONFIGS = {
    "MiniLM-L6-v2": {
        "factory": lambda: HuggingFaceEmbeddings(
            model_name="sentence-transformers/all-MiniLM-L6-v2",
            model_kwargs={"device": "cpu"},
            encode_kwargs={"normalize_embeddings": True},
        )
    },
    "MPNet-base-v2": {
        "factory": lambda: HuggingFaceEmbeddings(
            model_name="sentence-transformers/all-mpnet-base-v2",
            model_kwargs={"device": "cpu"},
            encode_kwargs={"normalize_embeddings": True},
        )
    },
    "BGE-base-en-v1.5": {
        "factory": lambda: HuggingFaceEmbeddings(
            model_name="BAAI/bge-base-en-v1.5",
            model_kwargs={"device": "cpu"},
            encode_kwargs={"normalize_embeddings": True},
        )
    },
}


# --------------------------------------------------
# 4. CHUNKING STRATEGIES
# --------------------------------------------------
def assign_chunk_ids(docs, strategy_name):
    out = []
    for i, doc in enumerate(docs, start=1):
        meta = dict(doc.metadata)
        meta["chunk_id"] = f"{strategy_name}_c{i}"
        out.append(Document(page_content=doc.page_content, metadata=meta))
    return out


def chunk_recursive(source_docs):
    splitter = RecursiveCharacterTextSplitter(chunk_size=220, chunk_overlap=40)
    docs = splitter.split_documents(source_docs)
    return assign_chunk_ids(docs, "recursive")


def chunk_token(source_docs):
    splitter = CharacterTextSplitter.from_tiktoken_encoder(
        encoding_name="cl100k_base", chunk_size=120, chunk_overlap=20
    )
    docs = splitter.split_documents(source_docs)
    return assign_chunk_ids(docs, "token")


def chunk_markdown_like(source_docs):
    splitter = MarkdownTextSplitter(chunk_size=220, chunk_overlap=30)
    docs = splitter.split_documents(source_docs)
    return assign_chunk_ids(docs, "markdown")


def chunk_semantic(source_docs):
    if not SEMANTIC_CHUNKER_AVAILABLE:
        raise RuntimeError("SemanticChunker unavailable. Install langchain-experimental.")
    helper_embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2",
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )
    splitter = SemanticChunker(helper_embeddings)
    docs = splitter.split_documents(source_docs)
    return assign_chunk_ids(docs, "semantic")


CHUNKING_BUILDERS = {
    "recursive": chunk_recursive,
    "token": chunk_token,
    "markdown": chunk_markdown_like,
}

if SEMANTIC_CHUNKER_AVAILABLE:
    CHUNKING_BUILDERS["semantic"] = chunk_semantic


# --------------------------------------------------
# 5. VECTOR STORES
# --------------------------------------------------
def cleanup_path(path):
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)


def build_faiss(docs, embeddings, run_id):
    start = time.perf_counter()
    vs = FAISS.from_documents(documents=docs, embedding=embeddings)
    build_s = time.perf_counter() - start
    return vs, build_s


def build_chroma(docs, embeddings, run_id):
    path = os.path.join("output", f"full_chroma_{run_id}")
    cleanup_path(path)
    start = time.perf_counter()
    vs = Chroma.from_documents(
        documents=docs,
        embedding=embeddings,
        collection_name=f"full_{run_id}",
        persist_directory=path,
    )
    build_s = time.perf_counter() - start
    return vs, build_s


def build_qdrant(docs, embeddings, run_id):
    start = time.perf_counter()
    client = QdrantClient(":memory:")
    vs = QdrantVectorStore.from_documents(
        documents=docs,
        embedding=embeddings,
        client=client,
        collection_name=f"full_{run_id}",
    )
    build_s = time.perf_counter() - start
    return vs, build_s


VECTOR_STORE_BUILDERS = {
    "FAISS": build_faiss,
    "Chroma": build_chroma,
    "Qdrant": build_qdrant,
}


# --------------------------------------------------
# 6. METRICS
# --------------------------------------------------
def detect_dimension(embedding_model):
    vec = embedding_model.embed_documents(["dimension check"])[0]
    return len(vec)


def is_relevant(doc, relevant_terms):
    text = doc.page_content.lower()
    return any(term.lower() in text for term in relevant_terms)


def recall_at_k(relevance_flags, k):
    return 1.0 if any(relevance_flags[:k]) else 0.0


def reciprocal_rank(relevance_flags):
    for idx, rel in enumerate(relevance_flags, start=1):
        if rel:
            return 1.0 / idx
    return 0.0


def avg_precision_at_k(relevance_flags, k):
    hits = 0
    precisions = []
    total_relevant_found = sum(relevance_flags[:k])
    for i, rel in enumerate(relevance_flags[:k], start=1):
        if rel:
            hits += 1
            precisions.append(hits / i)
    if not precisions:
        return 0.0
    denom = max(1, total_relevant_found)
    return sum(precisions) / denom


def ndcg_at_k(relevance_flags, k):
    dcg = 0.0
    for i, rel in enumerate(relevance_flags[:k], start=1):
        dcg += (1 if rel else 0) / math.log2(i + 1)
    ideal_rels = sorted([1 if r else 0 for r in relevance_flags], reverse=True)[:k]
    idcg = sum(rel / math.log2(i + 1) for i, rel in enumerate(ideal_rels, start=1))
    return dcg / idcg if idcg > 0 else 0.0


def search_docs(vectorstore, query, k=5):
    pairs = vectorstore.similarity_search_with_score(query, k=k)
    docs = []
    for doc, score in pairs:
        docs.append({
            "chunk_id": doc.metadata.get("chunk_id", ""),
            "score": float(score),
            "content": doc.page_content,
            "metadata": doc.metadata,
        })
    return docs


# --------------------------------------------------
# 7. BENCHMARK
# --------------------------------------------------
def run_full_benchmark():
    summary_rows = []
    detail_rows = []

    chunk_cache = {}
    for chunk_name, chunk_builder in CHUNKING_BUILDERS.items():
        chunk_start = time.perf_counter()
        chunked_docs = chunk_builder(SOURCE_DOCS)
        chunk_time_ms = (time.perf_counter() - chunk_start) * 1000
        chunk_cache[chunk_name] = {
            "docs": chunked_docs,
            "chunk_time_ms": chunk_time_ms,
        }
        print(f"Chunk strategy {chunk_name}: {len(chunked_docs)} chunks")

    for chunk_name, chunk_info in chunk_cache.items():
        docs = chunk_info["docs"]
        chunk_time_ms = round(chunk_info["chunk_time_ms"], 3)

        for emb_name, emb_cfg in EMBEDDING_CONFIGS.items():
            embeddings = emb_cfg["factory"]()
            dim = detect_dimension(embeddings)

            for store_name, store_builder in VECTOR_STORE_BUILDERS.items():
                run_id = f"{chunk_name}_{emb_name.replace('/', '_').replace(' ', '_')}_{store_name}"
                print(f"Running: {chunk_name} | {emb_name} | {store_name}")
                vectorstore, build_s = store_builder(docs, embeddings, run_id)

                r1s, r3s, r5s, mrrs, map5s, ndcg5s, latencies = [], [], [], [], [], [], []

                for item in EVAL_QUERIES:
                    q = item["query"]
                    relevant_terms = item["relevant_match_terms"]

                    t0 = time.perf_counter()
                    hits = search_docs(vectorstore, q, k=5)
                    latency_ms = (time.perf_counter() - t0) * 1000
                    latencies.append(latency_ms)

                    relevance_flags = [any(term.lower() in hit["content"].lower() for term in relevant_terms) for hit in hits]

                    r1 = recall_at_k(relevance_flags, 1)
                    r3 = recall_at_k(relevance_flags, 3)
                    r5 = recall_at_k(relevance_flags, 5)
                    mrr = reciprocal_rank(relevance_flags)
                    map5 = avg_precision_at_k(relevance_flags, 5)
                    ndcg5 = ndcg_at_k(relevance_flags, 5)

                    r1s.append(r1)
                    r3s.append(r3)
                    r5s.append(r5)
                    mrrs.append(mrr)
                    map5s.append(map5)
                    ndcg5s.append(ndcg5)

                    detail_rows.append({
                        "chunk_strategy": chunk_name,
                        "embedding_model": emb_name,
                        "vector_store": store_name,
                        "dimension": dim,
                        "query": q,
                        "relevant_terms": json.dumps(relevant_terms, ensure_ascii=False),
                        "top_hits": json.dumps(hits, ensure_ascii=False),
                        "Recall@1": round(r1, 4),
                        "Recall@3": round(r3, 4),
                        "Recall@5": round(r5, 4),
                        "MRR": round(mrr, 4),
                        "MAP@5": round(map5, 4),
                        "nDCG@5": round(ndcg5, 4),
                        "query_latency_ms": round(latency_ms, 3),
                    })

                summary_rows.append({
                    "chunk_strategy": chunk_name,
                    "num_chunks": len(docs),
                    "chunk_time_ms": chunk_time_ms,
                    "embedding_model": emb_name,
                    "vector_store": store_name,
                    "dimension": dim,
                    "build_time_ms": round(build_s * 1000, 3),
                    "avg_query_latency_ms": round(mean(latencies), 3),
                    "Recall@1": round(mean(r1s), 4),
                    "Recall@3": round(mean(r3s), 4),
                    "Recall@5": round(mean(r5s), 4),
                    "MRR": round(mean(mrrs), 4),
                    "MAP@5": round(mean(map5s), 4),
                    "nDCG@5": round(mean(ndcg5s), 4),
                })

    return summary_rows, detail_rows


if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)

    summary_rows, detail_rows = run_full_benchmark()

    summary_rows = sorted(summary_rows, key=lambda x: (-x["MRR"], -x["Recall@3"], x["avg_query_latency_ms"]))

    summary_path = "output/full_pipeline_benchmark_summary.csv"
    detail_path = "output/full_pipeline_benchmark_detailed.csv"
    report_path = "output/full_pipeline_benchmark_report.md"

    with open(summary_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
        writer.writeheader()
        writer.writerows(summary_rows)

    with open(detail_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(detail_rows[0].keys()))
        writer.writeheader()
        writer.writerows(detail_rows)

    best = summary_rows[0]
    lines = []
    lines.append("# Full Pipeline Benchmark Report\n")
    lines.append(f"- Chunk strategies tested: {len(CHUNKING_BUILDERS)}")
    lines.append(f"- Embedding models tested: {len(EMBEDDING_CONFIGS)}")
    lines.append(f"- Vector stores tested: {len(VECTOR_STORE_BUILDERS)}")
    lines.append(f"- Best config by MRR then Recall@3: **{best['chunk_strategy']} + {best['embedding_model']} + {best['vector_store']}**")
    lines.append("")
    lines.append("## Summary")
    lines.append("")
    lines.append("| Chunking | # Chunks | Embedding | Vector Store | Dim | Chunk ms | Build ms | Avg Query ms | Recall@1 | Recall@3 | Recall@5 | MRR | MAP@5 | nDCG@5 |")
    lines.append("|---|---:|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for row in summary_rows:
        lines.append(
            f"| {row['chunk_strategy']} | {row['num_chunks']} | {row['embedding_model']} | {row['vector_store']} | {row['dimension']} | {row['chunk_time_ms']} | {row['build_time_ms']} | {row['avg_query_latency_ms']} | {row['Recall@1']} | {row['Recall@3']} | {row['Recall@5']} | {row['MRR']} | {row['MAP@5']} | {row['nDCG@5']} |"
        )

    lines.append("")
    lines.append("## Notes")
    lines.append("")
    lines.append("1. Replace `SOURCE_DOCS` with your real source documents before chunking.")
    lines.append("2. Replace `EVAL_QUERIES` with your labeled evaluation questions.")
    lines.append("3. If `SemanticChunker` is installed, semantic chunking will be included automatically.")
    lines.append("4. This benchmark helps reduce the risk of locally optimal but globally suboptimal choices by testing chunking, embeddings, and vector stores jointly.")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("Saved files:")
    print(summary_path)
    print(detail_path)
    print(report_path)
    print("output/benchmark_rag_full_pipeline.py")
