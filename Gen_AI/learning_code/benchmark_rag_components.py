import os
import time
import json
import math
import shutil
from statistics import mean

from langchain_core.documents import Document
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_chroma import Chroma
from qdrant_client import QdrantClient
from langchain_qdrant import QdrantVectorStore


# --------------------------------------------------
# 1. SAMPLE CHUNKS
# Replace this with your own chunks if needed.
# --------------------------------------------------
RAW_CHUNKS = [
    {
        "content": "Figure 1: The Transformer model architecture. The Transformer uses stacked self-attention and point-wise fully connected layers for both encoder and decoder.",
        "metadata": {"chunk_id": "c1", "page_number": 3, "section": "architecture"},
    },
    {
        "content": "The encoder is composed of a stack of N = 6 identical layers. Each layer has two sub-layers. The first is a multi-head self-attention mechanism.",
        "metadata": {"chunk_id": "c2", "page_number": 3, "section": "encoder"},
    },
    {
        "content": "The second sub-layer in each encoder layer is a position-wise fully connected feed-forward network applied independently to each position.",
        "metadata": {"chunk_id": "c3", "page_number": 3, "section": "encoder"},
    },
    {
        "content": "The decoder is also composed of a stack of N = 6 identical layers. In addition to the two sub-layers in each encoder layer, the decoder inserts a third sub-layer.",
        "metadata": {"chunk_id": "c4", "page_number": 3, "section": "decoder"},
    },
    {
        "content": "This third sub-layer performs multi-head attention over the output of the encoder stack. Masking is used in the decoder self-attention to prevent positions from attending to subsequent positions.",
        "metadata": {"chunk_id": "c5", "page_number": 3, "section": "decoder"},
    },
    {
        "content": "Residual connections are employed around each of the two sub-layers, followed by layer normalization. The output of each sub-layer is LayerNorm(x + Sublayer(x)).",
        "metadata": {"chunk_id": "c6", "page_number": 3, "section": "residuals"},
    },
]


def chunks_to_documents(chunks):
    docs = []
    for item in chunks:
        content = item.get("content", "")
        metadata = item.get("metadata", {})
        if not content or not content.strip():
            continue
        if not isinstance(metadata, dict):
            metadata = {}
        docs.append(Document(page_content=content.strip(), metadata=metadata))
    return docs


documents = chunks_to_documents(RAW_CHUNKS)


# --------------------------------------------------
# 2. EVAL SET
# Edit this to fit your corpus.
# relevant_chunks = set of chunk_id values.
# --------------------------------------------------
EVAL_QUERIES = [
    {
        "query": "How many layers are there in the encoder?",
        "relevant_chunks": {"c2", "c4"},
    },
    {
        "query": "What is the first sub-layer in the encoder?",
        "relevant_chunks": {"c2"},
    },
    {
        "query": "What is the second sub-layer in the encoder?",
        "relevant_chunks": {"c3"},
    },
    {
        "query": "What extra sub-layer does the decoder add?",
        "relevant_chunks": {"c4", "c5"},
    },
    {
        "query": "Why is masking used in decoder self-attention?",
        "relevant_chunks": {"c5"},
    },
    {
        "query": "How are residual connections and normalization used?",
        "relevant_chunks": {"c6"},
    },
]


# --------------------------------------------------
# 3. EMBEDDING MODELS
# Add/remove models as you like.
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
# 4. METRICS
# --------------------------------------------------
def recall_at_k(ranked_ids, relevant_ids, k):
    return 1.0 if any(doc_id in relevant_ids for doc_id in ranked_ids[:k]) else 0.0


def reciprocal_rank(ranked_ids, relevant_ids):
    for idx, doc_id in enumerate(ranked_ids, start=1):
        if doc_id in relevant_ids:
            return 1.0 / idx
    return 0.0


def avg_precision_at_k(ranked_ids, relevant_ids, k):
    hits = 0
    precisions = []
    for i, doc_id in enumerate(ranked_ids[:k], start=1):
        if doc_id in relevant_ids:
            hits += 1
            precisions.append(hits / i)
    if not precisions:
        return 0.0
    denom = min(len(relevant_ids), k)
    return sum(precisions) / denom


def ndcg_at_k(ranked_ids, relevant_ids, k):
    dcg = 0.0
    for i, doc_id in enumerate(ranked_ids[:k], start=1):
        rel = 1 if doc_id in relevant_ids else 0
        dcg += rel / math.log2(i + 1)
    ideal_rels = [1] * min(len(relevant_ids), k)
    idcg = sum(rel / math.log2(i + 1) for i, rel in enumerate(ideal_rels, start=1))
    return dcg / idcg if idcg > 0 else 0.0


# --------------------------------------------------
# 5. HELPERS
# --------------------------------------------------
def detect_dimension(embedding_model):
    vec = embedding_model.embed_documents(["dimension check"])[0]
    return len(vec)


def cleanup_path(path):
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)


def search_ranked_ids(vectorstore, query, k=5):
    pairs = vectorstore.similarity_search_with_score(query, k=k)
    ranked_ids = []
    scored = []
    for doc, score in pairs:
        cid = doc.metadata.get("chunk_id", "")
        ranked_ids.append(cid)
        scored.append({
            "chunk_id": cid,
            "score": float(score),
            "page_content_preview": doc.page_content[:120],
        })
    return ranked_ids, scored


# --------------------------------------------------
# 6. VECTOR STORE FACTORIES
# --------------------------------------------------
def build_faiss(docs, embeddings, run_id):
    start = time.perf_counter()
    vs = FAISS.from_documents(documents=docs, embedding=embeddings)
    build_s = time.perf_counter() - start
    return vs, build_s


def build_chroma(docs, embeddings, run_id):
    path = os.path.join("output", f"chroma_{run_id}")
    cleanup_path(path)
    start = time.perf_counter()
    vs = Chroma.from_documents(
        documents=docs,
        embedding=embeddings,
        collection_name=f"bench_{run_id}",
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
        collection_name=f"bench_{run_id}",
    )
    build_s = time.perf_counter() - start
    return vs, build_s


VECTOR_STORE_BUILDERS = {
    "FAISS": build_faiss,
    "Chroma": build_chroma,
    "Qdrant": build_qdrant,
}


# --------------------------------------------------
# 7. MAIN BENCHMARK
# --------------------------------------------------
def run_benchmark(documents, eval_queries, top_k=5):
    summary_rows = []
    detailed_rows = []

    for emb_name, emb_cfg in EMBEDDING_CONFIGS.items():
        print(f"\n=== Embedding model: {emb_name} ===")
        embeddings = emb_cfg["factory"]()
        dim = detect_dimension(embeddings)
        print(f"Dimension: {dim}")

        for store_name, builder in VECTOR_STORE_BUILDERS.items():
            run_id = f"{emb_name.replace('/', '_').replace(' ', '_')}_{store_name}"
            print(f"  -> Vector store: {store_name}")
            vectorstore, build_time_s = builder(documents, embeddings, run_id)

            r1, r3, r5, mrrs, map5, ndcg5 = [], [], [], [], [], []
            query_latencies = []

            for item in eval_queries:
                query = item["query"]
                relevant = set(item["relevant_chunks"])

                q_start = time.perf_counter()
                ranked_ids, scored = search_ranked_ids(vectorstore, query, k=top_k)
                q_elapsed = time.perf_counter() - q_start
                query_latencies.append(q_elapsed)

                this_r1 = recall_at_k(ranked_ids, relevant, 1)
                this_r3 = recall_at_k(ranked_ids, relevant, 3)
                this_r5 = recall_at_k(ranked_ids, relevant, 5)
                this_mrr = reciprocal_rank(ranked_ids, relevant)
                this_map5 = avg_precision_at_k(ranked_ids, relevant, 5)
                this_ndcg5 = ndcg_at_k(ranked_ids, relevant, 5)

                r1.append(this_r1)
                r3.append(this_r3)
                r5.append(this_r5)
                mrrs.append(this_mrr)
                map5.append(this_map5)
                ndcg5.append(this_ndcg5)

                detailed_rows.append({
                    "embedding_model": emb_name,
                    "vector_store": store_name,
                    "dimension": dim,
                    "query": query,
                    "relevant_chunks": sorted(list(relevant)),
                    "retrieved_ranked_ids": ranked_ids,
                    "top_hits": json.dumps(scored, ensure_ascii=False),
                    "Recall@1": round(this_r1, 4),
                    "Recall@3": round(this_r3, 4),
                    "Recall@5": round(this_r5, 4),
                    "MRR": round(this_mrr, 4),
                    "MAP@5": round(this_map5, 4),
                    "nDCG@5": round(this_ndcg5, 4),
                    "query_latency_ms": round(q_elapsed * 1000, 3),
                })

            summary_rows.append({
                "embedding_model": emb_name,
                "vector_store": store_name,
                "dimension": dim,
                "build_time_ms": round(build_time_s * 1000, 3),
                "avg_query_latency_ms": round(mean(query_latencies) * 1000, 3),
                "Recall@1": round(mean(r1), 4),
                "Recall@3": round(mean(r3), 4),
                "Recall@5": round(mean(r5), 4),
                "MRR": round(mean(mrrs), 4),
                "MAP@5": round(mean(map5), 4),
                "nDCG@5": round(mean(ndcg5), 4),
            })

    return summary_rows, detailed_rows


if __name__ == "__main__":
    print("Running benchmark over chunked documents...")
    print(f"Documents: {len(documents)}")
    print(f"Queries: {len(EVAL_QUERIES)}")

    summary_rows, detailed_rows = run_benchmark(documents, EVAL_QUERIES, top_k=5)

    import csv

    summary_path = os.path.join("output", "benchmark_summary.csv")
    detailed_path = os.path.join("output", "benchmark_detailed.csv")
    report_path = os.path.join("output", "benchmark_report.md")

    with open(summary_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
        writer.writeheader()
        writer.writerows(summary_rows)

    with open(detailed_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(detailed_rows[0].keys()))
        writer.writeheader()
        writer.writerows(detailed_rows)

    sorted_summary = sorted(summary_rows, key=lambda x: (-x["MRR"], -x["Recall@3"], x["avg_query_latency_ms"]))
    best = sorted_summary[0]

    lines = []
    lines.append("# RAG Benchmark Report\n")
    lines.append(f"- Documents benchmarked: {len(documents)}")
    lines.append(f"- Queries benchmarked: {len(EVAL_QUERIES)}")
    lines.append(f"- Best configuration by MRR then Recall@3: **{best['embedding_model']} + {best['vector_store']}**")
    lines.append("")
    lines.append("## Summary")
    lines.append("")
    lines.append("| Embedding | Vector Store | Dim | Build ms | Avg Query ms | Recall@1 | Recall@3 | Recall@5 | MRR | MAP@5 | nDCG@5 |")
    lines.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for row in sorted_summary:
        lines.append(
            f"| {row['embedding_model']} | {row['vector_store']} | {row['dimension']} | {row['build_time_ms']} | {row['avg_query_latency_ms']} | {row['Recall@1']} | {row['Recall@3']} | {row['Recall@5']} | {row['MRR']} | {row['MAP@5']} | {row['nDCG@5']} |"
        )

    lines.append("")
    lines.append("## How to adapt this")
    lines.append("")
    lines.append("1. Replace `RAW_CHUNKS` with your real chunk list.")
    lines.append("2. Replace `EVAL_QUERIES` with your own labeled queries and relevant `chunk_id`s.")
    lines.append("3. Add or remove embedding models in `EMBEDDING_CONFIGS`.")
    lines.append("4. Add more vector stores by following the existing builder pattern.")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("\nTop configurations:")
    for row in sorted_summary[:5]:
        print(row)

    print(f"\nSaved files:\n- {summary_path}\n- {detailed_path}\n- {report_path}\n- output/benchmark_rag_components.py")
