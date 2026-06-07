
# Embedding Model Decision Table: Tradeoffs & When to Switch

## 📊 Decision Table: Embedding Model Tradeoffs

| Dimension | OpenAI 3-small | OpenAI 3-large | Cohere light-v3.0 | Cohere v3.0 | Voyage-3 | MiniLM-L6-v2 | MPNet-base-v2 | BGE-base-v1.5 |
|-----------|---------------|----------------|-------------------|-------------|----------|--------------|---------------|---------------|
| **Type** | API | API | API | API | API | Self-host | Self-host | Self-host |
| **Cost** | $0.05/1M tokens | $0.13/1M tokens | $0.15/1M tokens | $0.35/1M tokens | ~$0.10/1M tokens | $0 (infra) | $0 (infra) | $0 (infra) |
| **Latency** | 50–150 ms | 50–150 ms | 50–100 ms | 50–100 ms | 30–80 ms | 2–5 ms GPU | 5–15 ms GPU | 10–30 ms GPU |
| **Vector Dim** | 1536 (↓256) | 3072 (↓1024) | 1024 | 1024 | 1024–4096 | 384 | 768 | 768 |
| **Storage (1M docs)** | 6.1 MB | 12.3 MB | 4.1 MB | 4.1 MB | 4.1 MB | 1.5 MB | 3.1 MB | 3.1 MB |
| **MTEB Score** | 78–80 | 80–82 | 75–77 | 76–78 | 79–81 | 70–73 | 75–78 | 77–80 |
| **Multilingual** | Limited | Limited | 30+ langs | 30+ langs | Limited | Limited | Multi-cased | 100+ langs |
| **Max Context** | 8K | 8K | 4K | 4K | **32K** | 8K | 8K | 8K |
| **Fine-tuning** | ❌ | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ | ✅ |
| **Data Residency** | Vendor | Vendor | Vendor | Vendor | Vendor | **Full control** | **Full control** | **Full control** |
| **HIPAA/FedRAMP** | Requires DPA | Requires DPA | Some options | Some options | Limited | **Yes (self-host)** | **Yes (self-host)** | **Yes (self-host)** |
| **Security Risk** | **Medium-High** | **Medium-High** | Medium | Medium | Medium | **Low** | **Low** | **Low** |
| **Best For** | Low-cost RAG | High-accuracy RAG | Fast multilingual | General English | Long docs | Edge/CPU | Balanced RAG | High-accuracy RAG |

---

## 🔄 When to Switch from One Model to Another

### 1. API → Self-Hosted (OpenAI/Cohere/Voyage → Hugging Face)

**Switch when:**

| Trigger Condition | Threshold | Why Switch |
|-------------------|-----------|------------|
| **Monthly API cost** | > $500–$1000/month | Self-hosted eliminates token costs; only infra cost |
| **Data residency required** | HIPAA, GDPR, FedRAMP strict | No data egress; full compliance control |
| **Latency budget** | < 50 ms critical | Self-hosted GPU: 5–20 ms vs API 50–150 ms |
| **Need fine-tuning** | Domain-specific terminology | Customize model on your data |
| **Scalability** | > 1M documents/day | Avoid API rate limits, vendor lock-in |

**Recommended self-hosted upgrade path:**
```
MiniLM (CPU) → MPNet (CPU/GPU) → BGE (GPU) → e5-large-instruct (GPU)
```

---

### 2. MiniLM → MPNet/BGE (Low accuracy → High accuracy)

**Switch when:**

| Metric | Current Threshold | Target |
|--------|-------------------|--------|
| **Recall@1** | < 0.65 | ≥ 0.75 |
| **Recall@3** | < 0.75 | ≥ 0.85 |
| **MRR** | < 0.55 | ≥ 0.70 |
| **User satisfaction** | < 70% relevant | > 85% relevant |

**Use case:** Hallucination unacceptable (financial, medical, legal RAG).

---

### 3. MPNet/BGE → OpenAI 3-Large or Voyage-3 (Accuracy-critical)

**Switch when:**

| Business Impact | Condition |
|-----------------|-----------|
| **Revenue impact** | 1% accuracy gain = significant revenue difference |
| **Risk impact** | Wrong retrieval = regulatory violation, lawsuit, patient harm |
| **MTEB score** | Need top 1% (≥ 80 on MTEB leaderboard) |
| **Budget** | Can afford $0.10–0.15/1M tokens |

**Tradeoff:** You gain ~2–4% accuracy but lose data residency and incur ongoing costs.

---

### 4. English-only → Multilingual (Cohere v3, e5-large-instruct)

**Switch when:**

| Condition | Threshold |
|-----------|-----------|
| **Non-English queries** | > 20% of total queries |
| **Regulatory requirement** | Must support multiple languages |
| **User base expansion** | Entering new markets (Asia, Africa, LatAm) |

**Recommended:** `Cohere embed-multilingual-v3.0` or `multilingual-e5-large-instruct`

---

### 5. General → Domain-Specific Fine-tuned Model

**Switch when:**

| Condition | Threshold |
|-----------|-----------|
| **Domain accuracy** | General model < 70% on domain queries |
| **Specialized terminology** | Healthcare, legal, finance, code |
| **Resources** | Budget for fine-tuning + model maintenance |

**Approach:** Start with BGE-base or MPNet, fine-tune on your domain data.

---

## 🔒 Security Breach Risk Analysis

### Risk Levels by Provider Type

| Provider Type | Risk Level | Reasons | Mitigations |
|---------------|-----------|---------|-------------|
| **OpenAI** | **Medium-High** | Data leaves environment; provider may retain for training; breach history | Sign DPA; review ToS; avoid PII |
| **Cohere/Voyage/Google** | **Medium** | Data leaves environment; smaller providers = less breach history but still risk | Sign DPA; avoid sensitive data |
| **AWS Bedrock** | **Medium** | AWS has strong security but still managed service; HIPAA-ready | Use AWS KMS; VPC endpoints |
| **Self-hosted HF** | **Low** | No data egress; only your infra security | Secure GPU server; network security; encryption at rest |

---

### Embedding Inversion Attack Risk (ALL MODELS)

**Research findings:** [web:86][web:89]

| Attack Type | Risk | Recovery Rate | Affected Models |
|-------------|------|---------------|-----------------|
| **Text reconstruction** | Medium | Up to **92%** of original text | All models, larger models = higher risk |
| **PII extraction** | High | Names, passwords, company, health conditions | BGE-large, e5-large, 3-large |
| **Model inversion** | Medium | Sensitive attributes inferred | Larger models (more capacity) |

**Key insight:** Large embedding models (BGE-large, e5-large, 3-large) capture more information → higher privacy risk [web:89].

---

### Mitigation Strategies

| Strategy | Security Gain | Utility Cost |
|----------|---------------|--------------|
| **Self-hosting** | No data egress → eliminates provider breach risk | None |
| **Perturbation (noise)** | Reduces inversion attack success | 2–5 accuracy drop |
| **Encryption (homomorphic)** | Maximum protection | 10–100×² compute cost |
| **Data filtering** | Remove PII before embedding | Minimal |
| **Smaller models** | Less information captured | 3–5% accuracy drop |

**Recommendation:** For sensitive data (PII, source code, IP), **self-hosted models are mandatory** [web:86].

---

## 🎯 Recommendation for Your Use Case

Given your interests (semantic search, RAG, vector databases, IR metrics, evaluation):

### Starting Point

| Scenario | Recommended Model | Why |
|----------|-------------------|-----|
| **Production + cost-sensitive** | OpenAI text-embedding-3-small | Best cost/accuracy balance, easy to use |
| **Production + accuracy-critical** | BGE-base-en-v1.5 (self-hosted) | High accuracy, no data egress, fine-tunable |
| **Multilingual** | Cohere embed-multilingual-v3.0 or e5-large-instruct | 100+ languages, strong performance |
| **Edge/CPU-only** | MiniLM-L6-v2 | Fastest on CPU, smallest storage |
| **Long documents** | Voyage-3 | 32K context, top retrieval quality |

### When to Re-evaluate

1. **Benchmark on YOUR data** (not MTEB)
   - If Recall@3 < 0.75 → switch to MPNet or BGE
   - If MRR < 0.6 → switch to 3-large or Voyage-3

2. **Monitor costs**
   - If API costs > $500/month → migrate to self-hosted BGE/MPNet

3. **Compliance requirements change**
   - If HIPAA/FedRAMP needed → migrate to self-hosted

4. **Latency becomes critical**
   - If 50–150 ms API latency hurts UX → self-hosted GPU (5–20 ms)

---

## 📁 Files Created

- `embedding_model_decision_table.csv` — Full CSV with all models and dimensions
- `embedding_selection_summary.txt` — Detailed switching guide

