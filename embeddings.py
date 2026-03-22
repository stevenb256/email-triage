"""
embeddings.py — Semantic search via local embeddings (fastembed + numpy).
"""
import json
import threading
import numpy as np
from db import get_db

_model = None
_model_lock = threading.Lock()
EMBEDDING_DIM = 384  # BAAI/bge-small-en-v1.5


def _get_model():
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                from fastembed import TextEmbedding
                _model = TextEmbedding("BAAI/bge-small-en-v1.5")
    return _model


def _embed_texts(texts: list[str]) -> list[np.ndarray]:
    model = _get_model()
    return list(model.embed(texts))


def _email_text(row) -> str:
    parts = []
    subj = (row["subject"] or "").strip()
    if subj:
        parts.append(subj)
    frm = (row["from_name"] or row["from_address"] or "").strip()
    if frm:
        parts.append(f"from {frm}")
    body = (row["body_preview"] or "").strip()
    if body:
        parts.append(body)
    return " | ".join(parts) or "empty"


def init_embeddings_table():
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS email_embeddings (
            email_id    TEXT PRIMARY KEY,
            embedding   BLOB NOT NULL
        );
    """)
    db.commit()


def embed_missing(batch_size=64):
    """Generate embeddings for all emails that don't have one yet."""
    db = get_db()
    rows = db.execute("""
        SELECT e.id, e.subject, e.from_name, e.from_address, e.body_preview
        FROM emails e
        LEFT JOIN email_embeddings ee ON e.id = ee.email_id
        WHERE ee.email_id IS NULL
        LIMIT 500
    """).fetchall()
    if not rows:
        return 0
    texts = [_email_text(r) for r in rows]
    ids = [r["id"] for r in rows]
    total = 0
    for i in range(0, len(texts), batch_size):
        batch_texts = texts[i:i + batch_size]
        batch_ids = ids[i:i + batch_size]
        embeddings = _embed_texts(batch_texts)
        db.executemany(
            "INSERT OR IGNORE INTO email_embeddings(email_id, embedding) VALUES(?, ?)",
            [(eid, emb.astype(np.float32).tobytes()) for eid, emb in zip(batch_ids, embeddings)],
        )
        db.commit()
        total += len(batch_texts)
    return total


def semantic_search(query: str, limit: int = 50) -> list[dict]:
    """Search emails by semantic similarity. Returns list of email dicts with scores."""
    db = get_db()
    query_emb = _embed_texts([query])[0].astype(np.float32)
    query_emb = query_emb / (np.linalg.norm(query_emb) + 1e-9)

    # Load all embeddings into memory (fast for <50K emails)
    rows = db.execute("""
        SELECT ee.email_id, ee.embedding,
               e.id, e.subject, e.from_name, e.from_address,
               e.received_date_time, e.body_preview, e.folder,
               e.is_read, e.conversation_key
        FROM email_embeddings ee
        JOIN emails e ON e.id = ee.email_id
    """).fetchall()

    if not rows:
        return []

    # Build matrix and compute cosine similarities
    n = len(rows)
    mat = np.zeros((n, EMBEDDING_DIM), dtype=np.float32)
    for i, r in enumerate(rows):
        mat[i] = np.frombuffer(r["embedding"], dtype=np.float32)
    norms = np.linalg.norm(mat, axis=1, keepdims=True)
    norms[norms == 0] = 1e-9
    mat /= norms
    scores = mat @ query_emb

    # Get top results
    top_idx = np.argsort(scores)[::-1][:limit]
    results = []
    for idx in top_idx:
        score = float(scores[idx])
        if score < 0.15:  # relevance threshold
            break
        r = rows[idx]
        results.append({
            "id": r["id"],
            "subject": r["subject"],
            "from_name": r["from_name"],
            "from_address": r["from_address"],
            "received_date_time": r["received_date_time"],
            "body_preview": r["body_preview"],
            "folder": r["folder"],
            "is_read": r["is_read"],
            "conversation_key": r["conversation_key"],
            "score": round(score, 4),
        })
    return results
