"""
RAG Retriever — Dynamic Few-Shot Example Selection

Uses TF-IDF similarity to find the most relevant training examples
for a given user prompt. Injects them as few-shot context so the LLM
gets targeted guidance instead of a massive static system prompt.

No external ML dependencies — uses pure Python TF-IDF.

Usage:
    from rag_retriever import get_few_shot_context
    context = get_few_shot_context("Create an M8 hex nut", top_k=3)
"""

import json
import math
import re
import os
from pathlib import Path
from collections import Counter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

PROJECT_DIR = Path(__file__).parent
TRAINING_DIR = PROJECT_DIR / "training_examples"
GENERATED_DATA = PROJECT_DIR / "training_data.jsonl"

# Cache the index in memory
_INDEX = None


# ---------------------------------------------------------------------------
# Pure Python TF-IDF Engine (zero dependencies)
# ---------------------------------------------------------------------------

def _tokenize(text):
    """Simple tokenizer: lowercase, split on non-alphanumeric, remove stopwords."""
    tokens = re.findall(r'[a-z0-9]+', text.lower())
    stopwords = {'a', 'an', 'the', 'and', 'or', 'is', 'it', 'to', 'of', 'in',
                 'for', 'on', 'with', 'at', 'by', 'from', 'that', 'this', 'be',
                 'as', 'was', 'are', 'were', 'has', 'have', 'had', 'do', 'does',
                 'did', 'will', 'would', 'can', 'could', 'should', 'may', 'might',
                 'i', 'me', 'my', 'you', 'your', 'we', 'our', 'create', 'make',
                 'design', 'build', 'generate'}
    return [t for t in tokens if t not in stopwords and len(t) > 1]


def _compute_tf(tokens):
    """Term frequency: count / total tokens."""
    counts = Counter(tokens)
    total = len(tokens)
    if total == 0:
        return {}
    return {word: count / total for word, count in counts.items()}


class TFIDFIndex:
    """Simple TF-IDF index for finding similar prompts."""

    def __init__(self):
        self.documents = []     # List of (prompt, actions_json, source)
        self.doc_tokens = []    # Tokenized prompts
        self.doc_tf = []        # TF for each doc
        self.idf = {}           # IDF values
        self.vocab = set()

    def add_document(self, prompt, actions_json, source="unknown"):
        """Add a document to the index."""
        tokens = _tokenize(prompt)
        self.documents.append((prompt, actions_json, source))
        self.doc_tokens.append(tokens)
        self.doc_tf.append(_compute_tf(tokens))
        self.vocab.update(tokens)

    def build(self):
        """Compute IDF values after all documents are added."""
        n_docs = len(self.documents)
        if n_docs == 0:
            return

        # Count how many documents contain each term
        doc_freq = Counter()
        for tokens in self.doc_tokens:
            unique = set(tokens)
            for token in unique:
                doc_freq[token] += 1

        # IDF = log(N / df) + 1 (smoothed)
        self.idf = {
            word: math.log(n_docs / (df + 1)) + 1
            for word, df in doc_freq.items()
        }

    def search(self, query, top_k=3):
        """Find the top-k most similar documents to the query."""
        if not self.documents:
            return []

        query_tokens = _tokenize(query)
        query_tf = _compute_tf(query_tokens)

        # Compute TF-IDF cosine similarity for each document
        scores = []
        for i, doc_tf in enumerate(self.doc_tf):
            score = self._cosine_similarity(query_tf, doc_tf)
            scores.append((score, i))

        # Sort by score descending
        scores.sort(key=lambda x: -x[0])

        results = []
        for score, idx in scores[:top_k]:
            if score > 0.01:  # Skip near-zero matches
                prompt, actions_json, source = self.documents[idx]
                results.append({
                    "prompt": prompt,
                    "actions": actions_json,
                    "score": round(score, 4),
                    "source": source,
                })

        return results

    def _cosine_similarity(self, tf_a, tf_b):
        """Cosine similarity between two TF-IDF vectors."""
        # Get all terms
        all_terms = set(tf_a.keys()) | set(tf_b.keys())
        if not all_terms:
            return 0.0

        dot = 0.0
        norm_a = 0.0
        norm_b = 0.0

        for term in all_terms:
            idf = self.idf.get(term, 1.0)
            a = tf_a.get(term, 0) * idf
            b = tf_b.get(term, 0) * idf
            dot += a * b
            norm_a += a * a
            norm_b += b * b

        if norm_a == 0 or norm_b == 0:
            return 0.0

        return dot / (math.sqrt(norm_a) * math.sqrt(norm_b))


# ---------------------------------------------------------------------------
# Index Building
# ---------------------------------------------------------------------------

def _load_from_training_examples():
    """Load prompt→JSON pairs from training_examples/ directory."""
    examples = []
    if not TRAINING_DIR.exists():
        return examples

    for txt_file in TRAINING_DIR.glob("*.txt"):
        json_file = txt_file.with_suffix(".json")
        if json_file.exists():
            try:
                prompt = txt_file.read_text(encoding="utf-8").strip()
                actions = json_file.read_text(encoding="utf-8").strip()
                # Validate JSON
                json.loads(actions)
                if prompt:
                    examples.append((prompt, actions, "saved_run"))
            except (json.JSONDecodeError, IOError):
                continue

    return examples


def _load_from_jsonl():
    """Load examples from training_data.jsonl."""
    examples = []
    if not GENERATED_DATA.exists():
        return examples

    try:
        with open(GENERATED_DATA, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    row = json.loads(line)
                    messages = row.get("messages", [])
                    if len(messages) >= 3:
                        prompt = messages[1].get("content", "")
                        actions = messages[2].get("content", "")
                        # Validate it's valid JSON
                        json.loads(actions)
                        if prompt:
                            examples.append((prompt, actions, "generated"))
                except (json.JSONDecodeError, KeyError):
                    continue
    except IOError:
        pass

    return examples


def build_index(force=False):
    """
    Build the TF-IDF index from all available training data.
    Sources: training_examples/ dir + training_data.jsonl
    """
    global _INDEX
    if _INDEX is not None and not force:
        return _INDEX

    index = TFIDFIndex()

    # Source 1: Manually saved successful runs
    saved = _load_from_training_examples()
    for prompt, actions, source in saved:
        index.add_document(prompt, actions, source)

    # Source 2: Generated training data
    generated = _load_from_jsonl()
    for prompt, actions, source in generated:
        index.add_document(prompt, actions, source)

    index.build()
    _INDEX = index

    total = len(index.documents)
    if total > 0:
        print(f"    📚 RAG index: {total} examples ({len(saved)} saved + {len(generated)} generated)")
    else:
        print(f"    📚 RAG index: empty (run 'python generate_training_data.py' to populate)")

    return index


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def find_similar(query, top_k=3):
    """
    Find the top-k most similar training examples for a query.

    Returns:
        List of dicts with keys: prompt, actions (JSON string), score, source
    """
    index = build_index()
    return index.search(query, top_k=top_k)


def get_few_shot_context(query, top_k=3):
    """
    Generate few-shot context string from the most similar training examples.
    This string is injected into the system prompt to guide the LLM.

    Returns:
        String with formatted examples, or empty string if no examples found.
    """
    results = find_similar(query, top_k=top_k)

    if not results:
        return ""

    lines = ["### RELEVANT EXAMPLES (from training data):"]
    lines.append("Use these as reference for your response:\n")

    for i, result in enumerate(results, 1):
        lines.append(f"**Example {i}:** \"{result['prompt']}\"")
        # Compact the JSON for the prompt
        try:
            actions = json.loads(result["actions"])
            compact_json = json.dumps(actions, separators=(',', ':'))
        except json.JSONDecodeError:
            compact_json = result["actions"]
        lines.append(compact_json)
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CLI for testing
# ---------------------------------------------------------------------------

def main():
    """Test the RAG retriever with sample queries."""
    print("🔍 RAG Retriever — Test Mode")
    print("=" * 50)

    # Build index
    index = build_index(force=True)
    print(f"\nIndex size: {len(index.documents)} documents")

    if len(index.documents) == 0:
        print("\n⚠️  No training data found.")
        print("   Run 'python generate_training_data.py' first to populate the index.")
        return

    # Test queries
    test_queries = [
        "Create an M6 hex nut",
        "Make a table with 4 legs",
        "Create a coffee mug with handle",
        "Design a 50mm sphere",
        "Create a box with rounded edges",
        "Make a washer",
        "Create a vase",
    ]

    for query in test_queries:
        print(f"\n{'─'*50}")
        print(f"Query: \"{query}\"")
        results = find_similar(query, top_k=3)

        if not results:
            print("  No matches found")
            continue

        for r in results:
            try:
                n_steps = len(json.loads(r['actions']))
            except:
                n_steps = "?"
            print(f"  [{r['score']:.3f}] \"{r['prompt']}\" ({n_steps} steps, src: {r['source']})")


if __name__ == "__main__":
    main()
