import os
import json
from collections import Counter, defaultdict
from datasets import load_dataset
from tqdm import tqdm

from extractor import extract_gsm8k_gold, extract_model_prediction
from client_wrapper import query_llm_paths
from ca_asc_evaluator import evaluate_ca_asc_stream, compute_path_weight

RESULTS_DIR = "results"
RESULTS_FILE = os.path.join(RESULTS_DIR, "gsm8k_run.jsonl")

# ---------------------------------------------------------
# Voting & Aggregation Strategies
# ---------------------------------------------------------
def aggregate_majority_vote(predictions: list[str]) -> str:
    """Standard unweighted majority voting (Wang et al.)."""
    counts = Counter(predictions)
    return counts.most_common(1)[0][0]

def aggregate_weighted_vote(samples: list[dict]) -> str:
    """Confidence-weighted voting using path weights."""
    weighted_votes = defaultdict(float)
    for s in samples:
        weighted_votes[s["pred"]] += s["weight"]
    return max(weighted_votes.items(), key=lambda x: x[1])[0]

# ---------------------------------------------------------
# Evaluation Harness
# ---------------------------------------------------------
def run_evaluation(num_questions: int = 15, max_budget_n: int = 6, batch_size: int = 2, stop_margin: int = 2):
    print(f"Loading GSM8K benchmark test split...")
    dataset = load_dataset("openai/gsm8k", "main", split="test")

    stats = {
        "n1_greedy": {"correct": 0, "tokens": 0},
        "standard_sc": {"correct": 0, "tokens": 0},
        "adaptive_sc": {"correct": 0, "tokens": 0},
        "weighted_sc": {"correct": 0, "tokens": 0}
    }

    # Ensure results folder exists
    os.makedirs(RESULTS_DIR, exist_ok=True)
    log_file = open(RESULTS_FILE, "w", encoding="utf-8")

    print(f"\nStarting benchmark on first {num_questions} questions:")
    print(f"Parameters: N_max={max_budget_n}, Batch_Size={batch_size}, Stop_Margin={stop_margin}\n")

    for idx in tqdm(range(num_questions), desc="Evaluating"):
        item = dataset[idx]
        question = item["question"]
        gold = extract_gsm8k_gold(item["answer"])

        # 1. Sample full budget of paths upfront
        full_samples = query_llm_paths(question, n_samples=max_budget_n, temperature=0.7)
        for s in full_samples:
            s["pred"] = extract_model_prediction(s["text"])

        total_tokens_full = sum(s["tokens"] for s in full_samples)

        # Baseline A: Greedy / Single-Path (n=1)
        pred_n1 = full_samples[0]["pred"]
        if pred_n1 == gold:
            stats["n1_greedy"]["correct"] += 1
        stats["n1_greedy"]["tokens"] += full_samples[0]["tokens"]

        # Baseline B: Standard Self-Consistency (Wang et al. fixed N_max)
        all_preds = [s["pred"] for s in full_samples]
        pred_sc = aggregate_majority_vote(all_preds)
        if pred_sc == gold:
            stats["standard_sc"]["correct"] += 1
        stats["standard_sc"]["tokens"] += total_tokens_full

        # Strategy C: Confidence-Weighted SC (Fixed N_max)
        pred_weighted = aggregate_weighted_vote(full_samples)
        if pred_weighted == gold:
            stats["weighted_sc"]["correct"] += 1
        stats["weighted_sc"]["tokens"] += total_tokens_full

        # Strategy D: Confidence-Aware Adaptive Self-Consistency (CA-ASC)
        ca_asc_result = evaluate_ca_asc_stream(
            samples=full_samples,
            theta=stop_margin,
            tau=0.70,
            batch_size=batch_size,
            max_budget=max_budget_n
        )
        final_adaptive_pred = ca_asc_result["pred"]
        adaptive_tokens = ca_asc_result["tokens_used"]
        adaptive_samples_used = ca_asc_result["k_used"]

        if final_adaptive_pred == gold:
            stats["adaptive_sc"]["correct"] += 1
        stats["adaptive_sc"]["tokens"] += adaptive_tokens

        # Record question trace to JSONL with full sample caching
        log_entry = {
            "question_id": idx,
            "question": question,
            "gold": gold,
            "samples": full_samples,
            "predictions": {
                "n1_greedy": pred_n1,
                "standard_sc": pred_sc,
                "weighted_sc": pred_weighted,
                "adaptive_sc": final_adaptive_pred
            },
            "adaptive_samples_used": adaptive_samples_used,
            "tokens": {
                "n1": full_samples[0]["tokens"],
                "standard_sc": total_tokens_full,
                "adaptive_sc": adaptive_tokens
            }
        }
        log_file.write(json.dumps(log_entry) + "\n")
        log_file.flush()

    log_file.close()

    # ---------------------------------------------------------
    # Final Metric Output
    # ---------------------------------------------------------
    print("\n" + "=" * 76)
    print("EXPERIMENT BENCHMARK SUMMARY")
    print("=" * 76)
    print(f"{'Method':<22} | {'Accuracy (%)':<14} | {'Total Tokens':<14} | {'Savings vs SC'}")
    print("-" * 76)

    sc_tokens = stats["standard_sc"]["tokens"]
    for method, data in stats.items():
        acc = (data["correct"] / num_questions) * 100
        tokens = data["tokens"]
        savings = (1 - (tokens / sc_tokens)) * 100 if sc_tokens > 0 else 0.0
        savings_text = f"{savings:+.1f}%" if method != "standard_sc" else "0.0% (Baseline)"
        print(f"{method:<22} | {acc:>6.2f}%        | {tokens:>10}     | {savings_text}")
    print("=" * 76)
    print(f"Detailed question-level logs saved to: {RESULTS_FILE}\n")

if __name__ == "__main__":
    # Test on a small batch of 10 questions first to verify end-to-end flow
    run_evaluation(num_questions=30, max_budget_n=6, batch_size=2, stop_margin=2)