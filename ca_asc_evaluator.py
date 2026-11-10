"""
ca_asc_evaluator.py - Unified Confidence-Aware Adaptive Self-Consistency (CA-ASC) Evaluator
Implements the hybrid margin + log-probability confidence stopping policy:
    Stop if: Delta_t >= theta  AND  (sum_{i in lead} w_i) / (sum_{j=1}^t w_j) >= tau
where w_i = exp( (1 / |r_i|) * sum log P(t_k) ) is the length-normalized sequence likelihood.
Includes a high-throughput offline simulation & replay engine over cached inference paths.
"""

import os
import json
import math
from collections import Counter, defaultdict
from typing import List, Dict, Any, Tuple, Optional
import numpy as np

from extractor import extract_model_prediction, clean_number_string

DEFAULT_BENCHMARK_PATH = os.path.join("results", "full_paper_benchmark.jsonl")


def compute_path_weight(sample: Dict[str, Any]) -> float:
    """
    Computes sequence weight w_i = exp(avg_logprob).
    Falls back to explicitly stored 'weight' or 1.0.
    """
    if "weight" in sample and sample["weight"] is not None:
        return float(sample["weight"])
    if "avg_logprob" in sample and sample["avg_logprob"] is not None:
        lp = float(sample["avg_logprob"])
        return math.exp(lp) if lp != 0.0 else 1.0
    return 1.0


def evaluate_ca_asc_stream(
    samples: List[Dict[str, Any]],
    theta: int = 2,
    tau: float = 0.70,
    batch_size: int = 2,
    max_budget: int = 6
) -> Dict[str, Any]:
    """
    Simulates the online sequential CA-ASC decision process on a stream of candidate paths.
    
    Parameters:
    - samples: Full list of generated paths (up to max_budget)
    - theta: Discrete count margin threshold (Delta = c_(1) - c_(2))
    - tau: Confidence share threshold for the leading candidate answer
    - batch_size: Number of paths evaluated per step (k=1 for strictly sequential, k=2 for batch)
    - max_budget: Upper ceiling of paths (Wang et al. fixed N)
    
    Returns a dict with:
    - pred: chosen answer
    - k_used: number of paths evaluated before stopping
    - tokens_used: cumulative completion tokens consumed
    - stopped_early: boolean indicating if stopping criteria triggered before max_budget
    - final_margin: Delta at stopping
    - final_conf: confidence share at stopping
    """
    evaluated_samples = []
    tokens_consumed = 0
    final_pred = None
    stopped_early = False
    final_margin = 0
    final_conf = 0.0

    total_available = min(len(samples), max_budget)

    for start_idx in range(0, total_available, batch_size):
        batch = samples[start_idx:start_idx + batch_size]
        evaluated_samples.extend(batch)
        tokens_consumed += sum(s.get("tokens", 0) for s in batch)

        # Extract predictions and weights for evaluated paths
        preds = []
        weights = []
        for s in evaluated_samples:
            p = s.get("pred")
            if not p or p == "NO_PRED" or p == "NO_PREDICTION":
                p = extract_model_prediction(s.get("text", ""))
            preds.append(p)
            weights.append(compute_path_weight(s))

        # Filter out invalid predictions for consensus calculation
        valid_indices = [idx for idx, p in enumerate(preds) if p and p != "NO_PREDICTION"]
        if not valid_indices:
            continue

        valid_preds = [preds[i] for i in valid_indices]
        valid_weights = [weights[i] for i in valid_indices]

        # Discrete counts
        counts = Counter(valid_preds).most_common()
        lead_ans, lead_count = counts[0]
        runner_up_count = counts[1][1] if len(counts) > 1 else 0
        delta_t = lead_count - runner_up_count

        # Confidence share
        total_w = sum(valid_weights)
        lead_w = sum(valid_weights[i] for i, p in enumerate(valid_preds) if p == lead_ans)
        conf_share = (lead_w / total_w) if total_w > 0 else 0.0

        final_margin = delta_t
        final_conf = conf_share

        # CA-ASC Stopping Criteria: Delta_t >= theta AND conf_share >= tau
        if (delta_t >= theta and conf_share >= tau) or len(evaluated_samples) >= total_available:
            final_pred = lead_ans
            stopped_early = (len(evaluated_samples) < total_available)
            break

    if final_pred is None and evaluated_samples:
        counts = Counter([s.get("pred", "") for s in evaluated_samples]).most_common()
        final_pred = counts[0][0] if counts else "NO_PREDICTION"

    return {
        "pred": clean_number_string(final_pred),
        "k_used": len(evaluated_samples),
        "tokens_used": tokens_consumed,
        "stopped_early": stopped_early,
        "final_margin": final_margin,
        "final_conf": final_conf
    }


class OfflineReplayEngine:
    """
    High-Throughput Offline Replay Engine.
    Loads cached inference traces and performs instantaneous ablation sweeps
    across theta, batch sizes, and confidence thresholds tau without API overhead.
    """
    def __init__(self, trace_filepath: str = DEFAULT_BENCHMARK_PATH):
        self.trace_filepath = trace_filepath
        self.records = self._load_records()

    def _load_records(self) -> List[Dict[str, Any]]:
        if not os.path.exists(self.trace_filepath):
            raise FileNotFoundError(f"Benchmark trace not found at: {self.trace_filepath}")
        records = []
        with open(self.trace_filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    records.append(json.loads(line))
        return records

    def run_sweep(
        self,
        thetas: List[int] = [1, 2, 3],
        batch_sizes: List[int] = [1, 2],
        taus: List[float] = [0.50, 0.60, 0.70, 0.80],
        max_budget: int = 6
    ) -> List[Dict[str, Any]]:
        """
        Executes an exhaustive grid sweep over (theta, batch_size, tau).
        """
        results = []
        total_questions = len(self.records)

        # Baseline: Wang et al. Standard SC (Fixed N=max_budget)
        sc_correct = 0
        sc_tokens = 0
        for r in self.records:
            samples = r.get("samples", [])[:max_budget]
            preds = [s.get("pred") or extract_model_prediction(s.get("text", "")) for s in samples]
            gold = clean_number_string(r.get("gold", ""))
            mode_pred = Counter(preds).most_common(1)[0][0]
            if clean_number_string(mode_pred) == gold:
                sc_correct += 1
            sc_tokens += sum(s.get("tokens", 0) for s in samples)

        sc_acc = (sc_correct / total_questions) * 100
        sc_avg_tokens = sc_tokens / total_questions

        # Sweep over grid
        for b in batch_sizes:
            for th in thetas:
                for t_val in taus:
                    correct = 0
                    tokens_sum = 0
                    k_list = []
                    regrets = 0  # Standard SC was correct, but early stop was wrong

                    for r in self.records:
                        gold = clean_number_string(r.get("gold", ""))
                        samples = r.get("samples", [])
                        
                        # Full SC prediction for regret check
                        sc_pred = clean_number_string(
                            Counter([s.get("pred") or extract_model_prediction(s.get("text", "")) for s in samples[:max_budget]]).most_common(1)[0][0]
                        )
                        
                        eval_res = evaluate_ca_asc_stream(
                            samples=samples,
                            theta=th,
                            tau=t_val,
                            batch_size=b,
                            max_budget=max_budget
                        )

                        pred = eval_res["pred"]
                        is_correct = (pred == gold)
                        if is_correct:
                            correct += 1
                        elif sc_pred == gold and pred != gold:
                            regrets += 1

                        tokens_sum += eval_res["tokens_used"]
                        k_list.append(eval_res["k_used"])

                    acc = (correct / total_questions) * 100
                    avg_tokens = tokens_sum / total_questions
                    mean_k = float(np.mean(k_list))
                    savings = (1.0 - (avg_tokens / sc_avg_tokens)) * 100 if sc_avg_tokens > 0 else 0.0
                    regret_pct = (regrets / total_questions) * 100

                    results.append({
                        "batch_size": b,
                        "theta": th,
                        "tau": t_val,
                        "accuracy": acc,
                        "avg_tokens": avg_tokens,
                        "mean_k": mean_k,
                        "compute_savings_pct": savings,
                        "regret_rate_pct": regret_pct
                    })

        return results

    def print_summary_table(self, sweep_results: List[Dict[str, Any]]):
        """Prints a publication-ready markdown / terminal summary table."""
        print("\n" + "=" * 90)
        print("CONFIDENCE-AWARE ADAPTIVE SELF-CONSISTENCY (CA-ASC) - OFFLINE ABLATION SWEEP")
        print("=" * 90)
        header = f"{'Batch (b)':<10} | {'Theta (θ)':<10} | {'Tau (τ)':<9} | {'Acc (%)':<9} | {'Mean k':<8} | {'Avg Tok':<9} | {'Savings':<10} | {'Regret'}"
        print(header)
        print("-" * 90)
        for res in sweep_results:
            row = (
                f"{res['batch_size']:<10} | "
                f"{res['theta']:<10} | "
                f"{res['tau']:<9.2f} | "
                f"{res['accuracy']:>6.1f}%  | "
                f"{res['mean_k']:>6.2f}  | "
                f"{res['avg_tokens']:>7.0f}   | "
                f"{res['compute_savings_pct']:>+6.1f}%   | "
                f"{res['regret_rate_pct']:>5.1f}%"
            )
            print(row)
        print("=" * 90 + "\n")


if __name__ == "__main__":
    print("Testing CA-ASC Evaluator and High-Throughput Offline Replay Engine...")
    engine = OfflineReplayEngine()
    print(f"Loaded {len(engine.records)} benchmark trace questions from {DEFAULT_BENCHMARK_PATH}")
    
    # Run ablation sweep across theta in {1, 2, 3}, batch sizes in {1, 2}, and tau in {0.50, 0.70, 0.90}
    sweep = engine.run_sweep(
        thetas=[1, 2, 3],
        batch_sizes=[1, 2],
        taus=[0.50, 0.70, 0.90],
        max_budget=6
    )
    engine.print_summary_table(sweep)
