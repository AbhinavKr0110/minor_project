"""
ablation_study.py - High-Throughput Offline Policy & Ablation Sweeps
Performs local simulation sweeps on permanently cached benchmark traces without calling API.
"""

import os
from ca_asc_evaluator import OfflineReplayEngine, DEFAULT_BENCHMARK_PATH

def main():
    benchmark_file = DEFAULT_BENCHMARK_PATH if os.path.exists(DEFAULT_BENCHMARK_PATH) else "results/gsm8k_run.jsonl"
    print(f"Initializing Offline Replay Engine with: {benchmark_file}")
    engine = OfflineReplayEngine(trace_filepath=benchmark_file)

    # Sweep theta in {1, 2, 3}, batch sizes in {1, 2}, and tau in {0.50, 0.70, 0.90}
    sweep_results = engine.run_sweep(
        thetas=[1, 2, 3],
        batch_sizes=[1, 2],
        taus=[0.50, 0.70, 0.90],
        max_budget=6
    )
    engine.print_summary_table(sweep_results)

if __name__ == "__main__":
    main()
