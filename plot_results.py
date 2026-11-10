import os
import json
import matplotlib.pyplot as plt

RESULTS_FILE = os.path.join("results", "gsm8k_run.jsonl")
OUTPUT_FIG_PARETO = os.path.join("results", "figure_1_pareto_frontier.png")
OUTPUT_FIG_ALLOC = os.path.join("results", "figure_2_sample_allocation.png")

def load_data(filepath: str):
    records = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                records.append(json.loads(line))
    return records

def generate_paper_plots():
    if not os.path.exists(RESULTS_FILE):
        print(f"Error: {RESULTS_FILE} not found.")
        return

    records = load_data(RESULTS_FILE)
    total_q = len(records)
    print(f"Loaded {total_q} evaluated questions from {RESULTS_FILE}")

    methods = ["n1_greedy", "standard_sc", "weighted_sc", "adaptive_sc"]
    correct_counts = {m: 0 for m in methods}
    token_sums = {
        "n1_greedy": sum(r["tokens"]["n1"] for r in records),
        "standard_sc": sum(r["tokens"]["standard_sc"] for r in records),
        "weighted_sc": sum(r["tokens"]["standard_sc"] for r in records),
        "adaptive_sc": sum(r["tokens"]["adaptive_sc"] for r in records)
    }

    for r in records:
        gold = r["gold"]
        for m in methods:
            if r["predictions"].get(m) == gold:
                correct_counts[m] += 1

    accuracies = {m: (correct_counts[m] / total_q) * 100 for m in methods}
    avg_tokens = {m: token_sums[m] / total_q for m in methods}

    # -------------------------------------------------------------
    # Figure 1: Pareto Frontier (Accuracy vs. Compute)
    # -------------------------------------------------------------
    plt.figure(figsize=(8.5, 5.5), dpi=300)
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")

    markers = {
        "n1_greedy": ("o", "#4A5568", "Single-Path CoT (n=1)"),
        "standard_sc": ("s", "#E53E3E", "Standard SC (Fixed N=6)"),
        "weighted_sc": ("^", "#805AD5", "Weighted SC (Fixed N=6)"),
        "adaptive_sc": ("*", "#2B6CB0", "Adaptive SC (Ours)")
    }

    offsets = {
        "n1_greedy": (12, -5),
        "adaptive_sc": (12, 10),
        "standard_sc": (-160, -20),
        "weighted_sc": (-160, 12)
    }

    for m in methods:
        marker, color, label = markers[m]
        size = 220 if m == "adaptive_sc" else 110
        plt.scatter(avg_tokens[m], accuracies[m], color=color, s=size, marker=marker, label=label, zorder=5)
        plt.annotate(
            f"{label}\n({accuracies[m]:.1f}%, {avg_tokens[m]:.0f} tok)",
            (avg_tokens[m], accuracies[m]),
            textcoords="offset points",
            xytext=offsets[m],
            fontsize=9,
            weight="bold" if m == "adaptive_sc" else "normal"
        )

    # Connect baseline to adaptive to show Pareto improvement
    plt.plot(
        [avg_tokens["n1_greedy"], avg_tokens["adaptive_sc"], avg_tokens["standard_sc"]],
        [accuracies["n1_greedy"], accuracies["adaptive_sc"], accuracies["standard_sc"]],
        linestyle="--", color="#CBD5E0", zorder=1
    )

    plt.title("Accuracy vs. Inference Compute Trade-Off (GSM8K)", fontsize=13, weight="bold", pad=12)
    plt.xlabel("Average Completion Tokens per Question", fontsize=11)
    plt.ylabel("Accuracy (%)", fontsize=11)
    plt.ylim(min(accuracies.values()) - 8, max(accuracies.values()) + 10)
    plt.legend(frameon=True, loc="lower right")
    plt.tight_layout()
    plt.savefig(OUTPUT_FIG_PARETO)
    plt.close()
    print(f"Saved: {OUTPUT_FIG_PARETO}")

    # -------------------------------------------------------------
    # Figure 2: Dynamic Sample Allocation Distribution
    # -------------------------------------------------------------
    allocations = [r["adaptive_samples_used"] for r in records]
    all_bins = [2, 4, 6]
    counts = [allocations.count(b) for b in all_bins]

    plt.figure(figsize=(7, 4.5), dpi=300)
    bars = plt.bar([f"k = {b}" for b in all_bins], counts, color="#3182CE", edgecolor="#2B6CB0", width=0.5)
    plt.title("Dynamic Sample Allocation Distribution (Adaptive SC)", fontsize=13, weight="bold", pad=12)
    plt.xlabel("Reasoning Paths Evaluated Before Termination", fontsize=11)
    plt.ylabel("Number of Questions", fontsize=11)

    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2.0, yval + 0.1, int(yval), ha="center", va="bottom", fontsize=10)

    plt.ylim(0, max(counts) + 2)
    plt.tight_layout()
    plt.savefig(OUTPUT_FIG_ALLOC)
    plt.close()
    print(f"Saved: {OUTPUT_FIG_ALLOC}")

if __name__ == "__main__":
    generate_paper_plots()