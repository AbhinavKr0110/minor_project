import os
import re
import json
import asyncio
from collections import Counter
from datasets import load_dataset
from dotenv import load_dotenv
from openai import AsyncOpenAI
import matplotlib.pyplot as plt
import numpy as np

load_dotenv()

API_KEY = os.getenv("API_KEY")
BASE_URL = os.getenv("BASE_URL", "https://api.groq.com/openai/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "openai/gpt-oss-20b")

client = AsyncOpenAI(api_key=API_KEY, base_url=BASE_URL)

RESULTS_FILE = "results/full_paper_benchmark.jsonl"
os.makedirs("results", exist_ok=True)

COT_SYSTEM_PROMPT = (
    "You are a precise mathematical assistant. Solve the math problem step by step. "
    "At the very end, strictly output: 'The answer is [number]'."
)

def clean_num(s: str) -> str:
    s = re.sub(r"[,\$₹\s]", "", s).strip().rstrip(".")
    try:
        val = float(s)
        return str(int(val)) if val.is_integer() else str(val)
    except:
        return s

def extract_gold(text: str) -> str:
    match = re.search(r"####\s*(-?[\d,]+(?:\.\d+)?)", text)
    return clean_num(match.group(1)) if match else ""

def extract_pred(text: str) -> str:
    match = re.search(r"####\s*(-?[\$₹]?[\d,]+(?:\.\d+)?)", text)
    if match: return clean_num(match.group(1))
    
    match = re.search(r"(?:the\s+answer\s+is|final\s+answer\s+is|answer\s*[:=])\s*[\$₹]?\s*(-?[\d,]+(?:\.\d+)?)", text, re.I)
    if match: return clean_num(match.group(1))
    
    nums = re.findall(r"-?[\$₹]?[\d,]+(?:\.\d+)?", text)
    valid = [clean_num(n) for n in nums if any(c.isdigit() for c in n)]
    return valid[-1] if valid else "NO_PRED"

async def fetch_single_path(question: str, sem: asyncio.Semaphore):
    async with sem:
        for attempt in range(6):
            try:
                resp = await client.chat.completions.create(
                    model=MODEL_NAME,
                    messages=[
                        {"role": "system", "content": COT_SYSTEM_PROMPT},
                        {"role": "user", "content": question}
                    ],
                    temperature=0.7,
                    max_tokens=220
                )
                text = resp.choices[0].message.content or ""
                toks = resp.usage.completion_tokens if getattr(resp, "usage", None) else len(text.split()) * 4 // 3
                return {"text": text, "pred": extract_pred(text), "tokens": toks}
            except Exception as e:
                wait = 15 if "429" in str(e) else (2 ** attempt)
                await asyncio.sleep(wait)
        return {"text": "", "pred": "ERR", "tokens": 0}

async def benchmark(num_questions=30, max_n=6):
    ds = load_dataset("openai/gsm8k", "main", split="test")
    sem = asyncio.Semaphore(2) # Safe concurrent pool for free-tier TPM limits
    
    print(f"Running high-efficiency benchmark on {num_questions} problems (N={max_n})...")
    
    records = []
    for idx in range(num_questions):
        q = ds[idx]["question"]
        gold = extract_gold(ds[idx]["answer"])
        
        tasks = [fetch_single_path(q, sem) for _ in range(max_n)]
        samples = await asyncio.gather(*tasks)
        
        # Adaptive early stopping simulation (k=2, margin=2)
        adaptive_samples = []
        adaptive_tok = 0
        adaptive_pred = None
        for b in range(0, max_n, 2):
            batch = samples[b:b+2]
            adaptive_samples.extend(batch)
            adaptive_tok += sum(s["tokens"] for s in batch)
            counts = Counter([s["pred"] for s in adaptive_samples]).most_common()
            lead = counts[0][1]
            runner_up = counts[1][1] if len(counts) > 1 else 0
            if (lead - runner_up) >= 2 or len(adaptive_samples) >= max_n:
                adaptive_pred = counts[0][0]
                break
                
        rec = {
            "id": idx,
            "question": q,
            "gold": gold,
            "samples": samples,
            "adaptive_k": len(adaptive_samples),
            "adaptive_tokens": adaptive_tok,
            "adaptive_pred": adaptive_pred
        }
        records.append(rec)
        print(f"[{idx+1}/{num_questions}] Gold: {gold} | Adaptive (k={len(adaptive_samples)}): {adaptive_pred}")
        
    with open(RESULTS_FILE, "w") as f:
        for r in records:
            f.write(json.dumps(r) + "\n")
    print(f"\nAll experiments saved to {RESULTS_FILE}")

if __name__ == "__main__":
    asyncio.run(benchmark(num_questions=30, max_n=6))