import os
import time
import math
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Any
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

API_KEY = os.getenv("API_KEY")
BASE_URL = os.getenv("BASE_URL", "https://api.groq.com/openai/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "openai/gpt-oss-20b")

if not API_KEY:
    raise ValueError("API_KEY not found. Please set it in your .env file.")

client = OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

COT_SYSTEM_PROMPT = (
    "You are a precise mathematical reasoning assistant. "
    "Solve problems step-by-step. At the very end of your response, "
    "state your final conclusion strictly as: 'The answer is [number]'."
)

def _call_single_path(messages: list, temperature: float, max_retries: int = 6) -> Dict[str, Any]:
    """Calls the API with rate-limit backoff."""
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=messages,
                temperature=temperature,
                n=1,
                max_tokens=256,  # Kept concise to avoid OTPM exhaustion
            )
            
            choice = response.choices[0]
            text = choice.message.content or ""
            # Fallback for models outputting chain-of-thought into reasoning or reasoning_content
            if not text and getattr(choice.message, "reasoning", None):
                text = choice.message.reasoning or ""
            if not text and getattr(choice.message, "reasoning_content", None):
                text = choice.message.reasoning_content or ""
            
            tokens_used = (
                len(text.split()) * 4 // 3
                if not getattr(response, "usage", None)
                else response.usage.completion_tokens
            )

            avg_logprob = 0.0
            if getattr(choice, "logprobs", None) and choice.logprobs and choice.logprobs.content:
                token_lps = [t.logprob for t in choice.logprobs.content if t.logprob is not None]
                if token_lps:
                    avg_logprob = sum(token_lps) / len(token_lps)
            
            weight = math.exp(avg_logprob) if avg_logprob != 0.0 else 1.0

            return {
                "text": text,
                "avg_logprob": avg_logprob,
                "weight": weight,
                "tokens": tokens_used
            }

        except Exception as e:
            err_str = str(e)
            # If rate limited (429), pause for 15s to allow token bucket refill
            if "429" in err_str:
                wait_time = 15
                print(f"[Rate Limit 429] Token bucket full. Backing off for {wait_time}s...")
            else:
                wait_time = 2 ** attempt
                print(f"[Warning] API call failed ({e}). Retrying in {wait_time}s...")
            time.sleep(wait_time)

    raise RuntimeError(f"Failed to generate reasoning path after {max_retries} attempts.")


def query_llm_paths(
    prompt: str, 
    n_samples: int = 1, 
    temperature: float = 0.7
) -> List[Dict[str, Any]]:
    """Generates n paths using a safe worker pool to respect free tier limits."""
    messages = [
        {"role": "system", "content": COT_SYSTEM_PROMPT},
        {"role": "user", "content": prompt}
    ]

    # Use max_workers=2 so we do not flood the TPM quota simultaneously
    results = []
    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = [
            executor.submit(_call_single_path, messages=messages, temperature=temperature)
            for _ in range(n_samples)
        ]
        for f in futures:
            results.append(f.result())
            time.sleep(0.5)  # slight breathing room between completions
            
    return results


if __name__ == "__main__":
    print(f"Testing client connection to {BASE_URL} using model '{MODEL_NAME}'...")
    test_q = "If a box contains 3 red balls and 5 green balls, how many balls are there in total?"
    samples = query_llm_paths(test_q, n_samples=2, temperature=0.7)
    print(f"\nGenerated {len(samples)} path(s) successfully.")
    for i, s in enumerate(samples):
        print(f"Path {i+1}: {s['text'].strip()[:80]}... [Tokens: {s['tokens']}]")