import re

def clean_number_string(num_str: str) -> str:
    """
    Cleans a matched number string by:
    - Removing markdown formatting (e.g. **, *, _)
    - Removing thousand-separator commas (e.g., '1,200' -> '1200')
    - Removing currency signs ($ or ₹ or € or £ or ¥) and percent signs (%)
    - Stripping trailing punctuation (. , : ;)
    - Stripping trailing decimal zeros if it's an integer (e.g., '42.0' -> '42')
    """
    if not num_str:
        return ""
    
    # Strip markdown bold/italics
    s = re.sub(r"[\*\_]", "", str(num_str))
    
    # Isolate first numeric token if surrounded by trailing words (e.g. '104 dollars')
    match = re.search(r"(-?[\$₹€£¥]?\s*[\d,]+(?:\.\d+)?%?)", s)
    if match:
        s = match.group(1)

    # Remove whitespace, commas, currency symbols, and percentage signs
    cleaned = re.sub(r"[,\$₹€£¥%\s]", "", s).strip()
    
    # Strip trailing punctuation if accidentally captured
    cleaned = cleaned.rstrip(".,:;")

    # If it's a decimal like 42.0 or 15.00, normalize to integer string '42' or '15'
    try:
        val = float(cleaned)
        if val.is_integer():
            return str(int(val))
        return str(val)
    except ValueError:
        return cleaned


def extract_gsm8k_gold(gold_text: str) -> str:
    """
    GSM8K ground-truth answers in Hugging Face always end with:
    '#### <number>' (e.g., '#### 48' or '#### -12.5')
    """
    if not gold_text:
        return ""
    
    # Find number immediately following '####'
    match = re.search(r"####\s*(-?[\$₹€£¥]?\s*[\d,]+(?:\.\d+)?)", gold_text)
    if match:
        return clean_number_string(match.group(1))
    
    return ""


def extract_model_prediction(generated_text: str) -> str:
    """
    Extracts the numerical prediction from the model's Chain-of-Thought generation.
    Uses a priority cascade:
    1. Explicit GSM8K format (#### <num>)
    2. LaTeX \\boxed{<num>} (e.g. \\boxed{104}, \\boxed{$104})
    3. Explicit verbal conclusions ("the answer is <num>", "final answer: <num>")
    4. Markdown bolded numbers (e.g. **104**, **$104**)
    5. Written concluding statements ("The cost is 104 dollars", "Total: 104")
    6. Fallback: Last standalone number found in the generated text.
    """
    if not generated_text:
        return "NO_PREDICTION"

    text = generated_text.strip()
    if not text:
        return "NO_PREDICTION"

    # Rule 1: Model copied GSM8K format (#### 42)
    match = re.search(r"####\s*([^\n\r]+)", text)
    if match:
        c = clean_number_string(match.group(1))
        if c and any(ch.isdigit() for ch in c):
            return c

    # Rule 2: LaTeX \\boxed{...} (handles optional nested commands)
    boxed_matches = re.findall(r"\\boxed\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}", text)
    if boxed_matches:
        c = clean_number_string(boxed_matches[-1])
        if c and any(ch.isdigit() for ch in c):
            return c

    # Rule 3: Explicit verbal conclusion phrases (e.g. "The answer is 42", "Final answer: 42")
    verbal_patterns = [
        r"(?:the\s+answer\s+is|final\s+answer\s+is|the\s+answer\s+should\s+be|answer\s*[:=]|the\s+correct\s+answer\s+is)\s*[:=]?\s*([^\n\r]+)",
        r"(?:total\s*[:=]|the\s+total\s+is|total\s+cost\s*[:=]|the\s+cost\s+is)\s*[:=]?\s*([^\n\r]+)",
        r"(?:in\s+total[,\s]+(?:she|he|they|it)\s+(?:makes?|earns?|has|have|gets?|spent|pays?|cost))\s*[:=]?\s*([^\n\r]+)",
    ]
    for pattern in verbal_patterns:
        matches = list(re.finditer(pattern, text, re.IGNORECASE))
        if matches:
            c = clean_number_string(matches[-1].group(1))
            if c and any(ch.isdigit() for ch in c):
                return c

    # Rule 4: Markdown bolded numbers near conclusion (e.g. **104**, **$104**)
    bold_matches = re.findall(r"\*\*\s*([^\*]+?)\s*\*\*", text)
    if bold_matches:
        for b in reversed(bold_matches):
            c = clean_number_string(b)
            if c and any(ch.isdigit() for ch in c):
                return c

    # Rule 5: Written concluding statements ("The cost is 104 dollars", "Total: 104")
    concluding_stmt = re.findall(
        r"(?:cost|total|amount|sum|price|result|makes?|has|leaves?)\s+(?:is|of|=|was)?\s*[\$₹€£¥]?\s*([\d,]+(?:\.\d+)?)\s*(?:dollars?|cents?|eggs?|miles?|apples?)?",
        text,
        re.IGNORECASE
    )
    if concluding_stmt:
        c = clean_number_string(concluding_stmt[-1])
        if c and any(ch.isdigit() for ch in c):
            return c

    # Rule 6: Fallback - Grab the very last number appearing in the output
    all_numbers = re.findall(r"-?[\$₹€£¥]?\s*[\d,]+(?:\.\d+)?%?", text)
    if all_numbers:
        valid_candidates = []
        for num in all_numbers:
            c = clean_number_string(num)
            if c and any(ch.isdigit() for ch in c):
                valid_candidates.append(c)
        if valid_candidates:
            return valid_candidates[-1]

    return "NO_PREDICTION"


# -----------------------------------------------------------------------------
# Quick Unit Test (Runs only when executing: python extractor.py)
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    test_cases = [
        ("She bought 3 apples for $4 each. #### 12", "12"),
        ("After calculating 5 * 20, the final answer is 100.", "100"),
        ("The total cost comes out to \\boxed{1,250.00}.", "1250"),
        ("First we add 10 + 5 = 15. Then we subtract 3 to get 12.", "12"),
        ("Therefore, the answer is -4.50 degrees.", "-4.5"),
        ("So, the answer is **104**.", "104"),
        ("The cost is **$104** in total.", "104"),
        ("The cost is 104 dollars.", "104"),
        ("Total: 104", "104"),
        ("The result was 85%.", "85"),
        ("At $2 each, the daily earnings are \\(9 \\times 2 = \\$18\\).\n\nThe answer is 18", "18"),
    ]

    print("Running extractor tests...")
    for idx, (raw_text, expected) in enumerate(test_cases):
        pred = extract_model_prediction(raw_text)
        assert pred == expected, f"Test {idx + 1} failed: Expected '{expected}', got '{pred}'"
        print(f"Test {idx + 1} passed: '{raw_text[:35]}...' -> {pred}")

    print("\nAll parser tests passed successfully!")
