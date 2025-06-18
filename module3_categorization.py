import json
import argparse
from ollama import Client
from pathlib import Path
from tqdm import tqdm

# Connect to Ollama
ollama_client = Client(host='http://localhost:11434')

PROMPT_TEMPLATE = """
You are an expert business analyst. Categorize the following news summary into one of these six business themes:
Strategy and Management
Financials
Investments, M&A and Partnerships
Logistics and Operations
Commercials
ESG and Sustainability

News Summary:
"{summary}"

Respond ONLY with the category name.
"""

ALLOWED_CATEGORIES = [
    "Strategy and Management",
    "Financials",
    "Investments, M&A and Partnerships",
    "Logistics and Operations",
    "Commercials",
    "ESG and Sustainability"
]

def normalize_category(output):
    output = output.strip().lower()
    for cat in ALLOWED_CATEGORIES:
        if output == cat.lower():
            return cat
    # Fuzzy match: allow partials or close matches
    for cat in ALLOWED_CATEGORIES:
        if cat.lower() in output or output in cat.lower():
            return cat
    # Extra: handle plural/singular, typos, etc. (expand as needed)
    return "Uncategorised"

def categorize_summary(summary, model="mistral"):
    prompt = PROMPT_TEMPLATE + "\nRespond ONLY with a category from the list above. If unsure, pick the closest.".format(summary=summary)
    response = ollama_client.chat(model=model, messages=[{"role": "user", "content": prompt}])
    raw = response['message']['content'].strip()
    cat = normalize_category(raw)
    if cat == "Uncategorised":
        print(f"[WARN] LLM output not matched to allowed categories: '{raw}'")
    return cat

def process_articles(input_path, output_path, model="mistral"):
    with open(input_path, "r", encoding="utf-8") as f:
        articles = json.load(f)

    for idx, article in enumerate(tqdm(articles, desc="Classifying summaries")):
        try:
            print(f"\nArticle {idx + 1}/{len(articles)}: {article['title']}")
            summary_text = article.get('ai_summary') or article.get('original_summary') or article.get('summary') or ''
            category = categorize_summary(summary_text, model=model)
            print(f"Assigned Category: {category}")
        except Exception as e:
            print(f"Error processing article {idx + 1}: {e}")
            category = "Uncategorized"

        article['category'] = category

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(articles, f, indent=2)

    print(f"\nCategorized {len(articles)} articles. Output saved to: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Categorize summarized articles into business themes.")
    parser.add_argument("--input", required=True, help="Path to input JSON (Module 2 output)")
    parser.add_argument("--output", default="module3_categorized_output.json", help="Output path")
    parser.add_argument("--model", default="mistral", help="Ollama model to use (mistral or zephyr)")

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
    else:
        process_articles(input_path, output_path, model=args.model)
