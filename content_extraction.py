import json
import argparse
from transformers import BartTokenizer, BartForConditionalGeneration
import torch
from pathlib import Path
from tqdm import tqdm

# Load model and tokenizer once
tokenizer = BartTokenizer.from_pretrained("facebook/bart-large-cnn")
model = BartForConditionalGeneration.from_pretrained("facebook/bart-large-cnn")

def summarize_text(text, max_length=1024, summary_length=150):
    inputs = tokenizer.encode(text, return_tensors="pt", max_length=max_length, truncation=True)
    summary_ids = model.generate(inputs, max_length=summary_length, min_length=40, length_penalty=2.0, num_beams=4, early_stopping=True)
    summary = tokenizer.decode(summary_ids[0], skip_special_tokens=True)
    return summary

def process_articles(input_path, output_path, company_type):
    with open(input_path, "r", encoding="utf-8") as f:
        articles = json.load(f)

    processed = []
    for article in tqdm(articles, desc="Summarizing articles"):
        combined_text = article['title'] + ". " + article.get('summary', '')
        summary = summarize_text(combined_text)

        structured = {
            "title": article['title'],
            "link": article['link'],
            "published": article['published'],
            "source": article['source'],
            "summary": summary,
            "company_type": company_type.lower()  # 'competitor' or 'client'
        }
        processed.append(structured)

    # Save processed summaries
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(processed, f, indent=2)

    print(f"\nProcessed {len(processed)} articles saved to: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Summarize news articles and add company type.")
    parser.add_argument("--input", required=True, help="Path to input JSON from Module 1")
    parser.add_argument("--output", default="module2_summarized_output.json", help="Output path")
    parser.add_argument("--type", required=True, choices=["competitor", "client"], help="Company type")

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
    else:
        process_articles(input_path, output_path, args.type)
