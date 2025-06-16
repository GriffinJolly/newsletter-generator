import os
import sys
import json
from news_extraction import extract_news, save_articles_to_json
from content_extraction import process_articles as summarize_articles
from module3_categorization import process_articles as categorize_articles
from module4_ppt_generator import create_ppt_for_company, clean_filename


import requests

def check_ollama_server(url="http://localhost:11434"):
    try:
        r = requests.get(url + "/api/tags", timeout=3)
        if r.status_code == 200:
            return True
        else:
            return False
    except Exception:
        return False

def main():
    print("\n=== Automated News-to-PPT Pipeline ===\n")
    # Ollama server check
    if not check_ollama_server():
        print("\u274c Ollama server is not running or not reachable at http://localhost:11434. Please start it with 'ollama serve' and try again.")
        sys.exit(1)

    company_name = input("Enter the company name: ").strip()
    company_type = input("Is this a 'competitor' or 'potential customer'? (type exactly): ").strip().lower()
    assert company_type in ["competitor", "potential customer"], "Type must be 'competitor' or 'potential customer'"

    # Normalize company_type for downstream modules
    company_type_for_module2 = "competitor" if company_type == "competitor" else "client"

    output_base = os.path.join("outputs", "full_pipeline", clean_filename(company_name))
    os.makedirs(output_base, exist_ok=True)

    # Step 1: News Extraction
    print("\n[1/4] Extracting news articles...")
    try:
        articles = extract_news(company_name, max_articles=10, business_only=True)
        news_json = os.path.join(output_base, "module1_news.json")
        save_articles_to_json(articles, news_json)
        print(f"Saved news articles to {news_json}")
    except Exception as e:
        print(f"\u274c Error during news extraction: {e}")
        sys.exit(1)

    # Step 2: Content Extraction (Summarization)
    print("\n[2/4] Summarizing articles...")
    try:
        summarized_json = os.path.join(output_base, "module2_summarized.json")
        summarize_articles(news_json, summarized_json, company_type_for_module2)
        print(f"Saved summaries to {summarized_json}")
    except UnicodeDecodeError as ude:
        print(f"\u274c Unicode error during summarization: {ude}")
        sys.exit(1)
    except Exception as e:
        print(f"\u274c Error during summarization: {e}")
        sys.exit(1)

    # Step 3: Categorization
    print("\n[3/4] Categorizing articles...")
    try:
        categorized_json = os.path.join(output_base, "module3_categorized.json")
        categorize_articles(summarized_json, categorized_json)
        print(f"Saved categorized articles to {categorized_json}")
    except UnicodeDecodeError as ude:
        print(f"\u274c Unicode error during categorization: {ude}")
        sys.exit(1)
    except Exception as e:
        print(f"\u274c Error during categorization: {e}")
        sys.exit(1)

    # Step 4: PPT Generation
    print("\n[4/4] Generating PPT...")
    try:
        with open(categorized_json, "r", encoding="utf-8") as f:
            final_articles = json.load(f)
        ppt_dir = os.path.join(output_base, "ppt")
        os.makedirs(ppt_dir, exist_ok=True)
        ppt_filename = f"{clean_filename(company_name)}_{company_type.replace(' ', '_')}.pptx"
        ppt_path = os.path.join(ppt_dir, ppt_filename)
        create_ppt_for_company(company_name, final_articles, ppt_dir)
        print(f"\n\u2705 PPT generated: {os.path.abspath(ppt_path)}\n")
    except UnicodeEncodeError as uee:
        print(f"\u274c Unicode error during PPT generation: {uee}")
        sys.exit(1)
    except Exception as e:
        print(f"\u274c Error during PPT generation: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

