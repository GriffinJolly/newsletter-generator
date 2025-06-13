import feedparser
import argparse
import json
import os

def extract_news(company_name, max_articles=10, business_only=True):
    query = f'"{company_name}"'.replace(" ", "+")
    rss_url = f"https://news.google.com/rss/search?q={query}&hl=en-US&gl=US&ceid=US:en"

    feed = feedparser.parse(rss_url)
    articles = []

    business_keywords = [
        "revenue", "profit", "earnings", "loss", "forecast", "guidance", "financial", "strategy", "expansion",
        "layoffs", "hiring", "restructuring", "merger", "acquisition", "joint venture", "investment",
        "launch", "product", "supply chain", "manufacturing", "distribution", "logistics", "ceo", "executive",
        "partnership", "divest", "green", "emissions", "sustainability", "esg", "carbon", "environment", "ipo"
    ]

    for entry in feed.entries:
        title = entry.title
        summary = entry.summary if 'summary' in entry else ""
        text_to_check = (title + " " + summary).lower()

        if company_name.upper() == "UPS":
            ups_company_patterns = [
                "ups inc", "ups corp", "united parcel service", "ups stock", "ups shares",
                "ups earnings", "ups revenue", "ups delivery", "ups logistics", "ups package",
                "ups driver", "ups workers", "ups union", "ups ceo", "ups executive"
            ]
            if not any(pattern in text_to_check for pattern in ups_company_patterns):
                continue
        else:
            if company_name.lower() not in text_to_check:
                continue
            if len(company_name) <= 3 and company_name.upper() != "UPS":
                company_indicators = ["company", "corp", "corporation", "inc", "ltd", "llc", "stock", "shares", "earnings", "ceo"]
                if not any(indicator in text_to_check for indicator in company_indicators):
                    continue

        if business_only and not any(keyword in text_to_check for keyword in business_keywords):
            continue

        article = {
            "title": title,
            "link": entry.link,
            "published": entry.published if 'published' in entry else "N/A",
            "source": entry.source.title if 'source' in entry else "Unknown",
            "summary": summary
        }
        articles.append(article)

        if len(articles) >= max_articles:
            break

    return articles

def save_articles_to_json(articles, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(articles, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Articles saved to: {output_path}")

def display_articles(articles):
    for i, article in enumerate(articles, 1):
        print(f"\n{i}. {article['title']}")
        print(f"Published: {article['published']}")
        print(f"Source: {article['source']}")
        print(f"URL: {article['link']}")
        print(f"Summary: {article['summary'][:200]}...")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract news articles about a company.")
    parser.add_argument("--company", required=True, help="Name of the company to search for")
    parser.add_argument("--max", type=int, default=10, help="Max number of relevant articles to fetch")
    parser.add_argument("--all", action="store_true", help="Include all news, not just business-focused ones")

    args = parser.parse_args()

    print(f"\n🔍 Searching for news about: {args.company}")
    news = extract_news(args.company, max_articles=args.max, business_only=not args.all)

    if news:
        print(f"\n📡 Found {len(news)} relevant articles:")
        display_articles(news)

        # Save output to data/news_articles.json
        output_file = "data/news_articles.json"
        save_articles_to_json(news, output_file)
    else:
        print("❌ No relevant articles found.")
