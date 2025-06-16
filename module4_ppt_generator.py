import os
import json
import re
from collections import defaultdict
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from datetime import datetime

def clean_filename(name):
    return re.sub(r'\W+', '_', name).strip("_")

def extract_company_name(article):
    """Extract company name from article - improved logic"""
    title = article.get("title", "").lower()
    
    # Check for Slaughter and May variations
    if "slaughter" in title and "may" in title:
        return "Slaughter and May"
    
    # Check in summary as backup
    summary = article.get("summary", "").lower()
    if "slaughter" in summary and "may" in summary:
        return "Slaughter and May"
    
    # If not found, check if it's a competitor article about law firms
    if article.get("company_type") == "competitor":
        return "Slaughter and May"  # Assuming all competitor articles are about S&M
    
    return "Unknown Company"

def create_ppt_for_company(company_name, articles, output_dir):
    prs = Presentation()
    
    # Define title slide layout
    title_slide_layout = prs.slide_layouts[0]
    title_slide = prs.slides.add_slide(title_slide_layout)
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1]
    title.text = f"{company_name} News Report"
    subtitle.text = f"Generated on {datetime.today().strftime('%d-%m-%Y')} | {len(articles)} Articles"

    # Define bullet slide layout
    bullet_slide_layout = prs.slide_layouts[1]

    # Sort articles by date (newest first)
    sorted_articles = sorted(articles, key=lambda x: x.get('published', ''), reverse=True)

    for i, article in enumerate(sorted_articles, 1):
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes

        # Title of slide - include article number
        slide.shapes.title.text = f"{article.get('category', 'News')} - Article {i}/{len(articles)}"

        # Content
        content = shapes.placeholders[1]
        text_frame = content.text_frame
        text_frame.clear()

        # Add article title
        p = text_frame.paragraphs[0]
        p.text = f"📰 {article['title']}"
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue

        # Add article summary
        summary_para = text_frame.add_paragraph()
        summary_para.text = f"📝 Summary: {article['summary']}"
        summary_para.font.size = Pt(14)
        summary_para.level = 0

        # Add published date
        date_para = text_frame.add_paragraph()
        date_para.text = f"📅 Published: {article['published']}"
        date_para.font.size = Pt(12)
        date_para.font.italic = True
        date_para.level = 0

        # Add source
        source_para = text_frame.add_paragraph()
        source_para.text = f"📰 Source: {article['source']}"
        source_para.font.size = Pt(12)
        source_para.level = 0

        # Add link
        link_para = text_frame.add_paragraph()
        link_para.text = f"🔗 Link: {article['link']}"
        link_para.font.size = Pt(11)
        link_para.font.color.rgb = RGBColor(0, 0, 255)
        link_para.level = 0

    # Save the presentation
    output_path = os.path.join(output_dir, f"{clean_filename(company_name)}_consolidated_report.pptx")
    
    try:
        prs.save(output_path)
        print(f"✅ Successfully saved PowerPoint report for {company_name}")
        print(f"📁 Location: {output_path}")
        print(f"📊 Total articles: {len(articles)}")
        print(f"📄 Total slides: {len(prs.slides)} (including title slide)")
    except PermissionError:
        print(f"❌ Error: Could not save {output_path}")
        print("💡 Please close any open PowerPoint files and try again")
    except Exception as e:
        print(f"❌ Unexpected error: {str(e)}")

def generate_reports(json_file_path, output_dir="output_ppts"):
    """Generate consolidated report for Slaughter and May"""
    os.makedirs(output_dir, exist_ok=True)

    try:
        with open(json_file_path, "r", encoding="utf-8") as f:
            articles = json.load(f)
        print(f"📖 Loaded {len(articles)} articles from {json_file_path}")
    except FileNotFoundError:
        print(f"❌ Error: Could not find file {json_file_path}")
        return
    except json.JSONDecodeError:
        print(f"❌ Error: Invalid JSON in {json_file_path}")
        return

    # Group articles by company
    grouped_articles = defaultdict(list)
    for article in articles:
        company = extract_company_name(article)
        grouped_articles[company].append(article)

    print(f"🏢 Found companies: {list(grouped_articles.keys())}")
    
    # Create PPT for each company
    for company, company_articles in grouped_articles.items():
        print(f"\n🔄 Processing {company}...")
        create_ppt_for_company(company, company_articles, output_dir)

    print(f"\n✨ Report generation complete!")

# Alternative function if you want to force all articles into one company
def generate_single_company_report(json_file_path, company_name="Slaughter and May", output_dir="output_ppts"):
    """Generate a single consolidated report for specified company"""
    os.makedirs(output_dir, exist_ok=True)

    try:
        with open(json_file_path, "r", encoding="utf-8") as f:
            articles = json.load(f)
        print(f"📖 Loaded {len(articles)} articles from {json_file_path}")
    except FileNotFoundError:
        print(f"❌ Error: Could not find file {json_file_path}")
        return
    except json.JSONDecodeError:
        print(f"❌ Error: Invalid JSON in {json_file_path}")
        return

    print(f"🔄 Creating consolidated report for {company_name}...")
    create_ppt_for_company(company_name, articles, output_dir)
    print(f"✨ Single company report generation complete!")

# Run this script
if __name__ == "__main__":
    print("Choose an option:")
    print("1. Auto-detect companies from articles (recommended)")
    print("2. Force all articles into single Slaughter and May report")
    
    choice = input("Enter your choice (1 or 2): ").strip()
    
    if choice == "2":
        generate_single_company_report("outputs/module3_categorized_output.json", output_dir="outputs/ppt")
    else:
        generate_reports("outputs/module3_categorized_output.json", output_dir="outputs/ppt")