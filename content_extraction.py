"""
Improved Article Summarizer with Contextual Error Correction

Required installations:
pip install transformers torch beautifulsoup4 requests tqdm spacy
python -m spacy download en_core_web_sm

This version addresses contextual errors like misinterpreting company names as people.
"""

import json
import argparse
from transformers import BartTokenizer, BartForConditionalGeneration, pipeline
import torch
from pathlib import Path
from tqdm import tqdm
import requests
from bs4 import BeautifulSoup
import re
import time
import spacy
from collections import defaultdict

class ArticleSummarizer:
    def __init__(self, model_name="facebook/bart-large-cnn"):
        """Initialize with BART-CNN model which is specifically fine-tuned for summarization"""
        print("Loading model and tokenizer...")
        self.tokenizer = BartTokenizer.from_pretrained(model_name)
        self.model = BartForConditionalGeneration.from_pretrained(model_name)
        
        # Alternative: Use pipeline for easier handling
        self.summarizer_pipeline = pipeline(
            "summarization", 
            model=model_name, 
            tokenizer=model_name,
            device=0 if torch.cuda.is_available() else -1
        )
        
        # Load spaCy for NER
        try:
            self.nlp = spacy.load("en_core_web_sm")
        except OSError:
            print("Warning: spaCy model not found. Install with: python -m spacy download en_core_web_sm")
            self.nlp = None
        
        # Entity knowledge bases
        self.entity_corrections = self._load_entity_knowledge()
        self.entity_patterns = self._create_entity_patterns()

    def _load_entity_knowledge(self):
        """Load domain-specific entity knowledge"""
        # Law firms
        law_firms = {
            "Slaughter and May": "law firm",
            "Clifford Chance": "law firm", 
            "Linklaters": "law firm",
            "Freshfields": "law firm",
            "Allen & Overy": "law firm",
            "Skadden Arps": "law firm",
            "Davis Polk": "law firm",
            "Sullivan & Cromwell": "law firm",
            "Cravath": "law firm",
            "Wachtell Lipton": "law firm",
            "Kirkland & Ellis": "law firm",
            "Latham & Watkins": "law firm"
        }
        
        # Consulting firms
        consulting_firms = {
            "McKinsey & Company": "consulting firm",
            "Boston Consulting Group": "consulting firm",
            "Bain & Company": "consulting firm",
            "BCG": "consulting firm",
            "Deloitte": "consulting firm",
            "PwC": "consulting firm",
            "EY": "consulting firm",
            "KPMG": "consulting firm",
            "Accenture": "consulting firm"
        }
        
        # Investment banks
        investment_banks = {
            "Goldman Sachs": "investment bank",
            "Morgan Stanley": "investment bank",
            "JPMorgan Chase": "investment bank",
            "Bank of America": "investment bank",
            "Citigroup": "investment bank",
            "Barclays": "investment bank",
            "Deutsche Bank": "investment bank",
            "Credit Suisse": "investment bank",
            "UBS": "investment bank"
        }
        
        # Combine all entities
        entities = {}
        entities.update(law_firms)
        entities.update(consulting_firms)
        entities.update(investment_banks)
        return entities

    def _create_entity_patterns(self):
        """Create regex patterns for entity recognition"""
        patterns = []
        for entity, entity_type in self.entity_corrections.items():
            # Create pattern that matches the entity name
            escaped_entity = re.escape(entity)
            pattern = rf'\b{escaped_entity}\b'
            patterns.append((pattern, entity, entity_type))
        
        return patterns
    
    def preprocess_text_with_context(self, text):
        """Add context to text before summarization"""
        if not text:
            return text
        
        # Find entities and add context
        processed_text = text
        
        for pattern, entity, entity_type in self.entity_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                # Add context near the entity mention
                context_addition = f" (the {entity_type})"
                processed_text = re.sub(
                    pattern, 
                    f"{entity}{context_addition}", 
                    processed_text, 
                    flags=re.IGNORECASE,
                    count=1  # Only add context once
                )
        
        return processed_text
    
    def post_process_summary(self, summary):
        """Fix common contextual errors in summary"""
        if not summary:
            return summary
        
        # Fix common misinterpretations
        corrections = {
            # Law firm fixes
            r'\bSlaughter and May\b(?!\s+(?:law firm|firm|LLP))': 'Slaughter and May law firm',
            r'\bClifford Chance\b(?!\s+(?:law firm|firm|LLP))': 'Clifford Chance law firm',
            r'\bLinklaters\b(?!\s+(?:law firm|firm|LLP))': 'Linklaters law firm',
            
            # Consulting firm fixes  
            r'\bMcKinsey & Company\b(?!\s+(?:consulting|firm))': 'McKinsey & Company consulting firm',
            r'\bBoston Consulting Group\b(?!\s+(?:consulting|firm))': 'Boston Consulting Group',
            r'\bBain & Company\b(?!\s+(?:consulting|firm))': 'Bain & Company consulting firm',
            
            # Investment bank fixes
            r'\bGoldman Sachs\b(?!\s+(?:bank|investment))': 'Goldman Sachs investment bank',
            r'\bMorgan Stanley\b(?!\s+(?:bank|investment))': 'Morgan Stanley investment bank',
            
            # Fix person-like interpretations
            r'\b([\w\s]+) and ([\w\s]+) (?:survive|make it out|escape)': 
                lambda m: f"{m.group(1)} and {m.group(2)} (the law firm) survive" if any(
                    firm in f"{m.group(1)} and {m.group(2)}" for firm in self.entity_corrections.keys()
                ) else m.group(0),
            
            # Fix relationship interpretations
            r'\b(Slaughter and May)\s+(?:are|is)\s+(?:a couple|married|together)': r'\1 is a law firm',
        }
        
        processed_summary = summary
        for pattern, replacement in corrections.items():
            if callable(replacement):
                processed_summary = re.sub(pattern, replacement, processed_summary, flags=re.IGNORECASE)
            else:
                processed_summary = re.sub(pattern, replacement, processed_summary, flags=re.IGNORECASE)
        
        return processed_summary
    
    def extract_entities_with_spacy(self, text):
        """Extract entities using spaCy NER"""
        if not self.nlp:
            return []
        
        doc = self.nlp(text)
        entities = []
        
        for ent in doc.ents:
            entities.append({
                'text': ent.text,
                'label': ent.label_,
                'start': ent.start_char,
                'end': ent.end_char
            })
        
        return entities

    def clean_text(self, text):
        """Clean and preprocess text for better summarization"""
        if not text:
            return ""
        # Remove URLs
        text = re.sub(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', '', text)
        # Remove email addresses
        text = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '', text)
        # Remove excessive whitespace and newlines
        text = re.sub(r'\s+', ' ', text)
        # Remove common boilerplate phrases
        boilerplate_patterns = [
            r'click here.*',
            r'read more.*',
            r'subscribe.*',
            r'follow us.*',
            r'share this.*',
            r'comments.*below',
            r'cookie policy.*',
            r'privacy policy.*'
        ]
        for pattern in boilerplate_patterns:
            text = re.sub(pattern, '', text, flags=re.IGNORECASE)
        return text.strip()


    def extract_main_content(self, soup):
        """Better content extraction from HTML"""
        # Remove unwanted elements
        for element in soup(['script', 'style', 'nav', 'header', 'footer', 
                            'aside', 'advertisement', 'ads', 'comment']):
            element.decompose()
        
        # Try multiple strategies to find main content
        content_selectors = [
            'article',
            '[class*="content"]',
            '[class*="article"]',
            '[class*="story"]',
            '[class*="post"]',
            'main',
            '.entry-content',
            '#content'
        ]
        
        best_content = ""
        max_length = 0
        
        for selector in content_selectors:
            try:
                elements = soup.select(selector)
                for element in elements:
                    text = element.get_text(separator=' ', strip=True)
                    if len(text) > max_length:
                        max_length = len(text)
                        best_content = text
            except:
                continue
        
        # Fallback to all paragraphs if no better content found
        if len(best_content) < 200:
            paragraphs = soup.find_all('p')
            best_content = ' '.join(p.get_text() for p in paragraphs)
        
        return self.clean_text(best_content)

    def fetch_full_article_text(self, url, min_length=500):
        """Improved article text extraction"""
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }
            
            resp = requests.get(url, timeout=10, headers=headers)
            if resp.status_code != 200:
                return None
                
            soup = BeautifulSoup(resp.text, "html.parser")
            content = self.extract_main_content(soup)
            
            if len(content) < min_length:
                return None
                
            return content
            
        except Exception as e:
            print(f"Error fetching {url}: {str(e)}")
            return None

    def summarize_with_pipeline(self, text, max_input_length=1024, 
                               min_summary_length=50, max_summary_length=200):
        """Use pipeline for more reliable summarization with context preprocessing"""
        try:
            # Add context to text before summarization
            contextual_text = self.preprocess_text_with_context(text)
            
            # Truncate input if too long
            if len(contextual_text.split()) > max_input_length:
                contextual_text = ' '.join(contextual_text.split()[:max_input_length])
            
            # Use pipeline
            summary = self.summarizer_pipeline(
                contextual_text,
                min_length=min_summary_length,
                max_length=max_summary_length,
                do_sample=False,
                early_stopping=True
            )[0]['summary_text']
            
            # Post-process to fix contextual errors
            summary = self.post_process_summary(summary)
            
            return summary
            
        except Exception as e:
            print(f"Pipeline summarization failed: {str(e)}")
            return self.fallback_summarize(text, max_summary_length)

    def summarize_with_model(self, text, max_input_length=1024, 
                            min_summary_length=50, max_summary_length=200):
        """Direct model usage with better parameters and context preprocessing"""
        try:
            # Add context to text before summarization
            contextual_text = self.preprocess_text_with_context(text)
            
            # Prepare input
            inputs = self.tokenizer.encode(
                contextual_text, 
                return_tensors="pt", 
                max_length=max_input_length, 
                truncation=True,
                padding=True
            )
            
            # Generate summary with better parameters
            with torch.no_grad():
                summary_ids = self.model.generate(
                    inputs,
                    min_length=min_summary_length,
                    max_length=max_summary_length,
                    num_beams=6,  # Increased beam search
                    length_penalty=2.0,  # Encourage longer summaries
                    early_stopping=True,
                    no_repeat_ngram_size=3,  # Prevent repetition
                    do_sample=False,
                    temperature=0.7
                )
            
            summary = self.tokenizer.decode(summary_ids[0], skip_special_tokens=True)
            
            # Post-process to fix contextual errors
            summary = self.post_process_summary(summary)
            
            # Validate summary quality
            if self.is_valid_summary(summary, text):
                return summary
            else:
                return self.fallback_summarize(text, max_summary_length)
                
        except Exception as e:
            print(f"Model summarization failed: {str(e)}")
            return self.fallback_summarize(text, max_summary_length)

    def is_valid_summary(self, summary, original_text):
        """Check if the summary is actually a summary and not just copied text"""
        if not summary or len(summary.strip()) < 20:
            return False
        
        # Check if summary is just a URL or link
        if re.search(r'http[s]?://', summary):
            return False
        
        # Check if summary is too similar to original (basic check)
        summary_words = set(summary.lower().split())
        original_words = set(original_text.lower().split()[:100])  # First 100 words
        
        if len(summary_words) == 0:
            return False
            
        overlap = len(summary_words.intersection(original_words)) / len(summary_words)
        
        # If more than 80% overlap with beginning of original, likely not a good summary
        return overlap < 0.8

    def fallback_summarize(self, text, max_length=200):
        """Simple extractive summarization fallback"""
        sentences = re.split(r'[.!?]+', text)
        if len(sentences) < 2:
            return text[:max_length] + "..."
        
        # Take first few sentences up to max_length
        summary = ""
        for sentence in sentences[:3]:  # First 3 sentences
            if len(summary + sentence) > max_length:
                break
            summary += sentence.strip() + ". "
        
        return summary.strip()

    def process_articles(self, input_path, output_path, company_type, use_pipeline=True):
        """Process articles with improved summarization"""
        with open(input_path, "r", encoding="utf-8") as f:
            articles = json.load(f)

        processed = []
        failed_summaries = 0
        
        for article in tqdm(articles, desc="Summarizing articles"):
            try:
                # Get article content
                title = article.get('title', '')
                existing_summary = article.get('summary', '')
                url = article.get('link', '')
                
                # Try to fetch full article
                full_text = self.fetch_full_article_text(url)
                
                if full_text and len(full_text) > 200:
                    text_to_summarize = f"{title}. {full_text}"
                elif existing_summary and len(existing_summary) > 50:
                    text_to_summarize = f"{title}. {existing_summary}"
                else:
                    text_to_summarize = title
                
                # Generate summary
                if use_pipeline:
                    summary = self.summarize_with_pipeline(text_to_summarize)
                else:
                    summary = self.summarize_with_model(text_to_summarize)
                
                # Validate and fallback if needed
                if not self.is_valid_summary(summary, text_to_summarize):
                    summary = self.fallback_summarize(text_to_summarize)
                    failed_summaries += 1
                
                # Extract entities for debugging
                entities = self.extract_entities_with_spacy(text_to_summarize) if self.nlp else []
                
                structured = {
                    "title": title,
                    "link": url,
                    "published": article.get('published', ''),
                    "source": article.get('source', ''),
                    "original_summary": existing_summary,
                    "ai_summary": summary,
                    "company_type": company_type.lower(),
                    "content_source": "full_article" if full_text else "original_summary",
                    "detected_entities": entities  # For debugging
                }
                
                processed.append(structured)
                
                # Small delay to be respectful to servers
                time.sleep(0.1)
                
            except Exception as e:
                print(f"Error processing article: {str(e)}")
                continue

        # Save results
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(processed, f, indent=2, ensure_ascii=False)

        print(f"\nProcessed {len(processed)} articles")
        print(f"Failed summaries (used fallback): {failed_summaries}")
        print(f"Results saved to: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Summarize news articles with improved accuracy")
    parser.add_argument("--input", required=True, help="Path to input JSON from Module 1")
    parser.add_argument("--output", default="module2_summarized_output.json", help="Output path")
    parser.add_argument("--type", required=True, choices=["competitor", "client"], help="Company type")
    parser.add_argument("--model", default="facebook/bart-large-cnn", help="Model to use for summarization")
    parser.add_argument("--use-pipeline", action="store_true", default=True, help="Use transformers pipeline")
    parser.add_argument("--max-summary-length", type=int, default=200, help="Maximum summary length")
    parser.add_argument("--min-summary-length", type=int, default=50, help="Minimum summary length")

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
    else:
        summarizer = ArticleSummarizer(args.model)
        summarizer.process_articles(
            input_path, 
            output_path, 
            args.type,
            use_pipeline=args.use_pipeline
        )