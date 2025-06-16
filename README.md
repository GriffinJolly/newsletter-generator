# Newsletter Generator Pipeline

![MIT License](https://img.shields.io/badge/license-MIT-green)

A fully automated CLI pipeline to extract, summarize, categorize, and generate a PowerPoint report of business news for a given company. Built and maintained by [GriffinJolly](https://github.com/GriffinJolly).

## Features
- **End-to-end automation:** From news extraction to PPT generation in one command.
- **Business-focused:** Filters for business-relevant news and themes.
- **Flexible:** Works for both competitors and potential customers.
- **Robust Unicode support:** Handles non-English and special characters gracefully.

## Quickstart

### 1. Install Requirements
```bash
pip install -r requirements.txt
```

- You also need [Ollama](https://ollama.com/) (for local LLM inference) running on your machine. Start it with:
```bash
ollama serve
```

### 2. Run the Pipeline
```bash
python run_full_ppt_pipeline.py
```

You will be prompted for:
- Company name
- Whether it is a 'competitor' or 'potential customer'
- **How many articles to extract** (default: 10)

The script will:
1. Gather the specified number of news articles for the company
2. Summarize and extract key content
3. Categorize news into business themes (requires Ollama server)
4. Generate a PowerPoint (.pptx) report in `outputs/full_pipeline/<CompanyName>/ppt/`

## Project Structure
```
newsletter-generator/
├── content_extraction.py         # Module 2: Summarization
├── module3_categorization.py     # Module 3: Categorization (Ollama LLM)
├── module4_ppt_generator.py      # Module 4: PPT generation
├── news_extraction.py            # Module 1: News extraction
├── run_full_ppt_pipeline.py      # Unified CLI pipeline script
├── requirements.txt              # Python dependencies
├── outputs/                      # All generated outputs
└── README.md                     # This file
```

## Requirements
- Python 3.8+
- [Ollama](https://ollama.com/) (for local LLM API, e.g., mistral)
- Python packages in `requirements.txt` (transformers, pptx, requests, tqdm, etc.)

## Customization
- You can adjust the number of news articles, business keywords, or themes by editing the respective modules.
- To use a different LLM model, update the model name in `module3_categorization.py`.

## License
MIT

---

> Built by [GriffinJolly](https://github.com/GriffinJolly)
