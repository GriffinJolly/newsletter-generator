# Newsletter Generator Pipeline

![MIT License](https://img.shields.io/badge/license-MIT-green)

A fully automated CLI pipeline to extract, summarize, categorize, and generate a PowerPoint report of business news for a given company. Built and maintained by [GriffinJolly](https://github.com/GriffinJolly).

The app is deployed. Access through this link: https://newsletter-generator-frost.streamlit.app/

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

### 2. Run the Pipeline (CLI)
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

### 3. Run the GUI (Recommended)
A modern, browser-based interface is available using Streamlit:
```bash
pip install streamlit  # if not already installed
streamlit run app_streamlit.py
```
- Fill in the company name, type, and article count in the browser UI.
- Click "Run Pipeline 🚀" and download the generated PPT when ready.
- No console Unicode errors—fully robust and user-friendly!

## Project Structure
```
newsletter-generator/
├── app_streamlit.py              # Streamlit GUI for the pipeline
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
- For the GUI: `streamlit` (install with `pip install streamlit`)

## Customization
- You can adjust the number of news articles, business keywords, or themes by editing the respective modules.
- To use a different LLM model, update the model name in `module3_categorization.py`.

## Troubleshooting Unicode/Encoding Errors
- All CLI print statements are now ASCII-only. If you see encoding errors, update your terminal to use UTF-8 or use the Streamlit GUI (which avoids all console encoding issues).
- If you previously encountered errors like `UnicodeEncodeError: 'charmap' codec can't encode character ...`, simply update your codebase and re-run. All emojis have been removed from print output.

## License
MIT

---

> Built by [GriffinJolly](https://github.com/GriffinJolly)
