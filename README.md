# AI-Powered Product Spec Extractor & Market Analytics

An end-to-end AI Data Engineering pipeline that transforms unstructured e-commerce product titles into structured specifications using **Google Gemini API** + **Pydantic** schema enforcement, exports formatted multi-sheet Excel reports, and visualizes insights via an interactive **Streamlit** dashboard.

## Architecture

1. **Batch LLM Extraction (`enrich_pipeline.py`)**:
   - Processes raw product catalogs in configurable batches (`batch_size=20`) with automatic retry logic.
   - Enforces strict JSON structured output via **Pydantic** (`Brand`, `Device_Type`, `RAM_GB`, `Storage_GB`, `Screen_Inch`, `Color`) with zero-hallucination prompting rules.
   - Aggregates brand-level pricing statistics (`Brand_Analytics`) and formats multi-sheet Excel workbooks via `openpyxl`.
2. **Interactive BI Dashboard (`dashboard.py`)**:
   - Real-time filtering by keyword, brand, device type, and price range.
   - Interactive **Plotly** visualizations (Average Price by Brand & Price Distribution by Screen Size).
   - Supports custom `.xlsx`/`.csv` upload or falls back cleanly to the bundled sample dataset.

## Tech Stack

- **AI / Structured Output**: `google-genai` (`gemini-3.8-flash`), `pydantic`
- **Data Processing**: `pandas`, `openpyxl`
- **Web App & Visualization**: `streamlit`, `plotly`

## Run Locally

```bash
# 1. Setup virtual environment
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# 2. Launch Dashboard immediately (uses bundled sample dataset)
streamlit run dashboard.py

# 3. (Optional) Run AI Extraction Pipeline on your own input_products.xlsx
cp .env.example .env
# Add your GEMINI_API_KEY inside .env
python enrich_pipeline.py
```
