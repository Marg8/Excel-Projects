# Excel-Projects

Power Query scheduling logic in Streamlit with MA3 summary and BOM build analysis.

## Prerequisites

- Linux or macOS shell
- Python 3.12+

## One-time setup

```bash
cd /workspaces/Excel-Projects
python -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## Run the app (recommended)

```bash
cd /workspaces/Excel-Projects
./run_streamlit.sh
```

Then open one of the URLs shown in terminal:

- http://localhost:8501
- or your forwarded/external URL in Codespaces

## Stop the app

```bash
cd /workspaces/Excel-Projects
./stop_streamlit.sh
```

## Manual run (alternative)

```bash
cd /workspaces/Excel-Projects
source .venv/bin/activate
streamlit run app.py
```

## Health check

When running, this should return `ok`:

```bash
curl -sS http://localhost:8501/_stcore/health
```

## Troubleshooting

- If port 8501 is busy, run `./stop_streamlit.sh`, then `./run_streamlit.sh`.
- Run Streamlit with `streamlit run app.py`, not `python app.py`.
- If dependencies changed, reactivate `.venv` and rerun `pip install -r requirements.txt`.
