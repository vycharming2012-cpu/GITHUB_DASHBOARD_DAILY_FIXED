from pathlib import Path
ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
SOURCE = ROOT / "source"
DOCS.mkdir(exist_ok=True)
SOURCE.mkdir(exist_ok=True)
index = DOCS / "index.html"
if not index.exists():
    raise SystemExit("Missing docs/index.html")
print("OK: docs/index.html ready")
