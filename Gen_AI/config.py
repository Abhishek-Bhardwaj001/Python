import os
import sys
from pathlib import Path
print(os.getcwd())

repo_root = Path(__file__).resolve().parents[0] #0 for RAG_Systems and 1 for Gen_AI

# Append to sys.path
import sys
from pathlib import Path


from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parent
if str(REPO_ROOT) not in sys.path:
    sys.path.append(str(REPO_ROOT))

db_directory = REPO_ROOT / "rag_systems" / "vector_db"
db_directory.mkdir(parents=True, exist_ok=True)

docs_path = REPO_ROOT/ "docs"
pdf_path = REPO_ROOT / "docs" / "attention-is-all-you-need.pdf"