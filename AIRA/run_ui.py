"""Streamlit entry point.

Run with:  streamlit run run_ui.py

`streamlit run` puts the *script's* directory on sys.path, not the project root,
so pointing it straight at app/ui/streamlit/stream_app.py breaks every
`from app...` / `from config...` import. This shim fixes the path first.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.ui.streamlit import stream_app  # noqa: E402,F401
