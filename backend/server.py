from __future__ import annotations

import argparse
import os
from pathlib import Path

import uvicorn

from pdf_editor_service.app import create_app
from pdf_editor_service.document_store import DocumentStore


def resolve_template_library_root() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "PDF Desktop Editor" / "template-library"
    return Path.home() / ".pdf-desktop-editor" / "template-library"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--host", default="127.0.0.1")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    backend_root = Path(__file__).resolve().parent
    service_base_url = f"http://{args.host}:{args.port}"
    store = DocumentStore(
        runtime_root=backend_root / "runtime",
        service_base_url=service_base_url,
        template_library_root=resolve_template_library_root(),
    )
    app = create_app(store)
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
