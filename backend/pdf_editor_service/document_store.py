from __future__ import annotations

import shutil
from pathlib import Path

from .models import TextBlock
from .pdf_engine import DocumentSession, analyze_document, export_document, persist_draft, render_background_page
from .template_library import build_learned_template, load_user_templates, save_user_template


class DocumentStore:
    def __init__(self, runtime_root: Path, service_base_url: str, template_library_root: Path) -> None:
        self.runtime_root = runtime_root
        self.service_base_url = service_base_url
        self.template_library_root = template_library_root
        if self.runtime_root.exists():
            shutil.rmtree(self.runtime_root, ignore_errors=True)
        self.runtime_root.mkdir(parents=True, exist_ok=True)
        self._sessions: dict[str, DocumentSession] = {}

    def import_document(self, source_path: Path) -> DocumentSession:
        session = analyze_document(
            source_path,
            self.runtime_root,
            self.service_base_url,
            user_templates=load_user_templates(self.template_library_root),
        )
        stale_ids = [
            session_id
            for session_id, existing in self._sessions.items()
            if existing.source_path == session.source_path
        ]
        for session_id in stale_ids:
            previous = self._sessions.pop(session_id)
            if previous.work_dir.exists():
                shutil.rmtree(previous.work_dir, ignore_errors=True)
        self._sessions[session.model.id] = session
        return session

    def get(self, document_id: str) -> DocumentSession:
        try:
            return self._sessions[document_id]
        except KeyError as error:
            raise KeyError(f"Unbekanntes Dokument: {document_id}") from error

    def update_draft(self, document_id: str, updates: list[TextBlock]) -> Path:
        session = self.get(document_id)
        session.model.fields = updates
        return persist_draft(session)

    def learn_template(self, document_id: str, name: str, updates: list[TextBlock], description: str | None = None):
        session = self.get(document_id)
        session.model.fields = updates
        template = build_learned_template(session.source_path, session.model, name, description)
        save_result = save_user_template(self.template_library_root, template)
        return template, save_result

    def export(self, document_id: str, target_path: Path | None = None) -> Path:
        session = self.get(document_id)
        return export_document(session, target_path)

    def render_background(self, document_id: str, page_number: int, target_width: int | None = None) -> Path:
        session = self.get(document_id)
        return render_background_page(session, page_number, target_width)
