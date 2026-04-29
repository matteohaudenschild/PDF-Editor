from __future__ import annotations

import base64
from pathlib import Path

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from .document_store import DocumentStore
from .models import (
    DraftUpdateRequest,
    DraftUpdateResponse,
    ExportRequest,
    ExportResponse,
    ImportRequest,
    LearnTemplateRequest,
    LearnTemplateResponse,
    SupportReport,
    UploadImportRequest,
    WhiteboardExportRequest,
    WhiteboardExportResponse,
)
from .pdf_engine import export_whiteboard_pdf


def create_app(store: DocumentStore) -> FastAPI:
    app = FastAPI(title="PDF Editor Service")
    frontend_dir = Path(__file__).resolve().parents[2] / "app" / "src" / "renderer"
    frontend_available = frontend_dir.exists()

    def build_export_name(source_name: str, suffix: str = "-bearbeitet.pdf") -> str:
        base_name = Path(source_name or "dokument.pdf").name
        stem = Path(base_name).stem or "dokument"
        return f"{stem}{suffix}"

    def require_session(document_id: str):
        try:
            return store.get(document_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error

    if frontend_available:
        @app.get("/", include_in_schema=False)
        def open_frontend():
            return RedirectResponse(url="/app/")

    @app.get("/healthz")
    def healthz() -> dict[str, str]:
        return {"status": "ok"}

    @app.post("/documents/import")
    def import_document(payload: ImportRequest):
        source_path = Path(payload.sourcePath).expanduser()
        if not source_path.exists():
            raise HTTPException(status_code=404, detail="PDF-Datei wurde nicht gefunden.")
        session = store.import_document(source_path)
        return session.model

    @app.post("/documents/upload")
    def upload_document(payload: UploadImportRequest):
        uploads_dir = store.runtime_root / "uploads"
        uploads_dir.mkdir(parents=True, exist_ok=True)

        safe_name = Path(payload.fileName or "dokument.pdf").name or "dokument.pdf"
        upload_path = uploads_dir / safe_name
        counter = 1
        while upload_path.exists():
            upload_path = uploads_dir / f"{Path(safe_name).stem}-{counter}{Path(safe_name).suffix or '.pdf'}"
            counter += 1

        try:
            file_bytes = base64.b64decode(payload.fileDataBase64, validate=True)
        except ValueError as error:
            raise HTTPException(status_code=422, detail="Upload-Daten konnten nicht gelesen werden.") from error

        upload_path.write_bytes(file_bytes)
        session = store.import_document(upload_path)
        session.model.sourcePath = safe_name
        return session.model

    @app.get("/documents/{document_id}/pages/{page_number}/background")
    def get_background(document_id: str, page_number: int, width: int | None = Query(default=None, ge=1)):
        session = require_session(document_id)
        page_lookup = {page.pageNumber: page for page in session.model.pages}
        page = page_lookup.get(page_number)
        if not page:
            raise HTTPException(status_code=404, detail="Hintergrundbild nicht gefunden.")
        try:
            background_path = store.render_background(document_id, page_number, width)
        except ValueError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error
        return FileResponse(background_path, media_type="image/png")

    @app.get("/documents/{document_id}/fonts/{font_id}")
    def get_font(document_id: str, font_id: str):
        session = require_session(document_id)
        runtime = session.font_runtimes.get(font_id)
        if runtime is None or runtime.font_path is None:
            raise HTTPException(status_code=404, detail="Font nicht gefunden.")
        return FileResponse(runtime.font_path, media_type="application/octet-stream")

    @app.put("/documents/{document_id}/draft")
    def update_draft(document_id: str, payload: DraftUpdateRequest) -> DraftUpdateResponse:
        try:
            store.update_draft(document_id, payload.fields)
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error
        return DraftUpdateResponse(saved=True)

    @app.post("/documents/{document_id}/learn-template")
    def learn_template(document_id: str, payload: LearnTemplateRequest) -> LearnTemplateResponse:
        try:
            template, save_result = store.learn_template(
                document_id,
                payload.name,
                payload.fields,
                payload.description,
            )
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error
        except ValueError as error:
            raise HTTPException(status_code=422, detail=str(error)) from error
        return LearnTemplateResponse(
            saved=True,
            templateId=template.id,
            templateName=template.display_name or payload.name.strip(),
            fieldCount=len(template.learned_field_specs),
            templatePath=str(save_result.path),
            replacedExisting=save_result.replaced_existing,
        )

    @app.post("/documents/{document_id}/export")
    def export_document(document_id: str, payload: ExportRequest) -> ExportResponse:
        try:
            output_path = store.export(document_id, Path(payload.targetPath).expanduser() if payload.targetPath else None)
        except KeyError as error:
            raise HTTPException(status_code=404, detail=str(error)) from error
        except ValueError as error:
            raise HTTPException(status_code=422, detail=str(error)) from error
        return ExportResponse(exported=True, outputPath=str(output_path))

    @app.post("/documents/{document_id}/export-download")
    def export_document_download(document_id: str, payload: ExportRequest):
        session = require_session(document_id)
        downloads_dir = session.work_dir / "downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)
        target_path = downloads_dir / build_export_name(session.model.sourcePath)
        try:
            output_path = store.export(document_id, target_path)
        except ValueError as error:
            raise HTTPException(status_code=422, detail=str(error)) from error
        return FileResponse(
            output_path,
            media_type="application/pdf",
            filename=output_path.name,
        )

    @app.post("/whiteboard/export")
    def export_whiteboard(payload: WhiteboardExportRequest) -> WhiteboardExportResponse:
        try:
            output_path = export_whiteboard_pdf(
                image_data_url=payload.imageDataUrl,
                width=payload.width,
                height=payload.height,
                target_path=Path(payload.targetPath).expanduser() if payload.targetPath else None,
            )
        except ValueError as error:
            raise HTTPException(status_code=422, detail=str(error)) from error
        return WhiteboardExportResponse(exported=True, outputPath=str(output_path))

    @app.post("/whiteboard/export-download")
    def export_whiteboard_download(payload: WhiteboardExportRequest):
        downloads_dir = store.runtime_root / "whiteboard-downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)
        target_path = downloads_dir / build_export_name("whiteboard.pdf", suffix=".pdf")
        try:
            output_path = export_whiteboard_pdf(
                image_data_url=payload.imageDataUrl,
                width=payload.width,
                height=payload.height,
                target_path=target_path,
            )
        except ValueError as error:
            raise HTTPException(status_code=422, detail=str(error)) from error
        return FileResponse(
            output_path,
            media_type="application/pdf",
            filename=output_path.name,
        )

    @app.get("/documents/{document_id}/support-report")
    def support_report(document_id: str) -> SupportReport:
        session = require_session(document_id)
        return session.model.supportReport or SupportReport(
            documentClass=session.model.documentClass,
            supportMode=session.model.supportStatus.supportMode,
            reasons=session.model.supportStatus.reasons,
            warnings=session.model.supportStatus.warnings,
            reviewItems=session.model.reviewItems,
        )

    if frontend_available:
        app.mount("/app", StaticFiles(directory=frontend_dir, html=True), name="frontend")

    return app
