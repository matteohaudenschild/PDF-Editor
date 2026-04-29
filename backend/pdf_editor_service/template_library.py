from __future__ import annotations

import json
import math
import os
import re
from dataclasses import dataclass
from pathlib import Path

import pymupdf

from .document_templates import (
    DocumentTemplateSpec,
    TemplateFieldSpec,
    TemplateMarkerSpec,
    TemplatePageImageHashSpec,
    TemplatePageSizeSpec,
)
from .models import DocumentModel, TextBlock


TEMPLATE_LIBRARY_FILENAME = "learned-templates.json"


@dataclass(frozen=True)
class TemplateSaveResult:
    path: Path
    replaced_existing: bool


def default_template_library_root() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "PDF Desktop Editor" / "template-library"
    return Path.home() / ".pdf-desktop-editor" / "template-library"


def template_library_path(library_root: Path) -> Path:
    return library_root / TEMPLATE_LIBRARY_FILENAME


def load_user_templates(library_root: Path) -> tuple[DocumentTemplateSpec, ...]:
    path = template_library_path(library_root)
    if not path.exists():
        return ()

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return ()

    templates: list[DocumentTemplateSpec] = []
    for raw_template in payload.get("templates", []):
        try:
            templates.append(_template_from_dict(raw_template))
        except (KeyError, TypeError, ValueError):
            continue
    return tuple(templates)


def save_user_template(library_root: Path, template: DocumentTemplateSpec) -> TemplateSaveResult:
    library_root.mkdir(parents=True, exist_ok=True)
    path = template_library_path(library_root)
    existing_templates = list(load_user_templates(library_root))
    replaced_existing = False

    for index, existing in enumerate(existing_templates):
        if existing.id == template.id:
            existing_templates[index] = template
            replaced_existing = True
            break
    else:
        existing_templates.append(template)

    payload = {
        "version": 1,
        "templates": [_template_to_dict(entry) for entry in existing_templates],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return TemplateSaveResult(path=path, replaced_existing=replaced_existing)


def build_learned_template(source_path: Path, document: DocumentModel, name: str, description: str | None = None) -> DocumentTemplateSpec:
    display_name = str(name or "").strip()
    if not display_name:
        raise ValueError("Bitte einen Vorlagennamen angeben.")

    custom_blocks = sorted(
        (block for block in document.blocks if block.isCustom),
        key=lambda block: (block.page, block.bbox.y0, block.bbox.x0),
    )
    if not custom_blocks:
        raise ValueError("Es wurden noch keine manuellen Felder für die Vorlage gesetzt.")

    pages_by_number = {page.pageNumber: page for page in document.pages}
    field_specs: list[TemplateFieldSpec] = []
    for block in custom_blocks:
        page = pages_by_number.get(block.page)
        if page is None:
            continue
        field_specs.append(
            TemplateFieldSpec(
                page_number=block.page,
                source_page_width=round(page.width, 3),
                source_page_height=round(page.height, 3),
                x0=round(block.bbox.x0, 3),
                y0=round(block.bbox.y0, 3),
                x1=round(block.bbox.x1, 3),
                y1=round(block.bbox.y1, 3),
                font_family=block.fontFamily,
                font_key=block.fontKey,
                font_size=round(block.fontSize, 3),
                color=block.color,
                line_height=round(block.lineHeight, 3),
                align=block.align,
                rotation=round(block.rotation, 3),
                min_font_size=round(block.minFontSize, 3),
                css_font_family=block.cssFontFamily,
                font_asset_id=block.fontAssetId,
                font_weight=block.fontWeight,
                font_style=block.fontStyle,
                baseline=round(block.baseline, 3) if block.baseline is not None else None,
                is_checkbox=bool(block.isCheckbox),
                group_kind="generated-user-template-checkbox" if block.isCheckbox else "generated-user-template-field",
            )
        )

    if not field_specs:
        raise ValueError("Die manuellen Felder konnten nicht als Vorlage gespeichert werden.")

    document_markers = _select_document_markers(document)
    page_markers = _select_page_markers(document, document_markers)
    total_marker_count = len(document_markers) + len(page_markers)
    uses_marker_mode = total_marker_count >= 3 and sum(len(marker) for marker in document_markers) >= 24
    page_image_hashes = _build_page_image_hashes(source_path, max_pages=min(3, document.pageCount)) if not uses_marker_mode else ()

    minimum_marker_match_count = 0
    minimum_marker_match_ratio = 1.0
    if uses_marker_mode and total_marker_count:
        minimum_marker_match_count = max(2, math.ceil(total_marker_count * 0.55))
        minimum_marker_match_ratio = 0.55

    return DocumentTemplateSpec(
        id=f"user-{_slugify(display_name) or 'template'}",
        family="user_learned",
        kind="user_learned",
        description=description.strip() if description and description.strip() else f"Gelernte Vorlage: {display_name}",
        display_name=display_name,
        page_count=document.pageCount,
        page_sizes=tuple(
            TemplatePageSizeSpec(
                page_number=page.pageNumber,
                width=round(page.width, 3),
                height=round(page.height, 3),
            )
            for page in document.pages
        ),
        required_document_markers=document_markers,
        required_page_markers=page_markers,
        minimum_marker_match_count=minimum_marker_match_count,
        minimum_marker_match_ratio=minimum_marker_match_ratio,
        match_mode="markers" if uses_marker_mode else "image",
        page_image_hashes=page_image_hashes,
        learned_field_specs=tuple(field_specs),
        warning=f'Gelernte Vorlage "{display_name}" erkannt.',
    )


def _template_to_dict(template: DocumentTemplateSpec) -> dict:
    return {
        "id": template.id,
        "family": template.family,
        "kind": template.kind,
        "description": template.description,
        "displayName": template.display_name,
        "pageCount": template.page_count,
        "pageSizes": [
            {
                "pageNumber": page.page_number,
                "width": page.width,
                "height": page.height,
            }
            for page in template.page_sizes
        ],
        "requiredDocumentMarkers": list(template.required_document_markers),
        "requiredPageMarkers": [
            {
                "pageNumber": marker.page_number,
                "needle": marker.needle,
            }
            for marker in template.required_page_markers
        ],
        "minimumMarkerMatchCount": template.minimum_marker_match_count,
        "minimumMarkerMatchRatio": template.minimum_marker_match_ratio,
        "matchMode": template.match_mode,
        "pageImageHashes": [
            {
                "pageNumber": page.page_number,
                "hashHex": page.hash_hex,
                "maxDistance": page.max_distance,
            }
            for page in template.page_image_hashes
        ],
        "learnedFieldSpecs": [
            {
                "pageNumber": field.page_number,
                "sourcePageWidth": field.source_page_width,
                "sourcePageHeight": field.source_page_height,
                "x0": field.x0,
                "y0": field.y0,
                "x1": field.x1,
                "y1": field.y1,
                "fontFamily": field.font_family,
                "fontKey": field.font_key,
                "fontSize": field.font_size,
                "color": field.color,
                "lineHeight": field.line_height,
                "align": field.align,
                "rotation": field.rotation,
                "minFontSize": field.min_font_size,
                "cssFontFamily": field.css_font_family,
                "fontAssetId": field.font_asset_id,
                "fontWeight": field.font_weight,
                "fontStyle": field.font_style,
                "baseline": field.baseline,
                "isCheckbox": field.is_checkbox,
                "groupKind": field.group_kind,
            }
            for field in template.learned_field_specs
        ],
        "warning": template.warning,
    }


def _template_from_dict(raw_template: dict) -> DocumentTemplateSpec:
    return DocumentTemplateSpec(
        id=str(raw_template["id"]),
        family=str(raw_template.get("family") or "user_learned"),
        kind=str(raw_template.get("kind") or "user_learned"),
        description=str(raw_template.get("description") or "Gelernte Vorlage"),
        display_name=str(raw_template.get("displayName") or "") or None,
        page_count=int(raw_template["pageCount"]) if raw_template.get("pageCount") is not None else None,
        page_sizes=tuple(
            TemplatePageSizeSpec(
                page_number=int(page["pageNumber"]),
                width=float(page["width"]),
                height=float(page["height"]),
            )
            for page in raw_template.get("pageSizes", [])
        ),
        required_document_markers=tuple(str(marker) for marker in raw_template.get("requiredDocumentMarkers", [])),
        required_page_markers=tuple(
            TemplateMarkerSpec(
                page_number=int(marker["pageNumber"]),
                needle=str(marker["needle"]),
            )
            for marker in raw_template.get("requiredPageMarkers", [])
        ),
        minimum_marker_match_count=int(raw_template.get("minimumMarkerMatchCount") or 0),
        minimum_marker_match_ratio=float(raw_template.get("minimumMarkerMatchRatio") or 1.0),
        match_mode=str(raw_template.get("matchMode") or "markers"),
        page_image_hashes=tuple(
            TemplatePageImageHashSpec(
                page_number=int(page["pageNumber"]),
                hash_hex=str(page["hashHex"]),
                max_distance=int(page.get("maxDistance") or 24),
            )
            for page in raw_template.get("pageImageHashes", [])
        ),
        learned_field_specs=tuple(
            TemplateFieldSpec(
                page_number=int(field["pageNumber"]),
                source_page_width=float(field["sourcePageWidth"]),
                source_page_height=float(field["sourcePageHeight"]),
                x0=float(field["x0"]),
                y0=float(field["y0"]),
                x1=float(field["x1"]),
                y1=float(field["y1"]),
                font_family=str(field["fontFamily"]),
                font_key=str(field["fontKey"]),
                font_size=float(field["fontSize"]),
                color=str(field["color"]),
                line_height=float(field["lineHeight"]),
                align=str(field["align"]),
                rotation=float(field["rotation"]),
                min_font_size=float(field["minFontSize"]),
                css_font_family=str(field["cssFontFamily"]),
                font_asset_id=str(field["fontAssetId"]) if field.get("fontAssetId") else None,
                font_weight=str(field.get("fontWeight") or "400"),
                font_style=str(field.get("fontStyle") or "normal"),
                baseline=float(field["baseline"]) if field.get("baseline") is not None else None,
                is_checkbox=bool(field.get("isCheckbox")),
                group_kind=str(field.get("groupKind") or "generated-user-template-field"),
            )
            for field in raw_template.get("learnedFieldSpecs", [])
        ),
        warning=str(raw_template["warning"]) if raw_template.get("warning") else None,
    )


def _select_document_markers(document: DocumentModel) -> tuple[str, ...]:
    candidates: list[tuple[float, str]] = []
    for block in document.blocks:
        if block.isCustom or block.isCheckbox or block.page > min(3, document.pageCount):
            continue
        text = _normalize_marker_text(block.originalText)
        if not _is_marker_candidate(text):
            continue
        score = (min(len(text), 84) * 12.0) - (block.bbox.y0 * 0.08) - (block.bbox.x0 * 0.02)
        candidates.append((score, text))
    return _take_unique_markers(candidates, limit=4)


def _select_page_markers(document: DocumentModel, document_markers: tuple[str, ...]) -> tuple[TemplateMarkerSpec, ...]:
    blocked_markers = set(document_markers)
    page_markers: list[TemplateMarkerSpec] = []
    for page_number in range(1, min(3, document.pageCount) + 1):
        candidates: list[tuple[float, str]] = []
        for block in document.blocks:
            if block.page != page_number or block.isCustom or block.isCheckbox:
                continue
            text = _normalize_marker_text(block.originalText)
            if text in blocked_markers or not _is_marker_candidate(text):
                continue
            score = (min(len(text), 72) * 10.0) - (block.bbox.y0 * 0.06) - (block.bbox.x0 * 0.01)
            candidates.append((score, text))
        for marker in _take_unique_markers(candidates, limit=2):
            page_markers.append(TemplateMarkerSpec(page_number=page_number, needle=marker))
    return tuple(page_markers)


def _take_unique_markers(candidates: list[tuple[float, str]], *, limit: int) -> tuple[str, ...]:
    unique_markers: list[str] = []
    seen: set[str] = set()
    for _, text in sorted(candidates, key=lambda item: item[0], reverse=True):
        if text in seen:
            continue
        seen.add(text)
        unique_markers.append(text)
        if len(unique_markers) >= limit:
            break
    return tuple(unique_markers)


def _is_marker_candidate(text: str) -> bool:
    if len(text) < 4 or len(text) > 96:
        return False
    if text.count("_") >= 2:
        return False
    return sum(1 for char in text if char.isalpha()) >= 2


def _normalize_marker_text(value: str) -> str:
    normalized = re.sub(r"\s+", " ", str(value or "").strip())
    return normalized


def _slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.casefold()).strip("-")
    return slug[:72]


def _build_page_image_hashes(source_path: Path, *, max_pages: int) -> tuple[TemplatePageImageHashSpec, ...]:
    document = pymupdf.open(source_path)
    try:
        hashes: list[TemplatePageImageHashSpec] = []
        for page_index in range(min(max_pages, document.page_count)):
            page = document[page_index]
            hash_hex = _compute_page_image_hash(page)
            if not hash_hex:
                continue
            hashes.append(
                TemplatePageImageHashSpec(
                    page_number=page_index + 1,
                    hash_hex=hash_hex,
                    max_distance=28,
                )
            )
        return tuple(hashes)
    finally:
        document.close()


def _compute_page_image_hash(page: pymupdf.Page) -> str:
    scale_x = 16.0 / max(page.rect.width, 1.0)
    scale_y = 16.0 / max(page.rect.height, 1.0)
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale_x, scale_y), colorspace=pymupdf.csGRAY, alpha=False)
    values = list(pixmap.samples[: pixmap.width * pixmap.height])
    if not values:
        return ""
    average = sum(values) / len(values)
    bits = "".join("1" if value >= average else "0" for value in values)
    return f"{int(bits, 2):0{max(1, math.ceil(len(bits) / 4))}x}"
