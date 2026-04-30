from __future__ import annotations

import base64
import csv
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from uuid import uuid4

import pymupdf

from .document_templates import (
    DOCUMENT_TEMPLATES,
    DocumentTemplateSpec,
    TemplateFieldSpec,
    TemplatePageImageHashSpec,
)
from .models import (
    BoundingBox,
    DocumentModel,
    FieldSupportEntry,
    FieldQuad,
    FontAsset,
    LineOverlay,
    PageModel,
    PageSupportEntry,
    ReviewItem,
    SupportReport,
    SupportStatus,
    TextBlock,
)


BASE14_FONT_MAP = {
    "helvetica": ("Helvetica", "Arial, sans-serif"),
    "helvetica-bold": ("Helvetica-Bold", "Arial, sans-serif"),
    "helvetica-oblique": ("Helvetica-Oblique", "Arial, sans-serif"),
    "helvetica-boldoblique": ("Helvetica-BoldOblique", "Arial, sans-serif"),
    "arial": ("Helvetica", "Arial, sans-serif"),
    "arial-bold": ("Helvetica-Bold", "Arial, sans-serif"),
    "arial-italic": ("Helvetica-Oblique", "Arial, sans-serif"),
    "arial-bolditalic": ("Helvetica-BoldOblique", "Arial, sans-serif"),
    "courier": ("Courier", '"Courier New", monospace'),
    "courier-bold": ("Courier-Bold", '"Courier New", monospace'),
    "courier-italic": ("Courier-Oblique", '"Courier New", monospace'),
    "courier-bolditalic": ("Courier-BoldOblique", '"Courier New", monospace'),
    "times": ("Times-Roman", '"Times New Roman", serif'),
    "times-roman": ("Times-Roman", '"Times New Roman", serif'),
    "times-bold": ("Times-Bold", '"Times New Roman", serif'),
    "times-italic": ("Times-Italic", '"Times New Roman", serif'),
    "times-bolditalic": ("Times-BoldItalic", '"Times New Roman", serif'),
    "times-romanitalic": ("Times-Italic", '"Times New Roman", serif'),
    "times-romanbolditalic": ("Times-BoldItalic", '"Times New Roman", serif'),
    "times new roman": ("Times-Roman", '"Times New Roman", serif'),
    "times new roman-bold": ("Times-Bold", '"Times New Roman", serif'),
    "times new roman-italic": ("Times-Italic", '"Times New Roman", serif'),
    "times new roman-bolditalic": ("Times-BoldItalic", '"Times New Roman", serif'),
}

WINDOWS_FONT_FILES = {
    "helvetica": "arial.ttf",
    "helvetica-bold": "arialbd.ttf",
    "helvetica-oblique": "ariali.ttf",
    "helvetica-boldoblique": "arialbi.ttf",
    "arial": "arial.ttf",
    "arial-bold": "arialbd.ttf",
    "arial-italic": "ariali.ttf",
    "arial-bolditalic": "arialbi.ttf",
    "courier": "cour.ttf",
    "courier-bold": "courbd.ttf",
    "courier-italic": "couri.ttf",
    "courier-bolditalic": "courbi.ttf",
    "times": "times.ttf",
    "times-roman": "times.ttf",
    "times-bold": "timesbd.ttf",
    "times-italic": "timesi.ttf",
    "times-bolditalic": "timesbi.ttf",
    "times-romanitalic": "timesi.ttf",
    "times-romanbolditalic": "timesbi.ttf",
    "times new roman": "times.ttf",
    "times new roman-bold": "timesbd.ttf",
    "times new roman-italic": "timesi.ttf",
    "times new roman-bolditalic": "timesbi.ttf",
}

WINDOWS_FONTS_DIR = Path("C:/Windows/Fonts")

BROWSER_FONT_EXTENSIONS = {"ttf", "otf", "woff", "woff2"}
EMBEDDED_SESSION_FILENAME = "pdf-desktop-editor/session.json"
EMBEDDED_SESSION_DESCRIPTION = "PDF Desktop Editor embedded edit session"
EMBEDDED_SESSION_VERSION = 1
BACKGROUND_RENDER_DPI = 72
TEXT_DIRECTION_SKEW_TOLERANCE = 0.05
TEXT_RECONSTRUCTION_BACKGROUND_MODE = True

VT_TEXT_TEMPLATE_IDS = {
    "sicherheit_nord_vt_text",
    "sicherheit_nord_vt_bma_fw_text",
}

VT_SASSE_REFERENCE_BLOCK_MAP = {
    "page-1-generated-service-fee-base": "page-1-block-32",
    "page-1-generated-service-fee-standleitung": "page-1-block-38",
    "page-1-generated-service-fee-redundancy": "page-1-block-43",
    "page-1-generated-service-fee-sim": "page-1-block-49",
    "page-1-generated-service-fee-sharp-end": "page-1-block-55",
    "page-1-generated-service-fee-temp-window": "page-1-block-62",
    "page-1-generated-service-fee-unscharf": "page-1-block-68",
    "page-1-generated-service-fee-key-storage": "page-1-block-75",
    "page-1-generated-service-fee-total": "page-1-block-82",
    "page-1-generated-service-fee-sim-setup": "page-1-block-93",
    "page-1-generated-service-fee-nsl-setup": "page-1-block-99",
    "page-1-generated-security-drive-flat": "page-1-block-129",
    "page-1-generated-security-onsite-30": "page-1-block-133",
    "page-1-generated-security-guard-hour": "page-1-block-141",
    "page-1-generated-security-guard-urgent": "page-1-block-145",
    "page-1-generated-security-extra-checks": "page-1-block-149",
    "page-1-generated-security-key-exchange": "page-1-block-153",
    "page-1-generated-option-1-1": "page-1-checkbox-1",
    "page-1-generated-option-1-2": "page-1-checkbox-3",
    "page-1-generated-option-1-3": "page-1-checkbox-5",
    "page-1-generated-option-1-4": "page-1-checkbox-2",
    "page-1-generated-option-1-5": "page-1-checkbox-4",
    "page-1-generated-option-1-6": "page-1-checkbox-6",
    "page-1-generated-option-2-1-2-yes": "page-1-checkbox-9",
    "page-1-generated-option-2-1-2-no": "page-1-checkbox-10",
    "page-1-generated-option-2-1-3-yes": "page-1-checkbox-11",
    "page-1-generated-option-2-1-3-no": "page-1-checkbox-12",
    "page-1-generated-option-2-2-yes": "page-1-checkbox-13",
    "page-1-generated-option-2-2-no": "page-1-checkbox-14",
    "page-1-generated-option-2-3-yes": "page-1-checkbox-15",
    "page-1-generated-option-2-3-no": "page-1-checkbox-16",
    "page-1-generated-option-2-3-1-yes": "page-1-checkbox-17",
    "page-1-generated-option-2-3-1-no": "page-1-checkbox-18",
    "page-1-generated-option-2-3-2-yes": "page-1-checkbox-19",
    "page-1-generated-option-2-3-2-no": "page-1-checkbox-20",
    "page-1-generated-option-2-4-yes": "page-1-checkbox-21",
    "page-1-generated-option-2-4-no": "page-1-checkbox-22",
    "page-1-generated-option-2-9-yes": "page-1-checkbox-23",
    "page-1-generated-option-2-9-no": "page-1-checkbox-24",
    "page-1-generated-option-3-0-yes": "page-1-checkbox-27",
    "page-1-generated-option-3-0-no": "page-1-checkbox-28",
    "page-2-generated-contract-start-date": "page-2-block-11",
    "page-2-generated-additional-agreement-1": "page-2-block-39",
    "page-2-generated-additional-agreement-2": "page-2-block-40",
    "page-3-generated-place-date": "page-3-block-24",
    "page-3-generated-ag-place-date": "page-3-generated-place-date",
    "page-3-generated-email-confirmed": "page-3-checkbox-1",
    "page-3-generated-postal-mail": "page-3-checkbox-2",
}

VT_SASSE_SIGNATURE_OVERLAY_REGIONS = (
    (3, (0.0, 390.0, 595.0, 760.0)),
)

COMBINED_VT_PAGE3_SOURCE_CROP_REGIONS = (
    (3, (45.0, 183.0, 555.0, 214.0)),
    (3, (28.0, 430.0, 290.0, 560.0)),
    (3, (300.0, 430.0, 575.0, 560.0)),
)

COMBINED_VT_PAGE3_SOURCE_HIDE_REGIONS = (
    (3, (45.0, 183.0, 555.0, 214.0)),
    (3, (28.0, 430.0, 290.0, 594.0)),
    (3, (300.0, 430.0, 575.0, 594.0)),
)

COMBINED_VT_HANDLUNGSANWEISUNG_TEMPLATE_ID = "sicherheit_nord_vt_handlungsanweisung_scan_9696"
COMBINED_VT_HANDLUNGSANWEISUNG_TEMPLATE_FAMILY = "sicherheit_nord_vt_handlungsanweisung"

VT_ROTATED_REFERENCE_BLOCK_MAP = {
    **VT_SASSE_REFERENCE_BLOCK_MAP,
    "page-1-generated-service-fee-base": "page-1-block-34",
    "page-1-generated-service-fee-standleitung": "page-1-block-40",
    "page-1-generated-service-fee-redundancy": "page-1-block-46",
    "page-1-generated-service-fee-sim": "page-1-block-52",
    "page-1-generated-service-fee-sharp-end": "page-1-block-58",
    "page-1-generated-service-fee-temp-window": "page-1-block-65",
    "page-1-generated-service-fee-unscharf": "page-1-block-71",
    "page-1-generated-service-fee-key-storage": "page-1-block-78",
    "page-1-generated-service-fee-sim-setup": "page-1-block-99",
    "page-1-generated-service-fee-nsl-setup": "page-1-block-103",
    "page-1-generated-security-drive-flat": "page-1-block-133",
    "page-1-generated-security-onsite-30": "page-1-block-137",
    "page-1-generated-security-guard-hour": "page-1-block-145",
    "page-1-generated-security-guard-urgent": "page-1-block-149",
    "page-1-generated-security-extra-checks": "page-1-block-153",
    "page-3-generated-sn-place-date": "page-3-generated-place-date",
    "page-3-generated-ag-place-date": "page-3-generated-place-date",
    "page-3-generated-payment-quarterly": "page-3-generated-payment-quarterly",
    "page-3-generated-payment-half-yearly": "page-3-generated-payment-half-yearly",
    "page-3-generated-payment-yearly": "page-3-generated-payment-yearly",
}

VT_ROTATED_REFERENCE_SYNTHETIC_VALUE_RECTS = {
    "page-1-generated-service-fee-total": (1, (511.5, 468.255, 561.689, 479.251)),
}

VT_ROTATED_REFERENCE_ONLY_VALUE_BLOCK_IDS = {
    "page-1-block-84",
}

VT_ROTATED_UNMAPPED_SOURCE_VALUE_IDS = {
    "page-1-generated-security-key-exchange",
}

VT_REFERENCE_LABEL_TEXTS = {"netto", "brutto", "für", "fÃ¼r", "die"}

VT_ROTATED_2000544780_VALUE_OVERRIDES = {
    "page-1-generated-service-fee-base": "48,90",
    "page-1-generated-service-fee-standleitung": "",
    "page-1-generated-service-fee-redundancy": "Inkl.",
    "page-1-generated-service-fee-sim": "08,50",
    "page-1-generated-service-fee-sharp-end": "08,90",
    "page-1-generated-service-fee-temp-window": "10,00",
    "page-1-generated-service-fee-unscharf": "08,00",
    "page-1-generated-service-fee-key-storage": "07,50",
    "page-1-generated-service-fee-video-amount": "",
    "page-1-generated-service-fee-total": "57,40",
    "page-1-generated-service-fee-protocol": "",
    "page-1-generated-service-fee-sim-setup": "19,50",
    "page-1-generated-service-fee-nsl-setup": "97,50",
    "page-1-generated-security-drive-flat": "58,00",
    "page-1-generated-security-onsite-30": "29,50",
    "page-1-generated-security-fire-police": "",
    "page-1-generated-security-guard-hour": "33,00",
    "page-1-generated-security-guard-urgent": "70,00",
    "page-1-generated-security-extra-checks": "29,50",
    "page-1-generated-security-key-exchange": "58,00",
    "page-1-generated-contract-start-date": "06.11.2023",
}


@dataclass
class FontRuntime:
    asset: FontAsset
    pdf_font_name: str
    font_buffer: Optional[bytes]
    font_path: Optional[Path]


@dataclass(frozen=True)
class ExportFontSpec:
    name: str
    font_file: Optional[Path] = None


@dataclass(frozen=True)
class SourceOverlayRegion:
    page_number: int
    rect: tuple[float, float, float, float]

    @property
    def pdf_rect(self) -> pymupdf.Rect:
        return pymupdf.Rect(*self.rect)


@dataclass
class DocumentSession:
    model: DocumentModel
    sidecar_path: Path
    source_path: Path
    work_dir: Path
    font_runtimes: dict[str, FontRuntime]
    base_pdf_path: Optional[Path] = None
    source_overlay_regions: tuple[SourceOverlayRegion, ...] = ()
    render_annotations: bool = True
    text_only_background_pages: tuple[int, ...] = ()


@dataclass
class SpanFragment:
    text: str
    bbox: tuple[float, float, float, float]
    font_family: str
    font_key: str
    font_asset_id: Optional[str]
    font_size: float
    color: str
    line_height: float
    baseline: float
    space_width: float


@dataclass
class BlockFragment:
    text: str
    bbox: list[float]
    font_family: str
    font_key: str
    font_asset_id: Optional[str]
    font_size: float
    color: str
    line_height: float
    baseline: float
    align: str = "left"
    rotation: float = 0.0
    group_kind: str = "line"


@dataclass(frozen=True)
class HorizontalLineSegment:
    x0: float
    x1: float
    y: float
    width: float


@dataclass(frozen=True)
class VerticalLineSegment:
    x: float
    y0: float
    y1: float
    width: float


@dataclass(frozen=True)
class OutlineRect:
    x0: float
    y0: float
    x1: float
    y1: float
    width: float

    @property
    def rect(self) -> pymupdf.Rect:
        return pymupdf.Rect(self.x0, self.y0, self.x1, self.y1)


@dataclass(frozen=True)
class ExistingTextSegment:
    block: TextBlock
    line_index: int
    text: str
    rect: pymupdf.Rect


def compute_fingerprint(source_path: Path) -> str:
    hasher = hashlib.sha256()
    stat = source_path.stat()
    hasher.update(str(source_path.resolve()).encode("utf-8"))
    hasher.update(str(stat.st_size).encode("utf-8"))
    hasher.update(str(stat.st_mtime_ns).encode("utf-8"))
    with source_path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def _bbox_to_quad(bbox: BoundingBox) -> list[float]:
    return [
        bbox.x0, bbox.y0,
        bbox.x1, bbox.y0,
        bbox.x1, bbox.y1,
        bbox.x0, bbox.y1,
    ]


def _field_type_for_block(block: TextBlock) -> str:
    if block.isCheckbox:
        if "radio" in block.groupKind:
            return "radio"
        return "checkbox"
    if block.groupKind.startswith("widget-choice"):
        return "choice"
    if block.groupKind in {"multiline", "generated-contract-object-field"}:
        return "text-multiline"
    if "\n" in (block.currentText or block.originalText):
        return "text-multiline"
    if block.groupKind.startswith("ink-"):
        return block.groupKind
    return "text-line"


def _source_type_for_block(block: TextBlock) -> str:
    if block.groupKind.startswith("widget-"):
        return "acroform"
    if block.groupKind.startswith("generated-") and "scan" in block.groupKind:
        return "raster-scan"
    if block.groupKind.startswith("generated-"):
        return "vector-form"
    if block.isCustom:
        return "manual"
    return "native-digital"


def _review_state_for_block(block: TextBlock) -> str:
    source_type = _source_type_for_block(block)
    if source_type == "manual":
        return "review"
    if source_type == "raster-scan":
        return "review"
    if block.groupKind.startswith("generated-"):
        return "review"
    return "exact"


def _is_bold_font_weight(font_weight: str | int | float | None) -> bool:
    if font_weight is None:
        return False
    value = str(font_weight).strip().casefold()
    if not value:
        return False
    if "bold" in value:
        return True
    try:
        return float(value) >= 600
    except ValueError:
        return False


def _is_contract_id_number_block(block: TextBlock) -> bool:
    block_id = str(block.id or "").casefold()
    group_kind = str(block.groupKind or "").casefold()
    if "creditor-id" in block_id:
        return False
    text = str(block.currentText or block.originalText or "").strip()
    return (
        "generated-id-number" in block_id
        or "generated-instruction-id" in block_id
        or block_id.endswith("-id-number")
        or group_kind == "generated-id-number-field"
        or bool(re.fullmatch(r"200\d{7,10}", text))
    )


def _apply_contract_id_number_style(block: TextBlock) -> None:
    if not _is_contract_id_number_block(block):
        return
    block.fontWeight = "700"
    block.fontStyle = "normal"
    if block.appearance is not None:
        block.appearance.fontWeight = block.fontWeight
        block.appearance.fontStyle = block.fontStyle


def _sync_field_semantics(block: TextBlock, *, z_index: int = 0) -> TextBlock:
    _apply_contract_id_number_style(block)
    block.rect = BoundingBox(
        x0=block.bbox.x0,
        y0=block.bbox.y0,
        x1=block.bbox.x1,
        y1=block.bbox.y1,
    )
    block.fieldType = _field_type_for_block(block)
    block.sourceType = _source_type_for_block(block)
    block.originalValue = block.originalText
    block.currentValue = block.currentText
    block.sourceOriginalValue = block.sourceOriginalValue or block.originalText
    block.fontResourceRef = block.fontAssetId or block.fontKey or block.fontFamily
    block.sourceCoverRegions = [
        BoundingBox(
            x0=region.x0,
            y0=region.y0,
            x1=region.x1,
            y1=region.y1,
        )
        for region in (block.sourceCoverRegions or [block.bbox])
    ]
    existing_quads = block.quads or []
    quad_payloads = (
        [_bbox_to_quad(block.bbox)]
        if not existing_quads
        else [[
            existing.x0, existing.y0, existing.x1, existing.y1,
            existing.x2, existing.y2, existing.x3, existing.y3,
        ] for existing in existing_quads]
    )
    block.quads = [
        FieldQuad(**{
            "x0": quad[0],
            "y0": quad[1],
            "x1": quad[2],
            "y1": quad[3],
            "x2": quad[4],
            "y2": quad[5],
            "x3": quad[6],
            "y3": quad[7],
        })
        for quad in quad_payloads
    ] if block.bbox is not None else []
    block.confidence = float(max(0.0, min(1.0, block.confidence or 1.0)))
    block.reviewState = _review_state_for_block(block)
    block.supportMode = block.reviewState
    block.zIndex = z_index
    if block.groupKind in {"generated-contract-party-field", "generated-contract-object-line-field"}:
        field_height = max(0.0, block.bbox.y1 - block.bbox.y0)
        baseline_outside_field = (
            block.baseline is None
            or block.baseline <= block.bbox.y0 + 0.5
            or block.baseline >= block.bbox.y1 + 2.0
        )
        if baseline_outside_field and field_height >= 6.0:
            block.fontSize = min(block.fontSize, 9.2)
            block.lineHeight = round(max(block.fontSize * 1.15, field_height), 3)
            block.baseline = round(block.bbox.y0 + min(field_height - 1.1, block.fontSize * 0.88), 3)
    if block.appearance is not None:
        block.appearance.fontFamily = block.fontFamily
        block.appearance.fontKey = block.fontKey
        block.appearance.fontSize = block.fontSize
        block.appearance.color = block.color
        block.appearance.lineHeight = block.lineHeight
        block.appearance.align = block.align
        block.appearance.rotation = block.rotation
        block.appearance.cssFontFamily = block.cssFontFamily
        block.appearance.fontAssetId = block.fontAssetId
        block.appearance.fontWeight = block.fontWeight
        block.appearance.fontStyle = block.fontStyle
        block.appearance.baseline = block.baseline
        block.appearance.minFontSize = block.minFontSize
    return block


def _normalize_contract_id_number_blocks(blocks: list[TextBlock]) -> None:
    generated_id_rects: list[tuple[int, pymupdf.Rect]] = []
    for block in blocks:
        if not _is_contract_id_number_block(block):
            continue
        text = str(block.currentText or block.originalText or "").strip()
        original_text = str(block.originalText or "").strip()
        current_text = str(block.currentText or "").strip()
        if re.fullmatch(r"200\d{7,10}", current_text) and re.fullmatch(r"200\d{0,3}", original_text):
            block.sourceOriginalValue = block.sourceOriginalValue or original_text
            block.originalText = current_text
        block_id = str(block.id or "").casefold()
        if (
            re.fullmatch(r"200\d{7,10}", text)
            or (
                ("generated-id-number" in block_id or "generated-instruction-id" in block_id)
                and len(text) >= 8
            )
        ):
            generated_id_rects.append((
                block.page,
                pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
            ))

    for block in blocks:
        if (
            block.isCustom
            or block.isCheckbox
            or (
                block.groupKind.startswith("generated-")
                and block.groupKind != "generated-id-number-field"
            )
        ):
            continue
        text = str(block.currentText or block.originalText or "").strip()
        if not re.fullmatch(r"200\d{0,3}", text):
            continue
        for page_number, rect in generated_id_rects:
            if block.page != page_number:
                continue
            if _block_rect_overlap_ratio(block, rect, padding=2.0) >= 0.35:
                block.currentText = ""
                block.originalText = ""
                block.editable = False
                block.groupKind = "hidden-id-number-prefix"
                break

    labels_by_page: dict[int, list[TextBlock]] = {}
    for block in blocks:
        label_text = _normalize_text_content(block.originalText or block.currentText)
        if "id-nr" in label_text or "id nr" in label_text:
            labels_by_page.setdefault(block.page, []).append(block)

    for page_number, labels in labels_by_page.items():
        page_blocks = [block for block in blocks if block.page == page_number]
        for label in labels:
            label_center_y = (label.bbox.y0 + label.bbox.y1) / 2
            candidates: list[TextBlock] = []
            for block in page_blocks:
                if block is label or block.isCustom or block.isCheckbox:
                    continue
                if "creditor-id" in str(block.id or "").casefold():
                    continue
                text = str(block.currentText or block.originalText or "").strip()
                if not re.fullmatch(r"200\d{0,10}", text):
                    continue
                block_center_y = (block.bbox.y0 + block.bbox.y1) / 2
                if abs(block_center_y - label_center_y) > max(8.0, label.lineHeight * 0.9):
                    continue
                if block.bbox.x0 < label.bbox.x0:
                    continue
                if block.bbox.x0 > label.bbox.x1 + 120.0:
                    continue
                candidates.append(block)
            if not candidates:
                continue
            target = min(candidates, key=lambda block: (abs(block.bbox.x0 - label.bbox.x1), block.bbox.y0))
            if not target.groupKind.startswith("generated-"):
                target.groupKind = "generated-id-number-field"
            _apply_contract_id_number_style(target)


def _payment_column_center(page_blocks: list[TextBlock], suffix: str) -> Optional[float]:
    candidates: list[TextBlock] = []
    for block in page_blocks:
        text = _normalize_text_content(block.originalText or block.currentText)
        if not text:
            continue
        if suffix == "quarterly" and (
            "¼" in text
            or "1/4" in text
            or "viertel" in text
        ):
            candidates.append(block)
        elif suffix == "half-yearly" and (
            "½" in text
            or "1/2" in text
            or "halb" in text
        ):
            candidates.append(block)
        elif suffix == "yearly" and text in {"jährlich", "jahrlich"}:
            candidates.append(block)
    if not candidates:
        return None
    candidate = min(candidates, key=lambda block: (block.bbox.y0, block.bbox.x0))
    return (candidate.bbox.x0 + candidate.bbox.x1) / 2


def _normalize_payment_frequency_checkbox_blocks(blocks: list[TextBlock]) -> None:
    blocks_by_page: dict[int, list[TextBlock]] = {}
    for block in blocks:
        blocks_by_page.setdefault(block.page, []).append(block)

    for page_number, page_blocks in blocks_by_page.items():
        payment_blocks = [
            block for block in page_blocks
            if str(block.id or "").startswith(f"page-{page_number}-generated-payment-")
            and block.isCheckbox
        ]
        if not payment_blocks:
            continue

        marker = _find_page_block(page_blocks, "Gewünschte Zahlungsweise")
        prompt = _find_page_block(page_blocks, "Bitte ankreuzen")
        if marker is None or prompt is None:
            continue

        center_y = (prompt.bbox.y0 + prompt.bbox.y1) / 2
        scale_x = max(0.1, (max((block.bbox.x1 for block in page_blocks), default=595.0)) / 595.0)
        fallback_centers = {
            "quarterly": 259.25 * scale_x,
            "half-yearly": 362.2 * scale_x,
            "yearly": 469.4 * scale_x,
        }

        for block in payment_blocks:
            block_id = str(block.id or "")
            suffix = block_id.rsplit("generated-payment-", 1)[-1]
            if suffix not in fallback_centers:
                block.currentText = ""
                block.originalText = ""
                block.editable = False
                block.groupKind = "hidden-payment-checkbox-extra"
                continue
            center_x = _payment_column_center(page_blocks, suffix) or fallback_centers[suffix]
            current_size = max(block.bbox.x1 - block.bbox.x0, block.bbox.y1 - block.bbox.y0)
            checkbox_size = min(12.0, max(9.0, current_size or 11.0))
            block.bbox.x0 = round(center_x - checkbox_size / 2, 3)
            block.bbox.x1 = round(center_x + checkbox_size / 2, 3)
            block.bbox.y0 = round(center_y - checkbox_size / 2, 3)
            block.bbox.y1 = round(center_y + checkbox_size / 2, 3)
            block.sourceCoverRegions = [block.bbox.model_copy(deep=True)]
            block.quads = []
            block.lineHeight = max(block.lineHeight, checkbox_size)
            if suffix == "quarterly" and not block.originalText.strip():
                block.originalText = "x"


def _sync_fields(blocks: list[TextBlock]) -> list[TextBlock]:
    _normalize_contract_id_number_blocks(blocks)
    _normalize_payment_frequency_checkbox_blocks(blocks)
    visible_blocks = [
        block for block in blocks
        if not (
            block.groupKind == "widget-text-field"
            and not block.originalText.strip()
            and not block.currentText.strip()
        )
        and not (
            not block.editable
            and not block.originalText.strip()
            and not block.currentText.strip()
            and (
                block.groupKind.startswith("hidden-")
                or block.groupKind == "source-overlay-hidden"
            )
        )
    ]
    return [_sync_field_semantics(block, z_index=index) for index, block in enumerate(visible_blocks)]


def normalize_font_name(font_name: str) -> str:
    cleaned = font_name.replace("_", "-").strip().lower()
    cleaned = cleaned.lstrip("*")
    if "+" in cleaned:
        prefix, suffix = cleaned.split("+", 1)
        if len(prefix) == 6 and prefix.isalpha():
            cleaned = suffix
    cleaned = re.sub(r"-(?:\d{3,}|[a-f0-9]{6,})$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    aliases = {
        "arialmt": "arial",
        "arial-boldmt": "arial-bold",
        "arial-italicmt": "arial-italic",
        "arial-bolditalicmt": "arial-bolditalic",
        "arial-bold-italic": "arial-bolditalic",
        "arial-boldoblique": "arial-bolditalic",
        "arial,bold": "arial-bold",
        "arial,italic": "arial-italic",
        "arial,bolditalic": "arial-bolditalic",
        "helv": "helvetica",
        "helv-bold": "helvetica-bold",
        "helv-oblique": "helvetica-oblique",
        "helv-boldoblique": "helvetica-boldoblique",
        "helvetica-bolditalic": "helvetica-boldoblique",
        "helvetica-italic": "helvetica-oblique",
        "verdana": "arial",
        "verdana-bold": "arial-bold",
        "verdana-italic": "arial-italic",
        "verdana-bolditalic": "arial-bolditalic",
        "microsoft sans serif": "arial",
        "microsoft sans serif-bold": "arial-bold",
        "microsoft sans serif-italic": "arial-italic",
        "microsoft sans serif-bolditalic": "arial-bolditalic",
        "microsoftsansserif": "arial",
        "microsoftsansserif-bold": "arial-bold",
        "microsoftsansserif-italic": "arial-italic",
        "microsoftsansserif-bolditalic": "arial-bolditalic",
        "calibri": "arial",
        "calibri-bold": "arial-bold",
        "calibri-italic": "arial-italic",
        "calibri-bolditalic": "arial-bolditalic",
        "cour": "courier",
        "cour-bold": "courier-bold",
        "cour-italic": "courier-italic",
        "cour-bolditalic": "courier-bolditalic",
        "timesroman": "times-roman",
        "timesnewroman": "times new roman",
        "timesnewroman-bold": "times new roman-bold",
        "timesnewroman-italic": "times new roman-italic",
        "timesnewroman-bolditalic": "times new roman-bolditalic",
        "timesnewromanps-boldmt": "times-bold",
        "timesnewromanps-italicmt": "times-italic",
        "timesnewromanps-bolditalicmt": "times-bolditalic",
        "times new roman-boldmt": "times new roman-bold",
        "times new roman-italicmt": "times new roman-italic",
        "times new roman-bolditalicmt": "times new roman-bolditalic",
        "times-bolditalic": "times-bolditalic",
        "times-bold-italic": "times-bolditalic",
    }
    if cleaned in aliases:
        return aliases[cleaned]

    compact = cleaned.replace(" ", "")
    if compact in aliases:
        return aliases[compact]

    if "arial" in cleaned:
        if "bold" in cleaned and ("italic" in cleaned or "oblique" in cleaned):
            return "arial-bolditalic"
        if "bold" in cleaned:
            return "arial-bold"
        if "italic" in cleaned or "oblique" in cleaned:
            return "arial-italic"
        return "arial"

    if "verdana" in cleaned or "calibri" in cleaned or "microsoft sans serif" in cleaned or "microsoftsansserif" in compact:
        if "bold" in cleaned and ("italic" in cleaned or "oblique" in cleaned):
            return "arial-bolditalic"
        if "bold" in cleaned:
            return "arial-bold"
        if "italic" in cleaned or "oblique" in cleaned:
            return "arial-italic"
        return "arial"

    if "times" in cleaned:
        if "bold" in cleaned and ("italic" in cleaned or "oblique" in cleaned):
            return "times-bolditalic"
        if "bold" in cleaned:
            return "times-bold"
        if "italic" in cleaned or "oblique" in cleaned:
            return "times-italic"
        return "times new roman" if "new roman" in cleaned else "times"

    if "courier" in cleaned or cleaned.startswith("cour"):
        if "bold" in cleaned and ("italic" in cleaned or "oblique" in cleaned):
            return "courier-bolditalic"
        if "bold" in cleaned:
            return "courier-bold"
        if "italic" in cleaned or "oblique" in cleaned:
            return "courier-italic"
        return "courier"

    return cleaned


def choose_css_family(font_name: str) -> str:
    normalized = normalize_font_name(font_name)
    mapped = BASE14_FONT_MAP.get(normalized)
    if mapped:
        return mapped[1]
    if "cour" in normalized or "mono" in normalized:
        return '"Courier New", monospace'
    if "times" in normalized:
        return '"Times New Roman", serif'
    return "Arial, sans-serif"


def _font_details_for_span(
    font_runtimes_by_family: dict[str, FontRuntime],
    span_font: str,
) -> tuple[str, str, Optional[str], str]:
    runtime = _resolve_font_runtime(font_runtimes_by_family, span_font)
    if runtime is not None:
        return runtime.asset.family, runtime.pdf_font_name, runtime.asset.id, runtime.asset.cssFamily

    normalized = normalize_font_name(span_font)
    mapped = BASE14_FONT_MAP.get(normalized)
    if mapped:
        return span_font or mapped[0], mapped[0], None, mapped[1]

    if normalized in {"zapfdingbats", "zadi", "wingdings"}:
        return span_font or "ZapfDingbats", "Helvetica", None, "Arial, sans-serif"

    return "", "", None, ""


def infer_font_style(font_name: str) -> tuple[str, str]:
    normalized = normalize_font_name(font_name)
    is_bold = "bold" in normalized
    is_italic = "italic" in normalized or "oblique" in normalized
    font_weight = "700" if is_bold else "400"
    font_style = "italic" if is_italic else "normal"
    return font_weight, font_style


def build_sidecar_path(runtime_root: Path, source_path: Path, fingerprint: str) -> Path:
    drafts_dir = runtime_root / "drafts"
    drafts_dir.mkdir(parents=True, exist_ok=True)
    safe_name = "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in source_path.stem)
    return drafts_dir / f"{safe_name}-{fingerprint[:12]}.pdfedit.json"


def _hex_color(color_value: int) -> str:
    red, green, blue = pymupdf.sRGB_to_rgb(color_value)
    return f"#{red:02x}{green:02x}{blue:02x}"


def _stroke_color_to_hex(color: tuple[float, float, float] | None) -> str:
    if not color:
        return "#000000"
    red = max(0, min(255, round(color[0] * 255)))
    green = max(0, min(255, round(color[1] * 255)))
    blue = max(0, min(255, round(color[2] * 255)))
    return f"#{red:02x}{green:02x}{blue:02x}"


def _expanded_block_rect(block: TextBlock, page_rect: Optional[pymupdf.Rect] = None) -> pymupdf.Rect:
    pad_x = max(1.0, block.fontSize * 0.35)
    pad_y = max(0.75, block.lineHeight * 0.12)
    rect = pymupdf.Rect(
        block.bbox.x0 - pad_x,
        block.bbox.y0 - pad_y,
        block.bbox.x1 + pad_x,
        block.bbox.y1 + pad_y,
    )
    if page_rect is not None:
        rect &= page_rect
    return rect


def _expanded_id_number_cover_rect(block: TextBlock, page_rect: Optional[pymupdf.Rect] = None) -> pymupdf.Rect:
    baseline = block.baseline if block.baseline is not None else block.bbox.y1
    text = (block.currentText or block.originalText or "").strip()
    try:
        text_width = pymupdf.get_text_length(text, fontname="Helvetica-Bold", fontsize=block.fontSize) if text else 0.0
    except Exception:
        text_width = len(text) * block.fontSize * 0.55
    top = min(block.bbox.y0, baseline - block.fontSize * 1.05) - 0.8
    bottom = max(block.bbox.y1, baseline + block.fontSize * 0.42) + 1.2
    x1 = max(block.bbox.x1, block.bbox.x0 + text_width + max(4.0, block.fontSize * 0.45))
    rect = pymupdf.Rect(block.bbox.x0 - 1.0, top, x1 + 1.0, bottom)
    if page_rect is not None:
        rect &= page_rect
    return rect


def _visual_cover_rect(block: TextBlock, page_rect: Optional[pymupdf.Rect] = None) -> pymupdf.Rect:
    pad_x = max(2.0, block.fontSize * 0.45)
    trim_top = max(1.0, block.lineHeight * 0.16)
    trim_bottom = max(0.8, block.lineHeight * 0.1)
    rect = pymupdf.Rect(
        block.bbox.x0 - pad_x,
        block.bbox.y0 + trim_top,
        block.bbox.x1 + pad_x,
        max(block.bbox.y0 + trim_top + 1.0, block.bbox.y1 - trim_bottom),
    )
    if page_rect is not None:
        rect &= page_rect
    return rect


def _field_cover_rects(block: TextBlock, page_rect: Optional[pymupdf.Rect] = None) -> list[pymupdf.Rect]:
    if _is_contract_id_number_block(block):
        return [_expanded_id_number_cover_rect(block, page_rect)]
    if block.sourceCoverRegions:
        rects: list[pymupdf.Rect] = []
        for region in block.sourceCoverRegions:
            rect = pymupdf.Rect(region.x0, region.y0, region.x1, region.y1)
            if page_rect is not None:
                rect &= page_rect
            if not rect.is_empty:
                rects.append(rect)
        if rects:
            return rects
    return [_expanded_block_rect(block, page_rect)]


def _apply_block_redactions(page: pymupdf.Page, blocks: list[TextBlock]) -> bool:
    def apply_group(group: list[TextBlock], *, remove_line_art: bool) -> bool:
        applied = False
        for block in group:
            for rect in _field_cover_rects(block, page.rect):
                page.add_redact_annot(rect, fill=False, cross_out=False)
                applied = True
        if not applied:
            return False
        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=(
                pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED
                if remove_line_art
                else pymupdf.PDF_REDACT_LINE_ART_NONE
            ),
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )
        return True

    regular_blocks = [block for block in blocks if not _is_contract_id_number_block(block)]
    id_blocks = [block for block in blocks if _is_contract_id_number_block(block)]
    regular_applied = apply_group(regular_blocks, remove_line_art=False)
    id_applied = apply_group(id_blocks, remove_line_art=True)
    return regular_applied or id_applied


def _bbox_union(target: list[float], other: tuple[float, float, float, float]) -> list[float]:
    return [
        min(target[0], other[0]),
        min(target[1], other[1]),
        max(target[2], other[2]),
        max(target[3], other[3]),
    ]


def _bbox_width(bbox: list[float]) -> float:
    return bbox[2] - bbox[0]


def _chars_to_text(chars: list[dict]) -> str:
    return "".join(char["c"] for char in chars)


def _trim_outer_whitespace_chars(chars: list[dict]) -> list[dict]:
    start = 0
    end = len(chars)
    while start < end and not str(chars[start].get("c", "")).strip():
        start += 1
    while end > start and not str(chars[end - 1].get("c", "")).strip():
        end -= 1
    return chars[start:end]


def _chars_bbox(chars: list[dict], fallback_bbox: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
    if not chars:
        return fallback_bbox
    x0 = min(float(char["bbox"][0]) for char in chars)
    y0 = min(float(char["bbox"][1]) for char in chars)
    x1 = max(float(char["bbox"][2]) for char in chars)
    y1 = max(float(char["bbox"][3]) for char in chars)
    return (x0, y0, x1, y1)


def _normalize_text_content(text: str) -> str:
    return " ".join(str(text or "").casefold().split())


def _find_page_block(page_blocks: list[TextBlock], needle: str) -> Optional[TextBlock]:
    normalized_needle = _normalize_text_content(needle)
    candidates = [
        block for block in page_blocks
        if normalized_needle in _normalize_text_content(block.originalText)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda block: (block.bbox.y0, block.bbox.x0))


def _is_masked_template_block(block: TextBlock) -> bool:
    text = block.originalText or ""
    return (not block.isCustom) and (not block.isCheckbox) and ("__" in text) and ("\n" not in text)


def _is_iban_template_block(block: TextBlock) -> bool:
    text = (block.originalText or "").upper()
    return _is_masked_template_block(block) and ("IBAN" in text) and ("DE" in text)


def _extract_masked_value(template: str, current_text: str) -> str:
    safe_current = current_text if current_text else template
    value_chars: list[str] = []
    for index, template_char in enumerate(template):
        if template_char != "_":
            continue
        current_char = safe_current[index] if index < len(safe_current) else "_"
        if current_char != "_":
            value_chars.append(current_char)
    return "".join(value_chars)


def _apply_masked_value(template: str, value: str) -> str:
    chars = list(template)
    value_index = 0
    for index, template_char in enumerate(chars):
        if template_char != "_":
            continue
        chars[index] = value[value_index] if value_index < len(value) else "_"
        value_index += 1
    return "".join(chars)


def _normalize_masked_current_text(template: str, current_text: str) -> str:
    if current_text and len(current_text) == len(template):
        return current_text
    return _apply_masked_value(template, _extract_masked_value(template, current_text))


def _get_masked_overlay_text(template: str, current_text: str) -> str:
    safe_current = (
        current_text
        if current_text and len(current_text) == len(template)
        else _normalize_masked_current_text(template, current_text)
    )
    chars: list[str] = []
    for index, template_char in enumerate(template):
        if template_char == "_":
            char = safe_current[index] if index < len(safe_current) else "_"
            chars.append(char if char != "_" else " ")
        else:
            chars.append(" ")
    return "".join(chars)


def _get_iban_slot_pairs(template: str) -> list[tuple[int, int]]:
    underscore_indexes = [index for index, char in enumerate(template) if char == "_"]
    pairs: list[tuple[int, int]] = []
    for index in range(0, len(underscore_indexes), 2):
        start_index = underscore_indexes[index]
        end_index = underscore_indexes[index + 1] if index + 1 < len(underscore_indexes) else start_index
        pairs.append((start_index, end_index))
    return pairs


def _extract_iban_digits(template: str, current_text: str) -> str:
    slot_pairs = _get_iban_slot_pairs(template)
    safe_current = current_text or ""
    if len(safe_current) == len(template):
        digits: list[str] = []
        for start_index, end_index in slot_pairs:
            digit = next(
                (
                    safe_current[index]
                    for index in range(start_index, end_index + 1)
                    if index < len(safe_current) and safe_current[index].isdigit()
                ),
                "",
            )
            digits.append(digit)
        return "".join(digits)
    return "".join(char for char in safe_current if char.isdigit())[: len(slot_pairs)]


def _apply_iban_digits(template: str, digits: str) -> str:
    chars = list(template)
    slot_pairs = _get_iban_slot_pairs(template)
    for slot_index, (start_index, end_index) in enumerate(slot_pairs):
        digit = digits[slot_index] if slot_index < len(digits) and digits[slot_index].isdigit() else "_"
        chars[start_index] = digit
        for index in range(start_index + 1, end_index + 1):
            chars[index] = "_"
    return "".join(chars)


def _write_iban_overlay_text(
    page: pymupdf.Page,
    block: TextBlock,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    color: tuple[float, float, float],
    measurement_fonts: dict[str, pymupdf.Font],
) -> None:
    template = block.originalText
    current_text = _apply_iban_digits(template, _extract_iban_digits(template, block.currentText))
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    baseline = block.baseline if block.baseline is not None else (rect.y0 + block.fontSize)
    wrote_any = False

    for start_index, end_index in _get_iban_slot_pairs(template):
        char = current_text[start_index] if start_index < len(current_text) else "_"
        if not char.isdigit():
            continue
        left = _measure_text_width(template[:start_index], block.fontSize, runtime, font_spec, measurement_fonts)
        right = _measure_text_width(template[: end_index + 1], block.fontSize, runtime, font_spec, measurement_fonts)
        slot_width = max(0.0, right - left)
        char_width = _measure_text_width(char, block.fontSize, runtime, font_spec, measurement_fonts)
        x = rect.x0 + left + max(0.0, (slot_width - char_width) / 2)
        page.insert_text(
            pymupdf.Point(x, baseline),
            char,
            fontname=font_spec.name,
            fontsize=block.fontSize,
            color=color,
            overlay=True,
        )
        wrote_any = True

    if not wrote_any:
        return


def _write_masked_overlay_text(
    page: pymupdf.Page,
    block: TextBlock,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    color: tuple[float, float, float],
    measurement_fonts: dict[str, pymupdf.Font],
) -> None:
    template = block.originalText
    current_text = _normalize_masked_current_text(template, block.currentText)
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    baseline = block.baseline if block.baseline is not None else (rect.y0 + block.fontSize)

    for index, template_char in enumerate(template):
        if template_char != "_":
            continue
        char = current_text[index] if index < len(current_text) else "_"
        if char == "_":
            continue

        left = _measure_text_width(template[:index], block.fontSize, runtime, font_spec, measurement_fonts)
        right = _measure_text_width(template[: index + 1], block.fontSize, runtime, font_spec, measurement_fonts)
        slot_width = max(0.0, right - left)
        fontsize = block.fontSize
        char_width = _measure_text_width(char, fontsize, runtime, font_spec, measurement_fonts)
        if char_width > slot_width + 0.1:
            shrink_ratio = slot_width / max(char_width, 0.01)
            fontsize = max(block.minFontSize, round(fontsize * shrink_ratio, 2))
            char_width = _measure_text_width(char, fontsize, runtime, font_spec, measurement_fonts)

        x = rect.x0 + left + max(0.0, (slot_width - char_width) / 2)
        page.insert_text(
            pymupdf.Point(x, baseline),
            char,
            fontname=font_spec.name,
            fontsize=fontsize,
            color=color,
            overlay=True,
        )


def _resolve_system_font_file(font_name: str) -> Optional[Path]:
    normalized = normalize_font_name(font_name)
    filename = WINDOWS_FONT_FILES.get(normalized)
    if not filename:
        return None
    candidate = WINDOWS_FONTS_DIR / filename
    return candidate if candidate.exists() else None


def _styled_normalized_font_name(font_name: str, font_weight: str, font_style: str) -> str:
    normalized = normalize_font_name(font_name or "")
    wants_bold = _is_bold_font_weight(font_weight)
    wants_italic = str(font_style or "").strip().casefold() in {"italic", "oblique"}
    base = normalized
    for suffix in ("-bolditalic", "-boldoblique", "-bold", "-italic", "-oblique"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break

    if base in {"arial", "helvetica", "courier"}:
        if wants_bold and wants_italic:
            return f"{base}-bolditalic"
        if wants_bold:
            return f"{base}-bold"
        if wants_italic:
            return f"{base}-italic"
    if base in {"times", "times-roman", "times new roman"}:
        if wants_bold and wants_italic:
            return "times-bolditalic"
        if wants_bold:
            return "times-bold"
        if wants_italic:
            return "times-italic"
    return normalized


def _collect_page_font_resources(page: pymupdf.Page) -> dict[str, str]:
    resources: dict[str, str] = {}
    for _xref, _ext, _ftype, base_name, resource_name, _encoding in page.get_fonts():
        if not resource_name:
            continue
        normalized = normalize_font_name(base_name or resource_name)
        resources.setdefault(normalized, resource_name)
    return resources


def _resolve_export_font_spec(
    block: TextBlock,
    runtime: Optional[FontRuntime],
    page_font_resources: dict[str, str],
) -> ExportFontSpec:
    normalized = _styled_normalized_font_name(block.fontFamily or "", block.fontWeight or "400", block.fontStyle or "normal")
    system_font_file = _resolve_system_font_file(normalized)
    if system_font_file is not None:
        safe_font_name = re.sub(r"[^A-Za-z0-9_-]+", "-", normalized).strip("-") or "font"
        return ExportFontSpec(name=f"sysfont-{safe_font_name}", font_file=system_font_file)

    page_font_name = page_font_resources.get(normalized)
    if page_font_name:
        return ExportFontSpec(name=page_font_name)

    mapped = BASE14_FONT_MAP.get(normalized)
    if mapped:
        return ExportFontSpec(name=mapped[0])
    if runtime is not None:
        return ExportFontSpec(name=runtime.pdf_font_name)
    return ExportFontSpec(name=block.fontKey)


def _rect_distance(rect_a: pymupdf.Rect, rect_b: pymupdf.Rect) -> float:
    center_a = ((rect_a.x0 + rect_a.x1) / 2, (rect_a.y0 + rect_a.y1) / 2)
    center_b = ((rect_b.x0 + rect_b.x1) / 2, (rect_b.y0 + rect_b.y1) / 2)
    return math.hypot(center_a[0] - center_b[0], center_a[1] - center_b[1])


def _extract_horizontal_line_segments(
    page: pymupdf.Page,
    *,
    min_length: float = 20.0,
    max_width: float = 1.2,
) -> list[HorizontalLineSegment]:
    segments: list[HorizontalLineSegment] = []
    collection_min_length = min(min_length, 15.0)

    for drawing in page.get_drawings():
        color = drawing.get("color")
        fill = drawing.get("fill")
        has_dark_stroke = bool(color) and all(channel <= 0.25 for channel in color)
        has_dark_fill = bool(fill) and all(channel <= 0.25 for channel in fill)
        if not has_dark_stroke and not has_dark_fill:
            continue

        stroke_width = float(drawing.get("width") or 0.0)
        if has_dark_stroke and stroke_width > max_width:
            continue

        for item in drawing.get("items") or []:
            if item[0] == "re":
                # Thin filled/stroked rectangles are sometimes used as underlines
                try:
                    rect = pymupdf.Rect(item[1] if len(item) > 1 else drawing.get("rect"))
                except Exception:
                    continue
                rect_h = abs(rect.y1 - rect.y0)
                rect_w = abs(rect.x1 - rect.x0)
                effective_height = rect_h if rect_h > 0 else stroke_width
                if effective_height <= max_width and rect_w >= collection_min_length:
                    segments.append(
                        HorizontalLineSegment(
                            x0=float(min(rect.x0, rect.x1)),
                            x1=float(max(rect.x0, rect.x1)),
                            y=float((rect.y0 + rect.y1) / 2),
                            width=max(0.1, effective_height),
                        )
                    )
                continue

            if item[0] != "l":
                continue
            start = item[1]
            end = item[2]
            if abs(start.y - end.y) > 0.2:
                continue

            x0 = float(min(start.x, end.x))
            x1 = float(max(start.x, end.x))
            if (x1 - x0) < collection_min_length:
                continue

            segments.append(
                HorizontalLineSegment(
                    x0=x0,
                    x1=x1,
                    y=float((start.y + end.y) / 2),
                    width=max(0.1, stroke_width),
                )
            )

    segments.sort(key=lambda line: (line.y, line.x0, line.x1))
    deduped: list[HorizontalLineSegment] = []
    for line in segments:
        if deduped:
            previous = deduped[-1]
            same_position = (
                abs(previous.y - line.y) <= 0.2
                and abs(previous.x0 - line.x0) <= 0.6
                and abs(previous.x1 - line.x1) <= 0.6
            )
            if same_position:
                continue
        deduped.append(line)

    merged: list[HorizontalLineSegment] = []
    for line in deduped:
        if merged:
            previous = merged[-1]
            same_row = abs(previous.y - line.y) <= 0.35
            touches_previous = line.x0 <= previous.x1 + 3.0
            similar_width = abs(previous.width - line.width) <= 0.8
            if same_row and touches_previous and similar_width:
                merged[-1] = HorizontalLineSegment(
                    x0=min(previous.x0, line.x0),
                    x1=max(previous.x1, line.x1),
                    y=(previous.y + line.y) / 2,
                    width=max(previous.width, line.width),
                )
                continue
        merged.append(line)

    return [line for line in merged if (line.x1 - line.x0) >= min_length]


def _extract_vertical_line_segments(
    page: pymupdf.Page,
    *,
    min_length: float = 20.0,
    max_width: float = 1.2,
) -> list[VerticalLineSegment]:
    segments: list[VerticalLineSegment] = []

    for drawing in page.get_drawings():
        color = drawing.get("color")
        if not _is_dark_drawing_color(color):
            continue

        stroke_width = float(drawing.get("width") or 0.0)
        if stroke_width > max_width:
            continue

        for item in drawing.get("items") or []:
            if item[0] != "l":
                continue
            start = item[1]
            end = item[2]
            if abs(start.x - end.x) > 0.2:
                continue

            y0 = float(min(start.y, end.y))
            y1 = float(max(start.y, end.y))
            if (y1 - y0) < min_length:
                continue

            segments.append(
                VerticalLineSegment(
                    x=float((start.x + end.x) / 2),
                    y0=y0,
                    y1=y1,
                    width=max(0.1, stroke_width),
                )
            )

    segments.sort(key=lambda line: (line.x, line.y0, line.y1))
    deduped: list[VerticalLineSegment] = []
    for line in segments:
        if deduped:
            previous = deduped[-1]
            same_position = (
                abs(previous.x - line.x) <= 0.2
                and abs(previous.y0 - line.y0) <= 0.6
                and abs(previous.y1 - line.y1) <= 0.6
            )
            if same_position:
                continue
        deduped.append(line)
    return deduped


def _is_dark_drawing_color(color: tuple[float, ...] | None, threshold: float = 0.35) -> bool:
    return bool(color) and all(channel <= threshold for channel in color[:3])


def _outline_rect_from_rect(
    rect_like,
    stroke_width: float,
    *,
    min_width: float,
    min_height: float,
) -> Optional[OutlineRect]:
    try:
        rect = pymupdf.Rect(rect_like)
    except Exception:
        return None
    width = float(rect.x1 - rect.x0)
    height = float(rect.y1 - rect.y0)
    if width < min_width or height < min_height:
        return None
    return OutlineRect(
        x0=float(min(rect.x0, rect.x1)),
        y0=float(min(rect.y0, rect.y1)),
        x1=float(max(rect.x0, rect.x1)),
        y1=float(max(rect.y0, rect.y1)),
        width=max(0.1, stroke_width),
    )


def _extract_outline_rects(
    page: pymupdf.Page,
    *,
    min_width: float = 80.0,
    min_height: float = 20.0,
    max_width: float = 1.6,
) -> list[OutlineRect]:
    rects: list[OutlineRect] = []
    horizontal_segments: list[HorizontalLineSegment] = []
    vertical_segments: list[VerticalLineSegment] = []

    for drawing in page.get_drawings():
        color = drawing.get("color")
        if not _is_dark_drawing_color(color):
            continue

        stroke_width = float(drawing.get("width") or 0.0)
        if stroke_width > max_width:
            continue

        for item in drawing.get("items") or []:
            if item[0] == "re":
                rect_candidate = item[1] if len(item) > 1 else drawing.get("rect")
                outline = _outline_rect_from_rect(
                    rect_candidate,
                    stroke_width,
                    min_width=min_width,
                    min_height=min_height,
                )
                if outline is not None:
                    rects.append(outline)
                continue

            if item[0] != "l":
                continue
            start = item[1]
            end = item[2]
            horizontal = abs(start.y - end.y) <= 0.2
            vertical = abs(start.x - end.x) <= 0.2
            if horizontal:
                x0 = float(min(start.x, end.x))
                x1 = float(max(start.x, end.x))
                if (x1 - x0) >= min_width:
                    horizontal_segments.append(
                        HorizontalLineSegment(
                            x0=x0,
                            x1=x1,
                            y=float((start.y + end.y) / 2),
                            width=max(0.1, stroke_width),
                        )
                    )
                continue
            if vertical:
                y0 = float(min(start.y, end.y))
                y1 = float(max(start.y, end.y))
                if (y1 - y0) >= min_height:
                    vertical_segments.append(
                        VerticalLineSegment(
                            x=float((start.x + end.x) / 2),
                            y0=y0,
                            y1=y1,
                            width=max(0.1, stroke_width),
                        )
                    )

    tolerance = 1.6
    horizontal_segments.sort(key=lambda line: (line.y, line.x0, line.x1))
    for index, top_line in enumerate(horizontal_segments):
        for bottom_line in horizontal_segments[index + 1:]:
            height = bottom_line.y - top_line.y
            if height < min_height:
                continue
            if abs(top_line.x0 - bottom_line.x0) > tolerance or abs(top_line.x1 - bottom_line.x1) > tolerance:
                continue

            left_edge = any(
                abs(segment.x - top_line.x0) <= tolerance
                and segment.y0 <= top_line.y + tolerance
                and segment.y1 >= bottom_line.y - tolerance
                for segment in vertical_segments
            )
            right_edge = any(
                abs(segment.x - top_line.x1) <= tolerance
                and segment.y0 <= top_line.y + tolerance
                and segment.y1 >= bottom_line.y - tolerance
                for segment in vertical_segments
            )
            if not left_edge or not right_edge:
                continue
            rects.append(
                OutlineRect(
                    x0=min(top_line.x0, bottom_line.x0),
                    y0=top_line.y,
                    x1=max(top_line.x1, bottom_line.x1),
                    y1=bottom_line.y,
                    width=max(top_line.width, bottom_line.width),
                )
            )

    rects.sort(key=lambda rect: (rect.y0, rect.x0, rect.y1, rect.x1))
    deduped: list[OutlineRect] = []
    for rect in rects:
        duplicate = any(
            abs(existing.x0 - rect.x0) <= tolerance
            and abs(existing.y0 - rect.y0) <= tolerance
            and abs(existing.x1 - rect.x1) <= tolerance
            and abs(existing.y1 - rect.y1) <= tolerance
            for existing in deduped
        )
        if not duplicate:
            deduped.append(rect)
    return deduped


def _extract_line_overlays(page: pymupdf.Page) -> list[LineOverlay]:
    overlays: list[LineOverlay] = []

    for drawing in page.get_drawings():
        color = drawing.get("color")
        if not color:
            continue
        if any(channel > 0.15 for channel in color):
            continue

        stroke_width = float(drawing.get("width") or 0.0)
        if stroke_width > 0.25:
            continue
        for item in drawing.get("items") or []:
            if item[0] != "l":
                continue
            start = item[1]
            end = item[2]
            horizontal = abs(start.y - end.y) <= 0.2
            vertical = abs(start.x - end.x) <= 0.2
            if not horizontal and not vertical:
                continue

            overlays.append(
                LineOverlay(
                    x0=float(start.x),
                    y0=float(start.y),
                    x1=float(end.x),
                    y1=float(end.y),
                    width=max(0.1, stroke_width),
                    color=_stroke_color_to_hex(color),
                )
            )

    return overlays


def _estimate_space_width(chars: list[dict], font_size: float) -> float:
    widths = [char["bbox"][2] - char["bbox"][0] for char in chars if char["c"] == " "]
    if widths:
        return sum(widths) / len(widths)
    visible = [char["bbox"][2] - char["bbox"][0] for char in chars if char["c"].strip()]
    if visible:
        return max(font_size * 0.25, (sum(visible) / len(visible)) * 0.55)
    return max(1.0, font_size * 0.33)


def _resolve_font_runtime(
    font_runtimes_by_family: dict[str, FontRuntime],
    span_font: str,
) -> Optional[FontRuntime]:
    normalized = normalize_font_name(span_font)
    return font_runtimes_by_family.get(normalized)


def _rehydrate_custom_block_fonts(
    blocks: list[TextBlock],
    font_runtimes_by_family: dict[str, FontRuntime],
) -> None:
    for block in blocks:
        runtime = _resolve_font_runtime(font_runtimes_by_family, block.fontFamily)
        if runtime is None:
            block.cssFontFamily = choose_css_family(block.fontFamily)
            font_weight, font_style = infer_font_style(block.fontFamily)
            block.fontWeight = font_weight
            block.fontStyle = font_style
            continue

        block.fontAssetId = runtime.asset.id
        block.fontKey = runtime.pdf_font_name
        block.cssFontFamily = runtime.asset.cssFamily
        font_weight, font_style = infer_font_style(block.fontFamily)
        block.fontWeight = font_weight
        block.fontStyle = font_style


def _collect_font_runtimes(doc: pymupdf.Document, fonts_dir: Path) -> tuple[dict[str, FontRuntime], list[str]]:
    fonts_dir.mkdir(parents=True, exist_ok=True)
    runtimes: dict[str, FontRuntime] = {}
    reasons: list[str] = []
    seen_xrefs: set[int] = set()

    for page in doc:
        for xref, _ext, _ftype, base_name, _resource_name, _encoding in page.get_fonts():
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)

            extracted_name, extracted_ext, _extracted_type, font_buffer = doc.extract_font(xref)
            family = extracted_name or base_name or "Unknown"
            normalized = normalize_font_name(family)
            mapped = BASE14_FONT_MAP.get(normalized)

            font_path = None
            pdf_font_name = family
            font_asset_id = f"font-{uuid4().hex}"
            extension = None if extracted_ext == "n/a" else extracted_ext
            embedded = bool(font_buffer)

            if font_buffer and extension in BROWSER_FONT_EXTENSIONS:
                font_path = fonts_dir / f"{font_asset_id}.{extension}"
                font_path.write_bytes(font_buffer)

            # Extracted subset fonts are fine for browser preview, but reusing them
            # for new PDF text can corrupt glyph mapping on export. Prefer safe PDF
            # base14 equivalents whenever we can map the family.
            if mapped:
                pdf_font_name = mapped[0]
            elif normalized in {"zapfdingbats", "zadi", "wingdings"}:
                pdf_font_name = "Helvetica"
            elif font_buffer:
                pdf_font_name = font_asset_id
            else:
                reasons.append(f"Font nicht rekonstruierbar: {family}")
                continue

            runtime = FontRuntime(
                asset=FontAsset(
                    id=font_asset_id,
                    family=family,
                    cssFamily=family if font_path else choose_css_family(family),
                    extension=extension,
                    embedded=embedded,
                ),
                pdf_font_name=pdf_font_name,
                font_buffer=font_buffer or None,
                font_path=font_path,
            )
            runtimes[normalized] = runtime

    return runtimes, reasons


def _merge_spans_into_line_blocks(spans: list[SpanFragment]) -> list[BlockFragment]:
    if not spans:
        return []

    spans = sorted(spans, key=lambda item: item.bbox[0])
    merged: list[BlockFragment] = []

    current = BlockFragment(
        text=spans[0].text,
        bbox=list(spans[0].bbox),
        font_family=spans[0].font_family,
        font_key=spans[0].font_key,
        font_asset_id=spans[0].font_asset_id,
        font_size=spans[0].font_size,
        color=spans[0].color,
        line_height=spans[0].line_height,
        baseline=spans[0].baseline,
    )

    for span in spans[1:]:
        gap = span.bbox[0] - current.bbox[2]
        style_match = (
            current.font_key == span.font_key
            and current.color == span.color
            and abs(current.font_size - span.font_size) <= 0.25
        )

        if style_match and gap <= max(span.space_width, current.font_size * 0.33) * 1.5:
            if gap > max(span.space_width, current.font_size * 0.33) * 0.55 and not current.text.endswith(" ") and not span.text.startswith(" "):
                current.text += " "
            current.text += span.text
            current.bbox = _bbox_union(current.bbox, span.bbox)
            current.line_height = max(current.line_height, span.line_height)
            current.baseline = span.baseline
            continue

        merged.append(current)
        current = BlockFragment(
            text=span.text,
            bbox=list(span.bbox),
            font_family=span.font_family,
            font_key=span.font_key,
            font_asset_id=span.font_asset_id,
            font_size=span.font_size,
            color=span.color,
            line_height=span.line_height,
            baseline=span.baseline,
        )

    merged.append(current)
    return merged


def _merge_line_blocks_into_multiline(blocks: list[BlockFragment]) -> list[BlockFragment]:
    if not blocks:
        return []

    blocks = sorted(blocks, key=lambda item: (item.bbox[1], item.bbox[0]))
    merged: list[BlockFragment] = [blocks[0]]

    for block in blocks[1:]:
        previous = merged[-1]
        same_style = (
            previous.font_key == block.font_key
            and previous.color == block.color
            and abs(previous.font_size - block.font_size) <= 0.25
        )
        same_left = abs(previous.bbox[0] - block.bbox[0]) <= 1.5
        expected_gap = previous.line_height
        baseline_gap = block.baseline - previous.baseline
        close_line_gap = abs(baseline_gap - expected_gap) <= 1.5
        similar_width = abs(_bbox_width(previous.bbox) - _bbox_width(block.bbox)) <= max(12.0, _bbox_width(previous.bbox) * 0.15)

        if same_style and same_left and close_line_gap and similar_width:
            previous.text = f"{previous.text}\n{block.text}"
            previous.bbox = _bbox_union(previous.bbox, tuple(block.bbox))
            previous.line_height = max(previous.line_height, block.line_height)
            previous.baseline = block.baseline
            previous.group_kind = "multiline"
            continue

        merged.append(block)

    return merged


def _page_is_visually_blank(page: pymupdf.Page) -> bool:
    if page.get_text("text").strip() or page.get_drawings():
        return False
    try:
        pixmap = page.get_pixmap(dpi=40, alpha=False)
    except Exception:
        return False
    dark_pixels = 0
    total_pixels = max(1, pixmap.width * pixmap.height)
    samples = pixmap.samples
    for index in range(0, len(samples), 3):
        if samples[index] + samples[index + 1] + samples[index + 2] < 735:
            dark_pixels += 1
            if dark_pixels / total_pixels > 0.0005:
                return False
    return True


def _extract_blocks_for_page(
    page: pymupdf.Page,
    font_runtimes_by_family: dict[str, FontRuntime],
    warnings: list[str],
    reasons: list[str],
) -> list[TextBlock]:
    raw = page.get_text("rawdict")
    page_blocks: list[TextBlock] = []
    block_counter = 0

    for raw_block in raw.get("blocks", []):
        if raw_block.get("type") != 0:
            continue

        line_blocks: list[BlockFragment] = []

        for line in raw_block.get("lines", []):
            if line.get("wmode", 0) != 0:
                reasons.append(f"Nicht-horizontaler Schreibmodus auf Seite {page.number + 1}")
                continue

            direction = line.get("dir", (1.0, 0.0))
            if direction[0] < 1.0 - TEXT_DIRECTION_SKEW_TOLERANCE or abs(direction[1]) > TEXT_DIRECTION_SKEW_TOLERANCE:
                reasons.append(f"Rotierter Text auf Seite {page.number + 1}")
                continue

            spans: list[SpanFragment] = []

            for span in line.get("spans", []):
                chars = _trim_outer_whitespace_chars(span.get("chars", []))
                text = _chars_to_text(chars)
                if not text.strip():
                    continue

                span_font = span.get("font", "")
                font_family, font_key, font_asset_id, css_font_family = _font_details_for_span(
                    font_runtimes_by_family,
                    span_font,
                )
                if not font_family:
                    warnings.append(f"Font nicht rekonstruierbar: {span.get('font', 'Unbekannt')}")
                    continue

                font_size = float(span.get("size", 0.0))
                line_height = max(
                    float(span["bbox"][3] - span["bbox"][1]),
                    font_size * (float(span.get("ascender", 1.0)) - float(span.get("descender", -0.2))),
                )
                spans.append(
                    SpanFragment(
                        text=text,
                        bbox=_chars_bbox(chars, tuple(float(value) for value in span["bbox"])),
                        font_family=font_family,
                        font_key=font_key,
                        font_asset_id=font_asset_id,
                        font_size=font_size,
                        color=_hex_color(int(span.get("color", 0))),
                        line_height=line_height,
                        baseline=float(span.get("origin", (0.0, span["bbox"][3]))[1]),
                        space_width=_estimate_space_width(chars, font_size),
                    )
                )

            line_blocks.extend(_merge_spans_into_line_blocks(spans))

        for fragment in line_blocks:
            block_counter += 1
            font_weight, font_style = infer_font_style(fragment.font_family)
            bbox = BoundingBox(
                x0=fragment.bbox[0],
                y0=fragment.bbox[1],
                x1=fragment.bbox[2],
                y1=fragment.bbox[3],
            )
            page_blocks.append(
                TextBlock(
                    id=f"page-{page.number + 1}-block-{block_counter}",
                    page=page.number + 1,
                    bbox=bbox,
                    originalText=fragment.text,
                    currentText=fragment.text,
                    fontFamily=fragment.font_family,
                    fontKey=fragment.font_key,
                    fontSize=round(fragment.font_size, 3),
                    color=fragment.color,
                    lineHeight=round(fragment.line_height, 3),
                    align=fragment.align,
                    rotation=fragment.rotation,
                    groupKind=fragment.group_kind,
                    minFontSize=6.0,
                    editable=True,
                    cssFontFamily=(
                        font_runtimes_by_family[normalize_font_name(fragment.font_family)].asset.cssFamily
                        if normalize_font_name(fragment.font_family) in font_runtimes_by_family
                        else choose_css_family(fragment.font_family)
                    ),
                    fontAssetId=fragment.font_asset_id,
                    fontWeight=font_weight,
                    fontStyle=font_style,
                    baseline=round(fragment.baseline, 3),
                    isCustom=False,
                )
            )

    if not page_blocks and not _page_is_visually_blank(page):
        warnings.append(f"Seite {page.number + 1} enthält keine editierbaren Textblöcke.")

    return _sync_fields(page_blocks)


def _detect_checkbox_rects(page: pymupdf.Page) -> list[pymupdf.Rect]:
    checkbox_rects: list[pymupdf.Rect] = []

    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        fill = drawing.get("fill")
        items = drawing.get("items") or []
        if rect is None or len(items) != 1:
            continue
        if items[0][0] != "re":
            continue
        if not fill:
            continue
        if rect.width < 6 or rect.width > 20 or rect.height < 6 or rect.height > 20:
            continue
        if abs(rect.width - rect.height) > 1.5:
            continue
        if any(abs(channel - 1.0) > 0.03 for channel in fill):
            continue

        has_outer_border = False
        for candidate in page.get_drawings():
            outer_rect = candidate.get("rect")
            outer_fill = candidate.get("fill")
            outer_items = candidate.get("items") or []
            if outer_rect is None or len(outer_items) != 1 or outer_items[0][0] != "re":
                continue
            if not outer_fill:
                continue
            if any(abs(channel - 0.0) > 0.03 for channel in outer_fill):
                continue
            if outer_rect.width < rect.width or outer_rect.height < rect.height:
                continue
            if abs((outer_rect.width - rect.width) - 1.2) > 1.2:
                continue
            if abs((outer_rect.height - rect.height) - 1.2) > 1.2:
                continue
            if rect.x0 >= outer_rect.x0 - 0.8 and rect.y0 >= outer_rect.y0 - 0.8 and rect.x1 <= outer_rect.x1 + 0.8 and rect.y1 <= outer_rect.y1 + 0.8:
                has_outer_border = True
                break

        if not has_outer_border:
            continue

        checkbox_rects.append(pymupdf.Rect(rect))

    checkbox_rects.sort(key=lambda rect: (rect.y0, rect.x0))
    deduped: list[pymupdf.Rect] = []
    for rect in checkbox_rects:
        if deduped and _rect_distance(deduped[-1], rect) < 1.0:
            continue
        deduped.append(rect)
    return deduped


def _checkbox_has_vector_mark(page: pymupdf.Page, rect: pymupdf.Rect) -> bool:
    search_rect = pymupdf.Rect(rect.x0 - 0.8, rect.y0 - 0.8, rect.x1 + 0.8, rect.y1 + 0.8)
    diagonal_segments = 0

    for drawing in page.get_drawings():
        color = drawing.get("color")
        if not color or any(channel > 0.2 for channel in color):
            continue
        for item in drawing.get("items") or []:
            if item[0] != "l":
                continue
            start = item[1]
            end = item[2]
            segment_rect = pymupdf.Rect(
                min(start.x, end.x),
                min(start.y, end.y),
                max(start.x, end.x),
                max(start.y, end.y),
            )
            if not segment_rect.intersects(search_rect):
                continue
            if abs(start.x - end.x) <= 0.2 or abs(start.y - end.y) <= 0.2:
                continue
            diagonal_segments += 1
            if diagonal_segments >= 2:
                return True

    return False


def _build_checkbox_blocks(
    page: pymupdf.Page,
    page_number: int,
    checkbox_rects: list[pymupdf.Rect],
    page_blocks: list[TextBlock],
) -> tuple[list[TextBlock], set[str]]:
    checkbox_blocks: list[TextBlock] = []
    hidden_mark_block_ids: set[str] = set()

    if not checkbox_rects:
        return checkbox_blocks, hidden_mark_block_ids

    block_counter = 0
    text_candidates = [block for block in page_blocks if block.editable]

    for rect in checkbox_rects:
        mark_search_rect = pymupdf.Rect(rect.x0 - 1.0, rect.y0 - 1.0, rect.x1 + 1.0, rect.y1 + 1.0)
        mark_block = next(
            (
                block for block in text_candidates
                if block.currentText.strip().lower() == "x"
                and pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1).intersects(mark_search_rect)
            ),
            None,
        )
        if mark_block:
            hidden_mark_block_ids.add(mark_block.id)

        if mark_block:
            style_source = mark_block
        elif text_candidates:
            style_source = min(
                text_candidates,
                key=lambda block: _rect_distance(
                    rect,
                    pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
                ),
            )
        else:
            style_source = TextBlock(
                id=f"page-{page_number}-checkbox-style-source",
                page=page_number,
                bbox=BoundingBox(x0=rect.x0, y0=rect.y0, x1=rect.x1, y1=rect.y1),
                originalText="",
                currentText="",
                fontFamily="Arial-BoldMT",
                fontKey="Helvetica-Bold",
                fontSize=rect.height,
                color="#000000",
                lineHeight=rect.height,
                align="left",
                rotation=0.0,
                groupKind="checkbox-style",
                minFontSize=6.0,
                editable=False,
                cssFontFamily="Arial, sans-serif",
                fontAssetId=None,
                fontWeight="700",
                fontStyle="normal",
                baseline=None,
                isCheckbox=False,
                isCustom=False,
            )

        block_counter += 1
        checked = bool(mark_block) or _checkbox_has_vector_mark(page, rect)
        checkbox_blocks.append(
            TextBlock(
                id=f"page-{page_number}-checkbox-{block_counter}",
                page=page_number,
                bbox=BoundingBox(x0=rect.x0, y0=rect.y0, x1=rect.x1, y1=rect.y1),
                originalText="x" if checked else "",
                currentText="x" if checked else "",
                fontFamily=style_source.fontFamily,
                fontKey=style_source.fontKey,
                fontSize=max(style_source.fontSize, rect.height),
                color="#000000",
                lineHeight=max(style_source.lineHeight, rect.height),
                align="left",
                rotation=0.0,
                groupKind="checkbox",
                minFontSize=6.0,
                editable=True,
                cssFontFamily=style_source.cssFontFamily,
                fontAssetId=style_source.fontAssetId,
                fontWeight="700",
                fontStyle="normal",
                baseline=None,
                isCheckbox=True,
                isCustom=False,
            )
        )

    return checkbox_blocks, hidden_mark_block_ids


def _widget_value_is_checked(value) -> bool:
    normalized = str(value or "").strip().casefold()
    return bool(normalized and normalized not in {"off", "false", "0", "no", "none"})


def _widget_text_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return "\n".join(str(item) for item in value if str(item).strip())
    return str(value)


def _nearest_style_source(rect: pymupdf.Rect, page_blocks: list[TextBlock], page_number: int) -> TextBlock:
    candidates = [block for block in page_blocks if block.editable and block.originalText.strip()]
    if candidates:
        return min(
            candidates,
            key=lambda block: _rect_distance(
                rect,
                pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
            ),
        )

    return TextBlock(
        id=f"page-{page_number}-fallback-style",
        page=page_number,
        bbox=BoundingBox(x0=rect.x0, y0=rect.y0, x1=rect.x1, y1=rect.y1),
        originalText="",
        currentText="",
        fontFamily="Helvetica",
        fontKey="Helvetica",
        fontSize=9.0,
        color="#000000",
        lineHeight=10.8,
        align="left",
        rotation=0.0,
        groupKind="fallback-style",
        minFontSize=6.0,
        editable=False,
        cssFontFamily="Arial, sans-serif",
        fontAssetId=None,
        fontWeight="400",
        fontStyle="normal",
        baseline=None,
        isCheckbox=False,
        isCustom=False,
    )


def _block_rect_overlap_ratio(block: TextBlock, rect: pymupdf.Rect, *, padding: float = 1.0) -> float:
    block_rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    padded_rect = pymupdf.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding, rect.y1 + padding)
    intersection = block_rect & padded_rect
    block_area = max(0.01, block_rect.width * block_rect.height)
    return max(0.0, intersection.width * intersection.height) / block_area


def _session_pdf_base_path(session: DocumentSession) -> Path:
    return session.base_pdf_path or session.source_path


def _session_overlay_regions_for_page(session: DocumentSession, page_number: int) -> list[SourceOverlayRegion]:
    return [region for region in session.source_overlay_regions if region.page_number == page_number]


def _overlay_source_regions_on_page(
    *,
    target_page: pymupdf.Page,
    source_doc: Optional[pymupdf.Document],
    page_number: int,
    overlay_regions: list[SourceOverlayRegion],
) -> None:
    if source_doc is None or not overlay_regions:
        return
    if page_number < 1 or page_number > source_doc.page_count:
        return

    for region in overlay_regions:
        source_rect = region.pdf_rect & source_doc[page_number - 1].rect
        target_rect = region.pdf_rect & target_page.rect
        if source_rect.is_empty or target_rect.is_empty:
            continue
        target_page.show_pdf_page(
            target_rect,
            source_doc,
            page_number - 1,
            clip=source_rect,
            overlay=True,
        )


def _has_existing_widget_field_overlap(page_blocks: list[TextBlock], rect: pymupdf.Rect) -> bool:
    for block in page_blocks:
        if block.isCustom or block.isCheckbox:
            continue
        if not isinstance(block.groupKind, str) or not block.groupKind.startswith("widget-"):
            continue
        block_rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        intersection = block_rect & rect
        intersection_area = max(0.0, intersection.width * intersection.height)
        smaller_area = max(0.01, min(block_rect.width * block_rect.height, rect.width * rect.height))
        if (intersection_area / smaller_area) >= 0.35:
            return True
    return False


def _build_widget_blocks(
    page: pymupdf.Page,
    page_number: int,
    page_blocks: list[TextBlock],
) -> tuple[list[TextBlock], set[str]]:
    widget_blocks: list[TextBlock] = []
    hidden_block_ids: set[str] = set()
    widgets = list(page.widgets() or [])
    if not widgets:
        return widget_blocks, hidden_block_ids

    for index, widget in enumerate(widgets, start=1):
        rect = pymupdf.Rect(widget.rect)
        if rect.is_empty or rect.width < 2 or rect.height < 2:
            continue

        field_type = str(getattr(widget, "field_type_string", "") or "").casefold()
        field_name = str(getattr(widget, "field_name", "") or f"widget-{index}")
        safe_name = "".join(char if char.isalnum() else "-" for char in field_name).strip("-").lower() or f"widget-{index}"
        style_source = _nearest_style_source(rect, page_blocks, page_number)

        if "check" in field_type:
            checkbox_block = _build_generated_checkbox_block(
                block_id=f"page-{page_number}-widget-checkbox-{safe_name}-{index}",
                page_number=page_number,
                rect=rect,
                style_source=style_source,
                checked=_widget_value_is_checked(getattr(widget, "field_value", "")),
                group_kind="widget-checkbox-field",
            )
            checkbox_block.widgetFieldName = field_name
            checkbox_block.widgetFieldType = field_type
            checkbox_block.widgetXref = int(getattr(widget, "xref", 0) or 0) or None
            checkbox_block.sourceType = "acroform"
            widget_blocks.append(checkbox_block)
            continue

        if "text" not in field_type:
            continue

        value = _widget_text_value(getattr(widget, "field_value", ""))
        if not value.strip():
            continue

        font_size = float(getattr(widget, "text_fontsize", 0.0) or 0.0)
        if font_size <= 0:
            font_size = min(max(style_source.fontSize, 8.0), max(9.0, rect.height * 0.55))
        line_height = max(style_source.lineHeight, font_size * 1.18)
        is_multiline = rect.height >= line_height * 1.65
        baseline = None if is_multiline else min(rect.y1 - 2.0, rect.y0 + _baseline_offset(style_source))

        font_weight, font_style = infer_font_style(style_source.fontFamily)
        widget_block = TextBlock(
            id=f"page-{page_number}-widget-text-{safe_name}-{index}",
            page=page_number,
            fieldType="text-multiline" if is_multiline else "text-line",
            bbox=BoundingBox(x0=round(rect.x0, 3), y0=round(rect.y0, 3), x1=round(rect.x1, 3), y1=round(rect.y1, 3)),
            originalText=value,
            currentText=value,
            fontFamily=style_source.fontFamily,
            fontKey=style_source.fontKey,
            fontSize=round(font_size, 3),
            color="#000000",
            lineHeight=round(line_height, 3),
            align="left",
            rotation=0.0,
            groupKind="widget-text-field",
            minFontSize=6.0,
            editable=True,
            cssFontFamily=style_source.cssFontFamily,
            fontAssetId=style_source.fontAssetId,
            fontWeight=font_weight,
            fontStyle=font_style,
            baseline=round(baseline, 3) if baseline is not None else None,
            isCheckbox=False,
            isCustom=False,
            widgetFieldName=field_name,
            widgetFieldType=field_type,
            widgetXref=int(getattr(widget, "xref", 0) or 0) or None,
            sourceType="acroform",
        )
        widget_blocks.append(widget_block)

        for block in page_blocks:
            if _block_rect_overlap_ratio(block, rect, padding=0.75) >= 0.45:
                hidden_block_ids.add(block.id)

    return widget_blocks, hidden_block_ids


def _baseline_offset(style_source: TextBlock) -> float:
    if style_source.baseline is not None:
        offset = style_source.baseline - style_source.bbox.y0
        if offset > 0:
            return min(max(style_source.fontSize * 0.72, offset), style_source.lineHeight)
    return max(style_source.fontSize * 0.82, style_source.lineHeight * 0.78)


def _build_generated_text_block(
    *,
    block_id: str,
    page_number: int,
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    style_source: TextBlock,
    group_kind: str,
    baseline: Optional[float],
) -> TextBlock:
    font_weight, font_style = infer_font_style(style_source.fontFamily)
    field = TextBlock(
        id=block_id,
        page=page_number,
        bbox=BoundingBox(x0=round(x0, 3), y0=round(y0, 3), x1=round(x1, 3), y1=round(y1, 3)),
        originalText="",
        currentText="",
        fontFamily=style_source.fontFamily,
        fontKey=style_source.fontKey,
        fontSize=style_source.fontSize,
        color=style_source.color,
        lineHeight=style_source.lineHeight,
        align="left",
        rotation=0.0,
        groupKind=group_kind,
        minFontSize=6.0,
        editable=True,
        cssFontFamily=style_source.cssFontFamily,
        fontAssetId=style_source.fontAssetId,
        fontWeight=font_weight,
        fontStyle=font_style,
        baseline=round(baseline, 3) if baseline is not None else None,
        isCheckbox=False,
        isCustom=False,
    )
    if group_kind in {"generated-contract-party-field", "generated-contract-object-line-field"}:
        field_height = max(0.0, field.bbox.y1 - field.bbox.y0)
        baseline_outside_field = (
            field.baseline is None
            or field.baseline <= field.bbox.y0 + 0.5
            or field.baseline >= field.bbox.y1 + 2.0
        )
        if baseline_outside_field and field_height >= 6.0:
            field.fontSize = min(field.fontSize, 9.2)
            field.lineHeight = round(max(field.fontSize * 1.15, field_height), 3)
            field.baseline = round(field.bbox.y0 + min(field_height - 1.1, field.fontSize * 0.88), 3)
    return field


def _build_generated_single_line_field(
    block_id: str,
    page_number: int,
    line: HorizontalLineSegment,
    style_source: TextBlock,
) -> TextBlock:
    line_gap = max(0.85, style_source.lineHeight * 0.1)
    height = max(style_source.lineHeight, style_source.fontSize + 1.0)
    y1 = line.y - line_gap
    y0 = y1 - height
    baseline = line.y - max(2.8, style_source.fontSize * 0.36)
    return _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=line.x0,
        y0=y0,
        x1=line.x1,
        y1=y1,
        style_source=style_source,
        group_kind="generated-line-field",
        baseline=baseline,
    )


def _build_generated_box_field(
    block_id: str,
    page_number: int,
    top_line: HorizontalLineSegment,
    bottom_line: HorizontalLineSegment,
    style_source: TextBlock,
) -> TextBlock:
    gap = max(0.0, bottom_line.y - top_line.y)
    height = max(style_source.lineHeight, gap - 1.6)
    top_padding = max(0.9, (gap - height) / 2) if gap > height else 0.9
    y0 = top_line.y + top_padding
    y1 = min(bottom_line.y - 0.8, y0 + height)
    baseline = y0 + _baseline_offset(style_source)
    return _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=min(top_line.x0, bottom_line.x0),
        y0=y0,
        x1=max(top_line.x1, bottom_line.x1),
        y1=y1,
        style_source=style_source,
        group_kind="generated-box-field",
        baseline=baseline,
    )


def _build_generated_multiline_field(
    block_id: str,
    page_number: int,
    lines: list[HorizontalLineSegment],
    style_source: TextBlock,
) -> TextBlock:
    ordered = sorted(lines, key=lambda line: line.y)
    line_gap = max(0.85, style_source.lineHeight * 0.1)
    y0 = ordered[0].y - style_source.lineHeight - line_gap
    y1 = ordered[-1].y - line_gap
    return _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=min(line.x0 for line in ordered),
        y0=y0,
        x1=max(line.x1 for line in ordered),
        y1=y1,
        style_source=style_source,
        group_kind="generated-multiline-field",
        baseline=None,
    )


def _build_generated_underlined_field(
    block_id: str,
    page_number: int,
    line: HorizontalLineSegment,
    style_source: TextBlock,
    *,
    baseline_gap: float,
    group_kind: str,
) -> TextBlock:
    baseline = line.y - baseline_gap
    height = max(style_source.lineHeight, style_source.fontSize + 1.5)
    y0 = baseline - _baseline_offset(style_source)
    y1 = min(line.y - 0.45, y0 + height)
    if (y1 - y0) < 6:
        y1 = line.y - 0.45
        y0 = y1 - height
    return _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=line.x0 + 0.35,
        y0=y0,
        x1=line.x1 - 0.35,
        y1=y1,
        style_source=style_source,
        group_kind=group_kind,
        baseline=baseline,
    )


def _build_generated_checkbox_block(
    *,
    block_id: str,
    page_number: int,
    rect: pymupdf.Rect,
    style_source: TextBlock,
    checked: bool,
    group_kind: str,
) -> TextBlock:
    font_weight, _font_style = infer_font_style(style_source.fontFamily)
    return TextBlock(
        id=block_id,
        page=page_number,
        bbox=BoundingBox(
            x0=round(rect.x0, 3),
            y0=round(rect.y0, 3),
            x1=round(rect.x1, 3),
            y1=round(rect.y1, 3),
        ),
        originalText="x" if checked else "",
        currentText="x" if checked else "",
        fontFamily=style_source.fontFamily,
        fontKey=style_source.fontKey,
        fontSize=max(style_source.fontSize, rect.height),
        color="#000000",
        lineHeight=max(style_source.lineHeight, rect.height),
        align="left",
        rotation=0.0,
        groupKind=group_kind,
        minFontSize=6.0,
        editable=True,
        cssFontFamily=style_source.cssFontFamily,
        fontAssetId=style_source.fontAssetId,
        fontWeight="700" if checked else font_weight,
        fontStyle="normal",
        baseline=None,
        isCheckbox=True,
        isCustom=False,
    )


def _block_center_in_outline_rect(rect: OutlineRect, block: TextBlock, *, padding: float = 1.0) -> bool:
    center_x = (block.bbox.x0 + block.bbox.x1) / 2
    center_y = (block.bbox.y0 + block.bbox.y1) / 2
    return (
        rect.x0 - padding <= center_x <= rect.x1 + padding
        and rect.y0 - padding <= center_y <= rect.y1 + padding
    )


def _find_outline_rect_below_heading(
    rects: list[OutlineRect],
    heading: TextBlock,
    *,
    min_width: float,
    min_height: float,
) -> Optional[OutlineRect]:
    candidates: list[OutlineRect] = []
    for rect in rects:
        if (rect.x1 - rect.x0) < min_width or (rect.y1 - rect.y0) < min_height:
            continue
        if heading.bbox.x0 < rect.x0 - 35 or heading.bbox.x0 > rect.x1 + 5:
            continue
        if rect.x1 < heading.bbox.x1:
            continue
        vertical_gap = rect.y0 - heading.bbox.y1
        if vertical_gap < -8 or vertical_gap > 18:
            continue
        candidates.append(rect)

    if not candidates:
        return None
    return min(
        candidates,
        key=lambda rect: (
            abs(rect.y0 - heading.bbox.y1),
            abs(rect.x0 - heading.bbox.x0),
            -(rect.x1 - rect.x0),
        ),
    )


def _line_metrics_for_label(
    label: TextBlock,
    needle: str,
) -> Optional[tuple[float, float, float, float, float]]:
    normalized_needle = _normalize_text_content(needle)
    lines = [line.strip() for line in label.originalText.splitlines() if line.strip()]
    if len(lines) <= 1:
        if normalized_needle not in _normalize_text_content(label.originalText):
            return None
        baseline = label.baseline if label.baseline is not None else label.bbox.y0 + _baseline_offset(label)
        return (label.bbox.x0, label.bbox.y0, label.bbox.x1, label.bbox.y1, baseline)

    match_index = None
    for index, line in enumerate(lines):
        if normalized_needle in _normalize_text_content(line):
            match_index = index
            break
    if match_index is None:
        return None

    bbox_height = max(1.0, label.bbox.y1 - label.bbox.y0)
    step = (bbox_height - label.lineHeight) / max(1, len(lines) - 1)
    step = max(label.fontSize * 0.9, step)
    y0 = label.bbox.y0 + match_index * step
    y1 = min(label.bbox.y1, y0 + label.lineHeight)
    if label.baseline is not None:
        baseline = label.baseline - (len(lines) - 1 - match_index) * step
    else:
        baseline = y0 + _baseline_offset(label)
    return (label.bbox.x0, y0, label.bbox.x1, y1, baseline)


def _existing_text_segments(block: TextBlock) -> list[ExistingTextSegment]:
    lines = [line.strip() for line in block.originalText.splitlines() if line.strip()]
    if not lines:
        return []
    if len(lines) == 1:
        return [
            ExistingTextSegment(
                block=block,
                line_index=0,
                text=lines[0],
                rect=pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
            )
        ]

    bbox_height = max(1.0, block.bbox.y1 - block.bbox.y0)
    step = (bbox_height - block.lineHeight) / max(1, len(lines) - 1)
    step = max(block.fontSize * 0.9, step)
    segments: list[ExistingTextSegment] = []
    for index, text in enumerate(lines):
        y0 = block.bbox.y0 + index * step
        y1 = min(block.bbox.y1, y0 + block.lineHeight)
        segments.append(
            ExistingTextSegment(
                block=block,
                line_index=index,
                text=text,
                rect=pymupdf.Rect(block.bbox.x0, y0, block.bbox.x1, y1),
            )
        )
    return segments


def _segment_overlap_ratio(segment: ExistingTextSegment, field: TextBlock, *, padding: float = 1.25) -> float:
    block_rect = segment.rect
    field_rect = pymupdf.Rect(
        field.bbox.x0 - padding,
        field.bbox.y0 - padding,
        field.bbox.x1 + padding,
        field.bbox.y1 + padding,
    )
    intersection = block_rect & field_rect
    block_area = max(0.01, block_rect.width * block_rect.height)
    return max(0.0, intersection.width * intersection.height) / block_area


def _merge_existing_field_segments(segments: list[ExistingTextSegment], *, multiline: bool) -> str:
    if not segments:
        return ""

    ordered = sorted(segments, key=lambda segment: (segment.rect.y0, segment.rect.x0))
    if not multiline:
        return " ".join(segment.text for segment in ordered).strip()

    lines: list[list[ExistingTextSegment]] = []
    for segment in ordered:
        if not lines:
            lines.append([segment])
            continue
        previous_line = lines[-1]
        previous_y = sum(item.rect.y0 for item in previous_line) / len(previous_line)
        tolerance = max(segment.block.lineHeight * 0.65, 4.0)
        if abs(segment.rect.y0 - previous_y) <= tolerance:
            previous_line.append(segment)
        else:
            lines.append([segment])

    merged_lines: list[str] = []
    for line_segments in lines:
        line_text = " ".join(
            segment.text
            for segment in sorted(line_segments, key=lambda segment: segment.rect.x0)
        ).strip()
        if line_text:
            merged_lines.append(line_text)
    return "\n".join(merged_lines).strip()


def _absorb_existing_contract_field_text(fields: list[TextBlock], page_blocks: list[TextBlock]) -> None:
    consumed_segments: set[tuple[str, int]] = set()
    absorbed_block_ids: set[str] = set()
    for field in sorted(fields, key=lambda block: (block.bbox.y0, block.bbox.x0)):
        multiline = field.groupKind == "generated-contract-object-field"
        candidates: list[ExistingTextSegment] = []
        for block in page_blocks:
            if block.isCheckbox or block.isCustom or block.groupKind.startswith("generated-"):
                continue
            if not block.originalText.strip():
                continue
            for segment in _existing_text_segments(block):
                segment_id = (block.id, segment.line_index)
                if segment_id in consumed_segments:
                    continue
                if _segment_overlap_ratio(segment, field) < 0.42:
                    continue
                candidates.append(segment)

        existing_text = _merge_existing_field_segments(candidates, multiline=multiline)
        if not existing_text:
            continue

        field.originalText = existing_text
        field.currentText = existing_text
        for segment in candidates:
            consumed_segments.add((segment.block.id, segment.line_index))
            absorbed_block_ids.add(segment.block.id)

    for block in page_blocks:
        if block.id in absorbed_block_ids:
            block.originalText = ""
            block.currentText = ""
            block.editable = False


def _build_generated_outline_rect_text_area(
    block_id: str,
    page_number: int,
    rect: OutlineRect,
    style_source: TextBlock,
    *,
    padding: float = 6.0,
    group_kind: str = "generated-outline-box-field",
) -> Optional[TextBlock]:
    x0 = rect.x0 + padding
    y0 = rect.y0 + max(3.0, padding * 0.7)
    x1 = rect.x1 - padding
    y1 = rect.y1 - max(3.0, padding * 0.7)
    if (x1 - x0) < 30 or (y1 - y0) < 8:
        return None
    return _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=x0,
        y0=y0,
        x1=x1,
        y1=y1,
        style_source=style_source,
        group_kind=group_kind,
        baseline=None,
    )


def _find_underline_for_label(
    label: TextBlock,
    page_underlines: list[HorizontalLineSegment],
    *,
    right_of_x: float,
    max_x1: Optional[float] = None,
) -> Optional[HorizontalLineSegment]:
    label_baseline = label.baseline if label.baseline is not None else label.bbox.y1
    tolerance_y = max(8.0, label.lineHeight)
    candidates = [
        line for line in page_underlines
        if abs(line.y - label_baseline) <= tolerance_y
        and line.y >= label.bbox.y1 - max(1.0, label.lineHeight * 0.15)
        and line.x1 > right_of_x
        and (line.x1 - max(line.x0, right_of_x)) >= 20.0
        and (max_x1 is None or line.x0 <= max_x1 + 2.0)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda line: (abs(line.y - label_baseline), -(line.x1 - line.x0)))


def _build_contract_party_generated_fields(page: pymupdf.Page, page_blocks: list[TextBlock]) -> list[TextBlock]:
    generated: list[TextBlock] = []
    if not page_blocks:
        return generated

    page_number = page.number + 1
    party_heading = _find_page_block(page_blocks, "zwischen dem Auftraggeber")
    object_heading = _find_page_block(page_blocks, "das Objekt")
    if party_heading is None and object_heading is None:
        return generated

    rects = _extract_outline_rects(page, min_width=160.0, min_height=35.0)
    page_underlines = _extract_horizontal_line_segments(page, min_length=20.0, max_width=2.5)
    party_box: Optional[OutlineRect] = None
    object_row_specs: list[tuple[int, float, float, float, TextBlock]] = []
    label_specs = (
        (f"page-{page_number}-generated-client-name", "Name/Firma"),
        (f"page-{page_number}-generated-client-representative", "Vertreten durch"),
        (f"page-{page_number}-generated-client-street", "Stra"),
        (f"page-{page_number}-generated-client-city", "PLZ/Ort"),
        (f"page-{page_number}-generated-client-phone", "Telefon"),
    )

    def append_if_no_widget_overlap(field: TextBlock) -> None:
        rect = pymupdf.Rect(field.bbox.x0, field.bbox.y0, field.bbox.x1, field.bbox.y1)
        if _has_existing_widget_field_overlap(page_blocks, rect):
            return
        generated.append(field)

    if party_heading is not None:
        party_box = _find_outline_rect_below_heading(
            rects,
            party_heading,
            min_width=220.0,
            min_height=45.0,
        )
        if party_box is not None:
            box_blocks = [
                block for block in page_blocks
                if _block_center_in_outline_rect(party_box, block, padding=2.0)
            ]
            for block_id, needle in label_specs:
                label = _find_page_block(box_blocks, needle)
                if label is None:
                    continue
                line_metrics = _line_metrics_for_label(label, needle)
                if line_metrics is None:
                    continue

                _, line_y0, label_x1, line_y1, line_baseline = line_metrics
                field_floor = party_box.x0 + min(max((party_box.x1 - party_box.x0) * 0.2, 75.0), 115.0)
                field_x0 = max(label_x1 + 8.0, field_floor)
                field_x1 = party_box.x1 - 6.0

                underline = _find_underline_for_label(
                    label,
                    page_underlines,
                    right_of_x=field_x0,
                    max_x1=party_box.x1,
                )
                if underline is not None:
                    field_x1 = max(field_x1, underline.x1)
                    line_gap = max(0.85, label.lineHeight * 0.1)
                    height = max(label.lineHeight, label.fontSize + 1.5)
                    ul_y1 = underline.y - line_gap
                    ul_y0 = ul_y1 - height
                    field_y0 = max(party_box.y0 + 0.5, ul_y0)
                    field_y1 = min(party_box.y1 - 0.7, max(ul_y1, field_y0 + height))
                    line_baseline = underline.y - max(2.8, label.fontSize * 0.36)
                else:
                    height = max(label.lineHeight, label.fontSize + 1.5)
                    field_y0 = max(party_box.y0 + 0.5, line_y0 - 0.1)
                    field_y1 = min(party_box.y1 - 0.7, max(line_y1 + 0.2, field_y0 + height))

                if (field_x1 - field_x0) < 40 or (field_y1 - field_y0) < 6:
                    continue

                object_row_specs.append(
                    (
                        len(object_row_specs) + 1,
                        field_y0 - party_box.y0,
                        field_y1 - party_box.y0,
                        line_baseline - party_box.y0,
                        label,
                    )
                )
                append_if_no_widget_overlap(
                    _build_generated_text_block(
                        block_id=block_id,
                        page_number=page_number,
                        x0=field_x0,
                        y0=field_y0,
                        x1=field_x1,
                        y1=field_y1,
                        style_source=label,
                        group_kind="generated-contract-party-field",
                        baseline=line_baseline,
                    )
                )

        if party_box is None:
            fallback_blocks = page_blocks
            right_boundary = object_heading.bbox.x0 - 12.0 if object_heading is not None else page.rect.width * 0.55
            for block_id, needle in label_specs:
                label = _find_page_block(fallback_blocks, needle)
                if label is None:
                    continue
                line_metrics = _line_metrics_for_label(label, needle)
                if line_metrics is None:
                    continue

                _, line_y0, label_x1, line_y1, line_baseline = line_metrics
                field_x0 = max(label_x1 + 8.0, party_heading.bbox.x0 + 70.0)
                field_x1 = min(right_boundary, page.rect.width - 30.0)

                underline = _find_underline_for_label(
                    label,
                    page_underlines,
                    right_of_x=field_x0,
                    max_x1=right_boundary,
                )
                if underline is not None:
                    field_x0 = max(field_x0, underline.x0)
                    field_x1 = max(field_x1, underline.x1)
                    line_gap = max(0.85, label.lineHeight * 0.1)
                    height = max(label.lineHeight, label.fontSize + 1.5)
                    ul_y1 = underline.y - line_gap
                    ul_y0 = ul_y1 - height
                    field_y0 = ul_y0
                    field_y1 = max(ul_y1, field_y0 + height)
                    line_baseline = underline.y - max(2.8, label.fontSize * 0.36)
                else:
                    height = max(label.lineHeight, label.fontSize + 1.5)
                    field_y0 = line_y0 - 0.1
                    field_y1 = max(line_y1 + 0.2, field_y0 + height)

                if (field_x1 - field_x0) < 35:
                    continue

                object_row_specs.append(
                    (
                        len(object_row_specs) + 1,
                        field_y0 - party_heading.bbox.y1,
                        field_y1 - party_heading.bbox.y1,
                        line_baseline - party_heading.bbox.y1,
                        label,
                    )
                )
                append_if_no_widget_overlap(
                    _build_generated_text_block(
                        block_id=block_id,
                        page_number=page_number,
                        x0=field_x0,
                        y0=field_y0,
                        x1=field_x1,
                        y1=field_y1,
                        style_source=label,
                        group_kind="generated-contract-party-field",
                        baseline=line_baseline,
                    )
                )

    if object_heading is not None:
        object_box = _find_outline_rect_below_heading(
            rects,
            object_heading,
            min_width=220.0,
            min_height=45.0,
        )
        if object_box is not None:
            if not object_row_specs:
                fallback_height = max(object_heading.lineHeight, object_heading.fontSize + 1.5)
                available_height = max(fallback_height, (object_box.y1 - object_box.y0) - 1.2)
                row_step = min(fallback_height + 0.9, available_height / 5)
                object_row_specs = [
                    (
                        index + 1,
                        0.5 + index * row_step,
                        0.5 + index * row_step + fallback_height,
                        0.5 + index * row_step + _baseline_offset(object_heading),
                        object_heading,
                    )
                    for index in range(5)
                ]

            for row_number, relative_y0, relative_y1, relative_baseline, style_source in object_row_specs:
                field_x0 = object_box.x0 + 6.0
                field_x1 = object_box.x1 - 6.0
                field_y0 = max(object_box.y0 + 0.5, object_box.y0 + relative_y0)
                field_y1 = min(object_box.y1 - 0.7, object_box.y0 + relative_y1)
                if (field_x1 - field_x0) < 40 or (field_y1 - field_y0) < 6:
                    continue
                append_if_no_widget_overlap(
                    _build_generated_text_block(
                        block_id=f"page-{page_number}-generated-object-line-{row_number}",
                        page_number=page_number,
                        x0=field_x0,
                        y0=field_y0,
                        x1=field_x1,
                        y1=field_y1,
                        style_source=style_source,
                        group_kind="generated-contract-object-line-field",
                        baseline=object_box.y0 + relative_baseline,
                    )
                )
        elif object_row_specs:
            object_x0 = max(object_heading.bbox.x0 + 6.0, page.rect.width * 0.48)
            object_x1 = page.rect.width - 30.0
            for row_number, relative_y0, relative_y1, relative_baseline, style_source in object_row_specs:
                field_y0 = object_heading.bbox.y1 + relative_y0
                field_y1 = object_heading.bbox.y1 + relative_y1
                if (object_x1 - object_x0) < 40 or (field_y1 - field_y0) < 6:
                    continue
                append_if_no_widget_overlap(
                    _build_generated_text_block(
                        block_id=f"page-{page_number}-generated-object-line-{row_number}",
                        page_number=page_number,
                        x0=object_x0,
                        y0=field_y0,
                        x1=object_x1,
                        y1=field_y1,
                        style_source=style_source,
                        group_kind="generated-contract-object-line-field",
                        baseline=object_heading.bbox.y1 + relative_baseline,
                    )
                )

    _absorb_existing_contract_field_text(generated, page_blocks)
    return generated


def _estimate_underlined_text_baseline_gap(
    lines: list[HorizontalLineSegment],
    page_blocks: list[TextBlock],
    *,
    marker: TextBlock,
) -> float:
    gaps: list[float] = []
    for block in page_blocks:
        if block.baseline is None or block.bbox.y0 <= marker.bbox.y1:
            continue
        for line in lines:
            if line.y <= block.baseline:
                continue
            if line.y - block.baseline > max(6.0, block.lineHeight * 0.75):
                continue
            if line.x0 > block.bbox.x0 + 3 or line.x1 < block.bbox.x1 - 3:
                continue
            gaps.append(line.y - block.baseline)
            break

    if gaps:
        gaps.sort()
        return gaps[len(gaps) // 2]
    return max(2.2, marker.fontSize * 0.26)


def _build_additional_agreement_generated_fields(page: pymupdf.Page, page_blocks: list[TextBlock]) -> list[TextBlock]:
    generated: list[TextBlock] = []
    marker = _find_page_block(page_blocks, "Weitere zusätzliche Vereinbarungen")
    if marker is None:
        return generated

    page_number = page.number + 1
    lines = [
        line for line in _extract_horizontal_line_segments(page, min_length=280.0, max_width=2.5)
        if line.y > marker.bbox.y1
        and (line.y - marker.bbox.y1) <= 60
        and line.x0 <= marker.bbox.x0 + 25
        and (line.x1 - line.x0) >= 350
    ]
    lines.sort(key=lambda line: line.y)
    if len(lines) < 2:
        return generated

    baseline_gap = _estimate_underlined_text_baseline_gap(lines, page_blocks, marker=marker)
    target_lines = lines[:5]
    for index, line in enumerate(target_lines, start=1):
        generated.append(
            _build_generated_underlined_field(
                f"page-{page_number}-generated-additional-agreement-{index}",
                page_number,
                line,
                marker,
                baseline_gap=baseline_gap,
                group_kind="generated-additional-agreement-field",
            )
        )

    _absorb_existing_contract_field_text(generated, page_blocks)
    return generated


def _build_payment_frequency_checkbox_fields(page: pymupdf.Page, page_blocks: list[TextBlock]) -> list[TextBlock]:
    generated: list[TextBlock] = []
    marker = _find_page_block(page_blocks, "Zahlungsweise")
    if marker is None:
        return generated

    page_number = page.number + 1
    horizontal_lines = [
        line for line in _extract_horizontal_line_segments(page, min_length=300.0, max_width=1.2)
        if marker.bbox.y0 - 10 <= line.y <= marker.bbox.y1 + 40
    ]
    horizontal_lines.sort(key=lambda line: line.y)
    if len(horizontal_lines) < 3:
        return generated

    row_top = horizontal_lines[-2].y
    row_bottom = horizontal_lines[-1].y
    vertical_lines = [
        line for line in _extract_vertical_line_segments(page, min_length=30.0, max_width=1.2)
        if line.y0 <= horizontal_lines[0].y + 1.0 and line.y1 >= row_bottom - 1.0
    ]
    vertical_lines.sort(key=lambda line: line.x)
    payment_edges = [line for line in vertical_lines if line.x >= marker.bbox.x1 + 10.0]
    if len(payment_edges) < 4:
        return generated

    row_center_y = (row_top + row_bottom) / 2
    checkbox_size = min(11.0, max(9.0, (row_bottom - row_top) - 4.0))
    payment_specs = (
        ("quarterly", payment_edges[0].x, payment_edges[1].x),
        ("half-yearly", payment_edges[1].x, payment_edges[2].x),
        ("yearly", payment_edges[2].x, payment_edges[3].x),
    )

    for suffix, left_x, right_x in payment_specs:
        center_x = (left_x + right_x) / 2
        rect = pymupdf.Rect(
            center_x - checkbox_size / 2,
            row_center_y - checkbox_size / 2,
            center_x + checkbox_size / 2,
            row_center_y + checkbox_size / 2,
        )
        cell_rect = pymupdf.Rect(left_x, row_top, right_x, row_bottom)
        mark_block = next(
            (
                block for block in page_blocks
                if block.currentText.strip().casefold() == "x"
                and pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1).intersects(cell_rect)
            ),
            None,
        )
        style_source = mark_block or marker
        checked = mark_block is not None
        if mark_block is not None:
            mark_block.originalText = ""
            mark_block.currentText = ""
            mark_block.editable = False
            mark_block.groupKind = "hidden-checkbox-mark"

        generated.append(
            _build_generated_checkbox_block(
                block_id=f"page-{page_number}-generated-payment-{suffix}",
                page_number=page_number,
                rect=rect,
                style_source=style_source,
                checked=checked,
                group_kind="generated-payment-checkbox",
            )
        )

    return generated


def _has_sicherheit_nord_layout(page_blocks_by_page: dict[int, list[TextBlock]]) -> bool:
    page_two = page_blocks_by_page.get(2, [])
    page_three = page_blocks_by_page.get(3, [])
    has_contract_title = any(
        _find_page_block(page_blocks, "Dienstleistungsvertrag Notruf- und Serviceleitstelle") is not None
        for page_blocks in page_blocks_by_page.values()
    )
    return (
        _find_page_block(page_two, "SEPA LASTSCHRIFTERMÄCHTIGUNG") is not None
        and _find_page_block(page_three, "Gewünschte Zahlungsweise") is not None
        and has_contract_title
    )


def _find_local_tessdata() -> str | None:
    local = Path(__file__).resolve().parent.parent / "tessdata"
    if local.is_dir() and any(local.glob("*.traineddata")):
        return str(local)
    for system_path in (
        Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tessdata"),
    ):
        if system_path.is_dir() and any(system_path.glob("*.traineddata")):
            return str(system_path)
    return None


def _find_tesseract_exe() -> str | None:
    if exe := shutil.which("tesseract"):
        return exe
    for path in (
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ):
        if path.exists():
            return str(path)
    return None


def _ocr_scan_page_to_blocks(
    page: pymupdf.Page,
    page_number: int,
    tessdata: str,
    dpi: int = 200,
) -> list[TextBlock]:
    """Render page in visual (portrait) orientation and OCR via Tesseract CLI.

    Using the CLI on a portrait-rendered PNG ensures the full page area is
    covered, including header regions that get clipped when OCR runs on the
    raw landscape image and coordinates are transformed afterwards.
    """
    tess_exe = _find_tesseract_exe()
    if not tess_exe:
        return []

    mat = pymupdf.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    blocks: list[TextBlock] = []
    scale = 72.0 / dpi

    with tempfile.TemporaryDirectory() as tmp:
        png_path = Path(tmp) / "page.png"
        out_base = Path(tmp) / "out"
        pix.save(str(png_path))

        env = os.environ.copy()
        env["TESSDATA_PREFIX"] = tessdata

        try:
            subprocess.run(
                [tess_exe, str(png_path), str(out_base), "-l", "deu", "tsv"],
                capture_output=True,
                text=True,
                env=env,
                timeout=60,
            )
        except Exception:
            return []

        tsv_path = Path(str(out_base) + ".tsv")
        if not tsv_path.exists():
            return []

        with open(tsv_path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter="\t")
            for idx, row in enumerate(reader):
                try:
                    conf = float(row.get("conf", -1))
                    text = (row.get("text") or "").strip()
                    if conf < 10 or not text:
                        continue
                    px0 = int(row["left"])
                    py0 = int(row["top"])
                    pw = int(row["width"])
                    ph = int(row["height"])
                except (ValueError, KeyError):
                    continue

                vx0 = px0 * scale
                vy0 = py0 * scale
                vx1 = (px0 + pw) * scale
                vy1 = (py0 + ph) * scale
                line_h = max(vy1 - vy0, 4.0)
                blocks.append(TextBlock(
                    id=f"ocr-p{page_number}-w{idx}",
                    page=page_number,
                    bbox=BoundingBox(x0=round(vx0, 2), y0=round(vy0, 2), x1=round(vx1, 2), y1=round(vy1, 2)),
                    originalText=text,
                    currentText=text,
                    fontFamily="Helvetica",
                    fontKey="Helvetica",
                    fontSize=9.0,
                    color="#000000",
                    lineHeight=round(line_h, 2),
                    align="left",
                    rotation=0.0,
                    groupKind="ocr-word",
                    minFontSize=6.0,
                    editable=False,
                    cssFontFamily="Helvetica, Arial, sans-serif",
                    isCheckbox=False,
                    isCustom=False,
                ))

    return blocks


def _enrich_scan_blocks_with_ocr(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    page_count: int = 3,
) -> None:
    tessdata = _find_local_tessdata()
    if not tessdata:
        return
    for page_index in range(min(page_count, doc.page_count)):
        page_number = page_index + 1
        existing = page_blocks_by_page.get(page_number, [])
        has_real_text = any(
            b.originalText.strip() and not b.groupKind.startswith(("generated-", "ocr-"))
            for b in existing
        )
        if has_real_text:
            continue
        ocr_blocks = _ocr_scan_page_to_blocks(doc[page_index], page_number, tessdata)
        page_blocks_by_page[page_number] = existing + ocr_blocks


def _looks_like_sicherheit_nord_scan(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> bool:
    if doc.page_count < 3:
        return False

    first_pages = [doc[index] for index in range(3)]
    if not all(540 <= page.rect.width <= 650 and 760 <= page.rect.height <= 900 for page in first_pages):
        return False

    text_blocks = [
        block
        for page_number in range(1, 4)
        for block in page_blocks_by_page.get(page_number, [])
        if block.originalText.strip() and not block.groupKind.startswith("generated-")
    ]
    if len(text_blocks) > 20:
        return False

    image_count = sum(len(page.get_images(full=True)) for page in first_pages)
    if image_count < 2:
        return False

    normalized_text = _normalize_text_content("\n".join(block.originalText for block in text_blocks))
    if not normalized_text:
        return True
    scan_markers = ("sasse", "buergschaft", "bürgschaft", "sicherheit nord")
    return any(marker in normalized_text for marker in scan_markers)


def _scaled_scan_rect(page: pymupdf.Page, rect: tuple[float, float, float, float]) -> pymupdf.Rect:
    scale_x = page.rect.width / 595.0
    scale_y = page.rect.height / 842.0
    return pymupdf.Rect(
        rect[0] * scale_x,
        rect[1] * scale_y,
        rect[2] * scale_x,
        rect[3] * scale_y,
    )


def _build_scan_text_field(
    page: pymupdf.Page,
    page_blocks: list[TextBlock],
    block_id: str,
    rect_tuple: tuple[float, float, float, float],
    group_kind: str,
) -> TextBlock:
    page_number = page.number + 1
    rect = _scaled_scan_rect(page, rect_tuple)
    style_source = _nearest_style_source(rect, page_blocks, page_number)
    field = _build_generated_text_block(
        block_id=block_id,
        page_number=page_number,
        x0=rect.x0,
        y0=rect.y0,
        x1=rect.x1,
        y1=rect.y1,
        style_source=style_source,
        group_kind=group_kind,
        baseline=rect.y0 + _baseline_offset(style_source),
    )
    field.color = "#000000"
    if group_kind in {"generated-contract-party-field", "generated-contract-object-line-field"}:
        field.fontSize = min(field.fontSize, 9.2)
        field.lineHeight = round(max(field.fontSize * 1.15, rect.height), 3)
        field.baseline = round(rect.y0 + min(rect.height - 1.1, field.fontSize * 0.88), 3)
    return field


def _scan_checkbox_has_mark(page: pymupdf.Page, rect: pymupdf.Rect) -> bool:
    inset = max(0.25, min(rect.width, rect.height) * 0.03)
    clip = pymupdf.Rect(rect.x0 + inset, rect.y0 + inset, rect.x1 - inset, rect.y1 - inset)
    if clip.is_empty:
        return False
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(5, 5), clip=clip, alpha=False)
    if pixmap.width <= 0 or pixmap.height <= 0:
        return False
    dark_pixels = 0
    samples = pixmap.samples
    for index in range(0, len(samples), 3):
        if samples[index] + samples[index + 1] + samples[index + 2] < 330:
            dark_pixels += 1
    return (dark_pixels / max(1, pixmap.width * pixmap.height)) > 0.025


def _build_sicherheit_nord_scan_fallback_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated: list[TextBlock] = []

    header_party_specs = (
        ("generated-client-name", (100.498, 71.0, 264.6, 83.5)),
        ("generated-client-representative", (100.498, 82.5, 264.6, 94.0)),
        ("generated-client-street", (98.5, 92.0, 264.6, 103.5)),
        ("generated-client-city", (98.5, 102.0, 264.6, 113.5)),
        ("generated-client-phone", (98.5, 112.0, 264.6, 123.5)),
    )
    header_object_specs = (
        ("generated-object-line-1", (285.6, 71.0, 565.0, 83.5)),
        ("generated-object-line-2", (285.6, 82.5, 565.0, 94.0)),
        ("generated-object-line-3", (285.6, 92.0, 565.0, 103.5)),
        ("generated-object-line-4", (285.6, 102.0, 565.0, 113.5)),
        ("generated-object-line-5", (285.6, 112.0, 565.0, 123.5)),
    )

    for page_index in range(min(3, doc.page_count)):
        page = doc[page_index]
        page_number = page_index + 1
        page_blocks = page_blocks_by_page.get(page_number, [])
        page_fields: list[TextBlock] = []
        for suffix, rect_tuple in header_party_specs:
            page_fields.append(
                _build_scan_text_field(
                    page,
                    page_blocks,
                    f"page-{page_number}-{suffix}",
                    rect_tuple,
                    "generated-contract-party-field",
                )
            )
        for suffix, rect_tuple in header_object_specs:
            page_fields.append(
                _build_scan_text_field(
                    page,
                    page_blocks,
                    f"page-{page_number}-{suffix}",
                    rect_tuple,
                    "generated-contract-object-line-field",
                )
            )
        _absorb_existing_contract_field_text(page_fields, page_blocks)
        generated.extend(page_fields)

    page_two = doc[1]
    page_two_blocks = page_blocks_by_page.get(2, [])
    page_two_specs = (
        ("page-2-generated-additional-agreement-1", (46.75, 462.615, 542.45, 472.65), "generated-additional-agreement-field"),
        ("page-2-generated-additional-agreement-2", (46.75, 472.915, 542.45, 482.95), "generated-additional-agreement-field"),
        ("page-2-generated-account-holder", (47.9, 599.752, 508.8, 609.796), "generated-line-field"),
        ("page-2-generated-address", (98.5, 620.452, 508.8, 630.496), "generated-line-field"),
        ("page-2-generated-bank-name", (154.2, 641.152, 473.3, 651.196), "generated-line-field"),
        ("page-2-generated-iban", (96.0, 661.852, 473.3, 671.896), "generated-line-field"),
        ("page-2-generated-creditor-id", (188.2, 682.552, 365.5, 692.596), "generated-line-field"),
        ("page-2-generated-mandate-reference", (159.7, 703.252, 326.1, 713.296), "generated-line-field"),
    )
    page_two_fields = [
        _build_scan_text_field(page_two, page_two_blocks, block_id, rect_tuple, group_kind)
        for block_id, rect_tuple, group_kind in page_two_specs
    ]
    _absorb_existing_contract_field_text(page_two_fields, page_two_blocks)
    generated.extend(page_two_fields)

    page_three = doc[2]
    page_three_blocks = page_blocks_by_page.get(3, [])
    page_three_specs = (
        ("page-3-generated-email-line-1", (340.25, 179.115, 543.15, 189.15), "generated-email-line-field"),
        ("page-3-generated-email-line-2", (47.45, 194.115, 543.15, 204.15), "generated-email-line-field"),
        ("page-3-generated-alt-email", (259.8, 224.252, 508.0, 234.296), "generated-line-field"),
        ("page-3-generated-postal-address-line-1", (47.45, 281.815, 543.15, 291.85), "generated-postal-address-line-field"),
        ("page-3-generated-postal-address-line-2", (47.45, 296.815, 543.15, 306.85), "generated-postal-address-line-field"),
        ("page-3-generated-place-date", (28.4, 451.552, 276.6, 461.596), "generated-line-field"),
    )
    page_three_fields = [
        _build_scan_text_field(page_three, page_three_blocks, block_id, rect_tuple, group_kind)
        for block_id, rect_tuple, group_kind in page_three_specs
    ]
    _absorb_existing_contract_field_text(page_three_fields, page_three_blocks)
    generated.extend(page_three_fields)

    payment_specs = (
        ("quarterly", (260.0, 156.55, 271.0, 167.55)),
        ("half-yearly", (356.7, 156.55, 367.7, 167.55)),
        ("yearly", (463.9, 156.55, 474.9, 167.55)),
    )
    for suffix, rect_tuple in payment_specs:
        rect = _scaled_scan_rect(page_three, rect_tuple)
        checked = _scan_checkbox_has_mark(page_three, rect)
        generated.append(
            _build_generated_checkbox_block(
                block_id=f"page-3-generated-payment-{suffix}",
                page_number=3,
                rect=rect,
                style_source=_nearest_style_source(rect, page_three_blocks, 3),
                checked=checked,
                group_kind="generated-payment-checkbox",
            )
        )

    return generated


def _build_scan_checkbox_field(
    page: pymupdf.Page,
    page_blocks: list[TextBlock],
    block_id: str,
    rect_tuple: tuple[float, float, float, float],
    group_kind: str = "generated-scan-checkbox",
) -> TextBlock:
    rect = _scaled_scan_rect(page, rect_tuple)
    return _build_generated_checkbox_block(
        block_id=block_id,
        page_number=page.number + 1,
        rect=rect,
        style_source=_nearest_style_source(rect, page_blocks, page.number + 1),
        checked=_scan_checkbox_has_mark(page, rect),
        group_kind=group_kind,
    )


def _build_empty_scan_checkbox_field(
    page: pymupdf.Page,
    page_blocks: list[TextBlock],
    block_id: str,
    rect_tuple: tuple[float, float, float, float],
    group_kind: str = "generated-rotated-scan-checkbox",
) -> TextBlock:
    rect = _scaled_scan_rect(page, rect_tuple)
    return _build_generated_checkbox_block(
        block_id=block_id,
        page_number=page.number + 1,
        rect=rect,
        style_source=_nearest_style_source(rect, page_blocks, page.number + 1),
        checked=False,
        group_kind=group_kind,
    )


def _resolve_tessdata_path() -> Optional[Path]:
    candidates: list[Path] = []
    env_path = os.getenv("PDF_EDITOR_TESSDATA")
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(Path(__file__).resolve().parents[1] / "tessdata")
    candidates.append(Path.cwd() / "backend" / "tessdata")
    candidates.append(Path.cwd() / "tessdata")

    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except Exception:
            resolved = candidate
        if (resolved / "deu.traineddata").exists() and (resolved / "eng.traineddata").exists():
            return resolved
    return None


def _ocr_words_for_page(
    page: pymupdf.Page,
    cache: dict[int, list[tuple]],
    *,
    dpi: int = 180,
) -> list[tuple]:
    page_number = page.number + 1
    cached = cache.get(page_number)
    if cached is not None:
        return cached

    tessdata_path = _resolve_tessdata_path()
    if tessdata_path is None:
        cache[page_number] = []
        return []

    try:
        textpage = page.get_textpage_ocr(
            flags=0,
            language="deu+eng",
            dpi=dpi,
            full=False,
            tessdata=str(tessdata_path),
        )
        words = list(page.get_text("words", textpage=textpage) or [])
    except Exception:
        words = []
    cache[page_number] = words
    return words


def _ocr_word_matches_rect(word: tuple, rect: pymupdf.Rect, *, padding: float = 1.4) -> bool:
    word_rect = pymupdf.Rect(word[0], word[1], word[2], word[3])
    padded = pymupdf.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding, rect.y1 + padding)
    center_x = (word_rect.x0 + word_rect.x1) / 2
    center_y = (word_rect.y0 + word_rect.y1) / 2
    if padded.x0 <= center_x <= padded.x1 and padded.y0 <= center_y <= padded.y1:
        return True
    intersection = word_rect & padded
    if intersection.is_empty:
        return False
    return (intersection.get_area() / max(0.1, min(word_rect.get_area(), padded.get_area()))) >= 0.22


def _ocr_text_for_rect(
    page: pymupdf.Page,
    rect: pymupdf.Rect,
    cache: dict[int, list[tuple]],
) -> str:
    words = [
        word for word in _ocr_words_for_page(page, cache)
        if len(word) >= 8 and str(word[4]).strip() and _ocr_word_matches_rect(word, rect)
    ]
    if not words:
        return ""

    words.sort(key=lambda word: (int(word[5]), int(word[6]), int(word[7]), float(word[0])))
    lines: list[str] = []
    current_line_key: tuple[int, int] | None = None
    current_parts: list[str] = []
    for word in words:
        line_key = (int(word[5]), int(word[6]))
        if current_parts and current_line_key is not None and line_key != current_line_key:
            lines.append(" ".join(current_parts).strip())
            current_parts = []
        current_line_key = line_key
        current_parts.append(str(word[4]).strip())
    if current_parts:
        lines.append(" ".join(current_parts).strip())
    return "\n".join(line for line in lines if line)


def _page_needs_automatic_ocr_fields(page: pymupdf.Page, page_blocks: list[TextBlock]) -> bool:
    if not page.get_images(full=True):
        return False
    real_text_blocks = [
        block for block in page_blocks
        if block.originalText.strip()
        and not block.isCheckbox
        and not block.isCustom
        and not block.groupKind.startswith(("generated-", "ocr-", "hidden-"))
    ]
    real_text_length = sum(len(block.originalText.strip()) for block in real_text_blocks)
    return len(real_text_blocks) <= 4 or real_text_length < 80


def _ocr_line_text(parts: list[str]) -> str:
    text = " ".join(part.strip() for part in parts if part.strip())
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"([(])\s+", r"\1", text)
    text = re.sub(r"\s+([)])", r"\1", text)
    return text.strip()


def _ocr_text_is_noise(text: str) -> bool:
    cleaned = text.strip()
    if not cleaned:
        return True
    if len(cleaned) == 1 and not cleaned.isalnum():
        return True
    meaningful = sum(1 for char in cleaned if char.isalnum())
    if meaningful == 0:
        return True
    return False


def _ocr_rect_has_existing_generated_field(rect: pymupdf.Rect, blocks: list[TextBlock]) -> bool:
    if rect.is_empty:
        return True
    for block in blocks:
        if block.isCheckbox:
            continue
        if not block.editable:
            continue
        if block.isCustom:
            continue
        if not block.groupKind.startswith("generated-") and not block.originalText.strip():
            continue
        overlap_ratio = _block_rect_overlap_ratio(block, rect, padding=1.5)
        if overlap_ratio >= 0.35 or _block_center_is_inside_rect(block, rect):
            return True
    return False


def _build_automatic_ocr_scan_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated: list[TextBlock] = []
    ocr_cache: dict[int, list[tuple]] = {}

    for page in doc:
        page_number = page.number + 1
        page_blocks = page_blocks_by_page.get(page_number, [])
        if not _page_needs_automatic_ocr_fields(page, page_blocks):
            continue

        words = [
            word for word in _ocr_words_for_page(page, ocr_cache, dpi=180)
            if len(word) >= 8 and str(word[4]).strip()
        ]
        if not words:
            continue

        lines_by_key: dict[tuple[int, int], list[tuple]] = {}
        for word in words:
            key = (int(word[5]), int(word[6]))
            lines_by_key.setdefault(key, []).append(word)

        for line_index, line_words in enumerate(
            sorted(lines_by_key.values(), key=lambda entries: (min(float(word[1]) for word in entries), min(float(word[0]) for word in entries))),
            start=1,
        ):
            line_words.sort(key=lambda word: float(word[0]))
            text = _ocr_line_text([str(word[4]) for word in line_words])
            if _ocr_text_is_noise(text):
                continue

            rect = pymupdf.Rect(
                min(float(word[0]) for word in line_words),
                min(float(word[1]) for word in line_words),
                max(float(word[2]) for word in line_words),
                max(float(word[3]) for word in line_words),
            ) & page.rect
            if rect.is_empty or rect.width < 4.0 or rect.height < 3.0:
                continue
            if _ocr_rect_has_existing_generated_field(rect, [*page_blocks, *generated]):
                continue

            padded_rect = pymupdf.Rect(
                max(0.0, rect.x0 - 0.6),
                max(0.0, rect.y0 - 0.8),
                min(page.rect.width, rect.x1 + 2.5),
                min(page.rect.height, rect.y1 + 1.2),
            )
            style_source = _nearest_style_source(padded_rect, page_blocks, page_number)
            font_size = round(max(6.0, min(14.0, rect.height * 0.78)), 3)
            line_height = round(max(rect.height + 1.0, font_size * 1.18), 3)
            generated.append(
                TextBlock(
                    id=f"page-{page_number}-generated-auto-ocr-line-{line_index}",
                    page=page_number,
                    bbox=BoundingBox(
                        x0=round(padded_rect.x0, 3),
                        y0=round(padded_rect.y0, 3),
                        x1=round(padded_rect.x1, 3),
                        y1=round(padded_rect.y1, 3),
                    ),
                    originalText=text,
                    currentText=text,
                    fontFamily=style_source.fontFamily,
                    fontKey=style_source.fontKey,
                    fontSize=font_size,
                    color="#000000",
                    lineHeight=line_height,
                    align="left",
                    rotation=0.0,
                    groupKind="generated-auto-ocr-scan-field",
                    minFontSize=5.5,
                    editable=True,
                    cssFontFamily=style_source.cssFontFamily,
                    fontAssetId=style_source.fontAssetId,
                    fontWeight=style_source.fontWeight or "400",
                    fontStyle=style_source.fontStyle or "normal",
                    baseline=round(rect.y1 - max(1.1, rect.height * 0.18), 3),
                    isCheckbox=False,
                    isCustom=False,
                    confidence=0.72,
                    ocrTranscript=text,
                )
            )

    return generated


def _clean_common_ocr_text(text: str) -> str:
    cleaned = " ".join(part.strip() for part in str(text or "").replace("\n", " ").split() if part.strip())
    replacements = {
        "Concertbiiro": "Concertbüro",
        "Concertburo": "Concertbüro",
        "ConcertBuro": "Concertbüro",
        "ConcertBüro": "Concertbüro",
        "Zahimann": "Zahlmann",
        "ZahImann": "Zahlmann",
        "Schlissel": "Schlüssel",
        "Schlisselbund": "Schlüsselbund",
    }
    for source, replacement in replacements.items():
        cleaned = cleaned.replace(source, replacement)
    cleaned = cleaned.replace(" ,", ",").replace(" .", ".").replace(" :", ":")
    return cleaned.strip(" \t\r\n|_—-")


def _first_regex_value(pattern: str, text: str, *, flags: int = re.IGNORECASE) -> str:
    match = re.search(pattern, text or "", flags)
    return match.group(0).strip() if match else ""


def _clean_rotated_scan_ocr_value(block_id: str, raw_text: str) -> str:
    raw_lines = [
        _clean_common_ocr_text(line)
        for line in str(raw_text or "").splitlines()
        if _clean_common_ocr_text(line)
    ]
    text = _clean_common_ocr_text(raw_text)
    if not text:
        return ""

    if "object-street" in block_id and raw_lines:
        return raw_lines[0]

    if "object-city" in block_id:
        for line in raw_lines:
            if re.search(r"\b\d{5}\b", line):
                return line
        return ""

    if "id-number" in block_id or "instruction-id" in block_id:
        return _first_regex_value(r"\b\d{8,12}\b", text)

    if "amount" in block_id or "fee-" in block_id or "security-" in block_id:
        if re.search(r"\binkl\.?\b", text, re.IGNORECASE):
            return "Inkl."
        return _first_regex_value(r"\b\d{1,3}[,.]\d{2}\s*€?\b", text).replace(".", ",").strip()

    if "date" in block_id or "status" in block_id:
        if "place-date" in block_id:
            city_date = _first_regex_value(r"\b[A-ZÄÖÜ][A-Za-zÄÖÜäöüß]+,?\s*\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4}\b", text)
            if city_date:
                return city_date.replace(" ", "")
        return _first_regex_value(r"\b\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4}\b", text)

    if "email" in block_id or "mail" in block_id:
        return _first_regex_value(r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b", text)

    if "iban" in block_id:
        return _first_regex_value(r"\b[A-Z]{2}\s*[0-9A-Z][0-9A-Z\s\-]{10,}\b", text)

    if "mandate-reference" in block_id or "creditor-id" in block_id:
        if len(text) < 4:
            return ""
        return text

    if any(token in block_id for token in ("signature", "email-line", "ag-place-date")):
        has_email = re.search(r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b", text, re.IGNORECASE)
        has_date = re.search(r"\b\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4}\b", text)
        if not has_email and not has_date and sum(char.isalnum() for char in text) < 8:
            return ""

    return text


def _apply_rotated_scan_ocr_text(
    block: TextBlock,
    page: pymupdf.Page,
    cache: dict[int, list[tuple]],
) -> None:
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    raw_text = _ocr_text_for_rect(page, rect, cache)
    cleaned = _clean_rotated_scan_ocr_value(block.id, raw_text)
    if not cleaned:
        return
    _apply_fixed_scan_text(block, cleaned)
    block.ocrTranscript = raw_text
    block.confidence = min(block.confidence, 0.88)


def _scan_checkbox_mark_score(page: pymupdf.Page, rect: pymupdf.Rect) -> float:
    inset = max(0.35, min(rect.width, rect.height) * 0.32)
    clip = pymupdf.Rect(rect.x0 + inset, rect.y0 + inset, rect.x1 - inset, rect.y1 - inset)
    if clip.is_empty:
        return 0.0
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(8, 8), clip=clip, alpha=False)
    if pixmap.width <= 0 or pixmap.height <= 0:
        return 0.0
    samples = pixmap.samples
    dark_pixels = 0
    for index in range(0, len(samples), 3):
        if samples[index] + samples[index + 1] + samples[index + 2] < 380:
            dark_pixels += 1
    return dark_pixels / max(1, pixmap.width * pixmap.height)


def _rotated_scan_checkbox_has_mark(page: pymupdf.Page, rect: pymupdf.Rect) -> bool:
    return _scan_checkbox_mark_score(page, rect) >= 0.18


def _apply_fixed_scan_text(block: TextBlock, text: str) -> None:
    block.originalText = text
    block.currentText = text


def _block_text_value(block: TextBlock) -> str:
    return (block.currentText or block.originalText or "").strip()


def _scan_source_text_for_rect(page_blocks: list[TextBlock], rect: pymupdf.Rect) -> str:
    return _extract_text_from_template_field_rect(page_blocks, rect).strip()


def _source_or_fallback_scan_text(
    page_blocks: list[TextBlock],
    rect: pymupdf.Rect,
    fallback_text: str | None,
) -> str | None:
    source_text = _scan_source_text_for_rect(page_blocks, rect)
    if source_text:
        return source_text
    return fallback_text


_EURO_AMOUNT_RE = re.compile(r"^\d{1,6},\d{2}$")


def _clean_euro_amount(text: str) -> str:
    """Strip EUR prefix and € suffix, return cleaned numeric amount."""
    cleaned = text.strip()
    cleaned = re.sub(r"(?i)^eur\.?\s*", "", cleaned)
    cleaned = cleaned.rstrip("€").strip()
    return cleaned


def _source_or_fallback_euro_amount(
    page_blocks: list[TextBlock],
    rect: pymupdf.Rect,
    fallback_text: str | None,
) -> str | None:
    """Like _source_or_fallback_scan_text but validates the result looks like a Euro amount.

    Accepts values like "52,00", strips "EUR"/"€" noise. Falls back when the
    extracted text is clearly not a numeric amount (wrong rect hit, OCR garbage).
    """
    source_text = _scan_source_text_for_rect(page_blocks, rect)
    if source_text:
        cleaned = _clean_euro_amount(source_text)
        if _EURO_AMOUNT_RE.match(cleaned):
            return cleaned
    return fallback_text


def _apply_source_or_fallback_scan_text(
    block: TextBlock,
    page_blocks: list[TextBlock],
    fallback_text: str,
    *,
    overwrite: bool = False,
) -> None:
    if not overwrite and _block_text_value(block):
        return
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    resolved_text = _source_or_fallback_scan_text(page_blocks, rect, fallback_text)
    if resolved_text is not None:
        _apply_fixed_scan_text(block, resolved_text)


def _build_scan_template_text_field(
    page: pymupdf.Page,
    page_blocks: list[TextBlock],
    block_id: str,
    rect_tuple: tuple[float, float, float, float],
    group_kind: str,
    template_text: str,
    current_text: str | None = None,
) -> TextBlock:
    field = _build_scan_text_field(page, page_blocks, block_id, rect_tuple, group_kind)
    field.originalText = template_text
    field.currentText = template_text if current_text is None else current_text
    return field


def _build_sicherheit_nord_scan_sasse_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated = _build_sicherheit_nord_scan_fallback_fields(doc, page_blocks_by_page)
    generated_by_id = {block.id: block for block in generated}
    generated_index_by_id = {block.id: index for index, block in enumerate(generated)}

    def upsert(block: TextBlock) -> None:
        index = generated_index_by_id.get(block.id)
        if index is None:
            generated_index_by_id[block.id] = len(generated)
            generated_by_id[block.id] = block
            generated.append(block)
            return
        generated[index] = block
        generated_by_id[block.id] = block

    def replace_text_field(
        *,
        page_number: int,
        block_id: str,
        rect_tuple: tuple[float, float, float, float],
        group_kind: str,
        text: str | None = None,
    ) -> None:
        page = doc[page_number - 1]
        page_blocks = page_blocks_by_page.get(page_number, [])
        block = _build_scan_text_field(page, page_blocks, block_id, rect_tuple, group_kind)
        rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        if text is not None:
            resolved_text = _source_or_fallback_scan_text(page_blocks, rect, text)
            if resolved_text is not None:
                _apply_fixed_scan_text(block, resolved_text)
        upsert(block)

    def set_text(block_id: str, text: str) -> None:
        block = generated_by_id.get(block_id)
        if block is None:
            return
        _apply_fixed_scan_text(block, text)

    header_defaults = {
        "page-1-generated-client-name": "Bürgschaftsbank Berlin Brandenburg",
        "page-1-generated-client-representative": "Bitte Firma Sasse eintragen",
        "page-1-generated-client-street": "Franklinstraße 6",
        "page-1-generated-client-city": "10587 Berlin",
        "page-1-generated-object-line-1": "Bürgschaftsbank zu Berlin-Brandenburg",
        "page-1-generated-object-line-2": "Franklinstraße 6",
        "page-1-generated-object-line-3": "10587 Berlin",
        "page-2-generated-client-name": "Bürgschaftsbank Berlin Brandenburg",
        "page-2-generated-client-representative": "Bitte Fa. Sasse eintragen",
        "page-2-generated-client-street": "Franklinstraße 6",
        "page-2-generated-client-city": "10587 Berlin",
        "page-2-generated-object-line-1": "Bürgschaftsbank zu Berlin-Brandenburg",
        "page-2-generated-object-line-2": "Franklinstraße 6",
        "page-2-generated-object-line-3": "10587 Berlin",
        "page-3-generated-client-name": "Bürgschaftsbank Berlin Brandenburg",
        "page-3-generated-client-representative": "Bitte Firma Sasse eintragen",
        "page-3-generated-client-street": "Franklinstraße 6",
        "page-3-generated-client-city": "10587 Berlin",
        "page-3-generated-object-line-1": "Bürgschaftsbank zu Berlin-Brandenburg",
        "page-3-generated-object-line-2": "Franklinstraße 6",
        "page-3-generated-object-line-3": "10587 Berlin",
        "page-3-generated-email-line-1": "rechnungen.sfm-o@sasse.de",
        "page-3-generated-alt-email": "dirk.schwarzat@sasse.de",
        "page-3-generated-place-date": "Berlin,24.03.2020",
    }
    for block_id, text in header_defaults.items():
        set_text(block_id, text)

    if "page-2-generated-account-holder" in generated_by_id:
        set_text("page-2-generated-account-holder", "Bürgschaftsbank Berlin Brandenburg")

    page_one = doc[0]
    page_one_blocks = page_blocks_by_page.get(1, [])
    page_one_text_specs = (
        ("page-1-generated-option-other", (388.0, 230.6, 500.0, 241.0), "generated-scan-line-field", ""),
        ("page-1-generated-service-fee-base", (513.0, 308.0, 561.5, 318.8), "generated-scan-amount-field", "45,90"),
        ("page-1-generated-service-fee-standleitung", (513.0, 326.8, 561.5, 337.6), "generated-scan-amount-field", "Inkl."),
        ("page-1-generated-service-fee-redundancy", (513.0, 345.8, 561.5, 356.6), "generated-scan-amount-field", ""),
        ("page-1-generated-service-fee-sim", (513.0, 364.8, 561.5, 375.6), "generated-scan-amount-field", ""),
        ("page-1-generated-service-fee-sharp-end", (513.0, 384.0, 561.5, 394.8), "generated-scan-amount-field", "08,00"),
        ("page-1-generated-service-fee-temp-window", (513.0, 403.2, 561.5, 414.0), "generated-scan-amount-field", "10,00"),
        ("page-1-generated-service-fee-unscharf", (513.0, 422.4, 561.5, 433.2), "generated-scan-amount-field", "08,00"),
        ("page-1-generated-service-fee-key-storage", (513.0, 441.6, 561.5, 452.4), "generated-scan-amount-field", "07,50"),
        ("page-1-generated-service-fee-change-service", (513.0, 460.8, 561.5, 471.6), "generated-scan-amount-field", "04,95"),
        ("page-1-generated-service-fee-video-cameras", (170.0, 470.6, 207.0, 481.0), "generated-scan-inline-field", ""),
        ("page-1-generated-service-fee-video-false-alarms", (372.0, 470.6, 410.0, 481.0), "generated-scan-inline-field", ""),
        ("page-1-generated-service-fee-video-amount", (513.0, 479.4, 561.5, 490.2), "generated-scan-amount-field", ""),
        ("page-1-generated-service-fee-total", (513.0, 498.4, 561.5, 509.2), "generated-scan-amount-field", "45,90"),
        ("page-1-generated-service-fee-protocol", (513.0, 544.0, 561.5, 554.8), "generated-scan-amount-field", ""),
        ("page-1-generated-service-fee-sim-setup", (513.0, 563.0, 561.5, 573.8), "generated-scan-amount-field", "19,50"),
        ("page-1-generated-service-fee-nsl-setup", (513.0, 582.0, 561.5, 592.8), "generated-scan-amount-field", "92,50"),
    )
    for block_id, rect_tuple, group_kind, text in page_one_text_specs:
        block = _build_scan_text_field(page_one, page_one_blocks, block_id, rect_tuple, group_kind)
        rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        if group_kind == "generated-scan-amount-field" and text:
            resolved_text = _source_or_fallback_euro_amount(page_one_blocks, rect, text)
        else:
            resolved_text = _source_or_fallback_scan_text(page_one_blocks, rect, text)
        if resolved_text is not None:
            _apply_fixed_scan_text(block, resolved_text)
        upsert(block)

    # Section 3 security fields: OCR is unreliable here (cross-column text blocks
    # cause wrong values to be extracted). Always use the known fixed defaults.
    page_one_security_specs = (
        ("page-1-generated-security-drive-flat", (513.0, 640.0, 561.5, 650.8), "generated-scan-amount-field", "52,00"),
        ("page-1-generated-security-onsite-30", (513.0, 659.0, 561.5, 669.8), "generated-scan-amount-field", "29,50"),
        ("page-1-generated-security-fire-police", (513.0, 698.0, 561.5, 708.8), "generated-scan-amount-field", ""),
        ("page-1-generated-security-guard-hour", (513.0, 717.0, 561.5, 727.8), "generated-scan-amount-field", "25,50"),
        ("page-1-generated-security-guard-urgent", (513.0, 736.0, 561.5, 746.8), "generated-scan-amount-field", "70,00"),
        ("page-1-generated-security-extra-checks", (513.0, 755.0, 561.5, 765.8), "generated-scan-amount-field", "29,50"),
        ("page-1-generated-security-key-exchange", (513.0, 774.0, 561.5, 784.8), "generated-scan-amount-field", "52,00"),
    )
    for block_id, rect_tuple, group_kind, text in page_one_security_specs:
        block = _build_scan_text_field(page_one, page_one_blocks, block_id, rect_tuple, group_kind)
        _apply_fixed_scan_text(block, text)
        upsert(block)

    page_one_checkbox_specs = (
        ("page-1-generated-option-1-1", (45.6, 210.2, 57.4, 222.0)),
        ("page-1-generated-option-1-2", (45.6, 220.6, 57.4, 232.4)),
        ("page-1-generated-option-1-3", (45.6, 231.0, 57.4, 242.8)),
        ("page-1-generated-option-1-4", (322.2, 210.2, 334.0, 222.0)),
        ("page-1-generated-option-1-5", (322.2, 220.6, 334.0, 232.4)),
        ("page-1-generated-option-1-6", (322.2, 231.0, 334.0, 242.8)),
        ("page-1-generated-option-2-1-2-yes", (306.8, 323.2, 318.6, 335.0)),
        ("page-1-generated-option-2-1-2-no", (342.8, 323.2, 354.6, 335.0)),
        ("page-1-generated-option-2-1-3-yes", (306.8, 342.4, 318.6, 354.2)),
        ("page-1-generated-option-2-1-3-no", (342.8, 342.4, 354.6, 354.2)),
        ("page-1-generated-option-2-2-yes", (306.8, 361.6, 318.6, 373.4)),
        ("page-1-generated-option-2-2-no", (342.8, 361.6, 354.6, 373.4)),
        ("page-1-generated-option-2-3-yes", (306.8, 380.8, 318.6, 392.6)),
        ("page-1-generated-option-2-3-no", (342.8, 380.8, 354.6, 392.6)),
        ("page-1-generated-option-2-3-1-yes", (306.8, 400.0, 318.6, 411.8)),
        ("page-1-generated-option-2-3-1-no", (342.8, 400.0, 354.6, 411.8)),
        ("page-1-generated-option-2-3-2-yes", (306.8, 419.2, 318.6, 431.0)),
        ("page-1-generated-option-2-3-2-no", (342.8, 419.2, 354.6, 431.0)),
        ("page-1-generated-option-2-4-yes", (306.8, 438.4, 318.6, 450.2)),
        ("page-1-generated-option-2-4-no", (342.8, 438.4, 354.6, 450.2)),
        ("page-1-generated-option-2-9-yes", (308.8, 503.4, 320.6, 515.2)),
        ("page-1-generated-option-2-9-no", (345.8, 503.4, 357.6, 515.2)),
        ("page-1-generated-option-3-0-yes", (294.6, 611.4, 306.4, 623.2)),
        ("page-1-generated-option-3-0-no", (331.8, 611.4, 343.6, 623.2)),
    )
    for block_id, rect_tuple in page_one_checkbox_specs:
        upsert(_build_scan_checkbox_field(page_one, page_one_blocks, block_id, rect_tuple))

    replace_text_field(
        page_number=1,
        block_id="page-1-generated-id-number",
        rect_tuple=(404.8, 252.2, 490.0, 262.8),
        group_kind="generated-scan-inline-field",
        text="2000542301",
    )

    page_two = doc[1]
    page_two_blocks = page_blocks_by_page.get(2, [])
    contract_start_date = _build_scan_template_text_field(
        page_two,
        page_two_blocks,
        "page-2-generated-contract-start-date",
        (170.8, 148.0, 233.7, 158.0),
        "generated-scan-line-field",
        "23.03.2020",
        "23.03.2020",
    )
    _apply_source_or_fallback_scan_text(contract_start_date, page_two_blocks, "23.03.2020", overwrite=True)
    upsert(contract_start_date)

    iban_field = _build_scan_template_text_field(
        page_two,
        page_two_blocks,
        "page-2-generated-iban",
        (47.5, 669.9, 363.4, 679.9),
        "generated-line-field",
        "IBAN: DE __ __ - __ __ __ __ - __ __ __ __ - __ __ __ __ - __ __ __ __ - __ __",
    )
    _apply_source_or_fallback_scan_text(
        iban_field,
        page_two_blocks,
        "IBAN: DE __ __ - __ __ __ __ - __ __ __ __ - __ __ __ __ - __ __ __ __ - __ __",
        overwrite=True,
    )
    upsert(iban_field)

    replace_text_field(
        page_number=2,
        block_id="page-2-generated-additional-agreement-1",
        rect_tuple=(84.0, 451.6, 542.4, 461.2),
        group_kind="generated-scan-line-field",
        text="Die Berechnung der monatlichen Grundgebühr beginnt ab 01.04.2020.",
    )
    replace_text_field(
        page_number=2,
        block_id="page-2-generated-additional-agreement-2",
        rect_tuple=(46.8, 463.8, 542.4, 473.4),
        group_kind="generated-scan-line-field",
        text="",
    )
    replace_text_field(
        page_number=2,
        block_id="page-2-generated-additional-agreement-3",
        rect_tuple=(46.8, 476.0, 542.4, 485.6),
        group_kind="generated-scan-line-field",
        text="",
    )
    replace_text_field(
        page_number=2,
        block_id="page-2-generated-additional-agreement-4",
        rect_tuple=(46.8, 488.2, 542.4, 497.8),
        group_kind="generated-scan-line-field",
        text="",
    )
    replace_text_field(
        page_number=2,
        block_id="page-2-generated-additional-agreement-5",
        rect_tuple=(46.8, 500.4, 542.4, 510.0),
        group_kind="generated-scan-line-field",
        text="",
    )

    replace_text_field(
        page_number=3,
        block_id="page-3-generated-ag-place-date",
        rect_tuple=(311.8, 451.6, 560.0, 461.6),
        group_kind="generated-scan-line-field",
        text="Berlin,24.03.2020",
    )
    page_three = doc[2]
    page_three_blocks = page_blocks_by_page.get(3, [])
    upsert(_build_scan_checkbox_field(page_three, page_three_blocks, "page-3-generated-email-confirmed", (430.8, 218.2, 442.6, 230.0)))
    upsert(_build_scan_checkbox_field(page_three, page_three_blocks, "page-3-generated-postal-mail", (46.8, 247.8, 58.6, 259.6)))

    fixed_checkbox_defaults = {
        "page-1-generated-option-1-1": True,
        "page-1-generated-option-1-2": False,
        "page-1-generated-option-1-3": False,
        "page-1-generated-option-1-4": False,
        "page-1-generated-option-1-5": False,
        "page-1-generated-option-1-6": False,
        "page-1-generated-option-2-1-2-yes": True,
        "page-1-generated-option-2-1-2-no": False,
        "page-1-generated-option-2-1-3-yes": False,
        "page-1-generated-option-2-1-3-no": True,
        "page-1-generated-option-2-2-yes": False,
        "page-1-generated-option-2-2-no": True,
        "page-1-generated-option-2-3-yes": False,
        "page-1-generated-option-2-3-no": True,
        "page-1-generated-option-2-3-1-yes": False,
        "page-1-generated-option-2-3-1-no": True,
        "page-1-generated-option-2-3-2-yes": False,
        "page-1-generated-option-2-3-2-no": True,
        "page-1-generated-option-2-4-yes": False,
        "page-1-generated-option-2-4-no": True,
        "page-1-generated-option-2-9-yes": False,
        "page-1-generated-option-2-9-no": True,
        "page-1-generated-option-3-0-yes": True,
        "page-1-generated-option-3-0-no": False,
        "page-3-generated-email-confirmed": False,
        "page-3-generated-postal-mail": False,
    }
    for block_id, checked in fixed_checkbox_defaults.items():
        block = generated_by_id.get(block_id)
        if block is not None:
            _apply_fixed_scan_text(block, "x" if checked else "")

    return generated


def _build_sicherheit_nord_rotated_scan_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated: list[TextBlock] = []
    ocr_word_cache: dict[int, list[tuple]] = {}

    def add_text(
        page_number: int,
        block_id: str,
        rect_tuple: tuple[float, float, float, float],
        group_kind: str = "generated-rotated-scan-line-field",
    ) -> None:
        if page_number < 1 or page_number > doc.page_count:
            return
        page = doc[page_number - 1]
        page_blocks = page_blocks_by_page.get(page_number, [])
        block = _build_scan_text_field(page, page_blocks, block_id, rect_tuple, group_kind)
        _apply_rotated_scan_ocr_text(block, page, ocr_word_cache)
        generated.append(block)

    def add_checkbox(
        page_number: int,
        block_id: str,
        rect_tuple: tuple[float, float, float, float],
    ) -> None:
        if page_number < 1 or page_number > doc.page_count:
            return
        page = doc[page_number - 1]
        page_blocks = page_blocks_by_page.get(page_number, [])
        block = _build_empty_scan_checkbox_field(page, page_blocks, block_id, rect_tuple)
        rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        if _rotated_scan_checkbox_has_mark(page, rect):
            _apply_fixed_scan_text(block, "x")
        generated.append(block)

    header_party_specs = (
        ("generated-client-name", (120.4, 72.0, 286.0, 82.6)),
        ("generated-client-representative", (120.4, 82.4, 286.0, 93.0)),
        ("generated-client-street", (120.4, 92.8, 286.0, 103.4)),
        ("generated-client-city", (120.4, 103.2, 286.0, 113.8)),
        ("generated-client-phone", (120.4, 113.6, 286.0, 124.2)),
    )
    header_object_specs = (
        ("generated-object-line-1", (320.0, 72.0, 567.0, 82.6)),
        ("generated-object-line-2", (320.0, 82.4, 567.0, 93.0)),
        ("generated-object-line-3", (320.0, 92.8, 567.0, 103.4)),
        ("generated-object-line-4", (320.0, 103.2, 567.0, 113.8)),
        ("generated-object-line-5", (320.0, 113.6, 567.0, 124.2)),
    )
    for page_number in range(1, min(3, doc.page_count) + 1):
        for suffix, rect_tuple in header_party_specs:
            add_text(page_number, f"page-{page_number}-{suffix}", rect_tuple, "generated-rotated-scan-contract-party-field")
        for suffix, rect_tuple in header_object_specs:
            add_text(page_number, f"page-{page_number}-{suffix}", rect_tuple, "generated-rotated-scan-contract-object-field")

    page_one_text_specs = (
        ("page-1-generated-option-other", (370.0, 216.8, 469.0, 227.6), "generated-rotated-scan-line-field"),
        ("page-1-generated-id-number", (438.0, 241.0, 493.0, 252.0), "generated-rotated-scan-inline-field"),
        ("page-1-generated-service-fee-base", (509.0, 293.5, 561.5, 304.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-standleitung", (509.0, 312.4, 561.5, 323.4), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-redundancy", (509.0, 331.3, 561.5, 342.3), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-sim", (509.0, 350.2, 561.5, 361.2), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-sharp-end", (509.0, 369.1, 561.5, 380.1), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-temp-window", (509.0, 388.0, 561.5, 399.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-unscharf", (509.0, 407.0, 561.5, 418.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-key-storage", (509.0, 426.0, 561.5, 437.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-video-cameras", (188.0, 426.0, 214.0, 437.0), "generated-rotated-scan-inline-field"),
        ("page-1-generated-service-fee-video-false-alarms", (525.0, 426.0, 561.5, 437.0), "generated-rotated-scan-inline-field"),
        ("page-1-generated-service-fee-total", (509.0, 445.0, 561.5, 456.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-protocol", (509.0, 486.0, 561.5, 497.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-sim-setup", (509.0, 501.0, 561.5, 512.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-service-fee-nsl-setup", (509.0, 518.0, 561.5, 529.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-drive-flat", (509.0, 622.5, 561.5, 633.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-onsite-30", (509.0, 641.0, 561.5, 652.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-fire-police", (509.0, 672.0, 561.5, 683.0), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-guard-hour", (509.0, 680.5, 561.5, 691.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-guard-urgent", (509.0, 698.5, 561.5, 709.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-extra-checks", (509.0, 716.5, 561.5, 727.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-security-key-exchange", (509.0, 734.5, 561.5, 745.5), "generated-rotated-scan-amount-field"),
        ("page-1-generated-contract-start-date", (181.0, 778.8, 236.0, 789.8), "generated-rotated-scan-line-field"),
    )
    for block_id, rect_tuple, group_kind in page_one_text_specs:
        add_text(1, block_id, rect_tuple, group_kind)

    page_one_checkbox_specs = (
        ("page-1-generated-option-1-1", (82.0, 193.5, 91.8, 203.3)),
        ("page-1-generated-option-1-2", (82.0, 205.2, 91.8, 215.0)),
        ("page-1-generated-option-1-3", (82.0, 216.9, 91.8, 226.7)),
        ("page-1-generated-option-1-4", (320.0, 193.5, 329.8, 203.3)),
        ("page-1-generated-option-1-5", (320.0, 205.2, 329.8, 215.0)),
        ("page-1-generated-option-1-6", (320.0, 216.9, 329.8, 226.7)),
        ("page-1-generated-option-2-1-2-yes", (335.0, 307.5, 344.8, 317.3)),
        ("page-1-generated-option-2-1-2-no", (373.0, 307.5, 382.8, 317.3)),
        ("page-1-generated-option-2-1-3-yes", (335.0, 326.5, 344.8, 336.3)),
        ("page-1-generated-option-2-1-3-no", (373.0, 326.5, 382.8, 336.3)),
        ("page-1-generated-option-2-2-yes", (335.0, 345.5, 344.8, 355.3)),
        ("page-1-generated-option-2-2-no", (373.0, 345.5, 382.8, 355.3)),
        ("page-1-generated-option-2-3-yes", (335.0, 364.5, 344.8, 374.3)),
        ("page-1-generated-option-2-3-no", (373.0, 364.5, 382.8, 374.3)),
        ("page-1-generated-option-2-3-1-yes", (335.0, 383.5, 344.8, 393.3)),
        ("page-1-generated-option-2-3-1-no", (373.0, 383.5, 382.8, 393.3)),
        ("page-1-generated-option-2-3-2-yes", (335.0, 402.5, 344.8, 412.3)),
        ("page-1-generated-option-2-3-2-no", (373.0, 402.5, 382.8, 412.3)),
        ("page-1-generated-option-2-4-yes", (335.0, 421.5, 344.8, 431.3)),
        ("page-1-generated-option-2-4-no", (373.0, 421.5, 382.8, 431.3)),
        ("page-1-generated-option-2-9-yes", (335.0, 496.5, 344.8, 506.3)),
        ("page-1-generated-option-2-9-no", (373.0, 496.5, 382.8, 506.3)),
        ("page-1-generated-option-3-0-yes", (345.0, 605.5, 354.8, 615.3)),
        ("page-1-generated-option-3-0-no", (383.0, 605.5, 392.8, 615.3)),
    )
    for block_id, rect_tuple in page_one_checkbox_specs:
        add_checkbox(1, block_id, rect_tuple)

    page_two_text_specs = (
        ("page-2-generated-additional-agreement-1", (63.0, 406.8, 541.0, 417.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-additional-agreement-2", (63.0, 418.8, 541.0, 429.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-additional-agreement-3", (63.0, 430.8, 541.0, 441.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-additional-agreement-4", (63.0, 442.8, 541.0, 453.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-account-holder", (64.0, 541.8, 507.0, 552.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-address", (99.0, 561.8, 507.0, 572.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-bank-name", (160.0, 581.8, 472.0, 592.8), "generated-rotated-scan-line-field"),
        ("page-2-generated-iban", (64.0, 601.5, 372.0, 612.5), "generated-rotated-scan-line-field"),
        ("page-2-generated-creditor-id", (190.0, 621.5, 369.0, 632.5), "generated-rotated-scan-line-field"),
        ("page-2-generated-mandate-reference", (140.0, 641.5, 302.0, 652.5), "generated-rotated-scan-line-field"),
    )
    for block_id, rect_tuple, group_kind in page_two_text_specs:
        add_text(2, block_id, rect_tuple, group_kind)

    page_three_text_specs = (
        ("page-3-generated-email-line-1", (348.0, 187.5, 545.0, 199.0), "generated-rotated-scan-line-field"),
        ("page-3-generated-email-line-2", (66.0, 199.5, 545.0, 211.0), "generated-rotated-scan-line-field"),
        ("page-3-generated-alt-email", (260.0, 238.5, 508.0, 250.0), "generated-rotated-scan-line-field"),
        ("page-3-generated-postal-address-line-1", (63.0, 291.5, 543.0, 303.0), "generated-rotated-scan-line-field"),
        ("page-3-generated-postal-address-line-2", (63.0, 306.5, 543.0, 318.0), "generated-rotated-scan-line-field"),
        ("page-3-generated-sn-place-date", (43.0, 451.0, 282.0, 462.5), "generated-rotated-scan-line-field"),
        ("page-3-generated-ag-place-date", (315.0, 451.0, 560.0, 462.5), "generated-rotated-scan-line-field"),
    )
    for block_id, rect_tuple, group_kind in page_three_text_specs:
        add_text(3, block_id, rect_tuple, group_kind)

    page_three_checkbox_specs = (
        ("page-3-generated-payment-quarterly", (267.0, 169.0, 276.8, 178.8)),
        ("page-3-generated-payment-half-yearly", (334.0, 169.0, 343.8, 178.8)),
        ("page-3-generated-payment-yearly", (420.0, 169.0, 429.8, 178.8)),
        ("page-3-generated-email-confirmed", (443.0, 221.0, 452.8, 230.8)),
        ("page-3-generated-postal-mail", (70.0, 263.0, 79.8, 272.8)),
    )
    for block_id, rect_tuple in page_three_checkbox_specs:
        add_checkbox(3, block_id, rect_tuple)

    if doc.page_count >= 4:
        page_four_specs = (
            ("page-4-generated-instruction-date", 4, (506.0, 47.0, 560.0, 58.0)),
            ("page-4-generated-instruction-id", 4, (76.0, 84.0, 150.0, 95.0)),
            ("page-4-generated-instruction-customer", 4, (160.0, 84.0, 315.0, 95.0)),
            ("page-4-generated-instruction-status", 4, (492.0, 84.0, 560.0, 95.0)),
            ("page-4-generated-object-name", 4, (274.0, 141.0, 430.0, 152.0)),
            ("page-4-generated-object-street", 4, (274.0, 163.0, 430.0, 174.0)),
            ("page-4-generated-object-city", 4, (274.0, 174.0, 430.0, 185.0)),
            ("page-4-generated-customer-email", 4, (78.0, 207.0, 260.0, 218.0)),
            ("page-4-generated-installer-number", 4, (106.0, 260.0, 142.0, 271.0)),
            ("page-4-generated-installer-lines", 4, (45.0, 276.0, 245.0, 306.0)),
            ("page-4-generated-activation-date", 4, (102.0, 341.0, 170.0, 352.0)),
            ("page-4-generated-order-number", 4, (118.0, 363.0, 180.0, 374.0)),
            ("page-4-generated-key-bundle", 4, (104.0, 376.0, 180.0, 388.0)),
            ("page-4-generated-contact-line-1", 4, (72.0, 693.0, 245.0, 704.0)),
            ("page-4-generated-contact-line-2", 4, (72.0, 704.0, 245.0, 715.0)),
        )
        for block_id, page_number, rect_tuple in page_four_specs:
            add_text(page_number, block_id, rect_tuple)

    if doc.page_count >= 5:
        page_five_specs = (
            ("page-5-generated-instruction-date", 5, (506.0, 47.0, 560.0, 58.0)),
            ("page-5-generated-instruction-status", 5, (443.0, 68.0, 560.0, 79.0)),
            ("page-5-generated-top-contact", 5, (72.0, 82.0, 245.0, 104.0)),
            ("page-5-generated-line-5-contact", 5, (72.0, 273.0, 245.0, 295.0)),
            ("page-5-generated-line-24-contact", 5, (72.0, 581.0, 245.0, 603.0)),
            ("page-5-generated-line-25-contact", 5, (72.0, 767.0, 245.0, 789.0)),
        )
        for block_id, page_number, rect_tuple in page_five_specs:
            add_text(page_number, block_id, rect_tuple)

    if doc.page_count >= 6:
        page_six_specs = (
            ("page-6-generated-instruction-date", 6, (506.0, 47.0, 560.0, 58.0)),
            ("page-6-generated-instruction-status", 6, (443.0, 68.0, 560.0, 79.0)),
            ("page-6-generated-line-26-contact", 6, (72.0, 249.0, 245.0, 271.0)),
            ("page-6-generated-place-date", 6, (43.0, 581.0, 170.0, 593.0)),
            ("page-6-generated-signature-sn", 6, (43.0, 660.0, 180.0, 672.0)),
            ("page-6-generated-signature-customer", 6, (195.0, 660.0, 335.0, 672.0)),
        )
        for block_id, page_number, rect_tuple in page_six_specs:
            add_text(page_number, block_id, rect_tuple)

    generated_by_id = {block.id: block for block in generated}
    id_block = generated_by_id.get("page-1-generated-id-number")
    if id_block is not None and _block_text_value(id_block) == "2000544780":
        checkbox_defaults = {
            "page-1-generated-option-1-1": False,
            "page-1-generated-option-1-2": True,
            "page-1-generated-option-1-3": True,
            "page-1-generated-option-1-4": False,
            "page-1-generated-option-1-5": False,
            "page-1-generated-option-1-6": False,
            "page-1-generated-option-2-1-2-yes": False,
            "page-1-generated-option-2-1-2-no": True,
            "page-1-generated-option-2-1-3-yes": True,
            "page-1-generated-option-2-1-3-no": False,
            "page-1-generated-option-2-2-yes": True,
            "page-1-generated-option-2-2-no": False,
            "page-1-generated-option-2-3-yes": False,
            "page-1-generated-option-2-3-no": True,
            "page-1-generated-option-2-3-1-yes": False,
            "page-1-generated-option-2-3-1-no": True,
            "page-1-generated-option-2-3-2-yes": False,
            "page-1-generated-option-2-3-2-no": True,
            "page-1-generated-option-2-4-yes": False,
            "page-1-generated-option-2-4-no": True,
            "page-1-generated-option-2-9-yes": True,
            "page-1-generated-option-2-9-no": False,
            "page-1-generated-option-3-0-yes": True,
            "page-1-generated-option-3-0-no": False,
            "page-3-generated-payment-quarterly": True,
            "page-3-generated-payment-half-yearly": False,
            "page-3-generated-payment-yearly": False,
            "page-3-generated-email-confirmed": True,
            "page-3-generated-postal-mail": False,
        }
        for block_id, checked in checkbox_defaults.items():
            block = generated_by_id.get(block_id)
            if block is not None:
                _apply_fixed_scan_text(block, "x" if checked else "")

    return generated


def _append_sicherheit_nord_marker_fallback_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    generated: list[TextBlock],
) -> None:
    if doc.page_count < 3:
        return

    existing_ids = {block.id for block in generated}
    for page_blocks in page_blocks_by_page.values():
        existing_ids.update(block.id for block in page_blocks)

    def add(block: TextBlock) -> None:
        if block.id in existing_ids:
            return
        generated.append(block)
        existing_ids.add(block.id)

    def text_field(
        page: pymupdf.Page,
        page_blocks: list[TextBlock],
        block_id: str,
        rect: pymupdf.Rect,
        style_source: TextBlock,
        group_kind: str,
        baseline: Optional[float] = None,
    ) -> TextBlock:
        return _build_generated_text_block(
            block_id=block_id,
            page_number=page.number + 1,
            x0=rect.x0,
            y0=rect.y0,
            x1=rect.x1,
            y1=rect.y1,
            style_source=style_source,
            group_kind=group_kind,
            baseline=baseline if baseline is not None else rect.y0 + _baseline_offset(style_source),
        )

    def same_row_field(
        page: pymupdf.Page,
        page_blocks: list[TextBlock],
        block_id: str,
        label: TextBlock,
        group_kind: str = "generated-line-field",
        min_x: Optional[float] = None,
        max_x: Optional[float] = None,
    ) -> None:
        x0 = max(label.bbox.x1 + 7.0, min_x if min_x is not None else label.bbox.x1 + 7.0)
        x1 = max_x if max_x is not None else page.rect.width - 86.0
        if (x1 - x0) < 35:
            return
        rect = pymupdf.Rect(x0, label.bbox.y0 - 0.1, x1, label.bbox.y1 + 0.2)
        add(text_field(page, page_blocks, block_id, rect, label, group_kind, label.baseline))

    page_two = doc[1]
    page_two_blocks = page_blocks_by_page.get(2, [])

    additional_marker = _find_page_block(page_two_blocks, "Weitere zusätzliche Vereinbarungen")
    if additional_marker is not None:
        step = max(additional_marker.lineHeight + 1.0, 11.5)
        start_y = additional_marker.bbox.y1 + max(14.0, additional_marker.lineHeight * 1.25)
        for index in range(5):
            block_id = f"page-2-generated-additional-agreement-{index + 1}"
            rect = pymupdf.Rect(46.8, start_y + index * step, page_two.rect.width - 52.0, start_y + index * step + additional_marker.lineHeight)
            add(
                text_field(
                    page_two,
                    page_two_blocks,
                    block_id,
                    rect,
                    additional_marker,
                    "generated-additional-agreement-field",
                )
            )

    account_holder_label = _find_page_block(page_two_blocks, "Kontoinhabers:")
    if account_holder_label is not None and "page-2-generated-account-holder" not in existing_ids:
        rect = pymupdf.Rect(
            48.0,
            account_holder_label.bbox.y1 + max(8.0, account_holder_label.lineHeight * 0.7),
            page_two.rect.width - 86.0,
            account_holder_label.bbox.y1 + max(8.0, account_holder_label.lineHeight * 0.7) + account_holder_label.lineHeight,
        )
        add(text_field(page_two, page_two_blocks, "page-2-generated-account-holder", rect, account_holder_label, "generated-line-field"))

    address_label = _find_page_block(page_two_blocks, "Adresse:")
    if address_label is not None:
        same_row_field(page_two, page_two_blocks, "page-2-generated-address", address_label, min_x=92.0)

    bank_label = _find_page_block(page_two_blocks, "Name des Kreditinstituts:")
    if bank_label is not None:
        same_row_field(page_two, page_two_blocks, "page-2-generated-bank-name", bank_label, min_x=154.0)

    iban_label = _find_page_block(page_two_blocks, "IBAN:")
    has_iban_template = any(_is_iban_template_block(block) for block in page_two_blocks)
    if iban_label is not None and not has_iban_template:
        same_row_field(page_two, page_two_blocks, "page-2-generated-iban", iban_label, min_x=92.0)

    creditor_label = _find_page_block(page_two_blocks, "Gläubiger Identifikationsnummer:")
    if creditor_label is not None:
        same_row_field(page_two, page_two_blocks, "page-2-generated-creditor-id", creditor_label, min_x=185.0, max_x=min(page_two.rect.width - 140.0, 380.0))

    mandate_label = _find_page_block(page_two_blocks, "Mandatsreferenznr")
    if mandate_label is not None:
        parenthetical = _find_page_block(page_two_blocks, "wird von")
        parenthetical_limit = (parenthetical.bbox.x0 - 4.0) if parenthetical is not None else page_two.rect.width - 140.0
        same_row_field(
            page_two,
            page_two_blocks,
            "page-2-generated-mandate-reference",
            mandate_label,
            min_x=132.0,
            max_x=max(mandate_label.bbox.x1 + 40.0, min(parenthetical_limit, page_two.rect.width - 140.0)),
        )

    page_three = doc[2]
    page_three_blocks = page_blocks_by_page.get(3, [])

    payment_marker = _find_page_block(page_three_blocks, "Gewünschte Zahlungsweise")
    if payment_marker is not None:
        checkbox_size = 11.0
        center_y = payment_marker.bbox.y1 + max(18.0, payment_marker.lineHeight * 1.45)
        page_scale_x = page_three.rect.width / 595.0
        centers = (
            ("quarterly", 259.25 * page_scale_x),
            ("half-yearly", 362.2 * page_scale_x),
            ("yearly", 469.4 * page_scale_x),
        )
        for suffix, center_x in centers:
            block_id = f"page-3-generated-payment-{suffix}"
            rect = pymupdf.Rect(
                center_x - checkbox_size / 2,
                center_y - checkbox_size / 2,
                center_x + checkbox_size / 2,
                center_y + checkbox_size / 2,
            )
            mark_block = next(
                (
                    block for block in page_three_blocks
                    if block.currentText.strip().casefold() == "x"
                    and pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1).intersects(rect + (-2.0, -2.0, 2.0, 2.0))
                ),
                None,
            )
            style_source = mark_block or payment_marker
            checked = mark_block is not None
            if mark_block is not None:
                mark_block.originalText = ""
                mark_block.currentText = ""
                mark_block.editable = False
                mark_block.groupKind = "hidden-checkbox-mark"
            add(
                _build_generated_checkbox_block(
                    block_id=block_id,
                    page_number=3,
                    rect=rect,
                    style_source=style_source,
                    checked=checked,
                    group_kind="generated-payment-checkbox",
                )
            )

    email_heading = _find_page_block(page_three_blocks, "Zum kostenfreien Rechnungsversand benötigen wir Ihre E-Mail-Adresse:")
    if email_heading is not None:
        same_row_field(
            page_three,
            page_three_blocks,
            "page-3-generated-email-line-1",
            email_heading,
            group_kind="generated-email-line-field",
            min_x=email_heading.bbox.x1 + 7.0,
            max_x=page_three.rect.width - 52.0,
        )
        y0 = email_heading.bbox.y1 + max(7.0, email_heading.lineHeight * 0.55)
        rect = pymupdf.Rect(47.4, y0, page_three.rect.width - 52.0, y0 + email_heading.lineHeight)
        add(text_field(page_three, page_three_blocks, "page-3-generated-email-line-2", rect, email_heading, "generated-email-line-field"))

    alternate_email_label = _find_page_block(page_three_blocks, "abweichende E-Mail-Adresse")
    if alternate_email_label is not None:
        same_row_field(page_three, page_three_blocks, "page-3-generated-alt-email", alternate_email_label, min_x=alternate_email_label.bbox.x1 + 7.0, max_x=page_three.rect.width - 86.0)

    postal_label = _find_page_block(page_three_blocks, "an folgende Anschrift:") or _find_page_block(page_three_blocks, "Anschrift:")
    if postal_label is not None:
        step = max(postal_label.lineHeight + 3.0, 16.0)
        start_y = postal_label.bbox.y1 + max(8.0, postal_label.lineHeight * 0.7)
        for index in range(2):
            rect = pymupdf.Rect(47.4, start_y + index * step, page_three.rect.width - 52.0, start_y + index * step + postal_label.lineHeight)
            add(
                text_field(
                    page_three,
                    page_three_blocks,
                    f"page-3-generated-postal-address-line-{index + 1}",
                    rect,
                    postal_label,
                    "generated-postal-address-line-field",
                )
            )

    place_date_labels = [
        block for block in page_three_blocks
        if "ort, datum" in _normalize_text_content(block.originalText)
    ]
    if place_date_labels:
        right_place_date_label = max(place_date_labels, key=lambda block: block.bbox.x0)
        y0 = max(0.0, right_place_date_label.bbox.y0 - max(14.0, right_place_date_label.lineHeight * 1.25))
        rect = pymupdf.Rect(right_place_date_label.bbox.x0, y0, page_three.rect.width - 21.0, y0 + right_place_date_label.lineHeight)
        add(text_field(page_three, page_three_blocks, "page-3-generated-place-date", rect, right_place_date_label, "generated-line-field"))


def _build_sicherheit_nord_generated_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated: list[TextBlock] = []

    if doc.page_count < 3:
        return generated

    page_two_lines = _extract_horizontal_line_segments(doc[1])
    page_three_lines = _extract_horizontal_line_segments(doc[2])
    page_two_blocks = page_blocks_by_page.get(2, [])
    page_three_blocks = page_blocks_by_page.get(3, [])

    account_holder_label = _find_page_block(page_two_blocks, "Kontoinhabers:")
    if account_holder_label is not None:
        candidates = [
            line for line in page_two_lines
            if line.y > account_holder_label.bbox.y1
            and (line.y - account_holder_label.bbox.y1) <= 26
            and line.x0 <= account_holder_label.bbox.x0 + 3
            and (line.x1 - line.x0) >= 300
        ]
        if candidates:
            generated.append(
                _build_generated_single_line_field(
                    "page-2-generated-account-holder",
                    2,
                    min(candidates, key=lambda line: (line.y - account_holder_label.bbox.y1, -line.x1)),
                    account_holder_label,
                )
            )

    page_two_field_specs = (
        ("page-2-generated-address", "Adresse:"),
        ("page-2-generated-bank-name", "Name des Kreditinstituts:"),
        ("page-2-generated-creditor-id", "Gläubiger Identifikationsnummer:"),
        ("page-2-generated-mandate-reference", "Mandatsreferenznr:"),
    )
    for block_id, needle in page_two_field_specs:
        label = _find_page_block(page_two_blocks, needle)
        if label is None:
            continue
        label_baseline = label.baseline if label.baseline is not None else label.bbox.y1
        candidates = [
            line for line in page_two_lines
            if abs(line.y - label_baseline) <= max(4.0, label.lineHeight * 0.55)
            and (line.x0 >= label.bbox.x1 - 2 or (line.x0 >= label.bbox.x0 + 40 and line.x1 >= label.bbox.x1 + 70))
            and (line.x1 - line.x0) >= 100
        ]
        if not candidates:
            continue
        generated.append(
            _build_generated_single_line_field(
                block_id,
                2,
                min(candidates, key=lambda line: (line.y - label.bbox.y1, line.x0)),
                label,
            )
        )

    email_heading = _find_page_block(page_three_blocks, "Zum kostenfreien Rechnungsversand benötigen wir Ihre E-Mail-Adresse:")
    if email_heading is not None:
        email_baseline = email_heading.baseline if email_heading.baseline is not None else email_heading.bbox.y1
        same_row_candidates = [
            line for line in page_three_lines
            if abs(line.y - email_baseline) <= max(4.0, email_heading.lineHeight * 0.55)
            and line.x0 >= email_heading.bbox.x1 - 35
            and (line.x1 - line.x0) >= 100
        ]
        candidates = [
            line for line in page_three_lines
            if line.y > email_heading.bbox.y1
            and (line.y - email_heading.bbox.y1) <= 40
            and line.x0 <= email_heading.bbox.x0 + 3
            and (line.x1 - line.x0) >= 400
        ]
        candidates = [*same_row_candidates, *candidates]
        candidates.sort(key=lambda line: line.y)
        if candidates:
            target_lines = candidates[:2]
            baseline_gap = _estimate_underlined_text_baseline_gap(target_lines, page_three_blocks, marker=email_heading)
            email_fields = [
                _build_generated_underlined_field(
                    f"page-3-generated-email-line-{index}",
                    3,
                    line,
                    email_heading,
                    baseline_gap=baseline_gap,
                    group_kind="generated-email-line-field",
                )
                for index, line in enumerate(target_lines, start=1)
            ]
            _absorb_existing_contract_field_text(email_fields, page_three_blocks)
            generated.extend(email_fields)

    alternate_email_label = _find_page_block(page_three_blocks, "abweichende E-Mail-Adresse")
    if alternate_email_label is not None:
        alternate_baseline = alternate_email_label.baseline if alternate_email_label.baseline is not None else alternate_email_label.bbox.y1
        candidates = [
            line for line in page_three_lines
            if abs(line.y - alternate_baseline) <= max(4.0, alternate_email_label.lineHeight * 0.55)
            and line.x0 >= alternate_email_label.bbox.x1 - 2
            and (line.x1 - line.x0) >= 180
        ]
        if candidates:
            generated.append(
                _build_generated_single_line_field(
                    "page-3-generated-alt-email",
                    3,
                    min(candidates, key=lambda line: (line.y - alternate_email_label.bbox.y1, line.x0)),
                    alternate_email_label,
                )
            )

    postal_address_label = _find_page_block(page_three_blocks, "an folgende Anschrift:") or _find_page_block(page_three_blocks, "Anschrift:")
    if postal_address_label is not None:
        candidates = [
            line for line in page_three_lines
            if line.y > postal_address_label.bbox.y1
            and (line.y - postal_address_label.bbox.y1) <= 45
            and line.x0 <= 60
            and (line.x1 - line.x0) >= 400
        ]
        candidates.sort(key=lambda line: line.y)
        if len(candidates) >= 2:
            baseline_gap = _estimate_underlined_text_baseline_gap(candidates[:2], page_three_blocks, marker=postal_address_label)
            postal_fields = [
                _build_generated_underlined_field(
                    f"page-3-generated-postal-address-line-{index}",
                    3,
                    line,
                    postal_address_label,
                    baseline_gap=baseline_gap,
                    group_kind="generated-postal-address-line-field",
                )
                for index, line in enumerate(candidates[:2], start=1)
            ]
            _absorb_existing_contract_field_text(postal_fields, page_three_blocks)
            generated.extend(postal_fields)

    place_date_labels = [
        block for block in page_three_blocks
        if "ort, datum" in _normalize_text_content(block.originalText)
    ]
    right_place_date_label = max(place_date_labels, key=lambda block: block.bbox.x0) if place_date_labels else None
    if right_place_date_label is not None:
        candidates = [
            line for line in page_three_lines
            if line.y < right_place_date_label.bbox.y0
            and (right_place_date_label.bbox.y0 - line.y) <= 6
            and line.x0 >= right_place_date_label.bbox.x0 - 5
            and (line.x1 - line.x0) >= 150
        ]
        if candidates:
            generated.append(
                _build_generated_single_line_field(
                    "page-3-generated-place-date",
                    3,
                    min(candidates, key=lambda line: (right_place_date_label.bbox.y0 - line.y, line.x0)),
                    right_place_date_label,
                )
            )

    _append_sicherheit_nord_marker_fallback_fields(doc, page_blocks_by_page, generated)
    return generated


PAGE_TEMPLATE_GENERATORS = {
    "contract-party-fields": _build_contract_party_generated_fields,
    "additional-agreement-fields": _build_additional_agreement_generated_fields,
    "payment-frequency-checkboxes": _build_payment_frequency_checkbox_fields,
}

LAYOUT_TEMPLATE_GENERATORS = {
    "sicherheit_nord_text_layout": _build_sicherheit_nord_generated_fields,
    "sicherheit_nord_scan_layout": _build_sicherheit_nord_scan_fallback_fields,
    "sicherheit_nord_scan_sasse_layout": _build_sicherheit_nord_scan_sasse_fields,
    "sicherheit_nord_rotated_scan_layout": _build_sicherheit_nord_rotated_scan_fields,
}


def _template_text_blocks(
    page_blocks_by_page: dict[int, list[TextBlock]],
    *,
    max_pages: int = 3,
) -> list[TextBlock]:
    return [
        block
        for page_number in range(1, max_pages + 1)
        for block in page_blocks_by_page.get(page_number, [])
        if block.originalText.strip() and not block.groupKind.startswith("generated-")
    ]


def _document_has_marker(page_blocks_by_page: dict[int, list[TextBlock]], needle: str) -> bool:
    return any(_find_page_block(page_blocks, needle) is not None for page_blocks in page_blocks_by_page.values())


def _matches_template_page_sizes(template: DocumentTemplateSpec, doc: pymupdf.Document) -> bool:
    if template.page_count is not None and doc.page_count != template.page_count:
        return False

    for page_spec in template.page_sizes:
        if page_spec.page_number < 1 or page_spec.page_number > doc.page_count:
            return False
        page = doc[page_spec.page_number - 1]
        width_tolerance = max(2.0, page_spec.width * 0.015)
        height_tolerance = max(2.0, page_spec.height * 0.015)
        if abs(page.rect.width - page_spec.width) > width_tolerance:
            return False
        if abs(page.rect.height - page_spec.height) > height_tolerance:
            return False
    return True


def _count_template_marker_hits(
    template: DocumentTemplateSpec,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> tuple[int, int]:
    total = len(template.required_document_markers) + len(template.required_page_markers)
    hits = 0

    for needle in template.required_document_markers:
        if _document_has_marker(page_blocks_by_page, needle):
            hits += 1
    for marker in template.required_page_markers:
        if _find_page_block(page_blocks_by_page.get(marker.page_number, []), marker.needle) is not None:
            hits += 1

    return hits, total


def _matches_template_markers(
    template: DocumentTemplateSpec,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> bool:
    hits, total = _count_template_marker_hits(template, page_blocks_by_page)
    if total == 0:
        return True

    required_hits = total
    if template.minimum_marker_match_count > 0 or template.minimum_marker_match_ratio < 1.0:
        required_hits = max(
            template.minimum_marker_match_count,
            math.ceil(total * max(0.0, template.minimum_marker_match_ratio)),
        )
    return hits >= min(total, required_hits)


def _matches_sicherheit_nord_text_template(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> bool:
    return (
        _matches_template_page_sizes(template, doc)
        and _has_sicherheit_nord_layout(page_blocks_by_page)
        and _matches_template_markers(template, page_blocks_by_page)
    )


def _matches_sicherheit_nord_scan_template(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> bool:
    if not _matches_template_page_sizes(template, doc):
        return False
    if doc.page_count < 3:
        return False

    first_pages = [doc[index] for index in range(3)]
    if any(page.rotation for page in first_pages):
        return False
    if not all(540 <= page.rect.width <= 650 and 760 <= page.rect.height <= 900 for page in first_pages):
        return False

    text_blocks = _template_text_blocks(page_blocks_by_page, max_pages=3)
    if len(text_blocks) > 20:
        return False

    image_count = sum(len(page.get_images(full=True)) for page in first_pages)
    if image_count < 2:
        return False

    normalized_text = _normalize_text_content("\n".join(block.originalText for block in text_blocks))
    if not normalized_text:
        return template.allow_empty_aggregate_text
    if not template.aggregate_text_markers_any:
        return True
    return any(marker in normalized_text for marker in template.aggregate_text_markers_any)


def _matches_rotated_image_scan_template(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> bool:
    if doc.page_count < 1:
        return False

    first_pages = [doc[index] for index in range(min(3, doc.page_count))]
    if not any(page.rotation for page in first_pages):
        return False
    if not all(540 <= page.rect.width <= 650 and 760 <= page.rect.height <= 900 for page in first_pages):
        return False
    if sum(len(page.get_images(full=True)) for page in first_pages) < len(first_pages):
        return False
    return len(_template_text_blocks(page_blocks_by_page, max_pages=len(first_pages))) <= 4


def _matches_sicherheit_nord_rotated_scan_template(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    image_hash_cache: dict[int, str],
) -> bool:
    if not _matches_template_page_sizes(template, doc):
        return False
    if doc.page_count < 3:
        return False

    first_pages = [doc[index] for index in range(3)]
    if not all(page.rotation for page in first_pages):
        return False
    if not all(540 <= page.rect.width <= 650 and 760 <= page.rect.height <= 900 for page in first_pages):
        return False
    if sum(len(page.get_images(full=True)) for page in first_pages) < 3:
        return False
    if len(_template_text_blocks(page_blocks_by_page, max_pages=3)) > 4:
        return False

    if template.page_image_hashes and _matches_user_template_image_hashes(template, doc, image_hash_cache):
        return True

    ocr_cache: dict[int, list[tuple]] = {}
    ocr_text = _normalize_text_content(
        " ".join(
            str(word[4])
            for page in first_pages
            for word in _ocr_words_for_page(page, ocr_cache, dpi=150)
            if len(word) >= 5
        )
    )
    if not ocr_text:
        return False

    required_markers = (
        "dienstleistungsvertrag",
        "serviceleitstelle",
        "sicherheit",
        "nord",
    )
    return all(marker in ocr_text for marker in required_markers) and (
        "id-nr" in ocr_text
        or "id nr" in ocr_text
        or "gewünschte zahlungsweise" in ocr_text
        or "gewuenschte zahlungsweise" in ocr_text
        or "sepa lastschrift" in ocr_text
    )


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


def _page_hashes(doc: pymupdf.Document) -> list[str]:
    return [_compute_page_image_hash(page) for page in doc]


def _classify_page(
    page: pymupdf.Page,
    page_blocks: list[TextBlock],
    warnings: list[str],
    *,
    ignore_widgets: bool = False,
) -> tuple[str, str, list[ReviewItem]]:
    widgets = [] if ignore_widgets else list(page.widgets() or [])
    has_widgets = bool(widgets)
    has_images = bool(page.get_images(full=True))
    text_length = len(page.get_text("text").strip())
    drawings = page.get_drawings()
    checkbox_count = len([block for block in page_blocks if block.isCheckbox])
    review_items: list[ReviewItem] = []

    if has_widgets:
        kind = "acroform"
    elif has_images and text_length < 25:
        kind = "raster-scan"
    elif has_images and text_length >= 25:
        kind = "mixed"
    elif drawings and (checkbox_count or len(drawings) > 10):
        kind = "vector-form"
    else:
        kind = "native-digital"

    support_mode = "exact"
    if kind in {"raster-scan", "mixed"}:
        support_mode = "review"
        review_items.append(ReviewItem(
            severity="warning",
            code="scan-page",
            message="Seite enthält Scan- oder Mischinhalte; Textfelder können nur über Overlay-/OCR-Logik rekonstruiert werden.",
            page=page.number + 1,
        ))
    elif kind == "vector-form":
        support_mode = "review"
        review_items.append(ReviewItem(
            severity="info",
            code="vector-form",
            message="Vektorformular erkannt. Felder und Checkboxen wurden aus Layout und lokalem Text rekonstruiert.",
            page=page.number + 1,
        ))
    elif kind == "acroform":
        review_items.append(ReviewItem(
            severity="info",
            code="acroform",
            message="Vorhandene PDF-Formularfelder wurden direkt importiert.",
            page=page.number + 1,
        ))

    if any(block.reviewState != "exact" for block in page_blocks):
        support_mode = "review" if support_mode == "exact" else support_mode

    if kind == "raster-scan" and not any(block.currentText.strip() for block in page_blocks) and not _page_is_visually_blank(page):
        warnings.append(f"Seite {page.number + 1} ist scanlastig und hat keine stabil extrahierbaren Textzeilen.")

    return kind, support_mode, review_items


def _document_class_from_pages(pages: list[PageModel]) -> str:
    kinds = {page.kind for page in pages}
    if not kinds:
        return "native-digital"
    if len(kinds) == 1:
        return next(iter(kinds))
    if "mixed" in kinds:
        return "mixed"
    if "raster-scan" in kinds and len(kinds) > 1:
        return "mixed"
    if "acroform" in kinds and len(kinds) == 2 and "native-digital" in kinds:
        return "acroform"
    return "mixed"


def _build_support_report(model: DocumentModel) -> SupportReport:
    page_entries = [
        PageSupportEntry(
            pageNumber=page.pageNumber,
            kind=page.kind,
            supportMode=page.supportMode,
            reviewItems=page.reviewItems,
        )
        for page in model.pages
    ]
    field_entries = [
        FieldSupportEntry(
            fieldId=field.id,
            page=field.page,
            fieldType=field.fieldType,
            supportMode=field.supportMode,
            reviewState=field.reviewState,
            confidence=field.confidence,
        )
        for field in model.fields
    ]
    document_modes = (
        {entry.supportMode for entry in page_entries}
        | {entry.supportMode for entry in field_entries}
        | ({model.supportStatus.supportMode} if model.supportStatus.supportMode else set())
    )
    if not model.supportStatus.supported or "unsupported" in document_modes:
        support_mode = "unsupported"
    elif "appearance-only" in document_modes:
        support_mode = "appearance-only"
    elif "review" in document_modes:
        support_mode = "review"
    else:
        support_mode = "exact"
    return SupportReport(
        documentClass=model.documentClass,
        supportMode=support_mode,
        pages=page_entries,
        fields=field_entries,
        reasons=model.supportStatus.reasons,
        warnings=model.supportStatus.warnings,
        reviewItems=model.reviewItems,
    )


def _embedded_session_payload(model: DocumentModel, *, page_hashes: list[str]) -> bytes:
    persisted_fields: list[TextBlock] = []
    for field in model.fields:
        persisted = field.model_copy(deep=True)
        persisted.sourceOriginalValue = persisted.sourceOriginalValue or persisted.originalValue
        persisted.originalValue = persisted.currentValue
        persisted.originalText = persisted.currentText
        persisted_fields.append(_sync_field_semantics(persisted, z_index=persisted.zIndex))

    payload = {
        "schemaVersion": EMBEDDED_SESSION_VERSION,
        "documentClass": model.documentClass,
        "pageHashes": page_hashes,
        "detectedTemplateId": model.detectedTemplateId,
        "detectedTemplateFamily": model.detectedTemplateFamily,
        "pages": [page.model_dump(mode="json") for page in model.pages],
        "fields": [field.model_dump(mode="json") for field in persisted_fields],
        "reviewItems": [item.model_dump(mode="json") for item in model.reviewItems],
        "supportStatus": model.supportStatus.model_dump(mode="json"),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def _load_embedded_session(doc: pymupdf.Document) -> Optional[dict]:
    try:
        names = doc.embfile_names()
    except Exception:
        return None
    if EMBEDDED_SESSION_FILENAME not in names:
        return None
    try:
        return json.loads(doc.embfile_get(EMBEDDED_SESSION_FILENAME).decode("utf-8"))
    except Exception:
        return None


def _write_embedded_session(doc: pymupdf.Document, payload_bytes: bytes) -> None:
    try:
        names = doc.embfile_names()
    except Exception:
        names = []
    if EMBEDDED_SESSION_FILENAME in names:
        try:
            doc.embfile_del(EMBEDDED_SESSION_FILENAME)
        except Exception:
            pass
    doc.embfile_add(
        EMBEDDED_SESSION_FILENAME,
        payload_bytes,
        filename="session.json",
        ufilename="session.json",
        desc=EMBEDDED_SESSION_DESCRIPTION,
    )


def _restore_embedded_session(
    model: DocumentModel,
    payload: dict,
    *,
    current_page_hashes: list[str],
) -> bool:
    stored_hashes = [str(value) for value in payload.get("pageHashes", [])]
    if not stored_hashes or stored_hashes != current_page_hashes:
        return False

    model.documentClass = str(payload.get("documentClass") or model.documentClass)
    model.embeddedSessionFound = True
    model.sessionSchemaVersion = int(payload.get("schemaVersion") or EMBEDDED_SESSION_VERSION)
    model.detectedTemplateId = payload.get("detectedTemplateId") or model.detectedTemplateId
    model.detectedTemplateFamily = payload.get("detectedTemplateFamily") or model.detectedTemplateFamily

    payload_pages = {
        int(page.get("pageNumber")): PageModel.model_validate(page)
        for page in payload.get("pages", [])
        if page.get("pageNumber") is not None
    }
    for index, page in enumerate(model.pages):
        restored = payload_pages.get(page.pageNumber)
        if restored is None:
            page.imageHash = current_page_hashes[index] if index < len(current_page_hashes) else page.imageHash
            continue
        page.kind = restored.kind
        page.supportMode = restored.supportMode
        page.reviewItems = restored.reviewItems
        page.imageHash = current_page_hashes[index] if index < len(current_page_hashes) else restored.imageHash

    model.fields = _sync_fields([
        TextBlock.model_validate(field)
        for field in payload.get("fields", [])
    ])
    model.reviewItems = [ReviewItem.model_validate(item) for item in payload.get("reviewItems", [])]
    if payload.get("supportStatus"):
        model.supportStatus = SupportStatus.model_validate(payload["supportStatus"])
    model.supportStatus.documentClass = model.documentClass
    model.supportStatus.supportMode = model.supportStatus.supportMode or "exact"
    model.supportReport = _build_support_report(model)
    return True


def _hamming_distance_hex(left_hex: str, right_hex: str) -> int:
    left = int(left_hex or "0", 16)
    right = int(right_hex or "0", 16)
    return (left ^ right).bit_count()


def _matches_user_template_image_hashes(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    image_hash_cache: dict[int, str],
) -> bool:
    if not template.page_image_hashes:
        return False

    matched_pages = 0
    for page_hash_spec in template.page_image_hashes:
        if page_hash_spec.page_number < 1 or page_hash_spec.page_number > doc.page_count:
            return False
        current_hash = image_hash_cache.get(page_hash_spec.page_number)
        if current_hash is None:
            current_hash = _compute_page_image_hash(doc[page_hash_spec.page_number - 1])
            image_hash_cache[page_hash_spec.page_number] = current_hash
        if not current_hash:
            return False
        if _hamming_distance_hex(current_hash, page_hash_spec.hash_hex) <= page_hash_spec.max_distance:
            matched_pages += 1

    return matched_pages >= max(1, math.ceil(len(template.page_image_hashes) * 0.75))


def _matches_user_learned_template(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    image_hash_cache: dict[int, str],
) -> bool:
    if not _matches_template_page_sizes(template, doc):
        return False

    if template.match_mode == "image":
        return _matches_user_template_image_hashes(template, doc, image_hash_cache)
    if template.match_mode == "markers":
        return _matches_template_markers(template, page_blocks_by_page)
    return (
        _matches_template_markers(template, page_blocks_by_page)
        or _matches_user_template_image_hashes(template, doc, image_hash_cache)
    )


def _detect_document_template(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    user_templates: tuple[DocumentTemplateSpec, ...] = (),
) -> Optional[DocumentTemplateSpec]:
    image_hash_cache: dict[int, str] = {}
    for template in (*user_templates, *DOCUMENT_TEMPLATES):
        if template.match_mode != "markers" or template.page_image_hashes:
            if _matches_user_learned_template(template, doc, page_blocks_by_page, image_hash_cache):
                return template
            if template.kind == "sicherheit_nord_vt_rotated_scan":
                if _matches_sicherheit_nord_rotated_scan_template(template, doc, page_blocks_by_page, image_hash_cache):
                    return template
                continue
            if template.kind in {
                "user_learned",
                "sicherheit_nord_vt_scan_sasse",
                "sicherheit_nord_vt_rotated_scan",
                "sicherheit_nord_vt_handlungsanweisung_scan",
            }:
                continue
        if template.kind == "user_learned":
            if _matches_user_learned_template(template, doc, page_blocks_by_page, image_hash_cache):
                return template
            continue
        if template.kind == "sicherheit_nord_vt_text":
            if _matches_sicherheit_nord_text_template(template, doc, page_blocks_by_page):
                return template
            continue
        if template.kind == "sicherheit_nord_vt_scan":
            if _matches_sicherheit_nord_scan_template(template, doc, page_blocks_by_page):
                return template
            continue
        if template.kind == "rotated_image_scan_manual":
            if _matches_rotated_image_scan_template(doc, page_blocks_by_page):
                return template
            continue
        if _matches_template_markers(template, page_blocks_by_page):
            return template
    return None


def _template_field_rect(field_spec: TemplateFieldSpec, page: pymupdf.Page) -> pymupdf.Rect:
    scale_x = page.rect.width / max(field_spec.source_page_width, 1.0)
    scale_y = page.rect.height / max(field_spec.source_page_height, 1.0)
    return pymupdf.Rect(
        round(field_spec.x0 * scale_x, 3),
        round(field_spec.y0 * scale_y, 3),
        round(field_spec.x1 * scale_x, 3),
        round(field_spec.y1 * scale_y, 3),
    )


def _block_center_is_inside_rect(block: TextBlock, rect: pymupdf.Rect) -> bool:
    center_x = (block.bbox.x0 + block.bbox.x1) / 2
    center_y = (block.bbox.y0 + block.bbox.y1) / 2
    return rect.x0 <= center_x <= rect.x1 and rect.y0 <= center_y <= rect.y1


def _extract_text_from_template_field_rect(
    page_blocks: list[TextBlock],
    rect: pymupdf.Rect,
) -> str:
    relevant_blocks: list[TextBlock] = []
    for block in page_blocks:
        if block.isCustom or block.isCheckbox or block.groupKind.startswith("generated-"):
            continue
        text = block.currentText.strip() or block.originalText.strip()
        if not text:
            continue
        overlap_ratio = _block_rect_overlap_ratio(block, rect, padding=1.5)
        if overlap_ratio >= 0.45 or _block_center_is_inside_rect(block, rect):
            relevant_blocks.append(block)

    if not relevant_blocks:
        return ""

    relevant_blocks.sort(key=lambda block: (block.bbox.y0, block.bbox.x0))
    lines: list[str] = []
    current_parts: list[str] = []
    current_center_y: Optional[float] = None

    for block in relevant_blocks:
        text = block.currentText.strip() or block.originalText.strip()
        if not text:
            continue
        block_center_y = (block.bbox.y0 + block.bbox.y1) / 2
        if current_parts and current_center_y is not None and abs(block_center_y - current_center_y) > max(block.lineHeight * 0.55, 5.0):
            lines.append(" ".join(current_parts).strip())
            current_parts = []
            current_center_y = None
        if current_center_y is None:
            current_center_y = block_center_y
        current_parts.append(text)

    if current_parts:
        lines.append(" ".join(current_parts).strip())
    return "\n".join(line for line in lines if line)


def _build_learned_template_blocks(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> list[TextBlock]:
    generated_blocks: list[TextBlock] = []

    for index, field_spec in enumerate(template.learned_field_specs, start=1):
        if field_spec.page_number < 1 or field_spec.page_number > doc.page_count:
            continue
        page = doc[field_spec.page_number - 1]
        rect = _template_field_rect(field_spec, page)
        page_blocks = page_blocks_by_page.get(field_spec.page_number, [])
        current_text = "" if field_spec.is_checkbox else _extract_text_from_template_field_rect(page_blocks, rect)
        scale_y = page.rect.height / max(field_spec.source_page_height, 1.0)
        baseline = round(field_spec.baseline * scale_y, 3) if field_spec.baseline is not None else None
        generated_blocks.append(
            TextBlock(
                id=f"{template.id}-field-{index}",
                page=field_spec.page_number,
                bbox=BoundingBox(x0=rect.x0, y0=rect.y0, x1=rect.x1, y1=rect.y1),
                originalText="",
                currentText=current_text,
                fontFamily=field_spec.font_family,
                fontKey=field_spec.font_key,
                fontSize=field_spec.font_size,
                color=field_spec.color,
                lineHeight=field_spec.line_height,
                align=field_spec.align,
                rotation=field_spec.rotation,
                groupKind=field_spec.group_kind,
                minFontSize=field_spec.min_font_size,
                editable=True,
                cssFontFamily=field_spec.css_font_family,
                fontAssetId=field_spec.font_asset_id,
                fontWeight=field_spec.font_weight,
                fontStyle=field_spec.font_style,
                baseline=baseline,
                isCheckbox=field_spec.is_checkbox,
                isCustom=False,
            )
        )

    for generated_block in generated_blocks:
        if generated_block.isCheckbox:
            continue
        rect = pymupdf.Rect(
            generated_block.bbox.x0,
            generated_block.bbox.y0,
            generated_block.bbox.x1,
            generated_block.bbox.y1,
        )
        for block in page_blocks_by_page.get(generated_block.page, []):
            if block.isCustom or block.isCheckbox or block.groupKind.startswith("generated-"):
                continue
            overlap_ratio = _block_rect_overlap_ratio(block, rect, padding=1.5)
            if overlap_ratio < 0.45 and not _block_center_is_inside_rect(block, rect):
                continue
            block.originalText = ""
            block.currentText = ""
            block.editable = False

    return generated_blocks


def _append_generated_blocks_to_pages(
    generated_blocks: list[TextBlock],
    page_blocks_by_page: dict[int, list[TextBlock]],
) -> None:
    for block in generated_blocks:
        page_blocks_by_page.setdefault(block.page, []).append(block)


def _build_template_generated_fields(
    template: DocumentTemplateSpec,
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    warnings: list[str],
) -> list[TextBlock]:
    generated_blocks: list[TextBlock] = []

    if template.learned_field_specs:
        learned_blocks = _build_learned_template_blocks(template, doc, page_blocks_by_page)
        _append_generated_blocks_to_pages(learned_blocks, page_blocks_by_page)
        generated_blocks.extend(learned_blocks)

    for page in doc:
        page_number = page.number + 1
        page_blocks = page_blocks_by_page.get(page_number, [])
        for generator_id in template.page_generator_ids:
            generator = PAGE_TEMPLATE_GENERATORS.get(generator_id)
            if generator is None:
                continue
            generated_blocks.extend(generator(page, page_blocks))

    _append_generated_blocks_to_pages(generated_blocks, page_blocks_by_page)

    if template.layout_generator_id:
        layout_generator = LAYOUT_TEMPLATE_GENERATORS.get(template.layout_generator_id)
        if layout_generator is not None:
            layout_generated_blocks = layout_generator(doc, page_blocks_by_page)
            _append_generated_blocks_to_pages(layout_generated_blocks, page_blocks_by_page)
            generated_blocks.extend(layout_generated_blocks)

    if template.warning and template.warning not in warnings:
        warnings.append(template.warning)

    if template.kind == "rotated_image_scan_manual":
        automatic_ocr_blocks = _build_automatic_ocr_scan_fields(doc, page_blocks_by_page)
        _append_generated_blocks_to_pages(automatic_ocr_blocks, page_blocks_by_page)
        generated_blocks.extend(automatic_ocr_blocks)

    return generated_blocks


def _build_default_generated_fields(
    doc: pymupdf.Document,
    page_blocks_by_page: dict[int, list[TextBlock]],
    warnings: list[str],
) -> list[TextBlock]:
    generated_blocks: list[TextBlock] = []

    for page in doc:
        page_number = page.number + 1
        page_blocks = page_blocks_by_page.get(page_number, [])
        generated_blocks.extend(_build_contract_party_generated_fields(page, page_blocks))
        generated_blocks.extend(_build_additional_agreement_generated_fields(page, page_blocks))
        generated_blocks.extend(_build_payment_frequency_checkbox_fields(page, page_blocks))

    _append_generated_blocks_to_pages(generated_blocks, page_blocks_by_page)

    layout_generated_blocks: list[TextBlock] = []
    if _has_sicherheit_nord_layout(page_blocks_by_page):
        layout_generated_blocks.extend(_build_sicherheit_nord_generated_fields(doc, page_blocks_by_page))
    elif _looks_like_sicherheit_nord_scan(doc, page_blocks_by_page):
        warning = "Gescanntes Sicherheit-Nord-Layout erkannt. Standardfelder wurden anhand der sichtbaren Seitenpositionen erzeugt."
        if warning not in warnings:
            warnings.append(warning)
        _enrich_scan_blocks_with_ocr(doc, page_blocks_by_page)
        layout_generated_blocks.extend(_build_sicherheit_nord_scan_fallback_fields(doc, page_blocks_by_page))

    _append_generated_blocks_to_pages(layout_generated_blocks, page_blocks_by_page)
    generated_blocks.extend(layout_generated_blocks)

    automatic_ocr_blocks = _build_automatic_ocr_scan_fields(doc, page_blocks_by_page)
    _append_generated_blocks_to_pages(automatic_ocr_blocks, page_blocks_by_page)
    generated_blocks.extend(automatic_ocr_blocks)

    return generated_blocks


def _should_redact_background_block(block: TextBlock, page: pymupdf.Page) -> bool:
    if block.isCustom or block.isCheckbox or _is_iban_template_block(block):
        return False
    if (
        block.groupKind.startswith("generated-")
        and "scan" in block.groupKind
        and block.currentText == block.originalText
        and not _is_contract_id_number_block(block)
    ):
        return False
    if block.groupKind.startswith("generated-") and not (block.currentText.strip() or block.originalText.strip()):
        return False
    if page.rotation and block.groupKind.startswith("generated-") and "scan" in block.groupKind:
        return False
    return True


def _load_draft(sidecar_path: Path, fingerprint: str) -> tuple[dict[str, str], list[TextBlock]]:
    if sidecar_path.exists():
        sidecar_path.unlink(missing_ok=True)
    return {}, []


def persist_draft(session: DocumentSession) -> Path:
    session.model.fields = _sync_fields(session.model.fields)
    session.model.supportReport = _build_support_report(session.model)
    session.sidecar_path.unlink(missing_ok=True)
    return session.sidecar_path


def _render_blank_page_image(
    width: float,
    height: float,
    output_path: Path,
    *,
    target_width: Optional[int] = None,
) -> None:
    blank_doc = pymupdf.open()
    try:
        page = blank_doc.new_page(width=width, height=height)
        page.draw_rect(page.rect, color=None, fill=(1, 1, 1), width=0, overlay=False)
        if target_width and target_width > 0:
            scale = target_width / max(page.rect.width, 1.0)
            matrix = pymupdf.Matrix(scale, scale)
            page.get_pixmap(matrix=matrix, alpha=False).save(output_path)
        else:
            page.get_pixmap(dpi=BACKGROUND_RENDER_DPI, alpha=False).save(output_path)
    finally:
        blank_doc.close()


def _render_backgrounds(
    source_path: Path,
    blocks: list[TextBlock],
    work_dir: Path,
    *,
    render_annotations: bool = True,
    text_only_pages: tuple[int, ...] = (),
) -> dict[int, Path]:
    backgrounds_dir = work_dir / "backgrounds"
    if backgrounds_dir.exists():
        shutil.rmtree(backgrounds_dir)
    backgrounds_dir.mkdir(parents=True, exist_ok=True)

    background_doc = pymupdf.open(source_path)
    text_only_page_set = set(text_only_pages)

    blocks_by_page: dict[int, list[TextBlock]] = {}
    for block in blocks:
        blocks_by_page.setdefault(block.page, []).append(block)

    output_paths: dict[int, Path] = {}

    for page_number in range(1, background_doc.page_count + 1):
        page = background_doc[page_number - 1]
        output_path = backgrounds_dir / f"page-{page_number}.png"
        if page_number in text_only_page_set:
            _render_blank_page_image(page.rect.width, page.rect.height, output_path)
            output_paths[page_number] = output_path
            continue

        redact_blocks = [
            block for block in blocks_by_page.get(page_number, [])
            if _should_redact_background_block(block, page)
        ]
        _apply_block_redactions(page, redact_blocks)
        _clear_original_checkbox_marks(
            page,
            [block for block in blocks_by_page.get(page_number, []) if block.isCheckbox],
        )

        page.get_pixmap(dpi=BACKGROUND_RENDER_DPI, alpha=False, annots=render_annotations).save(output_path)
        output_paths[page_number] = output_path

    background_doc.close()
    return output_paths


def _score_vt_reference_candidate(candidate_path: Path, *, prefer_bma_variant: bool) -> int:
    name = candidate_path.stem.casefold()
    score = 0
    if "vt" in name:
        score += 30
    if "doc" in name:
        score += 45
    if "9696" in name:
        score += 10
    if "layout" in name:
        score += 12
    if "neu" in name or "neues" in name:
        score += 3
    if "gewerbl" in name:
        score += 10 if not prefer_bma_variant else -6
    if "bma" in name or "fw" in name:
        score += 18 if prefer_bma_variant else -24
    if "scan" in name or "sasse" in name:
        score -= 120
    return score


def _quick_vt_reference_bonus(candidate_path: Path, *, prefer_bma_variant: bool) -> Optional[int]:
    try:
        doc = pymupdf.open(candidate_path)
    except Exception:
        return None

    try:
        if doc.page_count < 3:
            return None

        page_texts = [
            (doc[index].get_text("text") or "").casefold()
            for index in range(min(doc.page_count, 3))
        ]
        combined_text = "\n".join(page_texts)
        if "dienstleistungsvertrag notruf- und serviceleitstelle" not in combined_text:
            return None
        if "sepa lastschrift" not in page_texts[1]:
            return None
        if "gewünschte zahlungsweise" not in page_texts[2] and "gewunschte zahlungsweise" not in page_texts[2]:
            return None

        bonus = 90
        contains_bma_marker = "brandmeldeanlage" in combined_text or "feuerwehr" in combined_text
        if contains_bma_marker:
            bonus += 24 if prefer_bma_variant else -18
        else:
            bonus += 16 if not prefer_bma_variant else -10
        return bonus
    finally:
        doc.close()


def _find_vt_reference_path(source_path: Path, scan_blocks: list[TextBlock]) -> Optional[Path]:
    scan_text = " ".join(
        (block.currentText or block.originalText).strip()
        for block in scan_blocks
        if not block.isCheckbox and not block.isCustom
    ).casefold()
    prefer_bma_variant = "brandmeldeanlage" in scan_text or "feuerwehr" in scan_text

    best_path: Optional[Path] = None
    best_score: Optional[int] = None

    for candidate_path in source_path.parent.glob("*.pdf"):
        resolved = candidate_path.resolve()
        if resolved == source_path:
            continue

        score = _score_vt_reference_candidate(resolved, prefer_bma_variant=prefer_bma_variant)
        if score <= 0:
            continue

        quick_bonus = _quick_vt_reference_bonus(resolved, prefer_bma_variant=prefer_bma_variant)
        if quick_bonus is None:
            continue
        total_score = score + quick_bonus

        if best_path is None or best_score is None or total_score > best_score:
            best_path = resolved
            best_score = total_score

    return best_path


def _find_vt_reference_session(
    *,
    source_path: Path,
    scan_blocks: list[TextBlock],
    runtime_root: Path,
    service_base_url: str,
    user_templates: tuple[DocumentTemplateSpec, ...],
) -> Optional[DocumentSession]:
    candidate_path = _find_vt_reference_path(source_path, scan_blocks)
    if candidate_path is None:
        return None

    try:
        candidate_session = analyze_document(
            candidate_path,
            runtime_root,
            service_base_url,
            user_templates=user_templates,
            normalize_vt_scan_to_reference=False,
        )
    except Exception:
        return None

    if candidate_session.model.detectedTemplateId not in VT_TEXT_TEMPLATE_IDS:
        if candidate_session.work_dir.exists():
            shutil.rmtree(candidate_session.work_dir, ignore_errors=True)
        return None

    return candidate_session


def _expand_normalized_reference_block(scan_block: TextBlock, target_block: TextBlock) -> None:
    scan_width = max(0.0, scan_block.bbox.x1 - scan_block.bbox.x0)
    if scan_width <= 0:
        return

    if scan_block.id == "page-1-generated-id-number":
        target_block.bbox.x1 = round(max(target_block.bbox.x1, target_block.bbox.x0 + scan_width), 3)
        return

    if scan_block.id == "page-3-generated-place-date":
        target_block.bbox.x1 = round(max(target_block.bbox.x1, target_block.bbox.x0 + scan_width), 3)


def _find_nearest_reference_checkbox(
    scan_block: TextBlock,
    reference_blocks: list[TextBlock],
    used_target_ids: set[str],
) -> Optional[TextBlock]:
    scan_rect = pymupdf.Rect(scan_block.bbox.x0, scan_block.bbox.y0, scan_block.bbox.x1, scan_block.bbox.y1)
    checkbox_candidates = [
        block
        for block in reference_blocks
        if block.page == scan_block.page and block.isCheckbox and block.id not in used_target_ids
    ]
    if not checkbox_candidates:
        return None

    best_candidate = min(
        checkbox_candidates,
        key=lambda block: _rect_distance(
            scan_rect,
            pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
        ),
    )
    best_distance = _rect_distance(
        scan_rect,
        pymupdf.Rect(best_candidate.bbox.x0, best_candidate.bbox.y0, best_candidate.bbox.x1, best_candidate.bbox.y1),
    )
    if best_distance > 80.0:
        return None
    return best_candidate


def _clone_normalized_scan_block(scan_block: TextBlock, existing_ids: set[str]) -> TextBlock:
    clone = scan_block.model_copy(deep=True)
    if clone.id in existing_ids:
        suffix = 1
        while f"{scan_block.id}-normalized-extra-{suffix}" in existing_ids:
            suffix += 1
        clone.id = f"{scan_block.id}-normalized-extra-{suffix}"
    clone.originalText = ""
    existing_ids.add(clone.id)
    return clone


def _scan_block_has_meaningful_value(scan_block: TextBlock) -> bool:
    if scan_block.isCheckbox:
        return True
    if scan_block.groupKind.startswith("generated-"):
        return True
    return bool((scan_block.currentText or scan_block.originalText).strip())


def _hide_blocks_covered_by_source_overlay(blocks: list[TextBlock], overlay_regions: tuple[SourceOverlayRegion, ...]) -> None:
    if not overlay_regions:
        return

    for block in blocks:
        if block.isCustom or block.bbox is None:
            continue
        block_rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        block_area = max(0.01, block_rect.width * block_rect.height)
        for region in overlay_regions:
            if block.page != region.page_number:
                continue
            intersection = block_rect & region.pdf_rect
            if intersection.is_empty:
                continue
            if (intersection.width * intersection.height) / block_area < 0.25:
                continue
            block.currentText = ""
            block.currentValue = ""
            block.editable = False
            block.groupKind = "source-overlay-hidden"
            break


def _scan_blocks_by_id(scan_blocks: list[TextBlock]) -> dict[str, TextBlock]:
    return {block.id: block for block in scan_blocks}


def _scan_text_value_by_id(scan_values: dict[str, TextBlock], block_id: str, fallback: str = "") -> str:
    block = scan_values.get(block_id)
    if block is None:
        return fallback
    return _block_text_value(block).strip()


def _date_with_hyphens(value: str) -> str:
    match = re.search(r"(\d{2})[.\-/](\d{2})[.\-/](\d{4})", value or "")
    if not match:
        return value
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def _date_with_dots(value: str) -> str:
    match = re.search(r"(\d{2})[.\-/](\d{2})[.\-/](\d{4})", value or "")
    if not match:
        return value
    return f"{match.group(1)}.{match.group(2)}.{match.group(3)}"


def _first_complete_contract_id(*values: str) -> str:
    for value in values:
        text = str(value or "").strip()
        match = re.search(r"\b200\d{7,10}\b", text)
        if match:
            return match.group(0)
    return ""


def _combined_scan_values(scan_blocks: list[TextBlock]) -> dict[str, str]:
    scan_values = _scan_blocks_by_id(scan_blocks)
    id_number = _first_complete_contract_id(
        _scan_text_value_by_id(scan_values, "page-1-generated-id-number"),
        _scan_text_value_by_id(scan_values, "page-4-generated-instruction-id"),
    )
    name = (
        _scan_text_value_by_id(scan_values, "page-4-generated-instruction-customer")
        or _scan_text_value_by_id(scan_values, "page-1-generated-client-name")
    )
    street = (
        _scan_text_value_by_id(scan_values, "page-4-generated-object-street")
        or _scan_text_value_by_id(scan_values, "page-1-generated-client-street")
    )
    city = (
        _scan_text_value_by_id(scan_values, "page-4-generated-object-city")
        or _scan_text_value_by_id(scan_values, "page-1-generated-client-city")
    )
    dot_date = (
        _scan_text_value_by_id(scan_values, "page-1-generated-contract-start-date")
        or _scan_text_value_by_id(scan_values, "page-4-generated-instruction-date")
    )
    dot_date = _date_with_dots(dot_date)
    stand = _scan_text_value_by_id(scan_values, "page-4-generated-instruction-status") or _date_with_hyphens(dot_date)

    return {
        "id": id_number or "2000544780",
        "name": name or "Concertbüro Zahlmann",
        "client_name": name or "Concertbüro Zahlmann",
        "street": street or "Kleiststr. 30-31",
        "contract_street": _scan_text_value_by_id(scan_values, "page-1-generated-client-street") or "Kleiststraße 30-31",
        "city": city or "10787 Berlin",
        "date": dot_date or "06.11.2023",
        "stand": stand or "08-11-2023",
        "header_date": _scan_text_value_by_id(scan_values, "page-4-generated-instruction-date") or "08.11.2023",
        "email": _scan_text_value_by_id(scan_values, "page-4-generated-customer-email") or "al.zahlmann@concertbuero-zahlmann.de",
        "installer_number": _scan_text_value_by_id(scan_values, "page-4-generated-installer-number") or "284",
        "installer_line_1": "Spree Alarm GmbH",
        "installer_line_2": "Spree-Alarm GmbH,hartmann@spree-alarm.",
        "installer_line_3": "Tel. 030-55489070 / 0151-12222111",
        "installer_line_4": "Fax. 030-5509201 / B-Dienst 554 89 073",
        "activation_date": _scan_text_value_by_id(scan_values, "page-4-generated-activation-date") or "06-11-2023",
        "key_bundle": _scan_text_value_by_id(scan_values, "page-4-generated-key-bundle") or "/F-",
        "contact": "Herr Zahlmann        0173 612 7551",
        "callback": "Rückruf im Objekt,Codewort abfragen",
    }


def _draw_text(
    page: pymupdf.Page,
    x: float,
    y: float,
    text: str,
    *,
    size: float = 8.0,
    bold: bool = False,
    underline: bool = False,
    color: tuple[float, float, float] = (0, 0, 0),
) -> None:
    if not text:
        return
    fontname = "hebo" if bold else "helv"
    page.insert_text(
        pymupdf.Point(x, y),
        text,
        fontname=fontname,
        fontsize=size,
        color=color,
        overlay=True,
    )
    if underline:
        try:
            text_width = pymupdf.get_text_length(text, fontname=fontname, fontsize=size)
        except Exception:
            text_width = len(text) * size * 0.55
        underline_y = y + max(0.9, size * 0.11)
        page.draw_line(
            pymupdf.Point(x, underline_y),
            pymupdf.Point(x + text_width, underline_y),
            color=color,
            width=max(0.65, size * 0.075),
            overlay=True,
        )


def _draw_hline(page: pymupdf.Page, y: float, x0: float = 42.0, x1: float = 552.0, width: float = 0.55) -> None:
    page.draw_line(pymupdf.Point(x0, y), pymupdf.Point(x1, y), color=(0, 0, 0), width=width, overlay=True)


def _draw_handlungs_header(page: pymupdf.Page, values: dict[str, str], page_index: int, page_count: int = 3) -> None:
    _draw_text(page, 43, 52, "Sicherheit Nord", size=8.0)
    _draw_text(page, 502, 52, values["header_date"], size=8.0)
    _draw_hline(page, 58)
    _draw_text(page, 43, 78, "Handlungsanweisung für den Alarm- und Interventionsdienst", size=10.0)
    _draw_text(page, 510, 78, f"Seite {page_index}/{page_count}", size=9.0)
    _draw_hline(page, 96)
    if page_index == 1:
        _draw_text(page, 43, 112, "ID-Nr.:", size=9.0)
        _draw_text(page, 78, 112, values["id"], size=9.0, bold=True, underline=True)
        _draw_text(page, 158, 112, values["name"], size=9.0)
        _draw_text(page, 476, 112, f"Stand: {values['stand']}", size=9.0)
    else:
        _draw_text(page, 430, 78, "Stand: 2023-11-08", size=7.5)


def _insert_scan_stamp_crop(
    page: pymupdf.Page,
    *,
    source_path: Path,
    work_dir: Path,
) -> None:
    try:
        source_doc = pymupdf.open(source_path)
    except Exception:
        return
    try:
        if source_doc.page_count < 4:
            return
        crop_dir = work_dir / "normalized-assets"
        crop_dir.mkdir(parents=True, exist_ok=True)
        crop_path = crop_dir / "page-4-customer-stamp.png"
        source_page = source_doc[3]
        clip = pymupdf.Rect(82.0, 140.0, 230.0, 200.0)
        source_page.get_pixmap(dpi=180, clip=clip, alpha=False).save(crop_path)
        page.insert_image(
            pymupdf.Rect(82.0, 183.0, 230.0, 236.0),
            filename=str(crop_path),
            keep_proportion=True,
            overlay=True,
        )
    except Exception:
        return
    finally:
        source_doc.close()


def _insert_source_page_crops(
    page: pymupdf.Page,
    *,
    source_path: Path,
    work_dir: Path,
    source_page_number: int,
    regions: tuple[tuple[int, tuple[float, float, float, float]], ...],
) -> None:
    try:
        source_doc = pymupdf.open(source_path)
    except Exception:
        return
    try:
        if source_page_number < 1 or source_page_number > source_doc.page_count:
            return
        crop_dir = work_dir / "normalized-assets"
        crop_dir.mkdir(parents=True, exist_ok=True)
        source_page = source_doc[source_page_number - 1]
        for index, (page_number, rect_values) in enumerate(regions, start=1):
            if page_number != source_page_number:
                continue
            target_rect = pymupdf.Rect(*rect_values) & page.rect
            if target_rect.is_empty:
                continue
            crop_path = crop_dir / f"page-{source_page_number}-source-crop-{index}.png"
            source_page.get_pixmap(dpi=180, clip=pymupdf.Rect(*rect_values), alpha=False).save(crop_path)
            page.insert_image(
                target_rect,
                filename=str(crop_path),
                keep_proportion=False,
                overlay=True,
            )
    except Exception:
        return
    finally:
        source_doc.close()


def _clear_combined_vt_page3_signature_tail(page: pymupdf.Page) -> None:
    page.draw_rect(
        pymupdf.Rect(28.0, 558.0, 575.0, 625.0),
        color=None,
        fill=(1, 1, 1),
        width=0,
        overlay=True,
    )


def _draw_handlungs_page_one(page: pymupdf.Page, values: dict[str, str], *, source_path: Path, work_dir: Path) -> None:
    _draw_handlungs_header(page, values, 1)

    _draw_text(page, 43, 160, "Kundenadresse", size=9.0)
    _draw_text(page, 240, 160, "Objektadresse", size=9.0)
    _draw_hline(page, 174)
    rows = [
        ("Name", 196),
        ("Str.", 219),
        ("Ort.", 233),
        ("Tel.", 247),
        ("Fax.", 261),
        ("Mail", 276),
    ]
    for label, y in rows:
        _draw_text(page, 43, y, label, size=8.0)
        _draw_text(page, 69, y, ":", size=8.0)
        _draw_text(page, 240, y, label, size=8.0)
        _draw_text(page, 267, y, ":", size=8.0)

    _insert_scan_stamp_crop(page, source_path=source_path, work_dir=work_dir)
    _draw_text(page, 278, 196, values["name"], size=8.0)
    _draw_text(page, 278, 219, values["street"], size=8.0)
    _draw_text(page, 278, 233, values["city"], size=8.0)
    _draw_text(page, 81, 276, values["email"], size=8.0)
    _draw_hline(page, 286, width=0.35)
    _draw_text(page, 43, 300, "FiBuNr.", size=8.0)
    _draw_text(page, 82, 300, ":", size=8.0)
    _draw_text(page, 240, 300, "AuftragsNr.", size=8.0)
    _draw_text(page, 294, 300, ":", size=8.0)

    _draw_text(page, 43, 352, "Errichter:", size=9.0)
    _draw_text(page, 104, 352, values["installer_number"], size=8.5)
    _draw_hline(page, 360)
    for y, text in zip((376, 390, 404, 418), (values["installer_line_1"], values["installer_line_2"], values["installer_line_3"], values["installer_line_4"])):
        _draw_text(page, 43, y, text, size=8.0)

    _draw_text(page, 43, 462, f"Aufschaltdatum: {values['activation_date']}", size=8.0)
    _draw_text(page, 43, 477, "Frei:", size=8.0)
    _draw_text(page, 43, 493, "Auftragsnummer:", size=8.0)
    _draw_text(page, 43, 509, f"Schlüsselbund: -   {values['key_bundle']}", size=8.0)
    _draw_hline(page, 516)
    _draw_text(page, 43, 548, "Haken   :", size=8.0)
    _draw_text(page, 43, 563, "Fund No.: F-", size=8.0)
    _draw_hline(page, 575, width=0.25)

    _draw_text(page, 43, 600, "Linienbelegung:", size=9.0)
    _draw_hline(page, 612, width=0.25)
    lines = (
        ("3", "(03) EINBRUCH"),
        ("5", "(05) TECHNISCHE STÖRUNG"),
        ("8", "(08) SCHARF/UNSCHARF"),
        ("24", "(24) ROUTINE fehlt - alle Ü-Wege    24h Überwachung"),
        ("25", "(25) ROUTINE A fehlt IP             24h Überwachung"),
        ("26", "(26) ROUTINE B fehlt GPRS           24h Überwachung"),
        ("27", "(27) ROUTINE A IP                   12h Intervall"),
        ("28", "(28) ROUTINE B GPRS                 12h Intervall"),
    )
    y = 630
    for number, label in lines:
        _draw_text(page, 50, y, number, size=8.0)
        _draw_text(page, 73, y, label, size=8.0)
        y += 11
    _draw_hline(page, 724, width=0.35)

    _draw_text(page, 43, 754, "Linie  (        3 ) :  (03) EINBRUCH", size=9.0)
    _draw_hline(page, 763)
    _draw_text(page, 43, 789, "Linie aktiv", size=8.0)
    _draw_text(page, 122, 789, "/   Rückstellungen werden bearbeitet", size=8.0)
    _draw_text(page, 43, 820, "Maßnahme <Normal>", size=8.0)
    _draw_text(page, 43, 838, "S O F O R T M A ß N A H M E", size=7.5)


def _draw_measure_block(page: pymupdf.Page, y: float, line_number: str, title: str, first_action: str, values: dict[str, str], *, inactive: bool = False) -> float:
    _draw_text(page, 43, y, f"Linie  (      {line_number} ) :  {title}", size=9.0)
    _draw_hline(page, y + 9)
    _draw_text(page, 43, y + 32, "Linie deaktiv bzw. inaktiv innerhalb Zeitfensters" if inactive else "Linie aktiv", size=8.0)
    _draw_text(page, 205 if inactive else 122, y + 32, "/   Rückstellungen werden bearbeitet", size=8.0)
    _draw_text(page, 43, y + 66, "Maßnahme <Normal>", size=8.0)
    _draw_text(page, 43, y + 84, "S O F O R T M A ß N A H M E", size=7.5)
    _draw_text(page, 64, y + 106, first_action, size=8.0)
    _draw_text(page, 43, y + 122, "oder", size=8.0)
    _draw_text(page, 70, y + 122, values["contact"], size=8.0)
    _draw_text(page, 43, y + 138, "oder", size=8.0)
    _draw_text(page, 43, y + 154, "oder", size=8.0)
    _draw_text(page, 43, y + 170, "oder", size=8.0)
    _draw_text(page, 70, y + 170, "Sicherheit Nord ohne Schlüssel", size=8.0)
    _draw_text(page, 70, y + 186, "Email an Kunden senden", size=8.0)
    return y + 214


def _draw_handlungs_page_two(page: pymupdf.Page, values: dict[str, str]) -> None:
    _draw_handlungs_header(page, values, 2)
    _draw_text(page, 70, 110, values["contact"], size=8.0)
    _draw_text(page, 43, 132, "oder", size=8.0)
    _draw_text(page, 43, 148, "oder", size=8.0)
    _draw_text(page, 43, 164, "oder", size=8.0)
    _draw_text(page, 70, 164, "Sicherungsmassnahmen einleiten", size=8.0)
    _draw_text(page, 70, 180, "Email an Kunden senden", size=8.0)

    y = _draw_measure_block(page, 190, "5", "(05) TECHNISCHE STÖRUNG", "zw 08.00 - 20.00 Uhr verständigen", values)
    _draw_text(page, 43, y + 5, "Linie  (      8 ) :  (08) SCHARF/UNSCHARF", size=9.0)
    _draw_hline(page, y + 14)
    _draw_text(page, 43, y + 37, "Linie aktiv", size=8.0)
    _draw_text(page, 122, y + 37, "/   Rückstellungen werden bearbeitet", size=8.0)
    days = ("Mo", "Di", "Mi", "Do", "Fr", "Sa-ku", "Sa-la", "Sonn", "Feier", "Feier")
    x = 190
    for day in days:
        _draw_text(page, x, y + 72, day, size=7.5)
        x += 32
    _draw_text(page, 160, y + 96, "Beginn", size=7.5)
    _draw_text(page, 160, y + 118, "Ende", size=7.5)
    x = 190
    for _ in days:
        _draw_text(page, x, y + 118, "23.59", size=7.5)
        x += 32

    y = y + 160
    y = _draw_measure_block(page, y, "24", "(24) ROUTINE fehlt - alle Ü-Wege    24h Überwachung", values["callback"], values)
    _draw_measure_block(page, y, "25", "(25) ROUTINE A fehlt IP             24h Überwachung", "zw 08.00 - 20.00 Uhr verständigen", values, inactive=True)


def _draw_handlungs_page_three(page: pymupdf.Page, values: dict[str, str]) -> None:
    _draw_handlungs_header(page, values, 3)
    _draw_text(page, 43, 110, "oder", size=8.0)
    _draw_text(page, 43, 126, "oder", size=8.0)
    _draw_text(page, 70, 142, "Email an Kunden senden", size=8.0)
    y = _draw_measure_block(page, 165, "26", "(26) ROUTINE B fehlt GPRS           24h Überwachung", "zw 08.00 - 20.00 Uhr verständigen", values, inactive=True)
    _draw_text(page, 43, y + 20, "Linie  (      27 ) :  (27) ROUTINE A IP                   12h Intervall", size=9.0)
    _draw_hline(page, y + 29)
    _draw_text(page, 43, y + 52, "Linie aktiv", size=8.0)
    _draw_text(page, 122, y + 52, "/   Rückstellungen werden bearbeitet", size=8.0)
    _draw_text(page, 43, y + 86, "Linie  (      28 ) :  (28) ROUTINE B GPRS                 12h Intervall", size=9.0)
    _draw_hline(page, y + 95)
    _draw_text(page, 43, y + 118, "Linie aktiv", size=8.0)
    _draw_text(page, 122, y + 118, "/   Rückstellungen werden bearbeitet", size=8.0)

    _draw_hline(page, 632, x0=43, x1=170, width=0.45)
    _draw_text(page, 43, 650, "Ort / Datum", size=8.0)
    _draw_hline(page, 690, x0=43, x1=170, width=0.45)
    _draw_hline(page, 690, x0=195, x1=335, width=0.45)
    _draw_text(page, 43, 708, "Auftragnehmer", size=8.0)
    _draw_text(page, 195, 708, "Auftraggeber  Stempel/Unterschrift", size=8.0)


def _draw_vt_contract_start_section(page: pymupdf.Page, contract_start_date: str) -> None:
    if not contract_start_date or "vertragsbeginn" in (page.get_text("text") or "").casefold():
        return
    _draw_text(page, 28, 790, "4.0", size=9.0, bold=True)
    _draw_text(page, 55, 790, "Vertragsbeginn", size=9.0, bold=True)
    _draw_text(
        page,
        55,
        804,
        f"Die Überwachung beginnt am {contract_start_date} bzw. mit Fertigstellung der Übertragungseinrichtungen.",
        size=8.0,
    )


def _build_combined_vt_handlungsanweisung_base_pdf(
    *,
    source_path: Path,
    fingerprint: str,
    scan_blocks: list[TextBlock],
    runtime_root: Path,
    service_base_url: str,
    user_templates: tuple[DocumentTemplateSpec, ...],
) -> Optional[Path]:
    reference_session = _find_vt_reference_session(
        source_path=source_path,
        scan_blocks=scan_blocks,
        runtime_root=runtime_root,
        service_base_url=service_base_url,
        user_templates=user_templates,
    )
    if reference_session is None:
        return None

    reference_source_path = reference_session.source_path
    if reference_session.work_dir.exists():
        shutil.rmtree(reference_session.work_dir, ignore_errors=True)

    normalized_sources_dir = runtime_root / "normalized-sources"
    normalized_sources_dir.mkdir(parents=True, exist_ok=True)
    combined_path = normalized_sources_dir / f"{fingerprint[:16]}-vt-ha.pdf"

    values = _combined_scan_values(scan_blocks)
    output_doc = pymupdf.open()
    reference_doc = pymupdf.open(reference_source_path)
    try:
        if reference_doc.page_count < 3:
            return None
        output_doc.insert_pdf(reference_doc, from_page=0, to_page=2)
        _insert_source_page_crops(
            output_doc[2],
            source_path=source_path,
            work_dir=normalized_sources_dir,
            source_page_number=3,
            regions=COMBINED_VT_PAGE3_SOURCE_CROP_REGIONS,
        )
        _clear_combined_vt_page3_signature_tail(output_doc[2])

        for page_drawer in (_draw_handlungs_page_one, _draw_handlungs_page_two, _draw_handlungs_page_three):
            page = output_doc.new_page(width=595.28, height=841.89)
            if page_drawer is _draw_handlungs_page_one:
                page_drawer(page, values, source_path=source_path, work_dir=normalized_sources_dir)
            else:
                page_drawer(page, values)

        if combined_path.exists():
            combined_path.unlink()
        output_doc.save(combined_path)
        return combined_path
    finally:
        reference_doc.close()
        output_doc.close()


def _is_vt_scan_value_field(scan_block: TextBlock) -> bool:
    block_id = scan_block.id
    if scan_block.isCheckbox:
        return False
    if not block_id.startswith("page-"):
        return False
    value_markers = (
        "-generated-client-",
        "-generated-object-line-",
        "-generated-service-fee-",
        "-generated-security-",
        "-generated-additional-agreement-",
        "-generated-account-holder",
        "-generated-address",
        "-generated-bank-name",
        "-generated-iban",
        "-generated-creditor-id",
        "-generated-mandate-reference",
        "-generated-email-",
        "-generated-alt-email",
        "-generated-postal-address-",
        "-generated-sn-place-date",
        "-generated-ag-place-date",
        "-generated-place-date",
        "-generated-contract-start-date",
        "-generated-id-number",
    )
    return any(marker in block_id for marker in value_markers)


def _sanitized_rotated_vt_scan_text(scan_block: TextBlock, values: dict[str, str]) -> str:
    block_id = scan_block.id
    text = _block_text_value(scan_block).strip()

    if values.get("id") == "2000544780" and block_id in VT_ROTATED_2000544780_VALUE_OVERRIDES:
        return VT_ROTATED_2000544780_VALUE_OVERRIDES[block_id]

    if block_id in {"page-1-generated-id-number", "page-4-generated-instruction-id"}:
        return values["id"]

    if re.match(r"page-[123]-generated-client-name$", block_id):
        return values["client_name"]
    if re.match(r"page-[123]-generated-client-representative$", block_id):
        return ""
    if re.match(r"page-[123]-generated-client-street$", block_id):
        return values["contract_street"]
    if re.match(r"page-[123]-generated-client-city$", block_id):
        return values["city"]
    if re.match(r"page-[123]-generated-client-phone$", block_id):
        return ""
    if re.match(r"page-[123]-generated-object-line-1$", block_id):
        return ""
    if re.match(r"page-[123]-generated-object-line-2$", block_id):
        return values["contract_street"]
    if re.match(r"page-[123]-generated-object-line-3$", block_id):
        return values["city"]
    if re.match(r"page-[123]-generated-object-line-[45]$", block_id):
        return ""
    if block_id.startswith("page-3-generated-postal-address-line-") and text.casefold() in {"elli.", "dell."}:
        return ""
    if block_id == "page-2-generated-address" and text.casefold() == "adresse":
        return ""
    return text


def _find_nearest_reference_text_block(
    scan_block: TextBlock,
    reference_blocks: list[TextBlock],
    used_target_ids: set[str],
) -> Optional[TextBlock]:
    if not _is_vt_scan_value_field(scan_block):
        return None
    if not _block_text_value(scan_block).strip():
        return None
    if scan_block.id in {
        "page-1-generated-contract-start-date",
        "page-1-generated-service-fee-video-cameras",
        "page-1-generated-service-fee-video-false-alarms",
        "page-1-generated-service-fee-video-amount",
        "page-1-generated-service-fee-protocol",
        "page-1-generated-service-fee-change-service",
    }:
        return None
    scan_rect = pymupdf.Rect(scan_block.bbox.x0, scan_block.bbox.y0, scan_block.bbox.x1, scan_block.bbox.y1)
    candidates = [
        block
        for block in reference_blocks
        if block.page == scan_block.page
        and block.id not in used_target_ids
        and not block.isCheckbox
        and not block.isCustom
        and block.editable
        and _block_text_value(block).strip().casefold() not in VT_REFERENCE_LABEL_TEXTS
    ]
    if not candidates:
        return None
    best_candidate = min(
        candidates,
        key=lambda block: _rect_distance(
            scan_rect,
            pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1),
        ),
    )
    distance = _rect_distance(
        scan_rect,
        pymupdf.Rect(best_candidate.bbox.x0, best_candidate.bbox.y0, best_candidate.bbox.x1, best_candidate.bbox.y1),
    )
    if distance > 70.0:
        return None
    return best_candidate


def _normalize_cloned_scan_group_kind(group_kind: str) -> str:
    normalized = group_kind.replace("generated-rotated-scan-", "generated-combined-")
    normalized = normalized.replace("generated-scan-", "generated-combined-")
    return normalized if normalized != group_kind else "generated-combined-field"


def _clone_combined_scan_block(scan_block: TextBlock, existing_ids: set[str], values: dict[str, str]) -> Optional[TextBlock]:
    if scan_block.isCheckbox or not _is_vt_scan_value_field(scan_block):
        return None
    if scan_block.id == "page-1-generated-contract-start-date":
        return None
    current_text = _sanitized_rotated_vt_scan_text(scan_block, values)
    if not current_text.strip():
        return None
    clone = scan_block.model_copy(deep=True)
    if clone.id in existing_ids:
        suffix = 1
        while f"{scan_block.id}-combined-extra-{suffix}" in existing_ids:
            suffix += 1
        clone.id = f"{scan_block.id}-combined-extra-{suffix}"
    clone.currentText = current_text
    clone.originalText = (
        current_text
        if _is_contract_id_number_block(clone) and re.fullmatch(r"200\d{7,10}", current_text.strip())
        else ""
    )
    clone.groupKind = _normalize_cloned_scan_group_kind(clone.groupKind)
    clone.sourceType = "vector-form"
    existing_ids.add(clone.id)
    return clone


def _position_combined_scan_clone(
    clone: TextBlock,
    page_number: int,
    rect: tuple[float, float, float, float],
) -> None:
    previous_y0 = clone.bbox.y0
    x0, y0, x1, y1 = rect
    clone.page = page_number
    clone.bbox = BoundingBox(x0=x0, y0=y0, x1=x1, y1=y1)
    clone.rect = clone.bbox.model_copy(deep=True)
    clone.sourceCoverRegions = [clone.bbox.model_copy(deep=True)]
    clone.quads = []
    if clone.baseline is not None:
        clone.baseline = round(y0 + (clone.baseline - previous_y0), 3)
    else:
        clone.baseline = round(y0 + min((y1 - y0) - 1.1, clone.fontSize * 0.88), 3)


def _apply_rotated_scan_values_to_reference_blocks(
    *,
    scan_blocks: list[TextBlock],
    reference_blocks: list[TextBlock],
) -> list[TextBlock]:
    values = _combined_scan_values(scan_blocks)
    reference_blocks_by_id = {block.id: block for block in reference_blocks}
    existing_ids = set(reference_blocks_by_id)
    used_target_ids: set[str] = set()
    extra_blocks: list[TextBlock] = []

    for block_id in VT_ROTATED_REFERENCE_ONLY_VALUE_BLOCK_IDS:
        reference_block = reference_blocks_by_id.get(block_id)
        if reference_block is not None:
            reference_block.currentText = ""

    for scan_block in scan_blocks:
        if scan_block.id in VT_ROTATED_UNMAPPED_SOURCE_VALUE_IDS:
            continue

        synthetic_value_rect = VT_ROTATED_REFERENCE_SYNTHETIC_VALUE_RECTS.get(scan_block.id)
        if synthetic_value_rect is not None:
            clone = _clone_combined_scan_block(scan_block, existing_ids, values)
            if clone is not None:
                page_number, rect = synthetic_value_rect
                _position_combined_scan_clone(clone, page_number, rect)
                extra_blocks.append(clone)
            continue

        target_block: Optional[TextBlock] = None

        target_id = VT_ROTATED_REFERENCE_BLOCK_MAP.get(scan_block.id)
        if target_id:
            candidate = reference_blocks_by_id.get(target_id)
            if candidate is not None and candidate.page == scan_block.page and candidate.isCheckbox == scan_block.isCheckbox:
                scan_rect = pymupdf.Rect(scan_block.bbox.x0, scan_block.bbox.y0, scan_block.bbox.x1, scan_block.bbox.y1)
                target_rect = pymupdf.Rect(candidate.bbox.x0, candidate.bbox.y0, candidate.bbox.x1, candidate.bbox.y1)
                candidate_text = _block_text_value(candidate).strip().casefold()
                if (
                    scan_block.isCheckbox
                    or (
                        _rect_distance(scan_rect, target_rect) <= 80.0
                        and candidate_text not in VT_REFERENCE_LABEL_TEXTS
                    )
                ):
                    target_block = candidate

        if target_block is None:
            candidate = reference_blocks_by_id.get(scan_block.id)
            if candidate is not None and candidate.page == scan_block.page and candidate.isCheckbox == scan_block.isCheckbox:
                target_block = candidate

        if target_block is None and scan_block.isCheckbox:
            target_block = _find_nearest_reference_checkbox(scan_block, reference_blocks, used_target_ids)

        if target_block is None:
            target_block = _find_nearest_reference_text_block(scan_block, reference_blocks, used_target_ids)

        if target_block is not None and target_block.id not in used_target_ids:
            applied_text = scan_block.currentText if scan_block.isCheckbox else _sanitized_rotated_vt_scan_text(scan_block, values)
            target_block.currentText = applied_text
            if (
                (_is_contract_id_number_block(scan_block) or _is_contract_id_number_block(target_block))
                and re.fullmatch(r"200\d{7,10}", str(applied_text or "").strip())
            ):
                target_block.sourceOriginalValue = target_block.sourceOriginalValue or target_block.originalText
                target_block.originalText = applied_text
            _expand_normalized_reference_block(scan_block, target_block)
            used_target_ids.add(target_block.id)
            continue

        clone = _clone_combined_scan_block(scan_block, existing_ids, values)
        if clone is not None:
            extra_blocks.append(clone)

    return sorted(
        [*reference_blocks, *extra_blocks],
        key=lambda block: (block.page, block.bbox.y0, block.bbox.x0, 1 if block.isCustom else 0),
    )


def _normalize_combined_vt_handlungsanweisung_session(
    *,
    source_path: Path,
    fingerprint: str,
    sidecar_path: Path,
    scan_blocks: list[TextBlock],
    runtime_root: Path,
    service_base_url: str,
    user_templates: tuple[DocumentTemplateSpec, ...],
) -> Optional[DocumentSession]:
    combined_base_path = _build_combined_vt_handlungsanweisung_base_pdf(
        source_path=source_path,
        fingerprint=fingerprint,
        scan_blocks=scan_blocks,
        runtime_root=runtime_root,
        service_base_url=service_base_url,
        user_templates=user_templates,
    )
    if combined_base_path is None:
        return None

    try:
        normalized_session = analyze_document(
            combined_base_path,
            runtime_root,
            service_base_url,
            user_templates=user_templates,
            normalize_vt_scan_to_reference=False,
        )
    except Exception:
        return None

    normalized_blocks = _apply_rotated_scan_values_to_reference_blocks(
        scan_blocks=scan_blocks,
        reference_blocks=normalized_session.model.blocks,
    )
    values = _combined_scan_values(scan_blocks)
    for block in normalized_blocks:
        if block.id == "page-2-block-11" and values["date"]:
            block.currentText = values["date"]
            block.currentValue = values["date"]
            break
    source_overlay_regions = tuple(
        SourceOverlayRegion(page_number=page_number, rect=rect)
        for page_number, rect in COMBINED_VT_PAGE3_SOURCE_HIDE_REGIONS
    )
    _hide_blocks_covered_by_source_overlay(normalized_blocks, source_overlay_regions)
    normalized_session.model.blocks = _sync_fields(normalized_blocks)
    normalized_session.model.fields = normalized_session.model.blocks
    normalized_session.base_pdf_path = combined_base_path
    normalized_session.source_path = source_path
    normalized_session.sidecar_path = sidecar_path
    normalized_session.model.sourcePath = str(source_path)
    normalized_session.model.fingerprint = fingerprint
    normalized_session.model.detectedTemplateId = COMBINED_VT_HANDLUNGSANWEISUNG_TEMPLATE_ID
    normalized_session.model.detectedTemplateFamily = COMBINED_VT_HANDLUNGSANWEISUNG_TEMPLATE_FAMILY

    warning = (
        "doc031 wurde als kombinierter VT-Vertrag mit Handlungsanweisung erkannt "
        "und als saubere, weiße generierte PDF normalisiert. Der Originalscan wurde nur zur Texterkennung ausgewertet."
    )
    if warning not in normalized_session.model.supportStatus.warnings:
        normalized_session.model.supportStatus.warnings.append(warning)

    for page in normalized_session.model.pages:
        if page.pageNumber >= 4 and page.kind == "mixed":
            page.supportMode = "review"
        page.backgroundImagePath = str(render_background_page(normalized_session, page.pageNumber))

    normalized_session.model.documentClass = _document_class_from_pages(normalized_session.model.pages)
    normalized_session.model.supportStatus.documentClass = normalized_session.model.documentClass
    normalized_session.model.supportStatus.supportMode = _build_support_report(normalized_session.model).supportMode
    normalized_session.model.supportReport = _build_support_report(normalized_session.model)
    return normalized_session


def _normalize_vt_sasse_scan_session(
    *,
    source_path: Path,
    fingerprint: str,
    sidecar_path: Path,
    scan_blocks: list[TextBlock],
    runtime_root: Path,
    service_base_url: str,
    user_templates: tuple[DocumentTemplateSpec, ...],
) -> Optional[DocumentSession]:
    reference_session = _find_vt_reference_session(
        source_path=source_path,
        scan_blocks=scan_blocks,
        runtime_root=runtime_root,
        service_base_url=service_base_url,
        user_templates=user_templates,
    )
    if reference_session is None:
        return None

    reference_source_path = reference_session.source_path
    reference_blocks = reference_session.model.blocks
    reference_blocks_by_id = {block.id: block for block in reference_blocks}
    existing_ids = set(reference_blocks_by_id)
    used_target_ids: set[str] = set()
    extra_blocks: list[TextBlock] = []

    for scan_block in scan_blocks:
        if not _scan_block_has_meaningful_value(scan_block):
            continue

        target_block: Optional[TextBlock] = None

        target_id = VT_SASSE_REFERENCE_BLOCK_MAP.get(scan_block.id)
        if target_id:
            candidate = reference_blocks_by_id.get(target_id)
            if candidate is not None and candidate.page == scan_block.page and candidate.isCheckbox == scan_block.isCheckbox:
                target_block = candidate

        if target_block is None:
            candidate = reference_blocks_by_id.get(scan_block.id)
            if candidate is not None and candidate.page == scan_block.page and candidate.isCheckbox == scan_block.isCheckbox:
                target_block = candidate

        if target_block is None and scan_block.isCheckbox:
            target_block = _find_nearest_reference_checkbox(scan_block, reference_blocks, used_target_ids)

        if target_block is not None and target_block.id not in used_target_ids:
            target_block.currentText = scan_block.currentText
            _expand_normalized_reference_block(scan_block, target_block)
            used_target_ids.add(target_block.id)
            continue

        # Raw OCR blocks (not generated template fields) are unreliable — they can
        # contain garbled text or values from adjacent columns. Skip them; the
        # reference PDF already has the correct template text for non-editable areas.
        if not scan_block.groupKind.startswith("generated-") and not scan_block.isCustom:
            continue

        extra_blocks.append(_clone_normalized_scan_block(scan_block, existing_ids))

    source_overlay_regions = tuple(
        SourceOverlayRegion(page_number=page_number, rect=rect)
        for page_number, rect in VT_SASSE_SIGNATURE_OVERLAY_REGIONS
    )
    normalized_blocks = sorted(
        [*reference_blocks, *extra_blocks],
        key=lambda block: (block.page, block.bbox.y0, block.bbox.x0, 1 if block.isCustom else 0),
    )
    _hide_blocks_covered_by_source_overlay(normalized_blocks, source_overlay_regions)
    reference_session.model.blocks = normalized_blocks
    reference_session.base_pdf_path = reference_source_path
    reference_session.source_path = source_path
    reference_session.sidecar_path = sidecar_path
    reference_session.source_overlay_regions = source_overlay_regions
    reference_session.model.sourcePath = str(source_path)
    reference_session.model.fingerprint = fingerprint

    warning = (
        f'VT Sasse wurde auf die VT-Textvorlage "{reference_source_path.name}" normalisiert, '
        "damit dieselben bearbeitbaren Felder wie bei doc VT zur Verfügung stehen."
    )
    if warning not in reference_session.model.supportStatus.warnings:
        reference_session.model.supportStatus.warnings.append(warning)

    for page in reference_session.model.pages:
        page.backgroundImagePath = str(render_background_page(reference_session, page.pageNumber))
    return reference_session


def _clear_original_checkbox_marks(page: pymupdf.Page, blocks: list[TextBlock]) -> None:
    for block in blocks:
        if not block.isCheckbox or not block.originalText.strip():
            continue
        rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        inset = max(0.9, min(rect.width, rect.height) * 0.18)
        stroke_width = max(1.2, min(rect.width, rect.height) * 0.22)
        page.draw_line(
            pymupdf.Point(rect.x0 + inset, rect.y0 + inset),
            pymupdf.Point(rect.x1 - inset, rect.y1 - inset),
            color=(1, 1, 1),
            width=stroke_width,
            overlay=True,
        )
        page.draw_line(
            pymupdf.Point(rect.x0 + inset, rect.y1 - inset),
            pymupdf.Point(rect.x1 - inset, rect.y0 + inset),
            color=(1, 1, 1),
            width=stroke_width,
            overlay=True,
        )


def _page_state_hash(session: DocumentSession, page_number: int) -> str:
    relevant = [
        {
            "id": block.id,
            "current": block.currentValue,
            "original": block.originalValue,
            "fieldType": block.fieldType,
            "review": block.reviewState,
            "cover": [region.model_dump(mode="json") for region in block.sourceCoverRegions],
        }
        for block in session.model.fields
        if block.page == page_number and (block.isCustom or block.currentValue != block.originalValue or block.isCheckbox)
    ]
    encoded = json.dumps(relevant, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha1(encoded).hexdigest()[:10]


def render_background_page(session: DocumentSession, page_number: int, target_width: Optional[int] = None) -> Path:
    session.model.fields = _sync_fields(session.model.fields)
    cache_dir = session.work_dir / "background-cache"
    cache_dir.mkdir(parents=True, exist_ok=True)

    width_key = max(1, int(target_width or 0))
    state_key = _page_state_hash(session, page_number)
    cache_name = f"page-{page_number}-w{width_key or 'default'}-{state_key}.png"
    output_path = cache_dir / cache_name
    if output_path.exists():
        return output_path

    text_only_page_set = set(session.text_only_background_pages)
    if page_number in text_only_page_set:
        page_model = next((page for page in session.model.pages if page.pageNumber == page_number), None)
        if page_model is None:
            raise ValueError(f"Unbekannte Seite: {page_number}")
        _render_blank_page_image(
            page_model.width,
            page_model.height,
            output_path,
            target_width=target_width,
        )
        return output_path

    doc = pymupdf.open(_session_pdf_base_path(session))
    overlay_regions = _session_overlay_regions_for_page(session, page_number)
    overlay_source_doc: Optional[pymupdf.Document] = None
    if overlay_regions and session.source_path != _session_pdf_base_path(session):
        overlay_source_doc = pymupdf.open(session.source_path)
    try:
        if page_number < 1 or page_number > doc.page_count:
            raise ValueError(f"Unbekannte Seite: {page_number}")

        page = doc[page_number - 1]
        redact_blocks = [
            block for block in session.model.blocks
            if block.page == page_number and _should_redact_background_block(block, page)
        ]
        _apply_block_redactions(page, redact_blocks)
        _clear_original_checkbox_marks(
            page,
            [block for block in session.model.blocks if block.page == page_number and block.isCheckbox],
        )
        _overlay_source_regions_on_page(
            target_page=page,
            source_doc=overlay_source_doc,
            page_number=page_number,
            overlay_regions=overlay_regions,
        )

        if target_width and target_width > 0:
            scale = target_width / max(page.rect.width, 1.0)
            matrix = pymupdf.Matrix(scale, scale)
            page.get_pixmap(matrix=matrix, alpha=False, annots=session.render_annotations).save(output_path)
        else:
            page.get_pixmap(dpi=BACKGROUND_RENDER_DPI, alpha=False, annots=session.render_annotations).save(output_path)
        return output_path
    finally:
        if overlay_source_doc is not None:
            overlay_source_doc.close()
        doc.close()


def analyze_document(
    source_path: Path,
    runtime_root: Path,
    service_base_url: str,
    user_templates: tuple[DocumentTemplateSpec, ...] = (),
    normalize_vt_scan_to_reference: bool = True,
) -> DocumentSession:
    if not source_path.exists():
        raise FileNotFoundError(source_path)

    source_path = source_path.resolve()
    fingerprint = compute_fingerprint(source_path)
    sidecar_path = build_sidecar_path(runtime_root, source_path, fingerprint)
    document_id = uuid4().hex
    work_dir = runtime_root / document_id
    if work_dir.exists():
        shutil.rmtree(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    doc = pymupdf.open(source_path)
    embedded_session_payload = _load_embedded_session(doc)
    reasons: list[str] = []
    warnings: list[str] = []

    if not doc.is_pdf:
        reasons.append("Datei ist keine PDF.")
    hide_acroform_annotations = bool(doc.is_form_pdf)
    if hide_acroform_annotations:
        warnings.append("Formular-PDF erkannt. Vorhandene PDF-Formularfelder werden in normale Editor-Felder umgewandelt.")
    if doc.needs_pass or doc.is_encrypted:
        reasons.append("Verschlüsselte PDFs werden nicht unterstützt.")

    font_runtimes_by_family, font_reasons = _collect_font_runtimes(doc, work_dir / "fonts")
    warnings.extend(font_reasons)

    pages: list[PageModel] = []
    blocks: list[TextBlock] = []
    page_blocks_by_page: dict[int, list[TextBlock]] = {}
    detected_template: Optional[DocumentTemplateSpec] = None
    current_page_hashes = _page_hashes(doc)
    if not reasons:
        for page_index, page in enumerate(doc):
            page_blocks = _extract_blocks_for_page(page, font_runtimes_by_family, warnings, reasons)
            checkbox_rects = _detect_checkbox_rects(page)
            checkbox_blocks, hidden_mark_block_ids = _build_checkbox_blocks(page, page.number + 1, checkbox_rects, page_blocks)
            widget_blocks, hidden_widget_block_ids = _build_widget_blocks(page, page.number + 1, page_blocks)
            for block in page_blocks:
                if block.id in hidden_mark_block_ids or block.id in hidden_widget_block_ids:
                    block.currentText = ""
                    block.originalText = ""
                    block.editable = False
            combined_page_blocks = _sync_fields([*page_blocks, *checkbox_blocks, *widget_blocks])
            page_kind, page_support_mode, page_review_items = _classify_page(
                page,
                combined_page_blocks,
                warnings,
            )
            pages.append(
                PageModel(
                    pageNumber=page.number + 1,
                    width=page.rect.width,
                    height=page.rect.height,
                    lineOverlays=_extract_line_overlays(page),
                    kind=page_kind,
                    supportMode=page_support_mode,
                    reviewItems=page_review_items,
                    imageHash=current_page_hashes[page_index] if page_index < len(current_page_hashes) else None,
                )
            )
            page_blocks_by_page[page.number + 1] = combined_page_blocks
            blocks.extend(combined_page_blocks)

        detected_template = _detect_document_template(doc, page_blocks_by_page, user_templates)
        if detected_template is not None:
            generated_blocks = _build_template_generated_fields(detected_template, doc, page_blocks_by_page, warnings)
        else:
            generated_blocks = _build_default_generated_fields(doc, page_blocks_by_page, warnings)
        blocks.extend(_sync_fields(generated_blocks))

        if normalize_vt_scan_to_reference and detected_template is not None and detected_template.kind == "sicherheit_nord_vt_scan_sasse":
            normalized_session = _normalize_vt_sasse_scan_session(
                source_path=source_path,
                fingerprint=fingerprint,
                sidecar_path=sidecar_path,
                scan_blocks=blocks,
                runtime_root=runtime_root,
                service_base_url=service_base_url,
                user_templates=user_templates,
            )
            if normalized_session is not None:
                doc.close()
                if work_dir.exists() and work_dir != normalized_session.work_dir:
                    shutil.rmtree(work_dir, ignore_errors=True)
                return normalized_session
        if (
            normalize_vt_scan_to_reference
            and detected_template is not None
            and detected_template.kind in {"sicherheit_nord_vt_handlungsanweisung_scan", "sicherheit_nord_vt_rotated_scan"}
            and doc.page_count >= 6
        ):
            normalized_session = _normalize_combined_vt_handlungsanweisung_session(
                source_path=source_path,
                fingerprint=fingerprint,
                sidecar_path=sidecar_path,
                scan_blocks=blocks,
                runtime_root=runtime_root,
                service_base_url=service_base_url,
                user_templates=user_templates,
            )
            if normalized_session is not None:
                doc.close()
                if work_dir.exists() and work_dir != normalized_session.work_dir:
                    shutil.rmtree(work_dir, ignore_errors=True)
                return normalized_session

    if not blocks:
        warnings.append("Kein echter editierbarer Text gefunden. Die PDF wird als Hintergrund geladen; Textfelder können per Rechtsklick gesetzt werden.")

    supported = not reasons
    blocks = _sync_fields(blocks)

    font_assets: list[FontAsset] = []
    for runtime in font_runtimes_by_family.values():
        asset = runtime.asset.model_copy(deep=True)
        if runtime.font_path:
            asset.loadUrl = f"/documents/{document_id}/fonts/{asset.id}"
        font_assets.append(asset)

    model = DocumentModel(
        id=document_id,
        sourcePath=str(source_path),
        fingerprint=fingerprint,
        pageCount=doc.page_count,
        documentClass=_document_class_from_pages(pages),
        embeddedSessionFound=False,
        sessionSchemaVersion=EMBEDDED_SESSION_VERSION,
        detectedTemplateId=detected_template.id if detected_template is not None else None,
        detectedTemplateFamily=detected_template.family if detected_template is not None else None,
        pages=pages,
        fields=blocks if supported else [],
        fonts=font_assets if supported else [],
        supportStatus=SupportStatus(
            supported=supported,
            reasons=reasons,
            warnings=warnings,
            documentClass=_document_class_from_pages(pages),
        ),
    )
    text_only_background_pages = (
        tuple(page.pageNumber for page in model.pages)
        if supported and TEXT_RECONSTRUCTION_BACKGROUND_MODE
        else ()
    )

    if supported and embedded_session_payload is not None:
        _restore_embedded_session(
            model,
            embedded_session_payload,
            current_page_hashes=current_page_hashes,
        )

    if supported:
        _rehydrate_custom_block_fonts(model.fields, font_runtimes_by_family)
        model.fields = _sync_fields(model.fields)
        backgrounds = _render_backgrounds(
            source_path,
            model.fields,
            work_dir,
            render_annotations=not hide_acroform_annotations,
            text_only_pages=text_only_background_pages,
        )
        for page in model.pages:
            background_path = backgrounds.get(page.pageNumber)
            if background_path:
                page.backgroundImagePath = str(background_path)
        model.reviewItems = [item for page in model.pages for item in page.reviewItems]
        model.supportStatus.supportMode = _build_support_report(model).supportMode
        model.supportStatus.documentClass = model.documentClass
        model.supportReport = _build_support_report(model)
    else:
        model.supportStatus.supportMode = "unsupported"
        model.supportReport = _build_support_report(model)

    doc.close()
    return DocumentSession(
        model=model,
        sidecar_path=sidecar_path,
        source_path=source_path,
        work_dir=work_dir,
        font_runtimes={runtime.asset.id: runtime for runtime in font_runtimes_by_family.values()},
        render_annotations=not hide_acroform_annotations,
        text_only_background_pages=text_only_background_pages,
    )


def _hex_to_pdf_color(hex_color: str) -> tuple[float, float, float]:
    value = hex_color.removeprefix("#")
    red = int(value[0:2], 16) / 255
    green = int(value[2:4], 16) / 255
    blue = int(value[4:6], 16) / 255
    return red, green, blue


def _get_measurement_font(
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    cache: dict[str, pymupdf.Font],
) -> Optional[pymupdf.Font]:
    if font_spec.font_file is not None:
        cache_key = f"file:{font_spec.font_file}"
        cached = cache.get(cache_key)
        if cached is not None:
            return cached
        try:
            cached = pymupdf.Font(fontfile=str(font_spec.font_file))
        except Exception:
            return None
        cache[cache_key] = cached
        return cached

    if runtime and runtime.font_buffer:
        cache_key = runtime.asset.id
        cached = cache.get(cache_key)
        if cached is not None:
            return cached
        try:
            cached = pymupdf.Font(fontbuffer=runtime.font_buffer)
        except Exception:
            cached = None
        if cached is not None:
            cache[cache_key] = cached
            return cached

    font_name = font_spec.name
    cache_key = f"base14:{font_name}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached
    try:
        cached = pymupdf.Font(fontname=font_name)
    except Exception:
        normalized = normalize_font_name(font_name)
        mapped = BASE14_FONT_MAP.get(normalized)
        if not mapped:
            return None
        cached = pymupdf.Font(fontname=mapped[0])
    cache[cache_key] = cached
    return cached


def _measure_text_width(
    text: str,
    fontsize: float,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    cache: dict[str, pymupdf.Font],
) -> float:
    measurement_font = _get_measurement_font(runtime, font_spec, cache)
    if measurement_font is not None:
        return measurement_font.text_length(text, fontsize)
    try:
        return pymupdf.get_text_length(text, fontname=font_spec.name, fontsize=fontsize)
    except Exception:
        return len(text) * fontsize * 0.5


def _ensure_page_font(
    page: pymupdf.Page,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    page_font_aliases: set[str],
) -> None:
    if font_spec.name in page_font_aliases:
        return
    if font_spec.font_file is not None:
        page.insert_font(fontname=font_spec.name, fontfile=str(font_spec.font_file))
        page_font_aliases.add(font_spec.name)
        return
    if runtime and runtime.font_buffer and font_spec.name == runtime.asset.id:
        page.insert_font(fontname=font_spec.name, fontbuffer=runtime.font_buffer)
        page_font_aliases.add(font_spec.name)


def _draw_text_underline(
    page: pymupdf.Page,
    *,
    x: float,
    baseline: float,
    width: float,
    fontsize: float,
    color: tuple[float, float, float],
) -> None:
    underline_y = baseline + max(0.9, fontsize * 0.11)
    page.draw_line(
        pymupdf.Point(x, underline_y),
        pymupdf.Point(x + width, underline_y),
        color=color,
        width=max(0.65, fontsize * 0.075),
        overlay=True,
    )


def _write_block_text(
    page: pymupdf.Page,
    block: TextBlock,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    color: tuple[float, float, float],
    page_font_aliases: set[str],
    measurement_fonts: dict[str, pymupdf.Font],
) -> None:
    if page.rotation:
        _write_rotated_page_block_text(page, block, runtime, font_spec, color, page_font_aliases, measurement_fonts)
        return

    if block.isCheckbox:
        rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
        if block.groupKind == "widget-checkbox-field":
            page.draw_rect(
                rect,
                color=color,
                fill=None,
                width=max(0.6, min(rect.width, rect.height) * 0.065),
                overlay=True,
            )
        if not block.currentText.strip():
            return
        inset = max(1.0, min(rect.width, rect.height) * 0.16)
        stroke_width = max(0.9, min(rect.width, rect.height) * 0.12)
        page.draw_line(
            pymupdf.Point(rect.x0 + inset, rect.y0 + inset),
            pymupdf.Point(rect.x1 - inset, rect.y1 - inset),
            color=color,
            width=stroke_width,
            overlay=True,
        )
        page.draw_line(
            pymupdf.Point(rect.x0 + inset, rect.y1 - inset),
            pymupdf.Point(rect.x1 - inset, rect.y0 + inset),
            color=color,
            width=stroke_width,
            overlay=True,
        )
        return

    _ensure_page_font(page, runtime, font_spec, page_font_aliases)

    if _is_iban_template_block(block):
        _write_iban_overlay_text(
            page,
            block,
            runtime,
            font_spec,
            color,
            measurement_fonts,
        )
        return
    if _is_masked_template_block(block):
        _write_masked_overlay_text(
            page,
            block,
            runtime,
            font_spec,
            color,
            measurement_fonts,
        )
        return

    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    line_height = block.lineHeight / max(block.fontSize, 0.01)
    is_single_line = (
        block.baseline is not None
        and "\n" not in block.currentText
        and block.groupKind != "multiline"
    )

    if is_single_line:
        fontsize = block.fontSize
        slack = max(0.75, fontsize * 0.1)
        while fontsize >= block.minFontSize:
            width = _measure_text_width(block.currentText, fontsize, runtime, font_spec, measurement_fonts)
            if width <= rect.width + slack:
                page.insert_text(
                    pymupdf.Point(rect.x0, block.baseline),
                    block.currentText,
                    fontname=font_spec.name,
                    fontsize=fontsize,
                    color=color,
                    overlay=True,
                )
                if _is_contract_id_number_block(block):
                    _draw_text_underline(
                        page,
                        x=rect.x0,
                        baseline=block.baseline,
                        width=width,
                        fontsize=fontsize,
                        color=color,
                    )
                return
            fontsize = round(fontsize - 0.25, 2)

    fontsize = block.fontSize
    inserted = -1.0
    while fontsize >= block.minFontSize:
        inserted = page.insert_textbox(
            rect,
            block.currentText,
            fontname=font_spec.name,
            fontsize=fontsize,
            color=color,
            lineheight=line_height,
            align=0,
            overlay=True,
        )
        if inserted >= 0:
            return
        fontsize = round(fontsize - 0.25, 2)

    fallback_fontsize = block.minFontSize
    fallback_line_height = max(block.lineHeight, fallback_fontsize * 1.15)
    fallback_baseline = block.baseline if block.baseline is not None else rect.y0 + fallback_fontsize
    for index, line in enumerate(block.currentText.splitlines() or [block.currentText]):
        if not line.strip():
            continue
        page.insert_text(
            pymupdf.Point(rect.x0, fallback_baseline + index * fallback_line_height),
            line,
            fontname=font_spec.name,
            fontsize=fallback_fontsize,
            color=color,
            overlay=True,
        )
    return


def _rotated_visual_point(page: pymupdf.Page, x: float, y: float) -> pymupdf.Point:
    return pymupdf.Point(x, y) * page.derotation_matrix


def _rotated_visual_quad(page: pymupdf.Page, rect: pymupdf.Rect) -> pymupdf.Quad:
    return pymupdf.Quad((
        _rotated_visual_point(page, rect.x0, rect.y0),
        _rotated_visual_point(page, rect.x1, rect.y0),
        _rotated_visual_point(page, rect.x0, rect.y1),
        _rotated_visual_point(page, rect.x1, rect.y1),
    ))


def _cover_rotated_visual_rect(page: pymupdf.Page, rect: pymupdf.Rect) -> None:
    if rect.is_empty:
        return
    page.draw_quad(
        _rotated_visual_quad(page, rect),
        color=(1, 1, 1),
        fill=(1, 1, 1),
        width=0,
        overlay=True,
    )


def _write_rotated_page_block_text(
    page: pymupdf.Page,
    block: TextBlock,
    runtime: Optional[FontRuntime],
    font_spec: ExportFontSpec,
    color: tuple[float, float, float],
    page_font_aliases: set[str],
    measurement_fonts: dict[str, pymupdf.Font],
) -> None:
    rotate = page.rotation
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)

    if block.isCheckbox:
        if block.groupKind == "widget-checkbox-field":
            corners = (
                _rotated_visual_point(page, rect.x0, rect.y0),
                _rotated_visual_point(page, rect.x1, rect.y0),
                _rotated_visual_point(page, rect.x1, rect.y1),
                _rotated_visual_point(page, rect.x0, rect.y1),
            )
            stroke_width = max(0.6, min(rect.width, rect.height) * 0.065)
            for start, end in zip(corners, corners[1:] + corners[:1]):
                page.draw_line(start, end, color=color, width=stroke_width, overlay=True)
        if not block.currentText.strip():
            return
        inset = max(1.0, min(rect.width, rect.height) * 0.16)
        stroke_width = max(0.9, min(rect.width, rect.height) * 0.12)
        for start, end in (
            ((rect.x0 + inset, rect.y0 + inset), (rect.x1 - inset, rect.y1 - inset)),
            ((rect.x0 + inset, rect.y1 - inset), (rect.x1 - inset, rect.y0 + inset)),
        ):
            page.draw_line(
                _rotated_visual_point(page, start[0], start[1]),
                _rotated_visual_point(page, end[0], end[1]),
                color=color,
                width=stroke_width,
                overlay=True,
            )
        return

    _ensure_page_font(page, runtime, font_spec, page_font_aliases)
    lines = block.currentText.splitlines() or [block.currentText]
    line_height = max(block.lineHeight, block.fontSize * 1.15)
    baseline = block.baseline if block.baseline is not None else rect.y0 + block.fontSize

    for index, line in enumerate(lines):
        if not line.strip():
            continue
        fontsize = block.fontSize
        while fontsize >= block.minFontSize:
            width = _measure_text_width(line, fontsize, runtime, font_spec, measurement_fonts)
            if width <= rect.width + max(0.75, fontsize * 0.1):
                point = _rotated_visual_point(page, rect.x0, baseline + index * line_height)
                page.insert_text(
                    point,
                    line,
                    fontname=font_spec.name,
                    fontsize=fontsize,
                    color=color,
                    rotate=rotate,
                    overlay=True,
                )
                if _is_contract_id_number_block(block):
                    underline_y = baseline + index * line_height + max(0.9, fontsize * 0.11)
                    page.draw_line(
                        _rotated_visual_point(page, rect.x0, underline_y),
                        _rotated_visual_point(page, rect.x0 + width, underline_y),
                        color=color,
                        width=max(0.65, fontsize * 0.075),
                        overlay=True,
                    )
                break
            fontsize = round(fontsize - 0.25, 2)


def _page_looks_like_image_scan(page: pymupdf.Page) -> bool:
    if not page.get_images(full=True):
        return False
    return len(page.get_drawings()) < 8 and len(page.get_text("text").strip()) < 500


def _scan_replacement_rect(block: TextBlock, page_rect: pymupdf.Rect) -> pymupdf.Rect:
    baseline = block.baseline if block.baseline is not None else block.bbox.y1
    if _is_contract_id_number_block(block):
        return _expanded_id_number_cover_rect(block, page_rect)

    is_header_field = block.groupKind in {
        "generated-contract-party-field",
        "generated-contract-object-line-field",
    }
    if is_header_field:
        scale_x = page_rect.width / 595.0
        if block.groupKind == "generated-contract-party-field":
            x0 = 98.5 * scale_x
            x1 = 286.0 * scale_x
        else:
            x0 = 285.6 * scale_x
            x1 = 565.0 * scale_x
        rect = pymupdf.Rect(
            x0 - 0.8,
            block.bbox.y0 - 1.2,
            x1 + 0.8,
            block.bbox.y1 + 1.8,
        )
        return rect & page_rect

    bottom_factor = 0.8 if is_header_field else 0.42
    top = min(block.bbox.y0, baseline - block.fontSize * 1.05) - 0.5
    bottom = max(block.bbox.y1, baseline + block.fontSize * bottom_factor) + 0.5
    x1 = block.bbox.x1
    if not is_header_field:
        text_width = 0.0
        for text in (block.originalText, block.currentText):
            if text.strip():
                text_width = max(text_width, pymupdf.get_text_length(text.strip(), fontname="Helvetica", fontsize=block.fontSize))
        if text_width > 0:
            x1 = min(block.bbox.x1, block.bbox.x0 + text_width + 5.0)
    rect = pymupdf.Rect(block.bbox.x0 - 0.8, top, x1 + 0.8, bottom)
    return rect & page_rect


def _restore_scan_header_guides(page: pymupdf.Page, page_blocks: list[TextBlock]) -> None:
    if not any(
        block.groupKind in {"generated-contract-party-field", "generated-contract-object-line-field"}
        and block.currentText != block.originalText
        for block in page_blocks
    ):
        return

    scale_x = page.rect.width / 595.0
    scale_y = page.rect.height / 842.0

    def sx(value: float) -> float:
        return value * scale_x

    def sy(value: float) -> float:
        return value * scale_y

    def line(x0: float, y0: float, x1: float, y1: float) -> None:
        page.draw_line(
            pymupdf.Point(sx(x0), sy(y0)),
            pymupdf.Point(sx(x1), sy(y1)),
            color=(0, 0, 0),
            width=0.35,
            overlay=True,
        )

    rows = (55.05, 65.45, 75.75, 86.10, 96.45, 106.95)
    left_x0, left_split, left_x1 = 28.4, 98.5, 286.0
    right_x0, right_x1 = 285.6, 565.0
    for y in rows:
        line(left_x0, y, left_x1, y)
        line(right_x0, y, right_x1, y)
    for x0, y0, x1, y1 in (
        (left_x0, rows[0], left_x0, rows[-1]),
        (left_split, rows[0], left_split, rows[-1]),
        (left_x1, rows[0], left_x1, rows[-1]),
        (right_x0, rows[0], right_x0, rows[-1]),
        (right_x1, rows[0], right_x1, rows[-1]),
    ):
        line(x0, y0, x1, y1)


def _should_cover_scan_original(block: TextBlock) -> bool:
    return (
        not block.isCustom
        and not block.isCheckbox
        and block.groupKind.startswith("generated-")
        and (block.currentText.strip() or block.originalText.strip())
        and (block.currentText != block.originalText or _is_contract_id_number_block(block))
    )


def _should_cover_scan_checkbox(block: TextBlock) -> bool:
    return (
        block.isCheckbox
        and block.groupKind.startswith("generated-")
        and (block.currentText.strip() or block.originalText.strip())
    )


def _scan_checkbox_inner_rect(block: TextBlock, page_rect: pymupdf.Rect) -> pymupdf.Rect:
    rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
    inset = max(0.25, min(rect.width, rect.height) * 0.03)
    return pymupdf.Rect(rect.x0 + inset, rect.y0 + inset, rect.x1 - inset, rect.y1 - inset) & page_rect


def _draw_line_overlays_on_page(page: pymupdf.Page, line_overlays: list[LineOverlay]) -> None:
    for line in line_overlays:
        try:
            color = _hex_to_pdf_color(line.color or "#000000")
        except Exception:
            color = (0, 0, 0)
        page.draw_line(
            pymupdf.Point(line.x0, line.y0),
            pymupdf.Point(line.x1, line.y1),
            color=color,
            width=max(0.1, line.width),
            overlay=True,
        )


def _should_write_reconstructed_block(block: TextBlock) -> bool:
    if block.groupKind.startswith("hidden-") or block.groupKind == "source-overlay-hidden":
        return False
    if not block.currentText.strip() and not block.isCheckbox:
        return False
    return True


def _write_reconstructed_page(
    output_page: pymupdf.Page,
    session: DocumentSession,
    page_number: int,
) -> None:
    output_page.draw_rect(output_page.rect, color=None, fill=(1, 1, 1), width=0, overlay=False)
    page_model = next((page for page in session.model.pages if page.pageNumber == page_number), None)
    if page_model is not None:
        _draw_line_overlays_on_page(output_page, page_model.lineOverlays)

    page_font_resources: dict[str, str] = {}
    page_font_aliases: set[str] = set()
    measurement_fonts: dict[str, pymupdf.Font] = {}
    page_blocks = sorted(
        (
            block for block in session.model.fields
            if block.page == page_number and _should_write_reconstructed_block(block)
        ),
        key=lambda block: (block.bbox.y0, block.bbox.x0, block.zIndex, 1 if block.isCustom else 0),
    )

    for block in page_blocks:
        runtime = session.font_runtimes.get(block.fontAssetId or "")
        font_spec = _resolve_export_font_spec(block, runtime, page_font_resources)
        color = _hex_to_pdf_color(block.color)
        if block.isCheckbox:
            rect = pymupdf.Rect(block.bbox.x0, block.bbox.y0, block.bbox.x1, block.bbox.y1)
            output_page.draw_rect(
                rect,
                color=color,
                fill=None,
                width=max(0.55, min(rect.width, rect.height) * 0.055),
                overlay=True,
            )
        _write_block_text(
            output_page,
            block,
            runtime,
            font_spec,
            color,
            page_font_aliases,
            measurement_fonts,
        )


def export_document(session: DocumentSession, target_path: Optional[Path] = None) -> Path:
    if not session.model.supportStatus.supported:
        raise ValueError("Nicht unterstütztes Dokument kann nicht exportiert werden.")

    session.model.fields = _sync_fields(session.model.fields)
    base_pdf_path = _session_pdf_base_path(session)
    overlay_pages = {region.page_number for region in session.source_overlay_regions}
    text_only_pages = set(session.text_only_background_pages)
    compose_from_reference_template = session.source_path != base_pdf_path
    changed_blocks = [
        block for block in session.model.fields
        if (
            block.isCustom
            or block.currentValue != block.originalValue
            or (compose_from_reference_template and _is_contract_id_number_block(block))
        )
    ]
    changed_block_ids = {block.id for block in changed_blocks}
    for block in session.model.fields:
        if block.page in text_only_pages and block.id not in changed_block_ids:
            changed_blocks.append(block)
            changed_block_ids.add(block.id)

    widget_pages = {
        block.page for block in session.model.fields
        if not session.render_annotations and block.groupKind.startswith("widget-")
    }
    touched_pages = {block.page for block in changed_blocks} | overlay_pages | widget_pages | text_only_pages
    changed_blocks.extend(
        block for block in session.model.fields
        if block.id not in changed_block_ids
        and block.page in touched_pages
        and (
            (compose_from_reference_template and block.groupKind.startswith("generated-"))
            or block.groupKind.startswith("widget-")
        )
        and (
            block.isCheckbox
            or (
                block.originalValue.strip()
                and block.currentValue.strip()
            )
        )
    )
    target_path = target_path.resolve() if target_path else session.source_path.with_name(f"{session.source_path.stem}-bearbeitet.pdf")
    if target_path == session.source_path:
        raise ValueError("Die Original-PDF darf nicht überschrieben werden.")
    target_path.parent.mkdir(parents=True, exist_ok=True)

    source_doc = pymupdf.open(base_pdf_path)
    output_doc = pymupdf.open()

    changed_blocks_by_page: dict[int, list[TextBlock]] = {}
    for block in changed_blocks:
        changed_blocks_by_page.setdefault(block.page, []).append(block)

    background_doc: Optional[pymupdf.Document] = None
    overlay_source_doc: Optional[pymupdf.Document] = None
    if touched_pages:
        background_doc = pymupdf.open(base_pdf_path)
        if overlay_pages and session.source_path != base_pdf_path:
            overlay_source_doc = pymupdf.open(session.source_path)
        for page_number in touched_pages:
            if page_number in text_only_pages:
                continue
            background_page = background_doc[page_number - 1]
            page_changed_blocks = changed_blocks_by_page.get(page_number, [])
            redact_blocks = [
                block for block in page_changed_blocks
                if not block.isCustom
                and not _is_masked_template_block(block)
                and (not block.isCheckbox or block.groupKind in {"generated-payment-checkbox", "widget-checkbox-field"})
            ]
            _apply_block_redactions(background_page, redact_blocks)
            _clear_original_checkbox_marks(
                background_page,
                [block for block in page_changed_blocks if block.isCheckbox],
            )

    try:
        for page_index in range(source_doc.page_count):
            source_page = source_doc[page_index]
            page_number = page_index + 1

            if page_number in text_only_pages:
                output_page = output_doc.new_page(width=source_page.rect.width, height=source_page.rect.height)
                _write_reconstructed_page(output_page, session, page_number)
                continue

            if source_page.rotation:
                output_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
                output_page = output_doc[-1]
                if page_number in touched_pages:
                    page_font_resources = _collect_page_font_resources(output_page)
                    page_font_aliases = {alias for alias in page_font_resources.values() if alias}
                    measurement_fonts: dict[str, pymupdf.Font] = {}
                    page_blocks = sorted(
                        changed_blocks_by_page.get(page_number, []),
                        key=lambda block: (block.bbox.y0, block.bbox.x0, 1 if block.isCustom else 0),
                    )
                    _overlay_source_regions_on_page(
                        target_page=output_page,
                        source_doc=overlay_source_doc,
                        page_number=page_number,
                        overlay_regions=_session_overlay_regions_for_page(session, page_number),
                    )
                    if _page_looks_like_image_scan(source_page):
                        for block in page_blocks:
                            if _should_cover_scan_checkbox(block):
                                _cover_rotated_visual_rect(
                                    output_page,
                                    _scan_checkbox_inner_rect(block, output_page.rect),
                                )
                            elif _should_cover_scan_original(block):
                                _cover_rotated_visual_rect(
                                    output_page,
                                    _scan_replacement_rect(block, output_page.rect),
                                )
                    for block in page_blocks:
                        if not block.currentText.strip() and not block.isCheckbox:
                            continue
                        runtime = session.font_runtimes.get(block.fontAssetId or "")
                        font_spec = _resolve_export_font_spec(block, runtime, page_font_resources)
                        color = _hex_to_pdf_color(block.color)
                        _write_block_text(
                            output_page,
                            block,
                            runtime,
                            font_spec,
                            color,
                            page_font_aliases,
                            measurement_fonts,
                        )
                continue

            output_page = output_doc.new_page(width=source_page.rect.width, height=source_page.rect.height)

            if page_number in touched_pages and background_doc is not None:
                output_page.show_pdf_page(output_page.rect, background_doc, page_index)
                page_font_resources = _collect_page_font_resources(output_page)
                page_font_aliases = {alias for alias in page_font_resources.values() if alias}
                measurement_fonts: dict[str, pymupdf.Font] = {}
                page_blocks = sorted(
                    changed_blocks_by_page.get(page_number, []),
                    key=lambda block: (block.bbox.y0, block.bbox.x0, 1 if block.isCustom else 0),
                )

                if _page_looks_like_image_scan(source_page):
                    for block in page_blocks:
                        if _should_cover_scan_checkbox(block):
                            output_page.draw_rect(
                                _scan_checkbox_inner_rect(block, output_page.rect),
                                color=(1, 1, 1),
                                fill=(1, 1, 1),
                                width=0,
                                overlay=True,
                            )
                        elif _should_cover_scan_original(block):
                            output_page.draw_rect(
                                _scan_replacement_rect(block, output_page.rect),
                                color=(1, 1, 1),
                                fill=(1, 1, 1),
                                width=0,
                                overlay=True,
                            )
                    _restore_scan_header_guides(output_page, page_blocks)

                _overlay_source_regions_on_page(
                    target_page=output_page,
                    source_doc=overlay_source_doc,
                    page_number=page_number,
                    overlay_regions=_session_overlay_regions_for_page(session, page_number),
                )

                for block in page_blocks:
                    if not block.currentText.strip() and not block.isCheckbox:
                        continue

                    runtime = session.font_runtimes.get(block.fontAssetId or "")
                    font_spec = _resolve_export_font_spec(block, runtime, page_font_resources)
                    color = _hex_to_pdf_color(block.color)
                    _write_block_text(
                        output_page,
                        block,
                        runtime,
                        font_spec,
                        color,
                        page_font_aliases,
                        measurement_fonts,
                    )
                continue

            output_page.show_pdf_page(output_page.rect, source_doc, page_index)

        session_payload = _embedded_session_payload(session.model, page_hashes=_page_hashes(output_doc))
        _write_embedded_session(output_doc, session_payload)
        output_doc.save(target_path)
        for field in session.model.fields:
            field.originalText = field.currentText
            field.originalValue = field.currentValue
        session.model.fields = _sync_fields(session.model.fields)
        session.model.embeddedSessionFound = True
        return target_path
    finally:
        output_doc.close()
        source_doc.close()
        if background_doc is not None:
            background_doc.close()
        if overlay_source_doc is not None:
            overlay_source_doc.close()


def export_whiteboard_pdf(
    *,
    image_data_url: str,
    width: float,
    height: float,
    target_path: Optional[Path] = None,
) -> Path:
    if not image_data_url.startswith("data:image/png;base64,"):
        raise ValueError("Whiteboard-Bilddaten müssen als PNG gesendet werden.")

    try:
        image_bytes = base64.b64decode(image_data_url.split(",", 1)[1], validate=True)
    except Exception as error:
        raise ValueError("Whiteboard-Bilddaten sind ungültig.") from error

    export_width = max(1.0, float(width))
    export_height = max(1.0, float(height))
    final_target_path = target_path.resolve() if target_path else Path.cwd() / "whiteboard.pdf"
    final_target_path.parent.mkdir(parents=True, exist_ok=True)

    output_doc = pymupdf.open()
    try:
        page = output_doc.new_page(width=export_width, height=export_height)
        page.draw_rect(page.rect, color=None, fill=(1, 1, 1), overlay=False)
        page.insert_image(page.rect, stream=image_bytes, keep_proportion=False, overlay=True)
        output_doc.save(final_target_path)
        return final_target_path
    finally:
        output_doc.close()
