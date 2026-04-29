from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class TemplateMarkerSpec:
    page_number: int
    needle: str


@dataclass(frozen=True)
class TemplatePageSizeSpec:
    page_number: int
    width: float
    height: float


@dataclass(frozen=True)
class TemplatePageImageHashSpec:
    page_number: int
    hash_hex: str
    max_distance: int = 24


@dataclass(frozen=True)
class TemplateFieldSpec:
    page_number: int
    source_page_width: float
    source_page_height: float
    x0: float
    y0: float
    x1: float
    y1: float
    font_family: str
    font_key: str
    font_size: float
    color: str
    line_height: float
    align: str
    rotation: float
    min_font_size: float
    css_font_family: str
    font_asset_id: Optional[str] = None
    font_weight: str = "400"
    font_style: str = "normal"
    baseline: Optional[float] = None
    is_checkbox: bool = False
    group_kind: str = "generated-user-template-field"


@dataclass(frozen=True)
class DocumentTemplateSpec:
    id: str
    family: str
    kind: str
    description: str
    display_name: Optional[str] = None
    page_generator_ids: tuple[str, ...] = ()
    layout_generator_id: Optional[str] = None
    page_count: Optional[int] = None
    page_sizes: tuple[TemplatePageSizeSpec, ...] = ()
    required_document_markers: tuple[str, ...] = ()
    required_page_markers: tuple[TemplateMarkerSpec, ...] = ()
    minimum_marker_match_count: int = 0
    minimum_marker_match_ratio: float = 1.0
    aggregate_text_markers_any: tuple[str, ...] = ()
    allow_empty_aggregate_text: bool = False
    match_mode: str = "markers"
    page_image_hashes: tuple[TemplatePageImageHashSpec, ...] = ()
    learned_field_specs: tuple[TemplateFieldSpec, ...] = ()
    warning: Optional[str] = None


VT_PAGE_GENERATORS = (
    "contract-party-fields",
    "additional-agreement-fields",
    "payment-frequency-checkboxes",
)


DOCUMENT_TEMPLATES: tuple[DocumentTemplateSpec, ...] = (
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_bma_fw_text",
        family="sicherheit_nord_vt_contract",
        kind="sicherheit_nord_vt_text",
        description="Sicherheit-Nord Vertragslayout mit BMA/Feuerwehr-Variante",
        page_generator_ids=VT_PAGE_GENERATORS,
        layout_generator_id="sicherheit_nord_text_layout",
        required_document_markers=(
            "Dienstleistungsvertrag Notruf- und Serviceleitstelle",
            "Brandmeldeanlage",
        ),
        required_page_markers=(
            TemplateMarkerSpec(2, "SEPA LASTSCHRIFTERM"),
            TemplateMarkerSpec(3, "Gewünschte Zahlungsweise"),
        ),
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_text",
        family="sicherheit_nord_vt_contract",
        kind="sicherheit_nord_vt_text",
        description="Sicherheit-Nord Vertragslayout mit editierbarem PDF-Text",
        page_generator_ids=VT_PAGE_GENERATORS,
        layout_generator_id="sicherheit_nord_text_layout",
        required_document_markers=(
            "Dienstleistungsvertrag Notruf- und Serviceleitstelle",
        ),
        required_page_markers=(
            TemplateMarkerSpec(2, "SEPA LASTSCHRIFTERM"),
            TemplateMarkerSpec(3, "Gewünschte Zahlungsweise"),
        ),
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_scan_sasse",
        family="sicherheit_nord_vt_contract",
        kind="sicherheit_nord_vt_scan_sasse",
        description="Spezifisches Sicherheit-Nord Scanlayout für VT Sasse",
        layout_generator_id="sicherheit_nord_scan_sasse_layout",
        match_mode="image",
        page_count=3,
        page_sizes=(
            TemplatePageSizeSpec(1, 595.0, 842.0),
            TemplatePageSizeSpec(2, 595.0, 842.0),
            TemplatePageSizeSpec(3, 595.0, 842.0),
        ),
        page_image_hashes=(
            TemplatePageImageHashSpec(1, "8079811f803f8107c3f7c03f81dfc7ffc7f7c07fc011c001c01fc01f803f9fff", 6),
            TemplatePageImageHashSpec(2, "8039811f8fff800b801f80018000800fc2ff83ff800fdfffcfff8003ffff9fff", 6),
            TemplatePageImageHashSpec(3, "8031800184a380038007dfffffffffff9f218381ffff8fffffffffffffff8eff", 6),
        ),
        warning="VT Sasse als gescannte Spezialvorlage erkannt. Alle sichtbaren Vertragsfelder wurden als editierbare Overlay-Felder geladen.",
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_handlungsanweisung_scan_9696",
        family="sicherheit_nord_vt_handlungsanweisung",
        kind="sicherheit_nord_vt_handlungsanweisung_scan",
        description="Gedrehter Sicherheit-Nord Kombi-Scan aus VT-Vertrag und Handlungsanweisung 9696",
        layout_generator_id="sicherheit_nord_rotated_scan_layout",
        match_mode="image",
        page_count=6,
        page_sizes=(
            TemplatePageSizeSpec(1, 595.0, 842.0),
            TemplatePageSizeSpec(2, 595.0, 842.0),
            TemplatePageSizeSpec(3, 595.0, 842.0),
            TemplatePageSizeSpec(4, 595.0, 842.0),
            TemplatePageSizeSpec(5, 595.0, 842.0),
            TemplatePageSizeSpec(6, 595.0, 842.0),
        ),
        page_image_hashes=(
            TemplatePageImageHashSpec(1, "fffb833f873fc007c3a7c03fc1dfc7ffc13fc09fc001c01fc01fc03fcfffdfff", 48),
            TemplatePageImageHashSpec(2, "fff1833f87bfc01fc01f8001c001c1ffc01f800fc0ffcfffc003ffffffff9fff", 48),
            TemplatePageImageHashSpec(3, "803180198697c0038003c003ffffffff93078103ff838fffffffffffffff8eff", 56),
            TemplatePageImageHashSpec(4, "dfff8071803f807f80ff83ff87ff8fffbfff81ff81ff83ff81ff87ff87ffffff", 56),
            TemplatePageImageHashSpec(5, "dfff800187ff81ff83ff83ffc7ff80038003807f81ff87ff87ff80ff87fff7ff", 56),
            TemplatePageImageHashSpec(6, "9fff800183ff804383ff87ff80ff80ff80ffffffffff9fff987fffffffffffff", 56),
        ),
        warning="Gedrehter Sicherheit-Nord Kombi-Scan erkannt. Der Vertrag wird auf eine saubere VT-/Handlungsanweisungs-Vorlage normalisiert; OCR dient nur zur Wertübernahme.",
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_rotated_scan_9696",
        family="sicherheit_nord_vt_contract",
        kind="sicherheit_nord_vt_rotated_scan",
        description="Gedrehtes Sicherheit-Nord VT-Scanlayout 9696",
        layout_generator_id="sicherheit_nord_rotated_scan_layout",
        match_mode="image",
        page_sizes=(
            TemplatePageSizeSpec(1, 595.0, 842.0),
            TemplatePageSizeSpec(2, 595.0, 842.0),
            TemplatePageSizeSpec(3, 595.0, 842.0),
        ),
        page_image_hashes=(
            TemplatePageImageHashSpec(1, "fffb833f873fc007c3a7c03fc1dfc7ffc13fc09fc001c01fc01fc03fcfffdfff", 48),
            TemplatePageImageHashSpec(2, "fff1833f87bfc01fc01f8001c001c1ffc01f800fc0ffcfffc003ffffffff9fff", 48),
            TemplatePageImageHashSpec(3, "803180198697c0038003c003ffffffff93078103ff838fffffffffffffff8eff", 56),
        ),
        warning="Gedrehtes Sicherheit-Nord VT-Scanlayout erkannt. Die bekannten Vertragsfelder wurden als editierbare Overlay-Felder geladen; vorhandene Scanwerte bleiben sichtbar, bis ein Feld geändert wird.",
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_vt_scan",
        family="sicherheit_nord_vt_contract",
        kind="sicherheit_nord_vt_scan",
        description="Sicherheit-Nord Vertragslayout als aufrechter Scan",
        layout_generator_id="sicherheit_nord_scan_layout",
        aggregate_text_markers_any=("sasse", "buergschaft", "bürgschaft", "sicherheit nord"),
        warning="Gescanntes Sicherheit-Nord-Layout erkannt. Standardfelder wurden anhand der sichtbaren Seitenpositionen erzeugt.",
    ),
    DocumentTemplateSpec(
        id="rotated_scan_manual",
        family="image_scan_manual",
        kind="rotated_image_scan_manual",
        description="Rotiertes Scan-PDF ohne verlässliche editierbare Texte",
        warning="Rotiertes Scan-PDF erkannt. Die Seiten werden als Hintergrund geladen; Textfelder können manuell gesetzt oder später als neue Vorlage gespeichert werden.",
    ),
    DocumentTemplateSpec(
        id="sicherheit_nord_handlungsanweisung",
        family="sicherheit_nord_handlungsanweisung",
        kind="handlungsanweisung_text",
        description="Handlungsanweisung für den Alarm- und Interventionsdienst",
        required_document_markers=(
            "Handlungsanweisung für den Alarm- und Interventionsdienst",
        ),
    ),
)
