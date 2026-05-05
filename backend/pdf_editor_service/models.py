from __future__ import annotations

from typing import Any, Optional

from pydantic import AliasChoices, BaseModel, ConfigDict, Field, model_validator


class BoundingBox(BaseModel):
    x0: float
    y0: float
    x1: float
    y1: float


class FieldQuad(BaseModel):
    x0: float
    y0: float
    x1: float
    y1: float
    x2: float
    y2: float
    x3: float
    y3: float


class ReviewItem(BaseModel):
    severity: str = "warning"
    message: str
    code: Optional[str] = None
    page: Optional[int] = None
    fieldId: Optional[str] = None


class SupportStatus(BaseModel):
    supported: bool
    reasons: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    supportMode: str = "exact"
    documentClass: Optional[str] = None


class FontAsset(BaseModel):
    id: str
    family: str
    cssFamily: str
    extension: Optional[str] = None
    embedded: bool = False
    loadUrl: Optional[str] = None


class LineOverlay(BaseModel):
    x0: float
    y0: float
    x1: float
    y1: float
    width: float
    color: str


class FieldAppearance(BaseModel):
    fontFamily: str
    fontKey: str
    fontSize: float
    color: str
    lineHeight: float
    align: str
    rotation: float
    cssFontFamily: str
    fontAssetId: Optional[str] = None
    fontWeight: str = "400"
    fontStyle: str = "normal"
    textDecoration: str = "none"
    baseline: Optional[float] = None
    minFontSize: float = 6.0


class PageModel(BaseModel):
    pageNumber: int
    width: float
    height: float
    backgroundImagePath: Optional[str] = None
    lineOverlays: list[LineOverlay] = Field(default_factory=list)
    kind: str = "native-digital"
    supportMode: str = "exact"
    reviewItems: list[ReviewItem] = Field(default_factory=list)
    imageHash: Optional[str] = None


class EditableField(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="ignore")

    id: str
    page: int
    fieldType: str = "text-line"
    rect: Optional[BoundingBox] = None
    bbox: Optional[BoundingBox] = None
    quads: list[FieldQuad] = Field(default_factory=list)
    sourceType: str = "native-digital"
    appearance: Optional[FieldAppearance] = None
    originalValue: str = ""
    currentValue: str = ""
    sourceOriginalValue: Optional[str] = None
    sourceCoverRegions: list[BoundingBox] = Field(default_factory=list)
    fontResourceRef: Optional[str] = None
    confidence: float = 1.0
    reviewState: str = "exact"
    supportMode: str = "exact"
    zIndex: int = 0
    ocrTranscript: Optional[str] = None
    inkPayload: Optional[dict[str, Any]] = None
    widgetFieldName: Optional[str] = None
    widgetFieldType: Optional[str] = None
    widgetXref: Optional[int] = None
    originalText: str = ""
    currentText: str = ""
    fontFamily: str
    fontKey: str
    fontSize: float
    color: str
    lineHeight: float
    align: str
    rotation: float
    groupKind: str
    minFontSize: float
    editable: bool
    cssFontFamily: str
    fontAssetId: Optional[str] = None
    fontWeight: str = "400"
    fontStyle: str = "normal"
    textDecoration: str = "none"
    baseline: Optional[float] = None
    imageDataUrl: Optional[str] = None
    imageWidth: Optional[float] = None
    imageHeight: Optional[float] = None
    isCheckbox: bool = False
    isCustom: bool = False

    @model_validator(mode="after")
    def sync_compatibility_fields(self) -> "EditableField":
        if self.bbox is None and self.rect is not None:
            self.bbox = self.rect.model_copy(deep=True)
        if self.rect is None and self.bbox is not None:
            self.rect = self.bbox.model_copy(deep=True)

        if not self.originalText and self.originalValue:
            self.originalText = self.originalValue
        if not self.currentText and self.currentValue:
            self.currentText = self.currentValue
        if not self.originalValue and self.originalText:
            self.originalValue = self.originalText
        if not self.currentValue and self.currentText:
            self.currentValue = self.currentText

        if self.fieldType in {"checkbox", "radio"}:
            self.isCheckbox = True

        if self.appearance is None:
            self.appearance = FieldAppearance(
                fontFamily=self.fontFamily,
                fontKey=self.fontKey,
                fontSize=self.fontSize,
                color=self.color,
                lineHeight=self.lineHeight,
                align=self.align,
                rotation=self.rotation,
                cssFontFamily=self.cssFontFamily,
                fontAssetId=self.fontAssetId,
                fontWeight=self.fontWeight,
                fontStyle=self.fontStyle,
                textDecoration=self.textDecoration,
                baseline=self.baseline,
                minFontSize=self.minFontSize,
            )
        else:
            self.fontFamily = self.appearance.fontFamily
            self.fontKey = self.appearance.fontKey
            self.fontSize = self.appearance.fontSize
            self.color = self.appearance.color
            self.lineHeight = self.appearance.lineHeight
            self.align = self.appearance.align
            self.rotation = self.appearance.rotation
            self.cssFontFamily = self.appearance.cssFontFamily
            self.fontAssetId = self.appearance.fontAssetId
            self.fontWeight = self.appearance.fontWeight
            self.fontStyle = self.appearance.fontStyle
            self.textDecoration = self.appearance.textDecoration
            self.baseline = self.appearance.baseline
            self.minFontSize = self.appearance.minFontSize

        if self.fontResourceRef is None:
            self.fontResourceRef = self.fontAssetId or self.fontKey or self.fontFamily

        if not self.sourceCoverRegions and self.bbox is not None:
            self.sourceCoverRegions = [self.bbox.model_copy(deep=True)]

        if not self.quads and self.bbox is not None:
            self.quads = [
                FieldQuad(
                    x0=self.bbox.x0,
                    y0=self.bbox.y0,
                    x1=self.bbox.x1,
                    y1=self.bbox.y0,
                    x2=self.bbox.x1,
                    y2=self.bbox.y1,
                    x3=self.bbox.x0,
                    y3=self.bbox.y1,
                )
            ]

        if not self.sourceOriginalValue:
            self.sourceOriginalValue = self.originalValue

        return self


class PageSupportEntry(BaseModel):
    pageNumber: int
    kind: str
    supportMode: str
    reviewItems: list[ReviewItem] = Field(default_factory=list)


class FieldSupportEntry(BaseModel):
    fieldId: str
    page: int
    fieldType: str
    supportMode: str
    reviewState: str
    confidence: float


class SupportReport(BaseModel):
    documentClass: str
    supportMode: str
    pages: list[PageSupportEntry] = Field(default_factory=list)
    fields: list[FieldSupportEntry] = Field(default_factory=list)
    reasons: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    reviewItems: list[ReviewItem] = Field(default_factory=list)


class DocumentModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="ignore")

    id: str
    sourcePath: str
    fingerprint: str
    pageCount: int
    documentClass: str = "native-digital"
    embeddedSessionFound: bool = False
    sessionSchemaVersion: int = 1
    detectedTemplateId: Optional[str] = None
    detectedTemplateFamily: Optional[str] = None
    pages: list[PageModel] = Field(default_factory=list)
    fields: list[EditableField] = Field(
        default_factory=list,
        validation_alias=AliasChoices("fields", "blocks"),
    )
    fonts: list[FontAsset] = Field(default_factory=list)
    supportStatus: SupportStatus
    reviewItems: list[ReviewItem] = Field(default_factory=list)
    supportReport: Optional[SupportReport] = None

    @property
    def blocks(self) -> list[EditableField]:
        return self.fields

    @blocks.setter
    def blocks(self, value: list[EditableField]) -> None:
        self.fields = value


class ImportRequest(BaseModel):
    sourcePath: str


class UploadImportRequest(BaseModel):
    fileName: str
    fileDataBase64: str


class DraftUpdateRequest(BaseModel):
    fields: list[EditableField] = Field(
        default_factory=list,
        validation_alias=AliasChoices("fields", "blocks"),
    )


class DraftUpdateResponse(BaseModel):
    saved: bool
    storage: str = "embedded-session"


class LearnTemplateRequest(BaseModel):
    name: str
    description: Optional[str] = None
    fields: list[EditableField] = Field(
        default_factory=list,
        validation_alias=AliasChoices("fields", "blocks"),
    )


class LearnTemplateResponse(BaseModel):
    saved: bool
    templateId: str
    templateName: str
    fieldCount: int
    templatePath: str
    replacedExisting: bool = False


class ExportRequest(BaseModel):
    targetPath: Optional[str] = None


class ExportResponse(BaseModel):
    exported: bool
    outputPath: str


class WhiteboardExportRequest(BaseModel):
    imageDataUrl: str
    width: float
    height: float
    targetPath: Optional[str] = None


class WhiteboardExportResponse(BaseModel):
    exported: bool
    outputPath: str


TextBlock = EditableField
