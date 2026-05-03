from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any, Dict, List, Literal, Optional

try:
    from pydantic import BaseModel, Field

    HAS_PYDANTIC = True
except ModuleNotFoundError:
    BaseModel = object  # type: ignore[assignment]
    HAS_PYDANTIC = False

    def Field(default: Any = None, default_factory: Any = None, **_: Any) -> Any:
        if default_factory is not None:
            return field(default_factory=default_factory)
        return default


class MonitorType(str, Enum):
    surface_water = "surface_water"
    groundwater = "groundwater"
    ambient_air = "ambient_air"
    acoustic = "acoustic"
    soil = "soil"
    wastewater = "wastewater"
    waste_gas = "waste_gas"
    unknown = "unknown"


if HAS_PYDANTIC:

    class DocumentChunk(BaseModel):
        chunk_id: str
        source_file: str
        kind: str
        text: str
        metadata: Dict[str, Any] = Field(default_factory=dict)


    class ExtractedRecord(BaseModel):
        source_type: Optional[str] = None
        monitor_type: MonitorType = MonitorType.unknown
        noise_type: Optional[str] = None
        noise_type_label: Optional[str] = None
        point: Optional[str] = None
        sample_date: Optional[str] = None
        factor: Optional[str] = None
        value: Optional[str] = None
        unit: Optional[str] = None
        standard_class: Optional[str] = None
        evidence: Any
        confidence: float = Field(default=0.5, ge=0.0, le=1.0)
        chunk_id: str
        source_file: str
        extraction_method: Literal["llm", "rule", "merged"] = "llm"
        needs_review: bool = False


    class ExtractionResult(BaseModel):
        contains_monitoring_data: bool = False
        records: List[ExtractedRecord] = Field(default_factory=list)
        notes: Optional[str] = None

else:

    @dataclass
    class _DumpMixin:
        def model_dump(self, mode: str = "json") -> Dict[str, Any]:
            data = asdict(self)
            if isinstance(data.get("monitor_type"), MonitorType):
                data["monitor_type"] = data["monitor_type"].value
            return data

        def dict(self) -> Dict[str, Any]:
            return self.model_dump()


    @dataclass
    class DocumentChunk(_DumpMixin):
        chunk_id: str
        source_file: str
        kind: str
        text: str
        metadata: Dict[str, Any] = field(default_factory=dict)


    @dataclass
    class ExtractedRecord(_DumpMixin):
        evidence: Any
        chunk_id: str
        source_file: str
        source_type: Optional[str] = None
        monitor_type: MonitorType | str = MonitorType.unknown
        noise_type: Optional[str] = None
        noise_type_label: Optional[str] = None
        point: Optional[str] = None
        sample_date: Optional[str] = None
        factor: Optional[str] = None
        value: Optional[str] = None
        unit: Optional[str] = None
        standard_class: Optional[str] = None
        confidence: float = 0.5
        extraction_method: Literal["llm", "rule", "merged"] = "llm"
        needs_review: bool = False

        def __post_init__(self) -> None:
            try:
                self.monitor_type = MonitorType(str(self.monitor_type))
            except ValueError:
                self.monitor_type = MonitorType.unknown
            try:
                confidence = float(self.confidence)
            except (TypeError, ValueError):
                confidence = 0.5
            self.confidence = max(0.0, min(1.0, confidence))
            if self.extraction_method not in {"llm", "rule", "merged"}:
                self.extraction_method = "llm"


    @dataclass
    class ExtractionResult(_DumpMixin):
        contains_monitoring_data: bool = False
        records: List[ExtractedRecord] = field(default_factory=list)
        notes: Optional[str] = None
