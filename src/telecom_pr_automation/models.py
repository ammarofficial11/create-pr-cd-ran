from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass(frozen=True)
class RawBOMLine:
    row_id: int
    raw_equipment: str
    quantity: float
    site_code: str = ""
    site_name: str = ""
    region: str = ""
    du_code: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class NormalizedBOMLine:
    row_id: int
    item_key: str
    quantity: float
    source: str
    site_code: str = ""
    site_name: str = ""
    region: str = ""
    du_code: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ItemKeyQuantity:
    item_key: str
    quantity: float


@dataclass(frozen=True)
class PRLineItem:
    item_key: str
    pr_description: str
    quantity: float
    unit: str
    metadata: Dict[str, Any]
