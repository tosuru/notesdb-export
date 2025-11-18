from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional

class LinkNote(BaseModel):
    raw: str
    server: Optional[str] = None
    replica: Optional[str] = None
    unid: Optional[str] = None
    query: Optional[str] = None
    resolved_url: Optional[str] = None

class NormalizedDoc(BaseModel):
    meta: Dict[str, Any]
    fields: Dict[str, Any]
    attachments: List[Dict[str, Any]] = Field(default_factory=list)
    links: Dict[str, Any] = Field(default_factory=dict)
    layout: Dict[str, Any] = Field(default_factory=dict)
