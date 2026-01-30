"""
Pydantic models for API request/response schemas
"""

from pydantic import BaseModel
from typing import List, Optional, Dict, Any


class ValidationResult(BaseModel):
    """Individual validation result"""
    rule_name: str
    sheet_name: str
    status: str  # "PASS" or "FAIL"
    expected: str
    actual: str
    location: str = ""
    validator_type: str = "QUOTE_WIN"  # "BOM" or "QUOTE_WIN"


class ValidationResponse(BaseModel):
    """Response from validation endpoint"""
    session_id: str
    results: List[ValidationResult]
    passed: int
    failed: int
    total: int
    file_name: str
    validator_type: str


class ExportRequest(BaseModel):
    """Request for export endpoints"""
    session_id: str
    results: List[ValidationResult]


class SaveFileRequest(BaseModel):
    """Request for save validated file endpoint"""
    session_id: str


class RuleInfo(BaseModel):
    """Rule information for documentation"""
    rule_num: str
    title: str
    description: str


class RulesResponse(BaseModel):
    """Response with rules information"""
    validator_type: str
    rules: List[RuleInfo]
