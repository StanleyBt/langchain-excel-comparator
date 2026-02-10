# models/schemas.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Dict, Optional, Any, Union


class SystemDataItem(BaseModel):
    """System data item structure"""
    details: Dict[str, Any] = Field(..., description="Nested details object with all payroll fields")
    employeeNumber: Union[str, int] = Field(..., description="Employee number at top level")
    
    @field_validator('employeeNumber', mode='before')
    @classmethod
    def convert_employee_number_to_string(cls, v):
        """Convert employee number to string if it's an integer"""
        if isinstance(v, (int, float)):
            return str(int(v))
        return str(v) if v is not None else None


class FormulaStep(BaseModel):
    """One step in a formula: firstOperand operator secondOperand (e.g. sgst + cgst)."""
    firstOperand: str = Field(..., description="Vendor header name or step result ref e.g. _0")
    operator: str = Field(..., description="Operation: +, -, *, or /")
    secondOperand: str = Field(..., description="Vendor header name or step result ref e.g. _1")

    @field_validator("operator")
    @classmethod
    def operator_allowed(cls, v: str) -> str:
        allowed = {"+", "-", "*", "/"}
        op = str(v).strip() if v else ""
        if op not in allowed:
            raise ValueError(f"operator must be one of {allowed}, got {v!r}")
        return op


class MatchedVendorDataItem(BaseModel):
    """Matched vendor data item from frontend.

    Uses mappedVendorHeaders (list). When formulaSteps is provided, vendor value
    is computed by applying each step's operator (+, -, *, /); otherwise
    single column is used as-is, multiple columns are summed.
    """
    systemColumn: str = Field(..., description="System column name")
    mappedVendorHeaders: List[str] = Field(..., min_length=1, description="Vendor header names (list)")
    formulaSteps: Optional[List[FormulaStep]] = Field(None, description="Optional formula steps to compute value from operands")
    status: Optional[str] = Field(None, description="Status of the match")
    issue: Optional[str] = Field(None, description="Any issues detected")
    matchType: Optional[str] = Field(None, description="Type of match: 'exact', 'semantic', or 'manual'")


class HeaderMatchingRequest(BaseModel):
    """Request model for header matching endpoint - supports both header matching and comparison"""
    organizationId: Optional[int] = None
    filesCount: Optional[int] = None
    month: Optional[int] = None
    processingStage: Optional[str] = None
    year: Optional[int] = None
    compensationId: Optional[int] = None
    paysheetComparisionId: Optional[int] = None
    systemData: List[SystemDataItem] = Field(..., description="System paysheet data array")
    vendorPaysheetData: List[Dict[str, Any]] = Field(..., description="Vendor paysheet data array (flat structure)")
    headerCheck: Optional[bool] = Field(default=True, description="If true, perform header matching. If false, perform comparison using matchedVendorData")
    matchedVendorData: Optional[List[MatchedVendorDataItem]] = Field(None, description="Matched vendor data for comparison mode. Required when headerCheck is false")
    constantHeaders: Optional[List[str]] = Field(None, description="List of constant headers to match (e.g. ['EMPLOYEE_NUMBER','INVOICE_VALUE','GENDER']). Caller should always send this.")


class HeaderMatch(BaseModel):
    """Individual header match information"""
    vendorHeader: str = Field(..., description="Header name from vendor paysheet")
    systemHeader: str = Field(..., description="Matched header name from system paysheet")
    matchType: str = Field(..., description="Type of match: 'exact', 'semantic', or 'manual'")


class HeaderMatchingResponse(BaseModel):
    """Response model for header matching endpoint"""
    matchedHeaders: List[HeaderMatch] = Field(..., description="List of successfully matched constant headers")
    unmatchedConstantHeaders: List[str] = Field(..., description="List of constant header names that couldn't be matched (e.g., ['BASE_VALUE', 'CONTRACTOR'])")
    unmatchedVendorHeaders: List[str] = Field(..., description="List of all vendor headers that don't match any constant header (available for manual mapping)")
    constantHeadersStatus: Dict[str, bool] = Field(..., description="Status of constant headers matching")
    message: Optional[str] = None
    month: Optional[int] = Field(None, description="Month from request")
    processingStage: Optional[str] = Field(None, description="Processing stage from request")
    year: Optional[int] = Field(None, description="Year from request")
    paysheetComparisionId: Optional[int] = Field(None, description="Paysheet comparison ID from request")
    compensationId: Optional[int] = Field(None, description="Compensation ID from request")


class ColumnComparison(BaseModel):
    """Comparison result for a single column"""
    columnName: str = Field(..., description="Column name being compared")
    vendorValue: Optional[Union[str, float, int]] = Field(None, description="Value from vendor paysheet")
    systemValue: Optional[Union[str, float, int]] = Field(None, description="Value from system paysheet")
    difference: Optional[float] = Field(None, description="Difference (vendor - system) for numeric values")
    isMatch: bool = Field(..., description="Whether values match")
    matchType: Optional[str] = Field(None, description="Type: 'exact', 'numeric_tolerance', 'text_normalized'")


class RowComparisonResult(BaseModel):
    """Comparison result for a single row (employee)"""
    employeeNumber: str = Field(..., description="Employee number (used as the key identifier)")
    columnComparisons: List[ColumnComparison] = Field(..., description="Comparison results for each column")
    overallMatch: bool = Field(..., description="Whether all columns match")
    rowStatus: str = Field(..., description="Status: 'matched', 'only_in_vendor', 'only_in_system'")


class ColumnMismatchDetail(BaseModel):
    """Detail of a top mismatch for a column"""
    employeeNumber: str = Field(..., description="Employee number with mismatch")
    vendorValue: Optional[Union[str, float, int]] = Field(None, description="Vendor value")
    systemValue: Optional[Union[str, float, int]] = Field(None, description="System value")
    difference: Optional[float] = Field(None, description="Difference (vendor - system)")


class ColumnStatistic(BaseModel):
    """Statistics for a single column"""
    columnName: str = Field(..., description="Column name being compared")
    totalComparisons: int = Field(..., description="Total number of comparisons for this column")
    matchedCount: int = Field(..., description="Number of matched comparisons")
    mismatchedCount: int = Field(..., description="Number of mismatched comparisons")
    matchRate: float = Field(..., description="Match rate (0.0 to 1.0)")
    totalDifference: Optional[float] = Field(None, description="Sum of all differences (for numeric columns)")
    averageDifference: Optional[float] = Field(None, description="Average difference (for numeric columns)")
    maxDifference: Optional[float] = Field(None, description="Maximum difference (for numeric columns)")
    minDifference: Optional[float] = Field(None, description="Minimum difference (for numeric columns)")
    mismatchPercentage: float = Field(..., description="Percentage of mismatches")
    topMismatches: List[ColumnMismatchDetail] = Field(default_factory=list, description="Top mismatches for this column")
    vendorPaysheetCount: Optional[Union[int, float]] = Field(None, description="Total sum of all vendor values for this column (numeric columns only). For Employee Head Count, this is the count (integer).")
    systemPaysheetCount: Optional[Union[int, float]] = Field(None, description="Total sum of all system values for this column (numeric columns only). For Employee Head Count, this is the count (integer).")
    headcount: Optional[int] = Field(None, description="Headcount/matched count (only for Employee Number column)")
    
    @field_validator('vendorPaysheetCount', 'systemPaysheetCount', mode='before')
    @classmethod
    def preserve_integer_type(cls, v):
        """Preserve integer type for Employee Head Count (don't convert to float)"""
        if v is not None:
            # If it's a whole number (like 1.0), convert to int to preserve integer type in JSON
            if isinstance(v, float) and v.is_integer():
                return int(v)
        return v


class KeyFinding(BaseModel):
    """A key finding or insight"""
    type: str = Field(..., description="Type: 'info', 'warning', 'error'")
    severity: str = Field(..., description="Severity: 'low', 'medium', 'high'")
    message: str = Field(..., description="Finding message")
    affectedEmployees: int = Field(..., description="Number of employees affected")
    affectedColumns: List[str] = Field(..., description="Columns affected")
    suggestion: Optional[str] = Field(None, description="Suggestion or recommendation")


class ColumnHealth(BaseModel):
    """Health status for a column"""
    columnName: str = Field(..., description="Column name")
    healthStatus: str = Field(..., description="Status: 'excellent', 'good', 'needs_attention', 'critical'")
    matchRate: float = Field(..., description="Match rate for this column")
    averageDifference: Optional[float] = Field(None, description="Average difference")
    issues: List[str] = Field(default_factory=list, description="List of issues detected")


class Insights(BaseModel):
    """Insights and analysis"""
    overallHealth: str = Field(..., description="Overall health: 'excellent', 'good', 'fair', 'poor'")
    healthScore: float = Field(..., description="Health score (0.0 to 1.0)")
    keyFindings: List[KeyFinding] = Field(default_factory=list, description="Key findings")
    columnHealth: List[ColumnHealth] = Field(default_factory=list, description="Health status per column")
    recommendations: List[str] = Field(default_factory=list, description="Recommendations")
    dataQuality: Dict[str, Any] = Field(default_factory=dict, description="Data quality metrics")


class QuickStats(BaseModel):
    """Quick statistics dashboard"""
    perfectMatches: int = Field(..., description="Employees with all columns matching")
    partialMatches: int = Field(..., description="Employees with some columns matching")
    noMatches: int = Field(..., description="Employees with no columns matching")
    mostProblematicColumn: Optional[str] = Field(None, description="Column with lowest match rate")
    bestMatchingColumn: Optional[str] = Field(None, description="Column with highest match rate")
    largestDiscrepancy: Optional[Dict[str, Any]] = Field(None, description="Largest discrepancy found")


class OverallSummary(BaseModel):
    """Simplified overall summary with essential statistics"""
    totalRowsValidated: int = Field(..., description="Total number of rows (employees) validated")
    totalMatchPercentage: float = Field(..., description="Overall match percentage across all columns (0.0 to 100.0)")
    columnMatchPercentages: Dict[str, float] = Field(..., description="Match percentage for each column (column name -> percentage)")


class ComparisonResponse(BaseModel):
    """Response model for comparison endpoint"""
    rowComparisons: List[RowComparisonResult] = Field(..., description="Comparison results for each row")
    summary: Dict[str, Any] = Field(..., description="Summary statistics")
    totalRows: int = Field(..., description="Total number of rows compared")
    matchedRows: int = Field(..., description="Number of rows with all columns matching")
    unmatchedRows: int = Field(..., description="Number of rows with mismatches")
    matchRate: float = Field(..., description="Overall match rate (0.0 to 1.0)")
    overallSummary: Optional[OverallSummary] = Field(None, description="Enhanced overall summary")
    columnStatistics: Optional[List[ColumnStatistic]] = Field(None, description="Statistics per column")
    insights: Optional[Insights] = Field(None, description="Insights and analysis")
    quickStats: Optional[QuickStats] = Field(None, description="Quick statistics dashboard")
    month: Optional[int] = Field(None, description="Month from request")
    processingStage: Optional[str] = Field(None, description="Processing stage from request")
    year: Optional[int] = Field(None, description="Year from request")
    paysheetComparisionId: Optional[int] = Field(None, description="Paysheet comparison ID from request")
    compensationId: Optional[int] = Field(None, description="Compensation ID from request")

