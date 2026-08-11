# main.py
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError
from starlette.middleware.base import BaseHTTPMiddleware
import uvicorn
import uuid
from models.schemas import (
    HeaderMatchingRequest,
    HeaderMatchingResponse,
    HeaderMatch,
    ComparisonResponse,
    MatchedVendorDataItem
)
from utils.data_processor import system_json_to_dataframe, vendor_json_to_dataframe
from utils.header_matching import match_headers_ai
from utils.comparison_engine import compare_rows
from utils.helpers import normalize_header_mapping, build_unmatched_vendor_header_list, is_not_applicable_item
from config import LOG_LEVEL, LOG_JSON, LOG_FILE, CORS_ORIGINS, compute_token_cost, azure_openai_configured, azure_openai_config_status, AZURE_DEPLOYMENT_NAME, AZURE_ENDPOINT, AZURE_API_KEY, AZURE_API_VERSION, AZURE_OPENAI_TIMEOUT
from utils.logger import setup_logging, get_logger

# Set up structured logging
logger = setup_logging(level=LOG_LEVEL, use_json=LOG_JSON, log_file=LOG_FILE)

app = FastAPI(
    title="Paysheet Comparator API",
    description="API for comparing vendor paysheets with system paysheets",
    version="1.0.0"
)

# Middleware to add correlation ID and handle request logging
class CorrelationIDMiddleware(BaseHTTPMiddleware):
    """Add correlation ID to requests for tracing"""
    async def dispatch(self, request: Request, call_next):
        # Generate correlation ID for request tracing
        correlation_id = str(uuid.uuid4())
        request.state.correlation_id = correlation_id
        
        response = await call_next(request)
        
        # Add correlation ID to response headers
        response.headers["X-Correlation-ID"] = correlation_id
        
        return response

# Add correlation ID middleware
app.add_middleware(CorrelationIDMiddleware)

# CORS middleware: use CORS_ORIGINS from config (set CORS_ORIGINS env in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Exception handler for validation errors
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """Custom handler for validation errors with GDPR-safe logging"""
    correlation_id = getattr(request.state, 'correlation_id', 'unknown')
    errors = exc.errors()
    
    # Log validation error WITHOUT sensitive data
    logger.warning(
        f"Request validation failed: {request.method} {request.url.path}",
        extra={
            "correlation_id": correlation_id,
            "method": request.method,
            "path": request.url.path,
            "validation_errors": errors,
            # DO NOT log request body - contains sensitive payroll data
        }
    )
    
    # Return the standard validation error response (without sensitive data)
    return JSONResponse(
        status_code=422,
        content={
            "detail": errors,
            "correlation_id": correlation_id,
            "message": "Validation error. Check error details."
        }
    )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    """
    Catch-all for unhandled exceptions. Log full details server-side;
    return a generic message to the client (no internal details).
    """
    correlation_id = getattr(request.state, "correlation_id", "unknown")
    # Log full exception and traceback (server-side only)
    logger.exception(
        "Unhandled exception processing request",
        extra={
            "correlation_id": correlation_id,
            "method": request.method,
            "path": request.url.path,
            "exception_type": type(exc).__name__,
        },
    )
    return JSONResponse(
        status_code=500,
        content={
            "detail": "An unexpected error occurred. Please try again or contact support.",
            "correlation_id": correlation_id,
            "message": "If the problem persists, provide the correlation_id when contacting support.",
        },
    )


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "Paysheet Comparator API"}


@app.get("/health/ai")
async def ai_health_check():
    """Check Azure OpenAI config and connectivity (no secrets returned)."""
    configured = azure_openai_configured()
    result = {
        "configured": configured,
        "envVars": azure_openai_config_status(),
        "deployment": AZURE_DEPLOYMENT_NAME if configured else None,
        "endpointHost": AZURE_ENDPOINT.split("//")[-1].split("/")[0] if AZURE_ENDPOINT else None,
        "reachable": False,
        "error": None,
    }
    if not configured:
        missing = [k for k, v in result["envVars"].items() if not v]
        result["error"] = f"Missing env vars: {', '.join(missing)}"
        return result
    try:
        from langchain_openai import AzureChatOpenAI
        from langchain.schema import HumanMessage
        llm = AzureChatOpenAI(
            deployment_name=AZURE_DEPLOYMENT_NAME,
            temperature=0,
            azure_endpoint=AZURE_ENDPOINT,
            api_key=AZURE_API_KEY,
            api_version=AZURE_API_VERSION,
            timeout=min(AZURE_OPENAI_TIMEOUT, 30),
            max_retries=0,
        )
        llm.invoke([HumanMessage(content='Reply with exactly: ok')])
        result["reachable"] = True
    except Exception as e:
        result["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    return result

@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "message": "Paysheet Comparator API",
        "version": "1.0.0",
        "docs": "/docs"
    }


@app.post("/api/v1/headers/match")
async def match_headers(http_request: Request, request: HeaderMatchingRequest):
    """
    Unified endpoint for header matching and comparison.
    
    Two modes:
    1. Header Check Mode (headerCheck: true, matchedVendorData: null):
       - Performs AI-powered header matching for constant headers
       - Returns matched headers response
       
    2. Comparison Mode (headerCheck: false, matchedVendorData: provided):
       - Uses matchedVendorData to perform row-by-row comparison
       - Auto-detects employee number columns
       - Matches employees by employee number
       - Compares each mapped column (numeric with tolerance, text normalized)
       - Returns comparison results
    """
    try:
        # Convert JSON to DataFrames
        df_system = system_json_to_dataframe(request.systemData)
        df_vendor = vendor_json_to_dataframe(request.vendorPaysheetData)
        
        # Determine mode based on headerCheck flag
        header_check_mode = request.headerCheck if request.headerCheck is not None else True
        correlation_id = getattr(http_request.state, 'correlation_id', 'unknown')

        # Mode 1: Header Matching (headerCheck: true)
        if header_check_mode:
            # Use constant headers from request (caller always sends constantHeaders)
            constant_headers = request.constantHeaders or []
            # Match ONLY constant headers using AI
            matched_headers_dict, unmatched_constant_header_names, constant_headers_status, token_usage = match_headers_ai(
                df_vendor, df_system, constant_headers=constant_headers
            )

            # Convert matched headers to response format (only constant headers)
            matched_headers_list = []
            for vendor_header, system_header in matched_headers_dict.items():
                # Determine match type
                match_type = "exact" if vendor_header == system_header else "semantic"
                matched_headers_list.append(
                    HeaderMatch(
                        vendorHeader=vendor_header,
                        systemHeader=system_header,
                        matchType=match_type
                    )
                )
            
            # Count how many constant headers were matched
            matched_count = sum(1 for status in constant_headers_status.values() if status)
            total_count = len(constant_headers_status)
            
            unmatched_vendor_headers = build_unmatched_vendor_header_list(request.vendorPaysheetData)
            response = HeaderMatchingResponse(
                matchedHeaders=matched_headers_list,
                unmatchedConstantHeaders=unmatched_constant_header_names,
                unmatchedVendorHeaders=unmatched_vendor_headers,
                constantHeadersStatus=constant_headers_status,
                message=f"Constant headers matching completed. {matched_count}/{total_count} constant headers matched. {len(unmatched_vendor_headers)} vendor headers.",
                month=request.month,
                processingStage=request.processingStage,
                year=request.year,
                paysheetComparisionId=request.paysheetComparisionId,
                compensationId=request.compensationId
            )
            
            log_extra = {
                "correlation_id": correlation_id,
                "matched": len(response.matchedHeaders),
                "unmatched": len(response.unmatchedConstantHeaders),
            }
            if token_usage:
                log_extra["tokens"] = token_usage["input_tokens"] + token_usage["output_tokens"]
                cost = compute_token_cost(token_usage["input_tokens"], token_usage["output_tokens"])
                if cost is not None:
                    log_extra["cost_usd"] = cost
            logger.info("Header match done", extra=log_extra)

            return response
        
        # Mode 2: Comparison (headerCheck: false)
        else:
            # Convert matchedVendorData to header mapping format
            # Key: tuple(mappedVendorHeaders). Value: { system_header, formula_steps } for formula-driven comparison
            header_mapping = {}
            if request.matchedVendorData:
                for item in request.matchedVendorData:
                    if is_not_applicable_item(item.status, item.mappedVendorHeaders):
                        continue
                    vendor_headers_tuple = tuple(item.mappedVendorHeaders)
                    mapping_value = {
                        "system_header": item.systemColumn,
                        "formula_steps": item.formulaSteps,
                    }
                    header_mapping[vendor_headers_tuple] = mapping_value
            
            if not header_mapping:
                raise HTTPException(
                    status_code=400,
                    detail="No valid header mappings found in matchedVendorData (all columns may be 'Not Applicable'). At least one mapped column is required for comparison.",
                )
            
            # Normalize header mapping to match DataFrame column names
            normalized_mapping = normalize_header_mapping(header_mapping, df_vendor, df_system)
            
            # Log warning if some mappings could not be normalized
            unmapped_count = len(header_mapping) - len(normalized_mapping)
            if unmapped_count > 0:
                logger.warning(
                    f"Could not find columns for {unmapped_count} mapping(s)",
                    extra={"correlation_id": correlation_id, "unmapped_count": unmapped_count}
                )
            
            if not normalized_mapping:
                raise HTTPException(
                    status_code=400,
                    detail="No valid header mappings found. Please ensure column names in matchedVendorData match the actual column names in the data."
                )
            
            # Perform row-by-row comparison
            comparison_result = compare_rows(
                df_vendor=df_vendor,
                df_system=df_system,
                header_mapping=normalized_mapping
            )
            
            # Extract results and unmatched counts
            row_comparisons = comparison_result["results"]
            only_in_vendor_count = comparison_result["only_in_vendor_count"]
            only_in_system_count = comparison_result["only_in_system_count"]
            
            # Calculate summary statistics
            from utils.statistics_calculator import (
                calculate_basic_summary,
                calculate_overall_summary,
                calculate_column_statistics,    
                calculate_insights,
                calculate_quick_stats
            )
            
            summary = calculate_basic_summary(
                row_comparisons, 
                len(normalized_mapping),
                only_in_vendor_count=only_in_vendor_count,
                only_in_system_count=only_in_system_count
            )
            total_rows = summary["totalRows"]
            matched_rows = summary["matchedRows"]
            unmatched_rows = summary["unmatchedRows"]
            match_rate = summary["matchRate"]
            
            column_statistics = calculate_column_statistics(row_comparisons, top_n=5)
            overall_summary = calculate_overall_summary(row_comparisons, column_statistics)
            insights = calculate_insights(row_comparisons, column_statistics)
            quick_stats = calculate_quick_stats(row_comparisons, column_statistics)
            
            # Filter to only show mismatched employees in response (statistics already calculated from full data)
            filtered_row_comparisons = [r for r in row_comparisons if not r.overallMatch]
            
            response = ComparisonResponse(
                rowComparisons=filtered_row_comparisons,
                summary=summary,
                totalRows=total_rows,
                matchedRows=matched_rows,
                unmatchedRows=unmatched_rows,
                matchRate=match_rate,
                overallSummary=overall_summary,
                columnStatistics=column_statistics,
                insights=insights,
                quickStats=quick_stats,
                month=request.month,
                processingStage=request.processingStage,
                year=request.year,
                paysheetComparisionId=request.paysheetComparisionId,
                compensationId=request.compensationId
            )
            
            logger.info(
                "Comparison done",
                extra={
                    "correlation_id": correlation_id,
                    "rows": total_rows,
                    "match_rate": round(match_rate, 4),
                    "columns": len(normalized_mapping),
                },
            )
            
            return response
        
        
    except HTTPException:
        raise


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)

