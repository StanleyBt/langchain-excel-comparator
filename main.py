# main.py
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response
import uvicorn
import json
import os
from datetime import datetime
from models.schemas import ( 
    HeaderMatchingRequest,
    HeaderMatchingResponse,
    HeaderMatch,
    ManualMappingRequest,
    ComparisonRequest,
    ComparisonResponse,
    MatchedVendorDataItem
)
from utils.data_processor import system_json_to_dataframe, vendor_json_to_dataframe
from utils.header_matching import match_headers_ai
from utils.comparison_engine import compare_rows
from utils.helpers import (
    filter_by_contractor, normalize_header_mapping, 
    calculate_match_rate, calculate_basic_summary
)
from config import CONSTANT_HEADERS

app = FastAPI(
    title="Paysheet Comparator API",
    description="API for comparing vendor paysheets with system paysheets",
    version="1.0.0"
)

# Middleware to store request body for exception handler
class RequestBodyLoggerMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        # Only process POST requests
        if request.method == "POST":
            # Read the request body
            body = await request.body()
            
            # Store body in request state for exception handler
            request.state.body = body
            try:
                request.state.body_str = body.decode('utf-8') if body else ""
            except Exception:
                request.state.body_str = ""
            
            # Recreate the request with the body (since we consumed it)
            async def receive():
                return {"type": "http.request", "body": body}
            request._receive = receive
        
        # Process the request and return response
        response = await call_next(request)
        return response

# Add request body logger middleware
app.add_middleware(RequestBodyLoggerMiddleware)

# CORS middleware for frontend communication
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure this based on your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Exception handler for validation errors
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """Custom handler to print validation errors for debugging"""
    import traceback
    
    print("\n" + "="*80)
    print("❌ ERROR: Request Validation Failed (422 Unprocessable Entity)")
    print("="*80)
    print(f"📍 Path: {request.url.path}")
    print(f"🔧 Method: {request.method}")
    print(f"🌐 URL: {request.url}")
    print("-"*80)
    print("📋 Validation Errors (Detailed):")
    errors = exc.errors()
    for i, error in enumerate(errors, 1):
        print(f"\n  Error {i}:")
        print(f"    Location: {error.get('loc', [])}")
        print(f"    Message: {error.get('msg', 'N/A')}")
        print(f"    Type: {error.get('type', 'N/A')}")
        if 'ctx' in error:
            print(f"    Context: {error.get('ctx')}")
    print("\n" + "-"*80)
    print("📦 Full Error JSON:")
    print(json.dumps(errors, indent=2))
    print("-"*80)
    print("📥 Request Body (if available):")
    body_str = ""
    body_json = None
    
    # Try to get body from request state (set by middleware)
    if hasattr(request.state, 'body_str') and request.state.body_str:
        body_str = request.state.body_str
    else:
        # Fallback: try to read body directly
        try:
            body = await request.body()
            body_str = body.decode('utf-8') if body else ""
        except Exception as read_err:
            print(f"(could not read body: {read_err})")
            body_str = ""
    
    # Try to parse as JSON
    if body_str:
        try:
            body_json = json.loads(body_str)
            print(json.dumps(body_json, indent=2, default=str))
        except:
            print("(Raw body - not valid JSON):")
            print(body_str[:1000] + ("..." if len(body_str) > 1000 else ""))
    else:
        print("(empty body)")
    
    # Save validation error to file for debugging
    try:
        logs_dir = "request_logs"
        os.makedirs(logs_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        safe_path = request.url.path.replace("/", "_").replace("\\", "_")
        filename = f"{logs_dir}/validation_error_{safe_path}_{timestamp}.json"
        
        error_log_data = {
            "timestamp": datetime.now().isoformat(),
            "method": request.method,
            "path": request.url.path,
            "url": str(request.url),
            "validation_errors": errors,
            "request_body": body_json if body_json else body_str if body_str else None
        }
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(error_log_data, f, indent=2, default=str, ensure_ascii=False)
        
        print(f"💾 Validation error saved to: {filename}")
    except Exception as save_error:
        print(f"⚠️  Failed to save validation error to file: {save_error}")
    
    print("="*80 + "\n")
    
    # Return the standard validation error response
    return JSONResponse(
        status_code=422,
        content={"detail": errors, "body": "Check server logs for full request body", "saved_to": filename if 'filename' in locals() else None}
    )

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "Paysheet Comparator API"}

@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "message": "Paysheet Comparator API",
        "version": "1.0.0",
        "docs": "/docs"
    }


@app.post("/api/v1/headers/match")
async def match_headers(request: HeaderMatchingRequest):
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
        
        print("\n" + "="*80)
        print(f"🔀 MODE DETECTION: headerCheck = {header_check_mode}")
        print("="*80)
        
        # Mode 1: Header Matching (headerCheck: true)
        if header_check_mode:
            print("✅ MODE 1: Header Matching Mode")
            print("="*80 + "\n")
            
            # Match ONLY constant headers using AI
            matched_headers_dict, unmatched_constant_header_names, unmatched_vendor_headers, constant_headers_status = match_headers_ai(
                df_vendor, df_system
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
            
            response = HeaderMatchingResponse(
                matchedHeaders=matched_headers_list,
                unmatchedConstantHeaders=unmatched_constant_header_names,
                unmatchedVendorHeaders=unmatched_vendor_headers,
                constantHeadersStatus=constant_headers_status,
                message=f"Constant headers matching completed. {matched_count}/{total_count} constant headers matched. {len(unmatched_vendor_headers)} vendor headers available for manual mapping.",
                month=request.month,
                processingStage=request.processingStage,
                year=request.year,
                paysheetComparisionId=request.paysheetComparisionId,
                compensationId=request.compensationId
            )
            
            # Log response summary
            print(f"✅ Header Matching Complete: {len(response.matchedHeaders)} matched, {len(response.unmatchedConstantHeaders)} unmatched")
            
            return response
        
        # Mode 2: Comparison (headerCheck: false)
        else:
            print("✅ MODE 2: Comparison Mode")
            if request.matchedVendorData:
                print(f"   - Using {len(request.matchedVendorData)} column mappings")
            print("="*80 + "\n")
            
            # Convert matchedVendorData to header mapping format
            # matchedVendorData format: {systemColumn, mappedVendorHeader, ...}
            # We need: {vendorHeader: systemHeader}
            header_mapping = {}
            if request.matchedVendorData:
                for item in request.matchedVendorData:
                    vendor_header = item.mappedVendorHeader
                    system_header = item.systemColumn
                    header_mapping[vendor_header] = system_header
            
            if not header_mapping:
                raise HTTPException(
                    status_code=400,
                    detail="No valid header mappings found in matchedVendorData. Cannot perform comparison."
                )
            
            # Log header mappings
            print(f"🔍 Using {len(header_mapping)} header mappings from matchedVendorData")
            
            # Filter by contractor if provided
            if request.contractorFilter:
                df_system = filter_by_contractor(df_system, request.contractorFilter)
                df_vendor = filter_by_contractor(df_vendor, request.contractorFilter)
            
            # Normalize header mapping to match DataFrame column names
            normalized_mapping = normalize_header_mapping(header_mapping, df_vendor, df_system)
            
            # Log warnings for unmapped columns
            normalized_keys = set(normalized_mapping.keys())
            for vendor_key, system_key in header_mapping.items():
                # Check if this mapping was successfully normalized
                vendor_found = any(str(v).strip().lower() == str(vendor_key).strip().lower() for v in normalized_keys)
                if not vendor_found:
                    print(f"⚠️  Warning: Could not find columns for mapping: '{vendor_key}' -> '{system_key}'")
            
            print(f"✅ Normalized {len(normalized_mapping)} header mappings")
            
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
            
            # Calculate enhanced statistics
            from utils.statistics_calculator import (
                calculate_overall_summary,
                calculate_column_statistics,    
                calculate_insights,
                calculate_quick_stats
            )
            
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
            
            # Log comparison summary
            print(f"✅ Comparison Complete:")
            print(f"   Total Rows: {total_rows} | Match Rate: {match_rate:.2%}")
            print(f"   Overall Match: {overall_summary.totalMatchPercentage:.2f}%")
            print(f"   Columns: {len(normalized_mapping)} | Perfect Matches: {quick_stats.perfectMatches}")
            
            return response
        
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing request: {str(e)}")


@app.post("/api/v1/headers/manual-map", response_model=HeaderMatchingResponse)
async def manual_map_headers(request: ManualMappingRequest):
    """
    Accept complete header mapping from frontend (includes both auto-matched and manually mapped headers).
    
    This endpoint validates the complete header mapping and returns it in the response format.
    """
    try:
        # Convert JSON to DataFrames
        df_system = system_json_to_dataframe(request.systemData)
        df_vendor = vendor_json_to_dataframe(request.vendorPaysheetData)
        
        # Normalize header mapping keys to match DataFrame column names
        vendor_headers_lower = {str(h).strip().lower(): h for h in df_vendor.columns}
        system_headers_lower = {str(h).strip().lower(): h for h in df_system.columns}
        
        # Process the complete header mapping from frontend
        matched_headers_dict = {}
        matched_headers_list = []
        
        for vendor_key, system_key in request.headerMapping.items():
            vendor_header = vendor_headers_lower.get(str(vendor_key).strip().lower())
            system_header = system_headers_lower.get(str(system_key).strip().lower())
            
            if vendor_header and system_header:
                matched_headers_dict[vendor_header] = system_header
                # Determine match type (check if it was auto-matched or manual)
                match_type = "manual"  # All from this endpoint are manual mappings
                matched_headers_list.append(
                    HeaderMatch(
                        vendorHeader=vendor_header,
                        systemHeader=system_header,
                        matchType=match_type
                    )
                )
        
        # Check constant headers status
        constant_headers_status = {}
        unmatched_constant_header_names = []
        
        for const_header in CONSTANT_HEADERS:
            # Check if this constant header is in the mapping
            const_found = False
            for vendor_h, system_h in matched_headers_dict.items():
                system_norm = str(system_h).strip().upper()
                if system_norm == const_header.upper():
                    constant_headers_status[const_header] = True
                    const_found = True
                    break
            
            if not const_found:
                constant_headers_status[const_header] = False
                unmatched_constant_header_names.append(const_header)
        
        return HeaderMatchingResponse(
            matchedHeaders=matched_headers_list,
            unmatchedConstantHeaders=unmatched_constant_header_names,
            constantHeadersStatus=constant_headers_status,
            message="Header mapping completed. All constant headers should be mapped before proceeding to comparison.",
            month=request.month,
            processingStage=request.processingStage,
            year=request.year,
            paysheetComparisionId=request.paysheetComparisionId,
            compensationId=request.compensationId
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error during manual header mapping: {str(e)}")


@app.post("/api/v1/compare", response_model=ComparisonResponse)
async def compare_paysheets(request: ComparisonRequest):
    """
    Perform row-by-row comparison between vendor and system paysheets.
    
    Requires:
    - systemData and vendorPaysheetData
    - Complete headerMapping (from AI matching + manual mapping)
    - Optional contractorFilter to filter by contractor name
    
    Returns:
    - Row-by-row comparison results with differences
    - Match status for each column
    - Summary statistics
    """
    try:
        # Convert JSON to DataFrames
        df_system = system_json_to_dataframe(request.systemData)
        df_vendor = vendor_json_to_dataframe(request.vendorPaysheetData)
        
        # Filter by contractor if provided
        if request.contractorFilter:
            df_system = filter_by_contractor(df_system, request.contractorFilter)
            df_vendor = filter_by_contractor(df_vendor, request.contractorFilter)
        
        # Normalize header mapping
        normalized_mapping = normalize_header_mapping(request.headerMapping, df_vendor, df_system)
        
        if not normalized_mapping:
            raise HTTPException(
                status_code=400, 
                detail="No valid header mappings found. Please ensure header mapping matches DataFrame columns."
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
        
        # Filter to only show mismatched employees in response (statistics already calculated from full data)
        filtered_row_comparisons = [r for r in row_comparisons if not r.overallMatch]
        
        return ComparisonResponse(
            rowComparisons=filtered_row_comparisons, 
            summary=summary,
            totalRows=total_rows,
            matchedRows=matched_rows,
            unmatchedRows=unmatched_rows,
            matchRate=match_rate,
            month=request.month,
            processingStage=request.processingStage,
            year=request.year,
            paysheetComparisionId=request.paysheetComparisionId,
            compensationId=request.compensationId
        )
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error during comparison: {str(e)}")

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)

