# excel_comparator/utils/header_matching.py
import pandas as pd
import json
import re
from typing import Dict, List, Tuple, Optional
from langchain_openai import AzureChatOpenAI
from langchain.schema import HumanMessage
from config import (
    AZURE_DEPLOYMENT_NAME, AZURE_ENDPOINT, AZURE_API_KEY, AZURE_API_VERSION,
    AZURE_OPENAI_TIMEOUT, AZURE_OPENAI_MAX_RETRIES
)
from config import CONSTANT_HEADERS, HEADER_MATCHING_PRIORITY, HEADER_VARIATIONS
import time
from utils.normalization import normalize_header_name
from utils.logger import get_logger

logger = get_logger("header_matching")

# Removed Streamlit-specific functions: find_header_row, detect_header_row, run_header_mapping
# These were only used by the Streamlit UI app which has been removed

# Alias for backward compatibility
def normalize_header_for_matching(header: str) -> str:
    """
    Normalize header name for matching comparison.
    Uses centralized normalization function.
    """
    return normalize_header_name(header, for_matching=True)


def match_headers_ai(
    df_vendor: pd.DataFrame, 
    df_system: pd.DataFrame,
    constant_headers: Optional[List[str]] = None
) -> Tuple[Dict[str, str], List[str], List[str], Dict[str, bool]]:
    """
    Match ONLY constant headers between vendor and system DataFrames using AI.
    Only matches the headers specified in CONSTANT_HEADERS config.
    
    Args:
        df_vendor: Vendor paysheet DataFrame
        df_system: System paysheet DataFrame
        constant_headers: List of constant headers to match (default: CONSTANT_HEADERS from config)
        
    Returns:
        Tuple of:
        - matched_headers: Dict mapping vendor_header -> system_header (ONLY constant headers)
        - unmatched_constant_header_names: List of constant header names that couldn't be matched (e.g., ['BASE_VALUE', 'CONTRACTOR'])
        - unmatched_vendor_headers: List of vendor headers that don't match any constant header
        - constant_headers_status: Dict showing which constant headers were matched
    """
    if constant_headers is None:
        constant_headers = CONSTANT_HEADERS
    
    # Get normalized headers (keep original for return)
    vendor_headers_original = [str(col) for col in df_vendor.columns]
    # Headers are already normalized to lowercase in data_processor
    vendor_headers = [str(col).strip().lower() for col in df_vendor.columns]
    system_headers = [str(col).strip().lower() for col in df_system.columns]
    
    matched_headers = {}
    constant_headers_status = {}
    unmatched_constant_header_names = []
    
    # Get variations and priority from config
    constant_mapping_variations = HEADER_VARIATIONS
    priority_keywords = HEADER_MATCHING_PRIORITY
    
    # Step 1: Match ONLY constant headers
    for const_header in constant_headers:
        const_normalized = normalize_header_for_matching(const_header)
        matched = False
        
        # Get variations ordered by priority
        variations = constant_mapping_variations.get(const_header, [const_normalized])
        
        # Phase 1: Create hints from variations (suggestions, not filters)
        # These hints help AI prioritize, but AI will analyze ALL headers
        high_confidence_hints = []  # List of (vendor_header, priority_score)
        
        for vendor_header in vendor_headers:
            vendor_norm = normalize_header_for_matching(vendor_header)
            
            # Check if vendor header matches any variation (for hint generation)
            matches_variation = False
            for variation in variations:
                if variation in vendor_norm or vendor_norm in variation:
                    matches_variation = True
                    break
            
            if matches_variation:
                # Calculate priority score based on keywords
                priority_score = 999  # Default low priority
                if const_header in priority_keywords and priority_keywords[const_header]:
                    # Use configured priority keywords
                    for priority, keyword_list in enumerate(priority_keywords[const_header]):
                        # Check if vendor header contains all keywords in this priority level
                        if all(keyword in vendor_norm for keyword in keyword_list):
                            priority_score = priority  # Lower number = higher priority
                            break
                
                # If no priority match found, use variation order as fallback
                if priority_score == 999:
                    for priority, variation in enumerate(variations):
                        if variation in vendor_norm or vendor_norm in variation:
                            priority_score = priority
                            break
                
                high_confidence_hints.append((vendor_header, priority_score))
        
        # Sort hints by priority (lower number = better match)
        high_confidence_hints.sort(key=lambda x: x[1])
        
        # Phase 2: Always call AI with ALL headers, passing hints as suggestions
        # AI will analyze all headers semantically and make the final decision
        ai_was_called = False
        ai_explicitly_rejected = False
        
        # Collect ALL vendor headers that haven't been matched yet
        all_vendor_headers = [vh for vh in vendor_headers if vh not in matched_headers]
        all_system_headers = list(system_headers)
        
        # Extract top hints (up to 5) as suggestions for AI
        hint_list = [h[0] for h in high_confidence_hints[:5]] if high_confidence_hints else None
        
        # Always call AI with all headers (not just filtered candidates)
        if all_vendor_headers and all_system_headers:
            try:
                ai_was_called = True
                ai_selected = _ai_semantic_matching_for_constant_header(
                    const_header, 
                    all_vendor_headers,  # ALL headers, not filtered
                    all_system_headers,  # ALL headers
                    variations,
                    hints=hint_list  # Pass hints as suggestions
                )
                if ai_selected:
                    for vendor_h, system_h in ai_selected.items():
                        if vendor_h not in matched_headers:
                            matched_headers[vendor_h] = system_h
                            constant_headers_status[const_header] = True
                            matched = True
                            break
                else:
                    # AI returned empty dict - it explicitly rejected all candidates
                    ai_explicitly_rejected = True
            except Exception as e:
                error_str = str(e).lower()
                is_timeout = "timeout" in error_str or "timed out" in error_str
                
                if is_timeout:
                    print(f"⚠️  Warning: AI request timed out after {AZURE_OPENAI_TIMEOUT}s (with {AZURE_OPENAI_MAX_RETRIES} retries). Falling back to priority-based matching.")
                else:
                    print(f"⚠️  Warning: AI validation failed, falling back to priority-based matching: {e}")
                # AI service failed, but we'll still try fallback
        
        # Phase 3: Fallback direct matching - Only if AI service failed (exception)
        # DO NOT fall back if AI explicitly rejected candidates (empty dict)
        if not matched and not ai_explicitly_rejected and ai_was_called:
            # AI was called but failed - try direct matching with hints as fallback
            # Try to match with system headers, starting with best vendor header from hints
            for vendor_header, _ in high_confidence_hints:
                # Check if this vendor header is already matched to a different constant header
                if vendor_header in matched_headers:
                    continue
                    
                vendor_norm = normalize_header_for_matching(vendor_header)
                
                # Find matching system header
                for system_header in system_headers:
                    system_norm = normalize_header_for_matching(system_header)
                    # Check if system header matches any variation
                    if any(var in system_norm or system_norm in var for var in variations):
                        matched_headers[vendor_header] = system_header
                        constant_headers_status[const_header] = True
                        matched = True
                        break
                if matched:
                    break
        
        # Phase 4: Fallback - Only needed if Phase 2 AI failed or was not called
        # Since Phase 2 now always calls AI with all headers, this phase is rarely needed
        # But keep it as a safety net for edge cases
        if not matched and not ai_was_called:
            # This should rarely happen now, but keep as fallback
            # Collect ALL vendor headers that haven't been matched yet
            potential_vendor = []
            for vendor_header in vendor_headers:
                # Skip if already matched to a different constant header
                if vendor_header not in matched_headers:
                    potential_vendor.append(vendor_header)
            
            # Collect ALL system headers
            potential_system = list(system_headers)
            
            # Use AI to autonomously analyze and match - AI will find matches based on semantic understanding
            # Pass hints if available
            hint_list = [h[0] for h in high_confidence_hints[:5]] if high_confidence_hints else None
            if potential_vendor and potential_system:
                try:
                    ai_matches = _ai_semantic_matching_for_constant_header(
                        const_header, potential_vendor, potential_system, variations, hints=hint_list
                    )
                    # Use AI's match if it found one
                    if ai_matches:
                        for vendor_h, system_h in ai_matches.items():
                            # Only use if vendor header not already matched
                            if vendor_h not in matched_headers:
                                matched_headers[vendor_h] = system_h
                                constant_headers_status[const_header] = True
                                matched = True
                                break
                except Exception as e:
                    error_str = str(e).lower()
                    is_timeout = "timeout" in error_str or "timed out" in error_str
                    
                    if is_timeout:
                        logger.warning(
                            f"AI request timed out for constant header",
                            extra={"timeout": AZURE_OPENAI_TIMEOUT, "retries": AZURE_OPENAI_MAX_RETRIES, "constant_header": const_header}
                        )
                    else:
                        logger.warning(
                            f"AI semantic matching failed for constant header",
                            extra={"error": str(e)[:100], "constant_header": const_header}  # Truncate error message
                        )
        
        if not matched:
            constant_headers_status[const_header] = False
            # Add the constant header name to unmatched list
            unmatched_constant_header_names.append(const_header)
    
    # Find all vendor headers that weren't matched to any constant header
    matched_vendor_headers_lower = set(matched_headers.keys())
    unmatched_vendor_headers = []
    
    for i, vendor_header_lower in enumerate(vendor_headers):
        if vendor_header_lower not in matched_vendor_headers_lower:
            # Return original header name (not normalized)
            unmatched_vendor_headers.append(vendor_headers_original[i])
    
    return matched_headers, unmatched_constant_header_names, unmatched_vendor_headers, constant_headers_status


def _get_semantic_context_for_header(const_header: str) -> str:
    """
    Get semantic context/description for a constant header to help AI understand its real-world meaning.
    """
    context_map = {
        "EMPLOYEE_NUMBER": """
BUSINESS MEANING: Unique identifier for an employee (e.g., "EMP001", "12345", "10007")
- This is an IDENTIFIER, not a name
- Usually text/string (may have leading zeros)
- Used to uniquely identify employees across systems
- Examples: "employee_number", "employee_id", "emp_no", "employee_code"
- REJECT: "employee_name", "employee_name_full" (these are names, not numbers/IDs)
""",
        "NET_PAY": """
BUSINESS MEANING: Employee's take-home pay AFTER all deductions (taxes, insurance, PF, etc.)
- This is the FINAL amount the employee receives
- Calculated as: Gross Pay - All Deductions = Net Pay
- Examples: "net_pay", "netpay", "net_salary", "take_home", "net_amount"
- REJECT: "gross_pay" (before deductions), "total_payable" (total to be paid, different concept), 
  "basic_pay" (base component), "total_pay" (could be gross)
- KEY DISTINCTION: "total_payable" means "total amount that needs to be paid" (could be gross or invoice amount),
  while "net_pay" means "what employee takes home after deductions" - these are DIFFERENT concepts
""",
        "INVOICE_VALUE": """
BUSINESS MEANING: The monetary value/amount on an invoice or bill
- This is the AMOUNT to be paid, not a count or number
- Examples: "invoice_value", "invoice_amount", "invoice_total", "bill_amount"
- REJECT: "invoice_number" (ID, not amount), "invoice_count" (quantity, not value), 
  "invoice_date" (date, not amount)
"""
    }
    
    return context_map.get(const_header, f"""
BUSINESS MEANING: {const_header}
- Analyze what this header represents in real-world payroll/paysheet systems
- Match only if vendor and system headers represent the same business concept
""")


def _ai_semantic_matching_for_constant_header(
    const_header: str,
    potential_vendor: List[str],
    potential_system: List[str],
    expected_variations: List[str],
    hints: Optional[List[str]] = None
) -> Dict[str, str]:
    """
    Use AI to semantically match headers for a specific constant header.
    AI will intelligently distinguish between semantically different headers.
    For example: "contractor name" vs "contractor id" - AI knows "name" is better for CONTRACTOR.
    
    Args:
        const_header: The constant header being matched (e.g., "CONTRACTOR")
        potential_vendor: List of ALL vendor headers to analyze (not filtered)
        potential_system: List of ALL system headers to analyze (not filtered)
        expected_variations: List of expected variations for this constant header
        hints: Optional list of high-confidence header hints from Phase 1 (suggestions, not requirements)
        
    Returns:
        Dictionary mapping vendor_header -> system_header (only if AI finds a good semantic match)
    """
    if not potential_vendor or not potential_system:
        return {}
    
    llm = AzureChatOpenAI(
        deployment_name=AZURE_DEPLOYMENT_NAME,
        temperature=0,
        azure_endpoint=AZURE_ENDPOINT,
        api_key=AZURE_API_KEY,
        api_version=AZURE_API_VERSION,
        timeout=AZURE_OPENAI_TIMEOUT,
        max_retries=0  # We handle retries manually with exponential backoff
    )
    
    # Get semantic context for the constant header
    semantic_context = _get_semantic_context_for_header(const_header)
    
    # Format hints for prompt
    hints_text = ", ".join(hints) if hints else "None - analyze all headers independently"
    
    # Create a smart prompt that helps AI understand semantic differences
    prompt = f"""
You are an intelligent data assistant matching column headers for payroll/paysheet data.
You have access to ALL headers and must autonomously find the best semantic match.

CONTEXT:
- System paysheet is the SOURCE OF TRUTH with standard column names (constant headers)
- Vendor paysheets come from different sources with varying naming conventions
- Your task: Match vendor columns to system's constant headers semantically

CONSTANT HEADER TO MATCH: "{const_header}"

{semantic_context}

Expected variations for this header: {expected_variations}

HIGH-CONFIDENCE HINTS (from Phase 1 - use as suggestions, not requirements):
These headers matched hardcoded variations and are likely candidates:
{hints_text}

ALL AVAILABLE VENDOR HEADERS (analyze ALL of these, not just hints):
{potential_vendor}

ALL AVAILABLE SYSTEM HEADERS (analyze ALL of these):
{potential_system}

MATCHING STRATEGY:
1. Consider hints as likely candidates, but analyze ALL headers independently
2. Understand business meaning, not just keywords
3. System headers are standard - vendor headers may vary significantly
4. Pick the BEST semantic match, even if not in hints
5. Reject matches that don't make business sense, even if keywords overlap

CORE PRINCIPLES:
1. Match ONLY if vendor and system headers represent the SAME business concept
2. Use the BUSINESS MEANING above - it tells you what to accept and reject
3. When in doubt, REJECT - return {{}} (empty object)
4. Better to require manual mapping than create incorrect matches

⚠️ Return output as JSON only, in this format:
{{"vendor_header": "system_header"}} or {{}} if no good match

IMPORTANT: The semantic context above already specifies what to accept/reject for "{const_header}".
Follow those guidelines. If you're unsure, return {{}} (empty object).
"""
    
    # Retry logic with exponential backoff for timeout errors
    last_error = None
    for attempt in range(AZURE_OPENAI_MAX_RETRIES + 1):
        try:
            response = llm.invoke([HumanMessage(content=prompt)])
            raw_output = response.content
            break  # Success, exit retry loop
        except Exception as e:
            last_error = e
            error_str = str(e).lower()
            is_timeout = "timeout" in error_str or "timed out" in error_str or "time" in error_str
            
            if attempt < AZURE_OPENAI_MAX_RETRIES and is_timeout:
                # Exponential backoff: wait 1s, 2s, 4s, etc.
                wait_time = 2 ** attempt
                logger.info(
                    f"AI request timeout, retrying",
                    extra={
                        "attempt": attempt + 1,
                        "max_retries": AZURE_OPENAI_MAX_RETRIES + 1,
                        "wait_time": wait_time,
                        "constant_header": const_header
                    }
                )
                time.sleep(wait_time)
            else:
                # Last attempt or non-timeout error - raise it
                raise
    
    # If we get here without breaking, it means all retries failed
    if last_error:
        raise last_error
    
    try:
        semantic_matches = json.loads(raw_output)
    except json.JSONDecodeError:
        match = re.search(r"{.*}", raw_output, re.DOTALL)
        if match:
            semantic_matches = json.loads(match.group())
        else:
            return {}
    
    # Filter out null/None values and "no match" strings
    return {
        k: v for k, v in semantic_matches.items() 
        if v and str(v).lower() not in ["null", "none", "no match", ""]
    }


