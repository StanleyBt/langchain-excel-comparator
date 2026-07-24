# utils/header_matching.py
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
from config import HEADER_MATCHING_PRIORITY, HEADER_VARIATIONS
import time
from utils.normalization import normalize_header_name
from utils.logger import get_logger

logger = get_logger("header_matching")


def match_headers_ai(
    df_vendor: pd.DataFrame, 
    df_system: pd.DataFrame,
    constant_headers: Optional[List[str]] = None
) -> Tuple[Dict[str, str], List[str], Dict[str, bool], Optional[Dict[str, int]]]:
    """
    Match ONLY constant headers between vendor and system DataFrames using AI.
    Only matches the headers specified in the constant_headers list (from request).

    Returns:
        Tuple of (matched_headers, unmatched_constant_header_names, constant_headers_status, token_usage).
        token_usage: {"input_tokens": int, "output_tokens": int} or None (for logging only, not in API response).
    """
    constant_headers = constant_headers or []

    vendor_headers = [str(col).strip().lower() for col in df_vendor.columns]
    system_headers = [str(col).strip().lower() for col in df_system.columns]

    matched_headers = {}
    used_system_headers = set()  # normalized — each system column can only be matched once
    constant_headers_status = {}
    unmatched_constant_header_names = []
    total_input_tokens = 0
    total_output_tokens = 0
    
    # Get variations and priority from config
    constant_mapping_variations = HEADER_VARIATIONS
    priority_keywords = HEADER_MATCHING_PRIORITY

    def _system_key(system_header: str) -> str:
        return normalize_header_name(system_header, for_matching=True)

    def _accept_match(vendor_h: str, system_h: str, const_header: str) -> bool:
        """Accept at most one vendor→system pair; system column must be unused."""
        if vendor_h in matched_headers:
            return False
        sys_key = _system_key(system_h)
        if sys_key in used_system_headers:
            return False
        if const_header == "EMPLOYEE_NUMBER" and not _is_valid_employee_number_vendor_match(vendor_h):
            return False
        matched_headers[vendor_h] = system_h
        used_system_headers.add(sys_key)
        constant_headers_status[const_header] = True
        return True
    
    # Step 1: Match ONLY constant headers (exactly one vendor column per constant / system header)
    for const_header in constant_headers:
        const_normalized = normalize_header_name(const_header, for_matching=True)
        matched = False
        
        # Get variations ordered by priority
        variations = constant_mapping_variations.get(const_header, [const_normalized])
        # Always include the constant's own normalized name as a variation
        if const_normalized not in variations:
            variations = [const_normalized] + list(variations)
        
        # Phase 1: Create hints from variations (suggestions, not filters)
        high_confidence_hints = []  # List of (vendor_header, priority_score)
        
        for vendor_header in vendor_headers:
            if vendor_header in matched_headers:
                continue
            vendor_norm = normalize_header_name(vendor_header, for_matching=True)
            
            matches_variation = False
            for variation in variations:
                var_norm = normalize_header_name(variation, for_matching=True)
                if var_norm == vendor_norm or var_norm in vendor_norm or vendor_norm in var_norm:
                    matches_variation = True
                    break
            
            if matches_variation:
                priority_score = 999
                if const_header in priority_keywords and priority_keywords[const_header]:
                    for priority, keyword_list in enumerate(priority_keywords[const_header]):
                        if all(keyword in vendor_norm for keyword in keyword_list):
                            priority_score = priority
                            break
                
                if priority_score == 999:
                    for priority, variation in enumerate(variations):
                        var_norm = normalize_header_name(variation, for_matching=True)
                        if var_norm == vendor_norm or var_norm in vendor_norm or vendor_norm in var_norm:
                            priority_score = priority
                            break
                
                high_confidence_hints.append((vendor_header, priority_score))
        
        high_confidence_hints.sort(key=lambda x: x[1])

        # Phase 1.5: Exact / near-exact match first (one vendor + one unused system column)
        # Prefer vendor name == system name == constant (e.g. da → da for constant DA)
        available_system = [sh for sh in system_headers if _system_key(sh) not in used_system_headers]
        for system_header in available_system:
            system_norm = _system_key(system_header)
            system_matches_const = (
                system_norm == const_normalized
                or any(
                    normalize_header_name(v, for_matching=True) == system_norm
                    or normalize_header_name(v, for_matching=True) in system_norm
                    or system_norm in normalize_header_name(v, for_matching=True)
                    for v in variations
                )
            )
            if not system_matches_const:
                continue
            # Prefer exact vendor name match to this system column
            exact_vendor = None
            for vendor_header in vendor_headers:
                if vendor_header in matched_headers:
                    continue
                if normalize_header_name(vendor_header, for_matching=True) == system_norm:
                    exact_vendor = vendor_header
                    break
            if exact_vendor and _accept_match(exact_vendor, system_header, const_header):
                matched = True
                break
            # Else take highest-priority hint that reasonably matches this system column
            if not matched:
                for vendor_header, _ in high_confidence_hints:
                    vendor_norm = normalize_header_name(vendor_header, for_matching=True)
                    if vendor_norm == system_norm or vendor_norm == const_normalized:
                        if _accept_match(vendor_header, system_header, const_header):
                            matched = True
                            break
            if matched:
                break
        
        # Phase 2: AI only if no exact match — unused vendor/system headers only
        ai_was_called = False
        ai_explicitly_rejected = False
        
        all_vendor_headers = [vh for vh in vendor_headers if vh not in matched_headers]
        all_system_headers = [sh for sh in system_headers if _system_key(sh) not in used_system_headers]
        
        hint_list = [h[0] for h in high_confidence_hints[:5]] if high_confidence_hints else None
        
        if not matched and all_vendor_headers and all_system_headers:
            try:
                ai_was_called = True
                ai_selected, usage = _ai_semantic_matching_for_constant_header(
                    const_header,
                    all_vendor_headers,
                    all_system_headers,
                    variations,
                    hints=hint_list
                )
                if usage:
                    total_input_tokens += usage.get("input_tokens", 0)
                    total_output_tokens += usage.get("output_tokens", 0)
                if ai_selected:
                    for vendor_h, system_h in ai_selected.items():
                        if _accept_match(vendor_h, system_h, const_header):
                            matched = True
                            break
                    if not matched:
                        # AI picked an already-used system column or invalid vendor — treat as no match
                        ai_explicitly_rejected = True
                else:
                    ai_explicitly_rejected = True
            except Exception as e:
                error_str = str(e).lower()
                is_timeout = "timeout" in error_str or "timed out" in error_str
                
                if is_timeout:
                    logger.warning(
                        "AI request timed out, falling back to priority-based matching",
                        extra={"timeout": AZURE_OPENAI_TIMEOUT, "retries": AZURE_OPENAI_MAX_RETRIES, "constant_header": const_header}
                    )
                else:
                    logger.warning(
                        "AI validation failed, falling back to priority-based matching",
                        extra={"error": str(e)[:200], "constant_header": const_header}
                    )
        
        # Phase 3: Fallback direct matching — only if AI service failed (exception)
        if not matched and not ai_explicitly_rejected and ai_was_called:
            for vendor_header, _ in high_confidence_hints:
                if vendor_header in matched_headers:
                    continue
                for system_header in system_headers:
                    if _system_key(system_header) in used_system_headers:
                        continue
                    system_norm = _system_key(system_header)
                    if any(
                        normalize_header_name(var, for_matching=True) in system_norm
                        or system_norm in normalize_header_name(var, for_matching=True)
                        for var in variations
                    ):
                        if _accept_match(vendor_header, system_header, const_header):
                            matched = True
                            break
                if matched:
                    break
        
        # Phase 4: Fallback if AI was never called
        if not matched and not ai_was_called:
            potential_vendor = [vh for vh in vendor_headers if vh not in matched_headers]
            potential_system = [sh for sh in system_headers if _system_key(sh) not in used_system_headers]
            hint_list = [h[0] for h in high_confidence_hints[:5]] if high_confidence_hints else None
            if potential_vendor and potential_system:
                try:
                    ai_matches, usage = _ai_semantic_matching_for_constant_header(
                        const_header, potential_vendor, potential_system, variations, hints=hint_list
                    )
                    if usage:
                        total_input_tokens += usage.get("input_tokens", 0)
                        total_output_tokens += usage.get("output_tokens", 0)
                    if ai_matches:
                        for vendor_h, system_h in ai_matches.items():
                            if _accept_match(vendor_h, system_h, const_header):
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
                            extra={"error": str(e)[:100], "constant_header": const_header}
                        )
        
        if not matched:
            constant_headers_status[const_header] = False
            unmatched_constant_header_names.append(const_header)

    token_usage = (
        {"input_tokens": total_input_tokens, "output_tokens": total_output_tokens}
        if (total_input_tokens or total_output_tokens) else None
    )
    return matched_headers, unmatched_constant_header_names, constant_headers_status, token_usage


# For EMPLOYEE_NUMBER: reject vendor headers that are clearly not identifiers (even if AI returned them).
_EMPLOYEE_NUMBER_REJECT_SUBSTRINGS = (
    "pf", "esi", "esic", "lwf", "pt_amount", "pt_", "ctc", "name", "arrear", "employer_",
    "contribution", "deduction", "allowance", "da", "hra", "basic", "gross", "payable",
    "mobile_charge", "canteen", "uniform", "transport", "bonus", "conveyance", "night_shift",
    "statutory", "compounding", "service_charge", "attendance_bonus", "holiday_days_amount",
)
# Vendor header must suggest identifier (number/id/no/code) to be valid for EMPLOYEE_NUMBER.
_EMPLOYEE_NUMBER_REQUIRE_ONE_OF = ("number", "id", "no", "code", "num")


def _is_valid_employee_number_vendor_match(vendor_header: str) -> bool:
    """
    Return False if vendor_header is clearly not an employee identifier (e.g. employee_pf, employee_esi).
    Used to reject wrong AI matches for EMPLOYEE_NUMBER.
    """
    if not vendor_header:
        return False
    v = str(vendor_header).strip().lower().replace(" ", "_")
    # Reject if it contains any known non-identifier concept
    for sub in _EMPLOYEE_NUMBER_REJECT_SUBSTRINGS:
        if sub in v:
            return False
    # Require at least one identifier-like keyword (number, id, no, code)
    for kw in _EMPLOYEE_NUMBER_REQUIRE_ONE_OF:
        if kw in v:
            return True
    # Allow "casper_id" / "emp_id" style (id is in _REQUIRE_ONE_OF). Also allow exact variation names.
    if v in ("employee_number", "employee_id", "employeenumber", "employeeid", "employee_no", "employeeno", "emp_no", "emp_id", "employee_code", "casper_id"):
        return True
    return False


# Optional overrides: only used when you want custom accept/reject rules for specific headers.
# If a header is not here, generic context (infer from name + synonyms) is used.
SEMANTIC_CONTEXT_OVERRIDES: Dict[str, str] = {
    "EMPLOYEE_NUMBER": """
BUSINESS MEANING: Unique identifier for an employee (e.g., "EMP001", "12345", "10007")
- This is an IDENTIFIER only: a code/number that uniquely identifies the person, not a name or any amount/deduction
- Usually text/string (may have leading zeros). Column holds values like employee IDs, not money or percentages
- MATCH only: "employee_number", "employee_id", "emp_no", "employee_code", "casper_id", "emp_id" (identifier-like names)
- REJECT: "employee_name", "employee_name_full" (these are names, not IDs)
- REJECT: "employee_pf", "employee_esic", "employee_lwf", "employer_pf", "employer_esic", "employer_lwf" (these are contribution/deduction amounts, NOT identifiers)
- REJECT: Any header that is clearly a deduction, contribution, or amount (pf, esi, lwf, pt, ctc, da, hra, allowance, etc.). EMPLOYEE_NUMBER must be the ID column only.
- If the only candidate vendor headers are things like employee_pf, employee_esi, etc., return {{}} (empty) - do NOT match. No match is better than a wrong match.
""",
    "NET_PAY": """
BUSINESS MEANING: Employee's take-home pay AFTER all deductions (taxes, insurance, PF, etc.)
- This is the FINAL amount the employee receives. Examples: "net_pay", "netpay", "take_home", "net_amount"
- REJECT: "gross_pay", "total_payable", "basic_pay" (different concepts)
""",
    "INVOICE_VALUE": """
BUSINESS MEANING: The monetary value/amount on an invoice or bill
- This is the AMOUNT to be paid, not a count or number
- Examples: "invoice_value", "invoice_amount", "invoice_total", "bill_amount"
- REJECT: "invoice_number" (ID), "invoice_count" (quantity), "invoice_date" (date)
""",
}


def _get_semantic_context_for_header(const_header: str) -> str:
    """
    Get semantic context for a constant header. Uses optional override if present,
    otherwise returns the generic context (infer meaning from header name + allow synonyms).
    No config is required for new columns.
    """
    if const_header in SEMANTIC_CONTEXT_OVERRIDES:
        return SEMANTIC_CONTEXT_OVERRIDES[const_header]
    # Generic context: works for any constant header without hardcoded config
    return f"""
BUSINESS MEANING: Infer from the constant header name "{const_header}".
- In payroll/HR data, this header likely represents a real-world concept (identifier, amount, category, date, etc.)
- The SAME concept can appear under different names: synonyms or equivalent terms (e.g. different words for the same idea), or different wording
- MATCH a vendor header to this constant if they represent the SAME business concept, even when the names differ
- REJECT only when the vendor header clearly means something different (e.g. name vs id, gross vs net)
- When the constant header name and a vendor header are synonyms or equivalent terms, that is a valid match
"""


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
        Tuple of (matches_dict, usage_dict).
        usage_dict: {"input_tokens": int, "output_tokens": int} or None.
    """
    if not potential_vendor or not potential_system:
        return {}, None
    
    # Use sorted copies so the prompt is identical for the same set of headers (deterministic AI behavior)
    potential_vendor = sorted(potential_vendor)
    potential_system = sorted(potential_system)
    
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
    
    # Optional hints (from config variations when available); AI must still consider ALL headers
    # Sort hints so prompt is deterministic for same inputs
    hints_text = ", ".join(sorted(hints)) if hints else "None - analyze all headers independently"
    variations_text = ", ".join(expected_variations) if expected_variations else "None - infer from constant header name"
    
    prompt = f"""
You are an intelligent data assistant matching column headers for payroll/paysheet data.
You must find the best semantic match for ONE constant header by analyzing ALL vendor and system headers.

CONSTANT HEADER TO MATCH: "{const_header}"

{semantic_context}

OPTIONAL (use only as hints; you may match a header not listed here):
- Expected variations: {variations_text}
- High-confidence hint headers: {hints_text}

ALL VENDOR HEADERS (you MUST consider every one; match can be any of these):
{potential_vendor}

ALL SYSTEM HEADERS (pick the system header that best corresponds to the constant "{const_header}"):
{potential_system}

RULES:
1. Infer the business meaning of "{const_header}" from its name and the guidance above
2. Return EXACTLY ONE pair: one vendor header → one system header for this constant only
3. Prefer an EXACT name match when available (e.g. constant "DA" / system "da" → vendor "da", NOT "el_days" or "arrear_da")
4. Do NOT map a different concept to this constant (e.g. do not map "fixed_basic" to "basic" when matching constant "FIXED_BASIC" if "fixed_basic" exists as its own system header)
5. The system header must be the one that corresponds to THIS constant (often the constant name itself or a close synonym), not a related sibling column
6. When unsure, return {{}} (empty object)

Return JSON only: {{"vendor_header": "system_header"}} for one match, or {{}} if no good match.
"""
    
    # Retry logic with exponential backoff for timeout errors
    last_error = None
    usage_out = None
    for attempt in range(AZURE_OPENAI_MAX_RETRIES + 1):
        try:
            response = llm.invoke([HumanMessage(content=prompt)])
            raw_output = response.content
            um = getattr(response, "usage_metadata", None) or (getattr(response, "response_metadata", None) or {}).get("usage_metadata")
            if um is not None:
                usage_out = {
                    "input_tokens": getattr(um, "input_tokens", None) if not isinstance(um, dict) else um.get("input_tokens"),
                    "output_tokens": getattr(um, "output_tokens", None) if not isinstance(um, dict) else um.get("output_tokens"),
                }
                if usage_out["input_tokens"] is None:
                    usage_out["input_tokens"] = 0
                if usage_out["output_tokens"] is None:
                    usage_out["output_tokens"] = 0
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
            return {}, usage_out

    result = {
        k: v for k, v in semantic_matches.items()
        if v and str(v).lower() not in ["null", "none", "no match", ""]
    }
    return result, usage_out


