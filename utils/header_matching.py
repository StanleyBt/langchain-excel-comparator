# excel_comparator/utils/header_matching.py
import pandas as pd
import json
import re
from typing import Dict, List, Tuple, Optional
from langchain_openai import AzureChatOpenAI
from langchain.schema import HumanMessage
from config import AZURE_DEPLOYMENT_NAME, AZURE_ENDPOINT, AZURE_API_KEY, AZURE_API_VERSION
from config import CONSTANT_HEADERS, HEADER_MATCHING_PRIORITY, HEADER_VARIATIONS, AI_MATCHING_KEYWORDS

# Removed Streamlit-specific functions: find_header_row, detect_header_row, run_header_mapping
# These were only used by the Streamlit UI app which has been removed

def normalize_header_for_matching(header: str) -> str:
    """
    Normalize header name for matching comparison.
    Removes spaces, converts to lowercase, handles common variations.
    """
    return str(header).strip().lower().replace(" ", "_").replace("-", "_").replace(".", "")


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
        
        # Find ALL potential vendor headers that could match (with priority)
        potential_matches = []  # List of (vendor_header, priority_score)
        
        for vendor_header in vendor_headers:
            vendor_norm = normalize_header_for_matching(vendor_header)
            
            # Check if vendor header matches any variation
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
                
                potential_matches.append((vendor_header, priority_score))
        
        # Sort by priority (lower number = better match)
        potential_matches.sort(key=lambda x: x[1])
        
        # Use AI to validate and select the best match when we have candidates
        # AI can distinguish semantic differences (e.g., "contractor name" vs "contractor id")
        ai_was_called = False
        ai_explicitly_rejected = False
        
        if potential_matches:
            # Get top candidates (same priority level or close) - include at least top 3 or all if fewer
            top_priority = potential_matches[0][1]
            top_candidates = [vh for vh, pri in potential_matches if pri <= top_priority + 1]
            
            # If we have candidates, use AI to validate semantic correctness
            # This ensures even single candidates are validated (e.g., "contractor id" will be rejected)
            try:
                ai_was_called = True
                ai_selected = _ai_semantic_matching_for_constant_header(
                    const_header, top_candidates, system_headers, variations
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
                print(f"Warning: AI validation failed, falling back to priority-based matching: {e}")
                # AI service failed, but we'll still try fallback
        
        # Only fall back to direct matching if:
        # 1. AI service failed (exception) OR
        # 2. AI was never called (no potential matches found)
        # DO NOT fall back if AI explicitly rejected candidates (empty dict)
        if not matched and not ai_explicitly_rejected:
            # Try to match with system headers, starting with best vendor header match
            for vendor_header, _ in potential_matches:
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
        
        # If not found with exact variations, use AI to intelligently match
        # AI can distinguish between semantically different headers (e.g., "contractor id" vs "contractor name")
        if not matched:
            # Collect all vendor headers that could potentially match (based on keywords)
            potential_vendor = []
            for vendor_header in vendor_headers:
                vendor_norm = normalize_header_for_matching(vendor_header)
                # Check if it's somewhat related (contains key words from config)
                words = AI_MATCHING_KEYWORDS.get(const_header, [])
                if any(word in vendor_norm for word in words):
                    potential_vendor.append(vendor_header)
            
            # Find potential system headers
            potential_system = []
            for system_header in system_headers:
                system_norm = normalize_header_for_matching(system_header)
                words = AI_MATCHING_KEYWORDS.get(const_header, [])
                if any(word in system_norm for word in words):
                    potential_system.append(system_header)
            
            # Use AI to intelligently match - AI will distinguish semantic differences
            if potential_vendor and potential_system:
                try:
                    ai_matches = _ai_semantic_matching_for_constant_header(
                        const_header, potential_vendor, potential_system, variations
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
                    print(f"Warning: AI semantic matching failed for {const_header}: {e}")
        
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
    expected_variations: List[str]
) -> Dict[str, str]:
    """
    Use AI to semantically match headers for a specific constant header.
    AI will intelligently distinguish between semantically different headers.
    For example: "contractor name" vs "contractor id" - AI knows "name" is better for CONTRACTOR.
    
    Args:
        const_header: The constant header being matched (e.g., "CONTRACTOR")
        potential_vendor: List of vendor headers that might match
        potential_system: List of system headers that might match
        expected_variations: List of expected variations for this constant header
        
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
        api_version=AZURE_API_VERSION
    )
    
    # Get semantic context for the constant header
    semantic_context = _get_semantic_context_for_header(const_header)
    
    # Create a smart prompt that helps AI understand semantic differences
    prompt = f"""
You are a data assistant matching column headers for payroll/paysheet data.

CONSTANT HEADER TO MATCH: "{const_header}"

{semantic_context}

Expected variations for this header: {expected_variations}

Vendor headers that might match:
{potential_vendor}

System headers that might match:
{potential_system}

CRITICAL SEMANTIC MATCHING RULES:
1. Match vendor headers to system headers ONLY if they represent the SAME REAL-WORLD CONCEPT
2. Understand the BUSINESS MEANING of "{const_header}" - not just keywords
3. STRICTLY REJECT semantically incorrect matches based on these patterns:

   TYPE MISMATCHES (REJECT these):
   - NAME vs IDENTIFIER: If constant header expects a NAME (text like "John Doe", "ABC Company"), 
     REJECT headers with: "id", "number", "code", "identifier", "ref", "reference"
     ✅ ACCEPT: "name", "title", "label", "description"
     ❌ REJECT: "id", "number", "code", "identifier"
   
   - IDENTIFIER vs NAME: If constant header expects an ID/NUMBER (like "EMP001", "12345"),
     REJECT headers with: "name", "title", "label", "description"
     ✅ ACCEPT: "id", "number", "no", "code", "identifier"
     ❌ REJECT: "name", "title", "label"
   
   - AMOUNT vs COUNT: If constant header expects an AMOUNT (monetary value),
     REJECT headers with: "count", "quantity", "qty", "number of", "total items"
     ✅ ACCEPT: "amount", "value", "price", "cost"
     ❌ REJECT: "count", "quantity", "qty"
   
   - NET vs GROSS vs TOTAL vs PAYABLE: Understand the payroll context:
     * NET_PAY = Employee's take-home pay AFTER all deductions (taxes, insurance, etc.)
     * GROSS_PAY = Employee's pay BEFORE deductions
     * TOTAL_PAYABLE = Total amount to be paid (could be gross, could include other components)
     * BASIC_PAY = Base salary component (before allowances/deductions)
     * If constant header is NET_PAY, REJECT: "gross", "total payable", "basic", "total pay"
     * If constant header is NET_PAY, ACCEPT: "net pay", "netpay", "net salary", "take home"
     * Understand: "total_payable" ≠ "net_pay" (total payable is what needs to be paid, net pay is after deductions)
   
   - PERCENTAGE vs DECIMAL: If constant header expects PERCENTAGE,
     REJECT headers with: "decimal", "ratio", "fraction" (unless they also say "percent")
     ✅ ACCEPT: "percentage", "percent", "%", "rate"
     ❌ REJECT: "decimal", "ratio" (without percent context)
   
   - DATE vs DATE_STRING: If constant header expects a DATE,
     REJECT headers with: "date string", "date text", "formatted date" (unless it's the only option)
     ✅ ACCEPT: "date", "dob", "doj", "timestamp"
     ⚠️ ACCEPT "date string" only if no better match exists

4. GENERAL RULES:
   - If vendor header contains "id"/"number"/"code" but constant header needs "name" → REJECT
   - If vendor header contains "name" but constant header needs "id"/"number" → REJECT
   - If vendor header contains "count" but constant header needs "amount" → REJECT
   - If vendor header contains "amount" but constant header needs "count" → REJECT
   - Only return matches that make semantic sense
   - If no good semantic match exists, return an empty JSON object {{}}

5. REAL-WORLD CONTEXT AWARENESS:
   - Think about what this column represents in actual payroll/paysheet systems
   - Consider business logic: Would these two columns have the same value in real payroll?
   - If the meanings are different, even if keywords overlap, REJECT the match
   - When in doubt, REJECT - it's better to require manual mapping than create incorrect matches

6. PAYROLL-SPECIFIC UNDERSTANDING:
   - Employee Number: Unique identifier (text, may have leading zeros) - NOT employee name
   - Net Pay: Take-home amount after ALL deductions - NOT gross, NOT total payable, NOT basic
   - Invoice Value: Amount on invoice/bill - NOT invoice count, NOT invoice number
   - Understand that payroll terms have specific meanings - don't match based on partial keyword overlap

⚠️ Return output as JSON only, in this format:
{{"vendor_header": "system_header"}} or {{}} if no good match

IMPORTANT: If you're unsure whether two headers represent the same concept, return {{}} (empty object).
It's better to require manual mapping than to create an incorrect automatic match.
"""
    
    response = llm.invoke([HumanMessage(content=prompt)])
    raw_output = response.content
    
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


def _ai_semantic_matching(unmatched_vendor: List[str], unmatched_system: List[str]) -> Dict[str, str]:
    """
    Use AI to semantically match unmatched headers.
    
    Args:
        unmatched_vendor: List of vendor headers that couldn't be matched
        unmatched_system: List of system headers available for matching
        
    Returns:
        Dictionary mapping vendor_header -> system_header
    """
    if not unmatched_vendor or not unmatched_system:
        return {}
    
    llm = AzureChatOpenAI(
        deployment_name=AZURE_DEPLOYMENT_NAME,
        temperature=0,
        azure_endpoint=AZURE_ENDPOINT,
        api_key=AZURE_API_KEY,
        api_version=AZURE_API_VERSION
    )
    
    prompt = f"""
You are a data assistant comparing column headers between two paysheet systems.

These are headers from a Vendor Paysheet that had no exact match:
{unmatched_vendor}

And these are remaining headers from the System Paysheet:
{unmatched_system}

Your task:
- Match each Vendor header to the most semantically similar System header.
- If no good match exists, set the value to null.
- Use domain knowledge of payroll systems. For example:
  - "employee id" could be "employee number", "cems employee id", or "blue tree id".
  - "gross salary" could be "fixed gross" or "ctc".
  - "contractor name" could be "contractor" or "vendor".

⚠️ Return output as JSON only, in this format:
{{"vendor_header_1": "system_header_1", "vendor_header_2": "system_header_2", ...}}
"""
    
    response = llm.invoke([HumanMessage(content=prompt)])
    raw_output = response.content
    
    try:
        semantic_matches = json.loads(raw_output)
    except json.JSONDecodeError:
        match = re.search(r"{.*}", raw_output, re.DOTALL)
        if match:
            semantic_matches = json.loads(match.group())
        else:
            raise ValueError("❌ GPT did not return valid JSON or parsable output.")
    
    # Filter out null/None values and "no match" strings
    return {
        k: v for k, v in semantic_matches.items() 
        if v and str(v).lower() not in ["null", "none", "no match", ""]
    }
