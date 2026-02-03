# utils/statistics_calculator.py
from typing import List, Dict, Any, Optional, Tuple
from models.schemas import (
    RowComparisonResult, ColumnComparison, OverallSummary, ColumnStatistic,
    ColumnMismatchDetail, Insights, KeyFinding, ColumnHealth, QuickStats
)
from utils.comparison_engine import is_text_only_column


def calculate_basic_summary(
    row_comparisons: List[RowComparisonResult], 
    columns_compared: int,
    only_in_vendor_count: int = 0,
    only_in_system_count: int = 0
) -> Dict[str, Any]:
    """
    Calculate basic summary statistics for comparison response.
    
    This function provides the summary structure needed by the API response,
    which is different from calculate_overall_summary (which provides enhanced stats).
    """
    total_rows = len(row_comparisons)
    matched_rows = sum(1 for r in row_comparisons if r.overallMatch)
    unmatched_rows = total_rows - matched_rows
    
    # Calculate match rate
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    total_comparisons = sum(len(r.columnComparisons) for r in matched_employees)
    matched_comparisons = sum(
        sum(1 for cc in r.columnComparisons if cc.isMatch)
        for r in matched_employees
    )
    match_rate = (matched_comparisons / total_comparisons) if total_comparisons > 0 else 0.0
    
    return {
        "totalRows": total_rows,
        "matchedRows": matched_rows,
        "unmatchedRows": unmatched_rows,
        "onlyInVendor": only_in_vendor_count,
        "onlyInSystem": only_in_system_count,
        "matchRate": match_rate,
        "columnsCompared": columns_compared
    }


def calculate_overall_summary(
    row_comparisons: List[RowComparisonResult], 
    column_statistics: List[ColumnStatistic]
) -> OverallSummary:
    """Calculate simplified overall summary statistics"""
    # Total rows validated (all employees processed)
    total_rows = len(row_comparisons)
    
    # Calculate overall match percentage
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    total_comparisons = sum(len(r.columnComparisons) for r in matched_employees)
    matched_comparisons = sum(
        sum(1 for cc in r.columnComparisons if cc.isMatch)
        for r in matched_employees
    )
    overall_match_percentage = (matched_comparisons / total_comparisons * 100) if total_comparisons > 0 else 0.0
    
    # Create column match percentages dictionary
    column_match_percentages = {
        col_stat.columnName: col_stat.matchRate * 100  # Convert to percentage
        for col_stat in column_statistics
    }
    
    return OverallSummary(
        totalRowsValidated=total_rows,
        totalMatchPercentage=round(overall_match_percentage, 2),
        columnMatchPercentages=column_match_percentages
    )


def calculate_column_statistics(
    row_comparisons: List[RowComparisonResult],
    top_n: int = 5
) -> List[ColumnStatistic]:
    """Calculate statistics for each column"""
    # Group comparisons by column name, storing employee number with each comparison
    column_data: Dict[str, List[Tuple[ColumnComparison, str]]] = {}  # (comparison, employee_number)
    
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    for row_comp in matched_employees:
        for col_comp in row_comp.columnComparisons:
            col_name = col_comp.columnName
            if col_name not in column_data:
                column_data[col_name] = []
            column_data[col_name].append((col_comp, row_comp.employeeNumber))
    
    column_stats = []
    for col_name, comparison_tuples in column_data.items():
        comparisons = [cc for cc, _ in comparison_tuples]
        total = len(comparisons)
        matched = sum(1 for cc in comparisons if cc.isMatch)
        mismatched = total - matched
        match_rate = matched / total if total > 0 else 0.0
        
        # Calculate numeric statistics (rounded to 2 decimal places)
        numeric_diffs = [cc.difference for cc in comparisons if cc.difference is not None]
        total_diff = round(sum(numeric_diffs), 2) if numeric_diffs else None
        avg_diff = round(sum(numeric_diffs) / len(numeric_diffs), 2) if numeric_diffs else None
        max_diff = round(max(numeric_diffs), 2) if numeric_diffs else None
        min_diff = round(min(numeric_diffs), 2) if numeric_diffs else None
        
        # Calculate vendor and system sums (for numeric columns only)
        # Skip text-only columns (e.g., Employee Number, Contractor)
        vendor_sum = None
        system_sum = None
        
        # Check if this is a text-only column - don't calculate sums for text columns
        is_text_column = is_text_only_column(col_name)
        
        if not is_text_column:
            # Try to calculate sums for numeric values (only for non-text columns)
            vendor_numeric_values = []
            system_numeric_values = []
            
            for cc in comparisons:
                # Try to convert vendor value to float
                if cc.vendorValue is not None:
                    try:
                        vendor_val = float(cc.vendorValue)
                        vendor_numeric_values.append(vendor_val)
                    except (ValueError, TypeError):
                        pass
                
                # Try to convert system value to float
                if cc.systemValue is not None:
                    try:
                        system_val = float(cc.systemValue)
                        system_numeric_values.append(system_val)
                    except (ValueError, TypeError):
                        pass
            
            # Only set sums if we have numeric values (indicates numeric column)
            # Round to 2 decimal places for numeric columns
            if vendor_numeric_values:
                vendor_sum = round(sum(vendor_numeric_values), 2)
            if system_numeric_values:
                system_sum = round(sum(system_numeric_values), 2)
        
        # Get top mismatches with employee numbers
        mismatch_tuples = [(cc, emp_id) for cc, emp_id in comparison_tuples if not cc.isMatch]
        mismatch_tuples.sort(key=lambda x: abs(x[0].difference) if x[0].difference is not None else 0, reverse=True)
        
        top_mismatches = [
            ColumnMismatchDetail(
                employeeNumber=emp_id,
                vendorValue=cc.vendorValue,
                systemValue=cc.systemValue,
                difference=round(cc.difference, 2) if cc.difference is not None else None
            )
            for cc, emp_id in mismatch_tuples[:top_n]
        ]
        
        # Determine if this is Employee Number/Head Count column
        col_name_lower = col_name.lower().strip()
        is_employee_number = (
            "employee" in col_name_lower and 
            ("number" in col_name_lower or "no" in col_name_lower or "id" in col_name_lower or "head count" in col_name_lower)
        )
        headcount = matched if is_employee_number else None
        
        # For Employee Head Count column, show count of employees instead of sum (as whole number)
        if is_employee_number:
            # Set counts to the number of employees (rows) - use integer, not float
            vendor_sum = int(total)  # Count of employees (whole number)
            system_sum = int(total)  # Count of employees (whole number, same for both vendor and system)
        
        column_stats.append(ColumnStatistic(
            columnName=col_name,
            totalComparisons=total,
            matchedCount=matched,
            mismatchedCount=mismatched,
            matchRate=match_rate,
            totalDifference=total_diff,
            averageDifference=avg_diff,
            maxDifference=max_diff,
            minDifference=min_diff,
            mismatchPercentage=(mismatched / total * 100) if total > 0 else 0.0,
            topMismatches=top_mismatches,
            vendorPaysheetCount=vendor_sum,
            systemPaysheetCount=system_sum,
            headcount=headcount
        ))
    
    return column_stats


def calculate_insights(
    row_comparisons: List[RowComparisonResult],
    column_stats: List[ColumnStatistic]
) -> Insights:
    """Calculate insights and recommendations"""
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    total_employees = len(matched_employees)
    
    # Calculate overall health score
    if total_employees == 0:
        health_score = 0.0
    else:
        perfect_matches = sum(1 for r in matched_employees if r.overallMatch)
        health_score = perfect_matches / total_employees
    
    # Determine overall health
    if health_score >= 0.95:
        overall_health = "excellent"
    elif health_score >= 0.80:
        overall_health = "good"
    elif health_score >= 0.60:
        overall_health = "fair"
    else:
        overall_health = "poor"
    
    # Generate key findings
    key_findings = []
    
    # Check for columns with high mismatch rates
    for col_stat in column_stats:
        if col_stat.mismatchPercentage > 20:
            affected_emps = col_stat.mismatchedCount
            key_findings.append(KeyFinding(
                type="warning",
                severity="high" if col_stat.mismatchPercentage > 50 else "medium",
                message=f"{col_stat.columnName} has {col_stat.mismatchPercentage:.1f}% mismatch rate ({col_stat.mismatchedCount} mismatches)",
                affectedEmployees=affected_emps,
                affectedColumns=[col_stat.columnName],
                suggestion=f"Review {col_stat.columnName} calculations and mappings"
            ))
        elif col_stat.matchRate == 1.0:
            key_findings.append(KeyFinding(
                type="info",
                severity="low",
                message=f"{col_stat.columnName} has 100% match rate",
                affectedEmployees=col_stat.totalComparisons,
                affectedColumns=[col_stat.columnName]
            ))
    
    # Check for large discrepancies
    for col_stat in column_stats:
        if col_stat.maxDifference and abs(col_stat.maxDifference) > 5000:
            key_findings.append(KeyFinding(
                type="warning",
                severity="medium",
                message=f"{col_stat.columnName} has significant discrepancies (max difference: {col_stat.maxDifference:.2f})",
                affectedEmployees=col_stat.mismatchedCount,
                affectedColumns=[col_stat.columnName],
                suggestion=f"Review {col_stat.columnName} for employees with differences >5000"
            ))
    
    # Calculate column health
    column_health_list = []
    for col_stat in column_stats:
        if col_stat.matchRate >= 0.95:
            health_status = "excellent"
        elif col_stat.matchRate >= 0.80:
            health_status = "good"
        elif col_stat.matchRate >= 0.60:
            health_status = "needs_attention"
        else:
            health_status = "critical"
        
        issues = []
        if col_stat.mismatchPercentage > 20:
            issues.append(f"High mismatch rate: {col_stat.mismatchPercentage:.1f}%")
        if col_stat.averageDifference and abs(col_stat.averageDifference) > 100:
            issues.append(f"Average difference exceeds threshold: {col_stat.averageDifference:.2f}")
        if col_stat.maxDifference and abs(col_stat.maxDifference) > 5000:
            issues.append(f"Large discrepancies detected: max {col_stat.maxDifference:.2f}")
        
        column_health_list.append(ColumnHealth(
            columnName=col_stat.columnName,
            healthStatus=health_status,
            matchRate=col_stat.matchRate,
            averageDifference=col_stat.averageDifference,
            issues=issues
        ))
    
    # Generate recommendations
    recommendations = []
    
    # Find problematic columns
    problematic_cols = [ch for ch in column_health_list if ch.healthStatus in ["needs_attention", "critical"]]
    if problematic_cols:
        recommendations.append(f"Review {len(problematic_cols)} column(s) with low match rates: {', '.join([c.columnName for c in problematic_cols[:3]])}")
    
    # Check for large differences
    large_diff_cols = [cs for cs in column_stats if cs.maxDifference and abs(cs.maxDifference) > 5000]
    if large_diff_cols:
        recommendations.append(f"Investigate large discrepancies in: {', '.join([c.columnName for c in large_diff_cols[:3]])}")
    
    # Check for missing data
    missing_data = sum(1 for r in matched_employees for cc in r.columnComparisons 
                      if cc.vendorValue is None or cc.systemValue is None)
    if missing_data > 0:
        recommendations.append(f"Address {missing_data} missing data points across all columns")
    
    # Data quality metrics
    total_comparisons = sum(len(r.columnComparisons) for r in matched_employees)
    null_values = sum(1 for r in matched_employees for cc in r.columnComparisons 
                     if cc.vendorValue is None or cc.systemValue is None)
    data_completeness = 1.0 - (null_values / total_comparisons) if total_comparisons > 0 else 0.0
    
    return Insights(
        overallHealth=overall_health,
        healthScore=health_score,
        keyFindings=key_findings,
        columnHealth=column_health_list,
        recommendations=recommendations if recommendations else ["All columns are within acceptable thresholds"],
        dataQuality={
            "missingDataCount": missing_data,
            "nullValuesCount": null_values,
            "dataCompleteness": data_completeness
        }
    )


def calculate_quick_stats(
    row_comparisons: List[RowComparisonResult],
    column_stats: List[ColumnStatistic]
) -> QuickStats:
    """Calculate quick statistics dashboard"""
    matched_employees = [r for r in row_comparisons if r.rowStatus == "matched"]
    
    perfect_matches = sum(1 for r in matched_employees if r.overallMatch)
    partial_matches = sum(1 for r in matched_employees if not r.overallMatch and any(cc.isMatch for cc in r.columnComparisons))
    no_matches = sum(1 for r in matched_employees if not any(cc.isMatch for cc in r.columnComparisons))
    
    # Find most problematic and best matching columns
    most_problematic = None
    best_matching = None
    
    if column_stats:
        most_problematic = min(column_stats, key=lambda x: x.matchRate)
        best_matching = max(column_stats, key=lambda x: x.matchRate)
    
    # Find largest discrepancy
    largest_discrepancy = None
    max_diff = 0.0
    for row_comp in matched_employees:
        for col_comp in row_comp.columnComparisons:
            if col_comp.difference is not None and abs(col_comp.difference) > abs(max_diff):
                max_diff = col_comp.difference
                largest_discrepancy = {
                    "employeeNumber": row_comp.employeeNumber,
                    "column": col_comp.columnName,
                    "difference": col_comp.difference,
                    "vendorValue": col_comp.vendorValue,
                    "systemValue": col_comp.systemValue
                }
    
    return QuickStats(
        perfectMatches=perfect_matches,
        partialMatches=partial_matches,
        noMatches=no_matches,
        mostProblematicColumn=most_problematic.columnName if most_problematic else None,
        bestMatchingColumn=best_matching.columnName if best_matching else None,
        largestDiscrepancy=largest_discrepancy
    )

