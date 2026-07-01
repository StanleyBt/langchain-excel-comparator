# utils/logger.py
"""
Structured logging utility with GDPR-safe data masking.

This module provides:
- Structured JSON logging
- Automatic masking of sensitive data (PII)
- Correlation IDs for request tracing
- GDPR-compliant logging practices
"""
import logging
import json
import sys
from typing import Any, Dict, Optional
from datetime import datetime


class SensitiveDataMasker:
    """Masks sensitive data from logs for GDPR compliance"""
    
    # Fields that contain sensitive data (PII)
    SENSITIVE_FIELDS = {
        'employeeNumber', 'employee_number', 'employee_id', 'emp_id', 'emp_no',
        'vendorValue', 'systemValue', 'vendor_value', 'system_value',
        'details', 'vendorPaysheetData', 'systemData',
        'request_body', 'body', 'data', 'payload'
    }
    
    # Patterns to detect sensitive data in nested structures
    SENSITIVE_PATTERNS = [
        'employee', 'pay', 'salary', 'amount', 'value', 'net_pay', 'invoice',
        'contractor', 'vendor', 'paysheet', 'details'
    ]
    
    @classmethod
    def mask_value(cls, value: Any, field_name: str = "") -> Any:
        """
        Mask sensitive values based on field name or content.
        
        Args:
            value: Value to potentially mask
            field_name: Name of the field (used to detect sensitive fields)
            
        Returns:
            Masked value or original value if not sensitive
        """
        if value is None:
            return None
        
        # Check if field name indicates sensitive data
        field_lower = field_name.lower() if field_name else ""
        is_sensitive_field = any(
            sensitive in field_lower for sensitive in cls.SENSITIVE_FIELDS
        )
        
        # Mask based on field name
        if is_sensitive_field:
            if isinstance(value, (dict, list)):
                return "[MASKED: Contains sensitive data]"
            elif isinstance(value, str) and len(value) > 0:
                # Show first 2 and last 2 characters, mask the rest
                if len(value) <= 4:
                    return "****"
                return f"{value[:2]}...{value[-2:]}"
            elif isinstance(value, (int, float)):
                return "[MASKED: Numeric value]"
            else:
                return "[MASKED]"
        
        # For nested structures, recursively mask sensitive fields
        if isinstance(value, dict):
            return {
                k: cls.mask_value(v, k) 
                for k, v in value.items()
            }
        elif isinstance(value, list):
            return [cls.mask_value(item, field_name) for item in value]
        
        return value
    
# LogRecord internals — never emit as top-level JSON fields
_STANDARD_RECORD_ATTRS = frozenset(
    logging.LogRecord("", 0, "", 0, "", (), None).__dict__.keys()
)


class StructuredFormatter(logging.Formatter):
    """Compact JSON formatter for structured logging"""

    def format(self, record: logging.LogRecord) -> str:
        log_data = {
            "time": datetime.utcnow().strftime("%H:%M:%S"),
            "level": record.levelname,
            "msg": record.getMessage(),
        }

        if hasattr(record, "correlation_id"):
            log_data["cid"] = record.correlation_id

        if record.exc_info:
            log_data["error"] = self.formatException(record.exc_info)

        masker = SensitiveDataMasker()
        for key, value in record.__dict__.items():
            if key not in _STANDARD_RECORD_ATTRS and key != "correlation_id":
                log_data[key] = masker.mask_value(value, key)

        return json.dumps(log_data, default=str, separators=(",", ":"))


def setup_logging(
    level: str = "INFO",
    use_json: bool = True,
    log_file: Optional[str] = None
) -> logging.Logger:
    """
    Set up structured logging for the application.
    
    Args:
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        use_json: Whether to use JSON formatting
        log_file: Optional file path for file logging
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger("paysheet_comparator")
    logger.setLevel(getattr(logging, level.upper()))
    
    # Remove existing handlers
    logger.handlers.clear()
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, level.upper()))
    
    if use_json:
        console_handler.setFormatter(StructuredFormatter())
    else:
        # Simple format for development
        simple_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        console_handler.setFormatter(simple_formatter)
    
    logger.addHandler(console_handler)
    
    # File handler if specified
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(getattr(logging, level.upper()))
        if use_json:
            file_handler.setFormatter(StructuredFormatter())
        else:
            file_handler.setFormatter(simple_formatter)
        logger.addHandler(file_handler)
    
    # Prevent propagation to root logger
    logger.propagate = False
    
    return logger


def get_logger(name: str = "paysheet_comparator") -> logging.Logger:
    """
    Get a logger instance.
    
    Args:
        name: Logger name
        
    Returns:
        Logger instance
    """
    return logging.getLogger(name)
