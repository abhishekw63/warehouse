import pandas as pd

def format_indian(number):
    """Format number in Indian numbering system (lakhs, crores)."""
    # Handle case where the number might already be a formatted string
    if isinstance(number, str):
        # Remove rupee symbol and existing commas
        number = number.replace('₹', '').replace(',', '').strip()
    
    try:
        s = str(int(float(number)))
    except (ValueError, TypeError):
        return str(number)  # Return as-is if cannot convert
    
    if len(s) <= 3:
        return s
    last3 = s[-3:]
    remaining = s[:-3]
    parts = []
    while len(remaining) > 2:
        parts.insert(0, remaining[-2:])
        remaining = remaining[:-2]
    if remaining:
        parts.insert(0, remaining)
    return ','.join(parts) + ',' + last3

def safe_ean_convert(x):
    """Convert EAN safely from scientific notation."""
    if pd.isna(x):
        return ""
    try:
        # Convert to string first, then to float, then to int
        return str(int(float(str(x))))
    except Exception:
        return str(x)
