import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
def get_base_path():
    """Get the base path for loading resources, compatible with PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        return Path(sys.executable).parent
    else:
        # Running in a normal Python environment
        # We assume this script is in src/marketplace_automation/
        # and the .env is in the project root.
        return Path(__file__).parent.parent.parent

# Determine path to .env file
base_path = get_base_path()
env_path = base_path / '.env'

# Load .env file
load_dotenv(dotenv_path=env_path)

class Config:
    """Application configuration."""

    # Email Configuration
    EMAIL_SENDER = os.getenv("EMAIL_SENDER", "")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
    SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
    DEFAULT_RECIPIENT = os.getenv("DEFAULT_RECIPIENT", "")

    # Process CC recipients list
    _cc_recipients_str = os.getenv("CC_RECIPIENTS", "")
    if _cc_recipients_str:
        CC_RECIPIENTS = [email.strip() for email in _cc_recipients_str.split(",") if email.strip()]
    else:
        CC_RECIPIENTS = []

    # App Expiration
    EXPIRY_DATE = os.getenv("EXPIRY_DATE", "31-06-2027")
