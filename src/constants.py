"""
    The Outlook API has a number of magic values where objects are referenced by an interger value
"""
from src.config import interface_dict

MAIL_ITEM = 0

# Set in configuration files or overwritten in cmdline

DISPLAY = interface_dict.get(
    "display", True
)  # instead of sending in background each email is created but you must press the send button yourself
