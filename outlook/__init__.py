"""Outlook support for the hermes-outlook plugin.

Talks to classic Outlook (desktop) through a Windows PowerShell COM bridge.
"""

from . import client
from .tools import HANDLERS

__all__ = ["HANDLERS", "client"]
