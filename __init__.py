"""hermes-outlook — read email from classic Outlook via a PowerShell COM bridge.

Registers the ``outlook_*`` tools and the bundled ``outlook-mail`` skill. Windows-only:
it drives classic Outlook (desktop) through its COM automation interface, so it works
against every mailbox the current Outlook profile can see (accounts, shared mailboxes,
PST archives) without any Graph/M365 admin consent.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict

from . import schemas
from .outlook import client
from .outlook.tools import HANDLERS

logger = logging.getLogger(__name__)

TOOLSET = "outlook"
SKILLS_DIR = Path(__file__).parent / "skills"

__all__ = ["TOOLSET", "register"]


def _schema_map() -> Dict[str, Dict[str, Any]]:
    return {schema["name"]: schema for schema in schemas.TOOL_SCHEMAS}


def _register_skills(ctx: Any) -> None:
    """Expose ``skills/*/SKILL.md`` as namespaced plugin skills."""
    if not SKILLS_DIR.is_dir():
        return
    for child in sorted(SKILLS_DIR.iterdir()):
        skill_md = child / "SKILL.md"
        if child.is_dir() and skill_md.exists():
            try:
                ctx.register_skill(child.name, skill_md)
            except Exception as error:  # noqa: BLE001 - skills are additive, never fatal
                logger.warning("hermes-outlook: could not register skill %s (%s)", child.name, error)


def register(ctx: Any) -> None:
    """Wire the plugin: configure the bridge and register all outlook_* tools."""
    client.configure(
        powershell_path=ctx.get_config("powershell_path", default=""),
        timeout_seconds=int(ctx.get_config("timeout_seconds", default=120) or 120),
    )

    handler_map = dict(HANDLERS)
    schema_map = _schema_map()

    missing_handlers = sorted(set(schema_map) - set(handler_map))
    missing_schemas = sorted(set(handler_map) - set(schema_map))
    if missing_handlers or missing_schemas:
        raise RuntimeError(
            "hermes-outlook: schema/handler mismatch — "
            f"schemas without handlers: {missing_handlers}; handlers without schemas: {missing_schemas}"
        )

    for name, schema in schema_map.items():
        ctx.register_tool(name=name, toolset=TOOLSET, schema=schema, handler=handler_map[name])

    _register_skills(ctx)
