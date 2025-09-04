from __future__ import annotations

import os
import re
from typing import Optional

import psycopg2
from dotenv import load_dotenv


_URL_PREFIX_RE = re.compile(r"^(postgres(?:ql)?)(\+[^:]+)(://)")


def _normalize_url(url: str) -> str:
    """Normalize SQLAlchemy-style URLs to psycopg2-friendly URIs.

    Examples:
    - postgresql+psycopg2://... -> postgresql://...
    - postgres+psycopg2://...   -> postgres://...
    """
    return _URL_PREFIX_RE.sub(lambda m: f"{m.group(1)}{m.group(3)}", url)


def get_dsn_from_env(env_key: str) -> Optional[str]:
    load_dotenv(override=False)
    val = os.getenv(env_key)
    if not val:
        return None
    if val.startswith("postgres"):
        return _normalize_url(val)
    # Raw DSN like "host=... dbname=..."
    return val


def connect_via_env_or_dsn(env_key: Optional[str], dsn: Optional[str]):
    """Return a psycopg2 connection from .env key or explicit dsn/uri.

    Precedence: explicit dsn > env_key > PG_DSN env > standard PG* env.
    """
    load_dotenv(override=False)

    use = dsn or (get_dsn_from_env(env_key) if env_key else None) or os.getenv("PG_DSN")
    if not use:
        # Fallback to libpq env vars if present (host/user/db etc.)
        # psycopg2.connect() with no args uses libpq defaults; allow that only if
        # at least PGHOST/PGDATABASE is set to avoid surprising localhost attempts.
        if os.getenv("PGHOST") or os.getenv("PGDATABASE"):
            return psycopg2.connect("")
        raise RuntimeError("No DSN/URI provided. Use --env-key, --dsn, or set PG_DSN.")

    if use.startswith("postgres"):
        use = _normalize_url(use)

    return psycopg2.connect(use)

