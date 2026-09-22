"""Public package interface for pgsql-export."""

from .db import connect_via_env_or_dsn

__all__ = [
    "connect_via_env_or_dsn",
]
