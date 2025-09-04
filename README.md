# pgsql-export

Export PostgreSQL schema to Excel (.xlsx), local-only.

Quick start:

```
uv sync
uv run pgsql-export --help
uv run pgsql-export --env-key INVENTORY_DB --out meta.xlsx
```

Notes:
- Copy `.env.example` to `.env` and fill credentials.
- Reads connection from `.env` (e.g., `INVENTORY_DB`) or `--dsn`.
- Outputs multiple sheets: Tables, Columns, Constraints, Indexes, Triggers, Functions, Views.

More details: `docs/usage.md`
