# pgsql-export

`pgsql-export` reads PostgreSQL schema metadata and writes it to a formatted
Excel workbook (`.xlsx`). It is a local command-line tool: the database is not
modified, and no metadata is sent to an external service.

## Requirements

- Python 3.13 or later
- Access to a PostgreSQL database
- [`uv`](https://docs.astral.sh/uv/) (recommended), or another Python package
  installer

## Installation

Clone the repository and install its locked dependencies:

```console
git clone https://github.com/ishiyamaism/pgsql-export.git
cd pgsql-export
uv sync
```

Alternatively, install the checkout into an existing Python 3.13+ environment:

```console
python -m pip install .
```

## Database connection

Copy `.env.example` to `.env` and replace the placeholder credentials. Then
select a variable with `--env-key`:

```dotenv
INVENTORY_DB=postgresql://USER:PASSWORD@localhost:5432/inventory
```

Connection URIs may use either `postgresql://` or the SQLAlchemy-style
`postgresql+psycopg2://` prefix. A libpq DSN such as
`host=localhost dbname=inventory user=reporter` is also accepted. Keep `.env`
out of version control; it may contain a database password.

## Usage

```console
# Read a connection URI from INVENTORY_DB in .env
uv run pgsql-export --env-key INVENTORY_DB --out inventory-schema.xlsx

# Or pass a connection URI/DSN directly
uv run pgsql-export \
  --dsn "postgresql://USER:PASSWORD@localhost:5432/inventory" \
  --out inventory-schema.xlsx

# Show every option
uv run pgsql-export --help
```

The default output path is `pg_metadata.xlsx`. PostgreSQL's internal schemas
(`information_schema` and names beginning with `pg_`) are omitted.

### Optional row counts

Add `--with-counts` to include a `row_count` column in the **Tables** sheet:

```console
uv run pgsql-export --env-key INVENTORY_DB --with-counts --out inventory-schema.xlsx
```

This runs an exact `COUNT(*)` against every exported table. On a large or busy
database it can take substantial time and I/O, and the connecting role must
have `SELECT` access to every counted table. Row counts are skipped by default.

## Workbook contents

The workbook contains seven sheets:

| Sheet | Contents |
| --- | --- |
| **Tables** | Schema, table name, owner, optional exact row count, and comment |
| **Columns** | Column order, data type, nullability, default, identity mode (`a` = always, `d` = by default), and comment |
| **Constraints** | Primary key, unique, foreign key, and check constraint definitions |
| **Indexes** | Index method, uniqueness, predicate, and full definition |
| **Triggers** | Trigger timing/events, function, and full definition |
| **Functions** | Arguments, return type, volatility, security mode, and comment |
| **Views** | View name, comment, and SQL definition |

Header rows are frozen, filters are enabled, and long definitions are wrapped.
Database text is preserved as text: values such as `=1+1` are not converted to
Excel formulas, and URL-looking values are not converted to hyperlinks.

## Fictional example

For an imaginary `Acme Books` database, parts of the workbook might look like
this:

**Tables**

| schema | table | owner | row_count | comment |
| --- | --- | --- | ---: | --- |
| public | authors | acme_app | 42 | Book authors |
| sales | orders | acme_app | 18,504 | Customer orders |

**Columns**

| schema | table | ordinal_position | column | data_type | nullable | default | identity |
| --- | --- | ---: | --- | --- | --- | --- | --- |
| sales | orders | 1 | id | bigint | false |  | d |
| sales | orders | 2 | placed_at | timestamp with time zone | false | now() |  |

**Triggers**

| schema | table | trigger_name | timing_events | function |
| --- | --- | --- | --- | --- |
| sales | orders | orders_audit | AFTER UPDATE | write_order_audit |

The example is illustrative only and contains no real database information.

## Development

Run the regression tests and type checker with:

```console
uv run python -m unittest discover -v
uv run mypy pgsql_export
```

Additional usage notes are available in [`docs/usage.md`](docs/usage.md).
