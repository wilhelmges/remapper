# AGENTS.md

# Project: Remapper (Excel → SQLite → Journal)

## Project purpose

This project is an internal automation utility for processing orders about property losses.

Business pipeline:

```text
Source Excel
 ↓
prepare_excel
 ↓
extract_orders
 ↓
SQLite
 ↓
fill_journal
 ↓
detect_filled_orders
 ↓
validate
```

The project is an operational tool used in production. Stability is more important than architectural purity.

AI agents must prioritize safe incremental improvements over large refactorings.

---

# General principles

## DO

* Keep the project runnable after every change.
* Prefer small commits.
* Refactor incrementally.
* Preserve existing business logic.
* Keep backward compatibility whenever possible.
* Minimize changes to Excel processing behavior.
* Prefer explicit code over abstractions.

## DO NOT

Do not rewrite the entire project.

Do not introduce enterprise architecture.

Do not introduce unnecessary layers.

Avoid overengineering.

Avoid adding frameworks.

Avoid introducing dependency injection frameworks.

Avoid introducing ORMs.

Avoid large "clean architecture" migrations.

---

# Architecture priorities

The order below is mandatory.

## Priority 1 (critical)

### 1. Remove import side effects

Modules must never execute business logic during import.

Forbidden:

```python
conn = sqlite3.connect(...)

wb = load_workbook(...)

print(...)
```

Allowed:

```python
def get_connection():
    ...

def open_workbook():
    ...
```

Nothing should execute on:

```python
import module
```

except constant definitions.

---

### 2. Centralize configuration

Create:

```text
config.py
```

Move all hardcoded values there.

Examples:

```python
SOURCE_FILE

OUTPUT_FILE

DB_FILE

DEFAULT_YEAR

START_ROW

HEADER_ROW
```

Do not hardcode:

```python
2026

5995

8
```

inside business logic.

---

### 3. Eliminate duplicate functions

Only one implementation must exist for:

* date validation
* order parsing
* comment parsing
* hashing

Move reusable code into:

```text
utils/
```

or

```text
domain/
```

---

### 4. Remove temporary files from project root

Examples:

```text
tmp.py
tmp00.py
hello_main.py
```

Move them to:

```text
scripts/experimental/
```

or delete them.

---

### 5. Replace print() with logging

Use:

```python
logger = logging.getLogger(__name__)
```

instead of:

```python
print(...)
```

Use levels:

```python
logger.info()

logger.warning()

logger.error()
```

---

# Priority 2 (high)

## Introduce dataclasses

Prefer structured objects instead of many independent variables.

Example:

```python
@dataclass
class OrderRecord:
    order_id: str
    date: date
    amount: Decimal
    department: str
    sheet_name: str
    changed: bool
```

Avoid:

```python
process(
    order_id,
    date,
    amount,
    department,
    sheet_name,
    changed
)
```

Prefer:

```python
process(order)
```

---

## Centralize SQLite access

Do not open SQLite in multiple files.

Create:

```text
db.py
```

or

```text
repositories/sqlite_repository.py
```

Example:

```python
get_orders()

save_order()

mark_processed()
```

---

## Centralize Excel access

Workbook operations must not be duplicated.

Create:

```text
excel_io.py
```

Examples:

```python
open_source_workbook()

open_journal()

save_workbook()
```

---

# Priority 3 (medium)

Split large modules.

Current examples:

```text
extract_orders.py

fill_journal.py

validate.py
```

Each module should have a single responsibility.

Services may be created:

```text
services/
```

Examples:

```text
extract_service.py

fill_service.py

validate_service.py
```

Avoid creating many tiny files.

---

# Priority 4 (low)

Introduce CLI.

Preferred commands:

```bash
python cli.py extract

python cli.py fill

python cli.py validate
```

or

```bash
python -m remapper extract
```

---

# Naming conventions

Use English names only.

Good:

```python
OrderRecord

JournalEntry

SQLiteRepository

ExcelReader

JournalWriter

WarningColor
```

Bad:

```python
worning_color

tmp00

get_sigmatures
```

Use:

```python
snake_case
```

for functions and variables.

Use:

```python
PascalCase
```

for classes.

Constants:

```python
UPPER_CASE
```

---

# Logging rules

Never use:

```python
print()
```

Use:

```python
logger.info()

logger.warning()

logger.error()
```

Configuration should exist only once.

Prefer:

```python
logging_config.py
```

or

```python
config.py
```

---

# Testing strategy

Test only business-critical logic.

Priority order:

1. parse_date()
2. parse_order()
3. get_order_from_comment()
4. department_to_sheet()
5. find_insert_position()

Avoid testing openpyxl internals.

Use small fixture Excel files.

---

# Folder structure target

```text
project/

config.py

cli.py

db.py

excel_io.py

models.py

utils/

services/

tests/

scripts/
```

Avoid unnecessary deep nesting.

Do not introduce complex enterprise architecture.

Keep the project simple.

---

# Refactoring rules

Every refactoring PR must satisfy:

1. The project still runs.
2. Business behavior is unchanged.
3. Existing Excel files are still supported.
4. No hardcoded years are introduced.
5. No duplicated code is added.
6. No side effects on import are introduced.

Prefer many small safe changes over one large rewrite.

The project goal is maintainability, not architectural perfection.
