---
name: vba-modernizer
description: Modernize any Excel VBA or VSTO codebase to a modern API-first architecture. Works on .xlsm, .xls, .xlam, .bas, .cls, .frm, or any exported VBA source. Produces a visual 4-phase modernization plan (MODERNIZATION_FLOW.html), a migration plan (MIGRATION_PLAN.md), and a complete API contract (API_PLAN.html) before writing a single line of migrated code. Use when the user asks to modernize, rewrite, migrate, or get off Excel VBA — for any domain (packing lists, finance, HR, inventory, manufacturing, etc.).
---

# VBA / VSTO Modernizer

Modernize Excel VBA or VSTO systems to modern API-first architecture through a structured 4-phase process.
Produces planning artifacts before touching any code — always Plan before Transform.

---

## When to Use This Skill

Trigger when the user mentions:
- VBA, VSTO, Excel macro, .xlsm, .xls, .xlam, .bas, .cls, .frm
- "modernize", "rewrite", "migrate", "get off Excel", "replace Excel tool"
- "build an API for this VBA", "make this a web app"
- Any reference to the 4-phase flow or MODERNIZATION_FLOW.html

---

## Four-Phase Flow

```
Phase 1          Phase 2          Phase 3               Phase 4
DISCOVERY   →    PLAN        →    API TRANSFORM    →    FRONTEND UI
─────────────    ────────────     ─────────────────     ────────────
Extract VBA      Migration        P1 Foundation         UserForms
Analyse risk     plan             P2 Lookup APIs        → React
Identify DBs     API contract     P3 Core Query         Worksheet events
Map modules      Decisions        P4 Edit/Write         → React state
                 D1–D5           P5 Validate/Import     Workbook init
                 Flow diagram     P6 Export/EDI         → App routing
```

**Run phases sequentially. After each phase — and each P-step inside Phase 3 — summarise in 1–2 sentences and stop for user confirmation before continuing.**

---

## Phase 1 — Discovery

> If `.vba-extracted/inventory.json` already exists, skip to step 1-3 and show summary.

### 1-1  Extract VBA source

```bash
# Install if missing
pip install oletools

# Extract from workbook
python path/to/extract_vba.py <file.xlsm> --out .vba-extracted/

# If user already has .bas/.cls/.frm files, accept them directly — skip extraction
```

Use the skill's `scripts/extract_vba.py` if available; otherwise use `olevba` directly:
```bash
olevba --reveal <file.xlsm> > .vba-extracted/raw_dump.txt
```

### 1-2  Analyse modules

```bash
python path/to/analyze_vba.py .vba-extracted/
```

Produces `inventory.json` and `call-graph.json`. If the script is unavailable, analyse manually by reading each `.bas`/`.cls` file and filling the inventory structure:

```json
{
  "modules": [
    {
      "name": "ModuleName",
      "type": "Standard|Class|UserForm|ThisWorkbook|Worksheet",
      "lines": 0,
      "public_api": [],
      "excel_features_used": [],
      "events": [],
      "win32_declares": [],
      "risk_patterns": [],
      "risk_score": 0
    }
  ]
}
```

### 1-3  Build system profile

Answer these questions from reading the code — do not guess:

**A. System purpose**
- What does this workbook do? (one sentence)
- Who are the users? (single user / team / automated batch)
- What is the main output? (report / data entry / calculation / file generation)

**B. Database inventory**
For each connection string found, record:
```
DB-N  Type (Oracle/SQL Server/Access/ODBC)
      Host / DSN
      Catalog / Schema
      Used by modules: [list]
      Purpose: [reads / writes / both]
      Credentials: hardcoded? [yes/no — flag if yes]
```
Identify DB type from SQL dialect:
- Oracle: `ADD_MONTHS`, `NVL`, `DECODE`, `ROWNUM`, `SYSDATE`, `DUAL`
- SQL Server: `CHARINDEX`, `ISNULL`, `GETDATE()`, `TOP n`, `WITH (NOLOCK)`
- Access: `#date#`, `IIF()`, `Nz()`

**C. Risk summary**
From `inventory.json`, surface:
- Modules with risk score ≥ 80 (CRITICAL) — list name + top pattern
- Total SQL string concatenation instances (CR-5) — SQL injection risk
- Total `On Error Resume Next` instances (HI-1) — silent error swallowing
- Total `MsgBox` in loops (HI-6) — cannot be directly translated
- Win32 API declares (CR-1) — Windows-only, may block web migration
- `CreateObject` / late-bound COM (HI-2) — external automation dependencies

**D. UserForm inventory**
For each UserForm, record:
- Name, control list (ComboBox / ListBox / TextBox / CommandButton / etc.)
- Which data does it load on open?
- What does it write on submit?

**E. Output files**
List every file the workbook writes:
- Excel sheets written to disk (SaveAs)
- Text / CSV / EDI output files
- Database writes (INSERT / UPDATE / stored procedures)
- Emails or other external sends

### 1-4  Phase 1 output

Create `.vba-extracted/SYSTEM_PROFILE.md` summarising findings A–E.

```
.vba-extracted/
├── *.bas / *.cls / *.frm    ← VBA source
├── inventory.json           ← module analysis
├── call-graph.json          ← procedure dependencies
└── SYSTEM_PROFILE.md        ← system profile (A–E)
```

---

## Phase 2 — Plan

> If `MIGRATION_PLAN.md` already exists, read it and ask if user wants to update or proceed to Phase 3.

### 2-1  Get decisions D-1 to D-5

**Do not proceed to generate plans until these are answered.**
Ask the user directly — present options, wait for answers.

| # | Decision | Options |
|---|---|---|
| **D-1** | Target architecture | A) React + Python backend &nbsp; B) React + Node backend &nbsp; C) Pure Python / Node (no UI) &nbsp; D) Other |
| **D-2** | Database strategy | A) Keep all existing DBs &nbsp; B) Migrate to single DB (separate project) &nbsp; C) Keep reads, migrate writes |
| **D-3** | Primary output format | A) Keep generating files (.xlsx / .txt / .csv) &nbsp; B) Web UI display + print/PDF &nbsp; C) Both |
| **D-4** | User settings storage (replaces Windows Registry) | A) Browser localStorage &nbsp; B) Backend DB table &nbsp; C) Config file (YAML/JSON) |
| **D-5** | Error/validation strategy (replaces MsgBox in loops) | A) Collect all errors, return batch &nbsp; B) Fail fast on first error &nbsp; C) Log and continue silently |

### 2-2  Generate MIGRATION_PLAN.md

Use `templates/migration-plan.md`. Must include:

- **Scope**: module count, line count, UserForm count, DB count
- **System purpose**: one paragraph from SYSTEM_PROFILE.md
- **Database inventory**: all DBs found with type, host, purpose, credential status
- **Wave order** (low-risk → high-risk):
  - Wave 1: zero-dependency utilities, config modules
  - Wave 2: data access / repository layer (all DB queries)
  - Wave 3: specialized queries / domain-specific data access
  - Wave 4: validation and write operations
  - Wave 5: core business logic (heaviest modules)
  - Wave 6: UI layer (UserForms, Worksheet events, ThisWorkbook)
- **Risk ledger**: each CRITICAL module with score and top pattern
- **Cut points**: which modules go to backend vs frontend
- **Out of scope**: Win32 declares, COM automation, items needing separate decisions

### 2-3  Generate API_PLAN.html

Produce an interactive HTML page listing every API endpoint.
Group endpoints by functional domain (derived from module groupings):

**Standard groups for any VBA system:**

| Group | Derived from | Typical endpoints |
|---|---|---|
| Reference / Lookup | Data-fetching modules with no writes | GET endpoints for drop-down data |
| Core Query | Main data-loading logic | POST /query or GET /{id} with complex assembly |
| Edit / Mutations | Row-level or record-level changes | POST / PUT / DELETE on data items |
| Import / Submit | Write-back to database (main action) | POST /import or POST /submit |
| Validation | Pre-submit checks, business rule validation | POST /validate |
| Export | File generation | GET /export/xlsx, /export/txt, /export/pdf |
| EDI / Integration | External system data exchange | GET + POST per external system |
| Settings | User preferences (replaces Registry) | GET + PUT /settings |

For each endpoint, document:
- Method + path
- Source VBA module.procedure
- Request parameters / body fields (derived from VBA function signatures)
- Response structure
- Which database it hits
- Any CRITICAL risk patterns being resolved

### 2-4  Generate MODERNIZATION_FLOW.html

Produce a 4-column visual flow diagram as an HTML file with:

- **Column 1 — Discovery**: status = ✓ Done (if inventory.json exists)
- **Column 2 — Plan**: status = ✓ Done (once MIGRATION_PLAN.md approved)
- **Column 3 — API Transform**: status = → Next, showing P1–P6 sub-steps with module names
- **Column 4 — Frontend UI**: status = Pending, showing UserForm → React mapping

Each column shows:
- Phase number and name
- Status badge (Done / Next / Pending)
- 3–5 step cards with module names and target descriptions
- Output artifacts box at the bottom
- DB legend at the page bottom (one card per database)

Use dark theme (`#0d1117` background). Arrows between columns. Responsive.

### 2-5  Phase 2 approval gate

**Stop here. Present the user with:**
1. Summary of D-1 to D-5 decisions
2. Wave order overview (which modules in which wave)
3. Total API endpoint count by group
4. List of CRITICAL modules that need special handling

**Do not write any migrated code until the user approves.**

---

## Phase 3 — API Transform (API First)

> **Gate:** D-1 to D-5 must be answered. MIGRATION_PLAN.md must be approved.
> **Rule:** Build and test all backend API endpoints before writing any frontend UI.

All migrated code goes in `migrated/backend/`.
Update `inventory.json` module status: `pending` → `in_progress` → `migrated` | `needs-review` | `blocked`.

### P1 — Foundation & Database Layer

**Deliverable:** All DB connections work, unified response format, error handling middleware.

1. Scaffold project structure per D-1 decision:
   - Python: `FastAPI` + `uvicorn` + `pydantic`
   - Node: `Express` + `TypeScript` + `zod`
2. Create DB connection pool for each database found in Phase 1:
   ```
   Python: cx_Oracle / pyodbc / sqlite3 / psycopg2
   Node:   oracledb / mssql / better-sqlite3 / pg
   ```
3. Move ALL hardcoded credentials to `.env` — never commit
4. Implement unified response wrapper:
   ```json
   { "success": true,  "data": {},    "error": null }
   { "success": false, "data": null,  "error": { "code": "", "message": "", "details": [] } }
   ```
5. Implement global error handler — replaces all `On Error Resume Next`

**Modules mapped:** connection management modules (mdlAdo or equivalent), Registry/config modules

---

### P2 — Lookup / Reference APIs

**Deliverable:** All read-only lookup endpoints return real data.

For each data-fetching function in repository modules:
1. Read original VBA SQL string
2. Rewrite as **parameterized query** — never string concatenation
3. Implement repository function
4. Wire to route handler
5. Write integration test against real DB

**Pattern translations (apply from `references/vba-sql-patterns.md`):**
- `"WHERE x=" & val` → `WHERE x = :x` (Oracle) or `WHERE x = @x` (SQL Server)
- `NVL(a, b)` → `COALESCE(a, b)`
- `DECODE(x, a, b, c)` → `CASE WHEN x=a THEN b ELSE c END`
- `CHARINDEX(s, str)` → parameterized, not concatenated

If a module contains **business rule lookups** (hardcoded SELECT UNION ALL or CASE WHEN tables), convert them to a config table or static data file — do not keep as hardcoded SQL.

---

### P3 — Core Query API

**Deliverable:** The main data-loading endpoint works end-to-end.

This typically maps to the heaviest module(s) — the equivalent of `InputData` or a main `Load` procedure. Decompose into a pipeline of named steps:

```
query_main()
  ├── fetch_step_1()    ← DB read
  ├── fetch_step_2()    ← DB read
  ├── ...
  ├── assemble()        ← pure logic, no DB (replaces Cells(r,c).Value assignments)
  └── return structured JSON
```

**Critical transformation rules:**
- `Cells(r,c).Value =` → append to JSON array (never write to a sheet)
- `Range("n_xxx").Value` → function parameter or config constant
- `MsgBox "..."` inside loops → append to `warnings[]` array (per D-5 decision)
- `On Error Resume Next` → try/except, log, append to `errors[]`, continue
- `GoTo ErrorHandler` → raise typed exception
- Global module-level variables (`Public g...`) → local variables or function parameters
- `Application.Wait` / `DoEvents` → async/await or remove entirely

**Response must include:**
```json
{ "data": { ...domain_data... }, "warnings": [], "errors": [] }
```

---

### P4 — Edit / Write APIs

**Deliverable:** Record-level create/update/delete endpoints work.

For each mutating procedure (InsertRow, DeleteRow, Update, etc.):
- Implement as POST / PUT / DELETE endpoint
- Pure logic helpers (formula recalculation, subtotal update) → internal service functions, not endpoints
- Each mutating endpoint returns the full updated data state (not just success/fail)

---

### P5 — Validation + Main Submit API

**Deliverable:** Complete submit flow works with transactions.

1. **Pre-submit validation endpoint(s)** — maps to any `CheckXxx` / `ValidateXxx` functions:
   - Collect ALL errors per D-5 decision
   - Return `{ "passed": false, "errors": [{ "field": ..., "message": ... }] }`
   - **Never** call database writes with unvalidated data

2. **Main import/submit endpoint** — maps to the primary action button's Click handler:
   - Wrap all DB writes in a single transaction (same DB)
   - Cross-DB writes (e.g., separate warehouse DB) are separate — log failures, do not roll back main transaction
   - VBA's `MsgBox "Continue?"` prompts → return `{ "requires_confirmation": true, "message": "..." }` from API; frontend re-calls with `{ "confirmed": true }`

---

### P6 — Export + Integration APIs

**Deliverable:** File export and external system integrations work.

1. **File export**: `.xlsx` / `.txt` / `.csv` / `.pdf`
   - `.xlsx`: use `openpyxl` (Python) or `exceljs` (Node) — reproduce original layout cell by cell
   - `.txt` fixed-width: replicate `genTXT` or equivalent logic exactly
   - Golden test: capture original VBA output as fixture; assert generated output matches

2. **External system integration**: EDI, SAP, ERP, etc.
   - Each external system gets its own route group
   - Parameterize all queries to external DBs

---

### Phase 3 Completion Gate

Before moving to Phase 4:
- [ ] All 38+ endpoints respond with HTTP 200 on happy path
- [ ] Integration tests pass against real DB
- [ ] OpenAPI / Swagger docs generated (`/docs` endpoint)
- [ ] `REVIEW_QUEUE.md` reviewed — no CRITICAL items unresolved
- [ ] Generate `MIGRATION_REPORT.md` with Phase 3 stats

---

## Phase 4 — Frontend UI

> **Gate:** All Phase 3 API tests must pass. Do not start UI until backend is verified.

Map every VBA UserForm and Worksheet event to a React component:

| VBA Artifact | React Target | Notes |
|---|---|---|
| UserForm + ComboBox/ListBox | `<SelectForm>` with cascading state | Calls Reference APIs on change |
| UserForm + DataGrid / ListBox showing records | `<DataGrid>` component | Calls Core Query API |
| UserForm + Submit button | `<SubmitDialog>` with validation display | Calls Validate then Import API |
| Worksheet_Change event | React `onChange` handler | Calls Edit API |
| Worksheet_SelectionChange | React `onFocus` / controlled input | Local state only if possible |
| Workbook_Open | `useEffect` on app mount | Calls Settings API |
| CommandBars / custom toolbars | `<Navbar>` + React Router | Route per major function |

**UI rules:**
- Zero business logic in React components — display and input only
- `MsgBox` confirmations → `<ConfirmDialog>` consuming `requires_confirmation` from API
- `MsgBox` batch errors → `<ValidationErrorList>` consuming `errors[]`
- Windows Registry `GetSetting/SaveSetting` → `localStorage` (D-4 A) or `GET/PUT /settings` (D-4 B)

---

## Output Layout

```
<project-root>/
├── <original-workbook>.xlsm      ← READ-ONLY, never modify
├── .vba-extracted/               ← Phase 1
│   ├── *.bas / *.cls / *.frm
│   ├── inventory.json
│   ├── call-graph.json
│   └── SYSTEM_PROFILE.md
├── migrated/
│   ├── backend/                  ← Phase 3
│   │   ├── db/                   ← connection pools
│   │   ├── repositories/         ← all SQL, parameterized
│   │   ├── services/             ← business logic from VBA
│   │   ├── routes/               ← API handlers
│   │   ├── tests/                ← integration tests (real DB)
│   │   ├── .env.example
│   │   └── main.py / index.ts
│   ├── frontend/                 ← Phase 4
│   │   └── src/components/
│   └── REVIEW_QUEUE.md
├── MODERNIZATION_FLOW.html       ← Phase 2: visual 4-phase diagram
├── MIGRATION_PLAN.md             ← Phase 2: detailed migration plan
├── MIGRATION_PLAN.html           ← Phase 2: rendered version
├── API_PLAN.html                 ← Phase 2: interactive endpoint reference
└── MIGRATION_REPORT.md           ← Phase 3 complete: stats + rollback
```

---

## Configuration File

If the project has `.vba-modernizer.yml`, honour it:

```yaml
target: react-python | react-node | python | node | generic
source: path/to/file.xlsm          # or directory of .bas/.cls files
output: migrated/
databases:
  - id: DB-1
    type: oracle
    env_var: ORACLE_DSN
  - id: DB-2
    type: sqlserver
    env_var: DW_CONN_STRING
decisions:
  d1: react-python
  d2: keep-existing
  d3: xlsx
  d4: localstorage
  d5: batch-errors
exclusions:
  - Module1          # skip empty/unused modules
risk_threshold: medium   # flag for review at >= this level
test_generation: true
```

---

## Hard Rules

- **Never modify the original workbook.** It is read-only at all times.
- **Never run macros during migration.** Static analysis only (olevba).
- **Never write frontend code before Phase 3 API tests pass.**
- **Never put SQL in route handlers.** All SQL belongs in `repositories/` only.
- **Never hardcode credentials.** All connection strings from `.env`.
- **Never translate `On Error Resume Next` to `except: pass`.** Always surface errors.
- **Never translate `MsgBox` inside loops to blocking UI.** Always use `errors[]` / `warnings[]` arrays.
- **Never inline-translate Excel formulas.** `=SUMPRODUCT(...)` → named function with unit test.
- **Never claim a phase complete with unresolved CRITICAL items in REVIEW_QUEUE.md.**
- **Always stop at D-1 to D-5 and at each phase gate. Never skip approvals.**
