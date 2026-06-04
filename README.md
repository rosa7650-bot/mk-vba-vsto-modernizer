# claude-skill-vba-modernizer

A Claude Code skill for modernizing Excel VBA / VSTO systems to modern API-first architecture.

Works on **any** VBA/VSTO project — packing lists, finance, HR, inventory, manufacturing, or any other domain.

---

## What It Does

Guides you through a structured **4-phase modernization process**:

```
Phase 1 Discovery  →  Phase 2 Plan  →  Phase 3 API Transform  →  Phase 4 Frontend UI
   Extract VBA           Migration          P1 Foundation              UserForms
   Analyse risk          plan               P2 Lookup APIs             → React
   Identify DBs          API contract       P3 Core Query
   Map modules           D1–D5 decisions    P4 Edit/Write
                         Flow diagram       P5 Validate/Import
                                            P6 Export/EDI
```

**API First strategy** — all backend APIs are built and tested before any frontend UI is written.

---

## Installation

```bash
# Global (all projects)
mkdir -p ~/.claude/skills
git clone https://github.com/<your-username>/claude-skill-vba-modernizer.git ~/.claude/skills/vba-modernizer

# Project-level (current project only)
mkdir -p .claude/skills
git clone https://github.com/<your-username>/claude-skill-vba-modernizer.git .claude/skills/vba-modernizer
```

Restart Claude Code after installing. Then use `/vba-modernizer` in any conversation.

---

## Requirements

```bash
pip install oletools   # for VBA extraction (Phase 1)
```

---

## Phases

### Phase 1 — Discovery
- Extracts VBA source from `.xlsm` / `.xls` / `.xlam` using `oletools`
- Analyses all modules: line count, Public API, risk score
- Identifies databases (Oracle / SQL Server / Access) from SQL dialect
- Flags hardcoded credentials, SQL injection risks, `On Error Resume Next`
- Produces: `.vba-extracted/`, `inventory.json`, `call-graph.json`, `SYSTEM_PROFILE.md`

### Phase 2 — Plan
- Collects 5 key decisions (D-1 to D-5): target stack, DB strategy, output format, settings storage, error strategy
- Generates `MIGRATION_PLAN.md` with Wave order (low-risk → high-risk)
- Generates `API_PLAN.html` — interactive endpoint reference grouped by functional domain
- Generates `MODERNIZATION_FLOW.html` — visual 4-phase flow diagram
- **Stops for user approval before any code is written**

### Phase 3 — API Transform (API First)
Builds all backend API endpoints in 6 sub-steps:
| Step | Focus | VBA source |
|---|---|---|
| P1 | Foundation & DB connections | Connection modules |
| P2 | Reference / Lookup APIs | Data-fetch modules |
| P3 | Core Query API | Main data-load logic |
| P4 | Edit / Write APIs | Row/record mutation |
| P5 | Validation + Submit | Pre-submit checks + DB writes |
| P6 | Export + Integrations | File output + external systems |

### Phase 4 — Frontend UI
Replaces UserForms and Worksheet events with React components.
**Only starts after all Phase 3 API tests pass.**

---

## Outputs

```
project/
├── .vba-extracted/           # Phase 1: raw VBA source + analysis
├── MODERNIZATION_FLOW.html   # Phase 2: visual 4-phase diagram
├── MIGRATION_PLAN.md         # Phase 2: detailed migration plan
├── API_PLAN.html             # Phase 2: interactive API reference
├── migrated/
│   ├── backend/              # Phase 3: API endpoints
│   │   ├── db/               # DB connection pools
│   │   ├── repositories/     # Parameterized SQL queries
│   │   ├── services/         # Business logic from VBA
│   │   ├── routes/           # API handlers
│   │   └── tests/            # Integration tests
│   └── frontend/             # Phase 4: React components
└── MIGRATION_REPORT.md       # Phase 3 complete: stats + rollback
```

---

## References Included

| File | Contents |
|---|---|
| `references/vba-sql-patterns.md` | SQL string concat → parameterized queries (Oracle + SQL Server) |
| `references/userform-to-react.md` | MSForms controls → React components with code examples |

---

## Templates Included

| File | Used for |
|---|---|
| `templates/migration-plan.md` | Generating `MIGRATION_PLAN.md` |
| `templates/modernization-flow.html` | Generating `MODERNIZATION_FLOW.html` |

---

## Hard Rules (enforced by the skill)

- Never modify the original workbook (read-only)
- Never run macros during migration (static analysis only)
- Never write frontend before Phase 3 API tests pass
- Never put SQL in route handlers (only in `repositories/`)
- Never hardcode credentials (always `.env`)
- Never translate `On Error Resume Next` to `except: pass`
- Never skip the D-1 to D-5 approval gate

---

## Configuration (optional)

Create `.vba-modernizer.yml` in your project root:

```yaml
target: react-python       # react-python | react-node | python | node | generic
source: path/to/file.xlsm
output: migrated/
decisions:
  d1: react-python
  d2: keep-existing
  d3: xlsx
  d4: localstorage
  d5: batch-errors
exclusions:
  - Module1
risk_threshold: medium
test_generation: true
```

---

## License

MIT
