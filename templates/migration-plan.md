# Migration Plan: {{workbook_name}} → Modern Architecture

**Date:** {{YYYY-MM-DD}}
**Source:** `{{path/to/workbook.xlsm}}`
**Target architecture:** {{D-1 answer}}
**Skill:** vba-modernizer

---

## Scope

| Metric | Count |
|---|---|
| Total modules | {{n}} (Standard {{n}}, Class {{n}}, UserForm {{n}}, Worksheet {{n}}, ThisWorkbook 1) |
| Lines of VBA | {{n}} |
| Public procedures | {{n}} |
| UserForms | {{n}} |
| Databases | {{n}} |
| Win32 API declares | {{n}} |
| SQL string concatenation | {{n}} (injection risk) |

---

## System Purpose

{{one paragraph from SYSTEM_PROFILE.md}}

---

## Database Inventory

| ID | Type | Host / DSN | Catalog | Modules | Credentials |
|---|---|---|---|---|---|
| DB-1 | {{Oracle/SQL Server/...}} | {{host}} | {{catalog}} | {{modules}} | {{hardcoded / env var}} |

---

## Decisions (D-1 to D-5)

| # | Decision | Answer |
|---|---|---|
| D-1 | Target architecture | {{answer}} |
| D-2 | Database strategy | {{answer}} |
| D-3 | Output format | {{answer}} |
| D-4 | User settings storage | {{answer}} |
| D-5 | Error/validation strategy | {{answer}} |

---

## Wave Order (Low Risk → High Risk)

| Wave | Modules | Risk | Description |
|---|---|---|---|
| Wave 1 | {{modules}} | Low | Utilities, config, zero-dependency |
| Wave 2 | {{modules}} | Medium | Data access / repository layer |
| Wave 3 | {{modules}} | Medium-High | Specialized queries, domain data |
| Wave 4 | {{modules}} | High | Validation, write operations |
| Wave 5 | {{modules}} | Critical | Core business logic |
| Wave 6 | {{modules}} | Critical | UI layer (UserForms, events) |

---

## Risk Ledger

| Module | Score | Top Pattern | Implication |
|---|---|---|---|
| {{name}} | {{score}} | {{pattern}} | {{what to do}} |

---

## Architecture Cut Points

```
Frontend (Vue 3 + TS)          Backend ({{Python/Node}})
─────────────────────          ──────────────────────────
{{UserForm1}}   ──API──►  {{/api/v1/endpoint}}
{{UserForm2}}   ──API──►  {{/api/v1/endpoint}}

                           {{mdlXxx}} (business logic)
                           {{mdlXxx}} (data access)
```

---

## Out of Scope

| Item | Reason |
|---|---|
| {{Win32 declares}} | Windows-only, no web equivalent |
| {{COM automation}} | Requires separate integration design |
| {{empty modules}} | No logic to migrate |

---

## Approval

**Status:** [ ] Pending  [ ] Approved  [ ] Needs revision

**Notes:** _____

