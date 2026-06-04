# VBA SQL Patterns → Parameterized Queries

## Never do this (string concatenation = SQL injection risk)
```vb
sSQL = "SELECT * FROM orders WHERE id=" & nId
sSQL = "WHERE name='" & sName & "'"
```

## Always do this
```python
# Python (cx_Oracle)
cursor.execute("SELECT * FROM orders WHERE id = :id", {"id": n_id})

# Python (pyodbc / SQL Server)
cursor.execute("SELECT * FROM orders WHERE id = ?", (n_id,))

# Node (oracledb)
conn.execute("SELECT * FROM orders WHERE id = :id", { id: nId })

# Node (mssql)
request.input('id', sql.Int, nId)
request.query("SELECT * FROM orders WHERE id = @id")
```

## Oracle-specific translations

| VBA / Oracle SQL | Modern equivalent |
|---|---|
| `NVL(a, b)` | `COALESCE(a, b)` |
| `DECODE(x, a, b, c)` | `CASE WHEN x=a THEN b ELSE c END` |
| `ADD_MONTHS(SYSDATE, -n)` | `SYSDATE - INTERVAL 'n' MONTH` |
| `TO_NUMBER(x)` | `CAST(x AS NUMBER)` or `int(x)` in Python |
| `INITCAP(x)` | `x.title()` in Python or SQL function |
| `UPPER(x)` | `x.upper()` or `UPPER(x)` |
| `ROWNUM <= n` | `FETCH FIRST n ROWS ONLY` (Oracle 12c+) |
| `FROM DUAL` | Remove — just `SELECT 'value' FROM DUAL` → `SELECT 'value'` |

## SQL Server-specific translations

| VBA / SQL Server SQL | Modern equivalent |
|---|---|
| `ISNULL(a, b)` | `COALESCE(a, b)` |
| `CHARINDEX(sub, str)` | parameterize sub: `CHARINDEX(@sub, str)` |
| `GETDATE()` | `GETDATE()` (keep) or pass as parameter |
| `TOP n` | `SELECT TOP (@n)` with parameterized n |
| `WITH (NOLOCK)` | Keep if needed for read performance |
| `CAST(x AS VARCHAR(n))` | Keep |

## ADODB → Python/Node

| VBA ADODB | Python cx_Oracle | Node oracledb |
|---|---|---|
| `New ADODB.Connection` | `cx_Oracle.connect(dsn, user, pwd)` | `oracledb.getConnection({...})` |
| `conn.Execute(sql)` | `cursor.execute(sql, params)` | `conn.execute(sql, binds)` |
| `New ADODB.Command` + SP | `cursor.callproc("PKG.PROC", [...])` | `conn.execute("BEGIN PKG.PROC(...); END;", binds)` |
| `rs.RecordCount` | `len(rows)` or `cursor.rowcount` | `result.rows.length` |
| `rs.MoveFirst` / `rs.MoveNext` | `for row in cursor:` | `for (const row of result.rows)` |
| `rs(0)` / `rs("col")` | `row[0]` / `row["col"]` | `row[0]` / `row.col` |
| `rs.EOF` | `if not rows:` | `if (!result.rows.length)` |
| `oCmd.Parameters.Append` | positional or named bind vars | positional or named bind vars |
| `adParamReturnValue` | first element in out-bind list | `{ dir: oracledb.BIND_OUT }` |

## On Error patterns

| VBA Pattern | Migrated Pattern |
|---|---|
| `On Error Resume Next` | `try: ... except Exception as e: errors.append(str(e))` |
| `On Error GoTo ErrorHandler` | `try: ... except SpecificError as e: raise` |
| `Err.Number <> 0` | `except Exception as e:` |
| `On Error GoTo 0` | Remove (Python/Node errors propagate by default) |
