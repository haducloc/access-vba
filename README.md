## Access VBA Framework

Build scalable Access front-end applications that connect to remote SQL databases (SQL Server, MySQL, PostgreSQL, Oracle, etc.) using ADO.

- Access is the **UI layer only** (forms/reports).
- Business logic and data access are handled in VBA modules/classes.
- Database storage is a remote SQL database.
- Framework files start with `X` (example: `XInputState`, `XStateCollection`, `XAdoUtil`).
- App-specific files do **not** start with `X` (example: `Ticket_MainForm`, `Ticket_EditForm`).

---

## 1) Repository structure and responsibilities

### Framework (`X*.vba`)

Reusable building blocks:

- Input parsing and validation: `XInputState`, `XInputUtil`, `XStateCollection`
- ADO connection/query helpers: `XAdoUtil`, `XAdoCrud`, `XAdoDropdown`
- Form helpers: `XFormUtil`, `XDatasheetDelegate`
- Utility/support modules: parsing, error, assertions, strings, DB type mapping, etc.

### App-specific (`Ticket_*.vba`, `App_*.vba`)

Implementation for one table/use case:

- `Ticket_MainForm`: search page + subform container
- `Ticket_Datasheet`: result grid (embedded subform)
- `Ticket_EditForm`: single-record add/update/delete
- `Ticket_Ado`: SQL query for ticket search
- `App_DbConnection`: database connection string and connection factory

---

## 2) Standard form pattern for each database table

For each table, create 3 forms with this naming convention:

- `Table_MainForm`
- `Table_Datasheet`
- `Table_EditForm`

Example for table `Ticket`:

- `Ticket_MainForm`
- `Ticket_Datasheet`
- `Ticket_EditForm`

### 2.1 `Table_MainForm`

Single Form used for searching.

Typical UI:

- Top section: search controls (`txt...`, `cbo...`, `chk...`)
- Bottom section: embedded `Table_Datasheet` subform
- Optional: `btnAddNew`

Key behavior:

- `Form_Load` calls `ConfigCustomForm Me`
- Calls search routine to load initial data

Search routine:

1. Reads search controls into `XInputState`
2. Adds them to `XStateCollection`
3. If invalid input exists -> bind empty recordset
4. If valid -> run ADO search, apply sort, bind result to datasheet subform

---

### 2.2 `Table_Datasheet`

Datasheet Form used only as an embedded subform.

Design rules:

- Set form `DefaultView = Datasheet`
- Add controls for each query field (textbox/checkbox as needed)
- No labels required
- For each displayed field, set:
  - `Control Source` (field name from SQL)
  - `Datasheet Caption`

Runtime rules:

- Let `XDatasheetDelegate` handle sort/double-click/error plumbing
- Implement only:
  - `RefreshFromParent(sortByAdo As String)`
  - `OpenEditForm(selectedRow As Object)`

---

### 2.3 `Table_EditForm`

Single Form used for creating/updating one row.

Design rules:

- Set form `DefaultView = Single Form`
- Add controls freely (Label/TextBox/ComboBox/CheckBox/Button)
- Name controls with convention:
  - `txtName`, `txtDescription`, `chkIsDone`, `cboTypeID`, `btnSave`, `btnDelete`

Runtime rules:

In `Form_Load`:

- `ConfigCustomForm Me`
- Initialize dropdowns
- Parse PK from `OpenArgs`
- Switch UI between add mode and edit mode
- Load row if editing

In `btnSave_Click`:

- Convert controls to `XInputState`
- Validate via `XStateCollection`
- Map to dictionary and call `InsertRowAdo` or `UpdateRowAdo`

In `btnDelete_Click`:

- Confirm then call `DeleteRowAdo`

In `Form_Close`:

- Close connection and refresh parent main form

---

## 3) Core framework concepts

### 3.1 `XInputState`: one control -> one typed state

`XInputState` stores:

- `FieldName` (DB field name)
- `ValueType` (int, string, bool, date, etc.)
- `Value` (converted value or Null)
- `IsValid` (conversion/required validation result)
- `ErrorMessage` (human-readable message)

You create states using helpers in `XInputUtil`:

- `GetString`, `GetCode`
- `GetInt2`, `GetInt4`, `GetInt8`, `GetByte`, `GetUByte`
- `GetBool`
- `GetDate`, `GetTime`, `GetDateTime`
- `GetFloat`, `GetDouble`, `GetDecimal`

Example:

```vb
Dim stName As XInputState
Set stName = GetString(Me.txtName, "Name", True)

If Not stName.IsValid Then
    MsgBox stName.ErrorMessage
End If
```

---

### 3.2 `XStateCollection`: validate many controls together

Use `XStateCollection` when a form has multiple inputs.

#### Common flow

```vb
Dim states As XStateCollection: Set states = New XStateCollection
states.AddStates stName, stDescription, stIsDone

If Not states.AllValid Then
    MsgBox states.ToErrorString
    Exit Sub
End If

Dim values As Object
Set values = states.ToValuesDict
```

#### Benefits

- Aggregated validation (`AllValid`)
- Centralized error display (`ToErrorString`)
- Easy DB payload creation (`ToValuesDict`)

---

### 3.3 `XAdoUtil`: parameterized ADO command helpers

`XAdoUtil` gives safe command helpers.

Core helpers:

- `CreateCommandAdo(cn, sql)`
- `ExecuteQueryAdo(cmd, disconnect)`
- `ExecuteUpdateAdo(cmd)`

Typed parameter helpers:

- `ParamInt4Ado`
- `ParamBoolAdo`
- `ParamDateAdo`
- `ParamVarcharAdo`
- etc.

LIKE helpers:

- `ParamLikeAdo`
- `ParamNLikeAdo`

This keeps SQL parameterized and DB-provider-aware.

---

### 3.4 `XAdoCrud`: generic insert/update/delete/select by PK

Main APIs:

- `InsertRowAdo`
- `UpdateRowAdo`
- `DeleteRowAdo`
- `GetRowByPkAdo`
- `ExistsByPkAdo`

Pattern:

- Pass `fieldsCsv` + `typesCsv`
- Pass values dictionary (usually from `XStateCollection.ToValuesDict`)
- Framework builds SQL and appends typed parameters in correct order

---

### 3.5 `XFormUtil` + `XDatasheetDelegate`

`XFormUtil`:

- `ConfigCustomForm` enforces Single Form settings
- `ConfigCustomDatasheet` enforces Datasheet settings
- `TryOpenForm` prevents opening duplicate edit forms
- `TryCallForm` calls public method on loaded form

`XDatasheetDelegate` handles:

- Datasheet sorting handoff to ADO recordset
- Double-click row -> open edit form
- Datasheet filter/sort edge cases

---

## 4) Ticket sample walkthrough

### 4.1 Connection setup (`App_DbConnection`)

- Configure provider + server + database in `GetDbConString`
- Use `GetConnection(ByRef cn)` to lazily create/open a cached ADODB connection
- Reuse this pattern in all forms/modules

---

### 4.2 Main search form (`Ticket_MainForm`)

#### SearchTickets flow

Read search controls:

```vb
GetInt4(Me.txtTicketID)
GetString(Me.txtName)
```

Validate with `XStateCollection`.

If invalid -> use `CreateEmptyRsAdo`.

If valid -> call `SearchTicketAdo(GetConn(), ...)`.

Apply sort using `BuildRsSortByAdo`.

Bind result to:

```vb
Me.Ticket_Datasheet.Form.Recordset
```

---

### 4.3 Search SQL module (`Ticket_Ado`)

- Uses one parameterized query
- Filters are optional using `(? IS NULL OR ...)` pattern
- Uses `ParamInt4Ado` and `ParamLikeAdo`
- Returns disconnected recordset via `ExecuteQueryAdo(cmd, True)`

---

### 4.4 Datasheet form (`Ticket_Datasheet`)

- Contains thin glue code only
- Delegate (`XDatasheetDelegate`) handles events

`OpenEditForm`:

- Reads selected `TicketID`
- Opens `Ticket_EditForm` with `OpenArgs`

---

### 4.5 Edit form (`Ticket_EditForm`)

- Determines add/edit mode from `OpenArgs`
- Loads dropdown options from SQL via  
  `ExecuteDropdownOptionsSqlAdo + XDropdownOptions.ToValueList`

On save:

- Build states (`GetString`, `GetBool`, `GetInt4`, `GetDate`)
- Validate
- Insert or update via `XAdoCrud`

On delete:

- Build PK dictionary
- Call `DeleteRowAdo`

On close:

- Calls `Ticket_MainForm.RefreshTickets`

---

## 5) How to create a new CRUD module for another table

Use this checklist.

### Create forms

- `YourTable_MainForm`
- `YourTable_Datasheet`
- `YourTable_EditForm`

### Design datasheet form

- Controls mapped to query columns via Control Source
- Set Datasheet Caption
- Keep it subform-only

### Design edit form

- Add input controls and action buttons
- Use clear names (`txt...`, `cbo...`, `chk...`, `btn...`)

### Create search SQL module

- `YourTable_Ado`
- Write parameterized query
- Return disconnected recordset

### Implement MainForm logic

- `ConfigCustomForm`
- Collect search states
- Validate and bind resultset to datasheet subform

### Implement Datasheet glue

- Add delegate hooks
- Implement `RefreshFromParent` + `OpenEditForm`

### Implement EditForm logic

- Parse `OpenArgs` PK
- Load row by PK for edit mode
- Save using `InsertRowAdo` / `UpdateRowAdo`
- Delete using `DeleteRowAdo`
- Refresh parent after close

In edit form close event:

```vb
TryCallForm "YourTable_MainForm", "Refresh..."
```

---

## 6) Practical beginner tips

- Always use `Option Explicit`
- Prefer framework parsing methods (`GetInt4`, `GetDate`, etc.) over manual conversion
- Always validate all input states before DB write operations
- Keep SQL parameterized (never string-concatenate user input into SQL)
- Keep datasheet forms thin; put business logic in modules/forms above them
- Reuse `X*` framework modules instead of copy-pasting custom helper code
- Close objects (`Recordset`, `Connection`) in form close events
