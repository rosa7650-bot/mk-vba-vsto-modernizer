# UserForm Controls → React Components

## Control Mapping

| MSForms Control | React Equivalent | Notes |
|---|---|---|
| `ComboBox` | `<select>` or `<Combobox>` (Headless UI) | `List(n, col)` → array of options |
| `ListBox` | `<select multiple>` or custom list | Multi-select if `MultiSelect <> 0` |
| `TextBox` | `<input type="text">` | `Me.txt___.Text` → controlled input value |
| `CommandButton` | `<button>` | `_Click` event → `onClick` handler |
| `CheckBox` | `<input type="checkbox">` | `Me.chk___.Value` → boolean state |
| `OptionButton` | `<input type="radio">` | Group by frame/name |
| `Frame` | `<fieldset>` or `<div>` | Visual grouping only |
| `Label` | `<label>` | Static text |
| `SpinButton` | `<input type="number" step>` | `SpinButton.Value` → numeric state |
| `Image` | `<img>` | Static display |
| `MultiPage` | `<Tabs>` | Each page → one tab |
| `TabStrip` | `<Tabs>` | Similar to MultiPage |

## Event Mapping

| VBA Event | React equivalent |
|---|---|
| `UserForm_Initialize` / `_Activate` | `useEffect(() => { fetchData() }, [])` |
| `CommandButton_Click` | `onClick={handleSubmit}` |
| `ComboBox_Change` | `onChange={handleChange}` → triggers cascade |
| `TextBox_Change` | `onChange` on controlled input |
| `CheckBox_Click` | `onChange` toggling boolean state |
| `UserForm_QueryClose` | `onClose` callback or modal `onDismiss` |
| `ListBox_DblClick` | `onDoubleClick` on list item |

## Cascading ComboBox pattern

VBA:
```vb
Private Sub cmbCustomer_Change()
    Call InitStyle(Me.cmbCustomer.Text)
End Sub
```

React:
```tsx
const [customer, setCustomer] = useState('')
const [styles, setStyles] = useState([])

useEffect(() => {
  if (!customer) return
  fetchStyles(customer).then(setStyles)
}, [customer])
```

## Form submission with validation

VBA:
```vb
Private Sub cmdImport_Click()
    If Me.cmbMaker.ListIndex = -1 Then MsgBox "Choose maker!": Exit Sub
    If Val(Me.txtQty.Text) = 0 Then MsgBox "Qty is zero!": Exit Sub
    ' ... 50 more validations ...
    Call DoImport(...)
End Sub
```

React:
```tsx
const handleSubmit = async () => {
  // 1. client-side pre-check (required fields only)
  if (!maker) { setError('Choose maker'); return }
  
  // 2. server-side validation
  const validation = await api.post('/validate', formData)
  if (!validation.passed) { setErrors(validation.errors); return }
  
  // 3. confirmation if needed
  if (validation.requires_confirmation) {
    setConfirmMessage(validation.message)
    setShowConfirm(true)
    return   // wait for user to confirm
  }
  
  // 4. submit
  const result = await api.post('/import', { ...formData, confirmed: true })
  if (result.success) onSuccess(result.data)
}
```

## MsgBox → React dialog

| VBA | React |
|---|---|
| `MsgBox "Error: " & msg` | `<Toast type="error" message={msg}>` |
| `MsgBox "Warning: ...", vbYesNo` | `<ConfirmDialog message={...} onConfirm={...}>` |
| Multiple MsgBox in loop | `<ValidationErrorList errors={errors}>` |
| `MsgBox "Done!"` | `<Toast type="success" message="Done">` |
