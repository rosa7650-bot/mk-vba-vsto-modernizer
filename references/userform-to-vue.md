# UserForm Controls → Vue 3 Components

## Control Mapping

| MSForms Control | Vue 3 Equivalent | Notes |
|---|---|---|
| `ComboBox` | `<select v-model>` or `<Combobox>` (Headless UI) | `List(n, col)` → array of options |
| `ListBox` | `<select multiple v-model>` or custom list | Multi-select if `MultiSelect <> 0` |
| `TextBox` | `<input type="text" v-model>` | `Me.txt___.Text` → reactive ref value |
| `CommandButton` | `<button @click>` | `_Click` event → `@click` handler |
| `CheckBox` | `<input type="checkbox" v-model>` | `Me.chk___.Value` → boolean ref |
| `OptionButton` | `<input type="radio" v-model>` | Group by same v-model name |
| `Frame` | `<fieldset>` or `<div>` | Visual grouping only |
| `Label` | `<label>` | Static text |
| `SpinButton` | `<input type="number" :step v-model>` | `SpinButton.Value` → numeric ref |
| `Image` | `<img>` | Static display |
| `MultiPage` | `<Tabs>` (e.g. Headless UI / Vuetify) | Each page → one tab |
| `TabStrip` | `<Tabs>` | Similar to MultiPage |

## Event Mapping

| VBA Event | Vue 3 Equivalent |
|---|---|
| `UserForm_Initialize` / `_Activate` | `onMounted(() => { fetchData() })` |
| `CommandButton_Click` | `<button @click="handleSubmit">` |
| `ComboBox_Change` | `<select @change="handleChange">` or `watch(customer, ...)` |
| `TextBox_Change` | `<input v-model="field">` + `watch(field, ...)` |
| `CheckBox_Click` | `<input type="checkbox" v-model="flag">` |
| `UserForm_QueryClose` | `@close` emit or `onBeforeUnmount` |
| `ListBox_DblClick` | `@dblclick` on list item |

## Cascading ComboBox Pattern

VBA:
```vb
Private Sub cmbCustomer_Change()
    Call InitStyle(Me.cmbCustomer.Text)
End Sub
```

Vue 3 (Composition API):
```vue
<script setup lang="ts">
import { ref, watch } from 'vue'
import { fetchStyles } from '@/api/ref'

const customer = ref('')
const styles = ref([])

watch(customer, async (val) => {
  if (!val) return
  styles.value = await fetchStyles(val)
})
</script>

<template>
  <select v-model="customer">
    <option v-for="c in customers" :key="c.id" :value="c.code">{{ c.name }}</option>
  </select>
  <select v-model="style">
    <option v-for="s in styles" :key="s.id" :value="s.id">{{ s.name }}</option>
  </select>
</template>
```

## Form Submission with Validation

VBA:
```vb
Private Sub cmdImport_Click()
    If Me.cmbMaker.ListIndex = -1 Then MsgBox "Choose maker!": Exit Sub
    If Val(Me.txtQty.Text) = 0 Then MsgBox "Qty is zero!": Exit Sub
    ' ... 50 more validations ...
    Call DoImport(...)
End Sub
```

Vue 3:
```vue
<script setup lang="ts">
const errors = ref([])
const showConfirm = ref(false)
const confirmMessage = ref('')

const handleSubmit = async () => {
  // 1. client-side pre-check (required fields only)
  if (!maker.value) { errors.value = ['請選擇工廠']; return }

  // 2. server-side validation
  const validation = await api.post('/validate', formData.value)
  if (!validation.passed) { errors.value = validation.errors; return }

  // 3. confirmation if needed (replaces MsgBox vbYesNo)
  if (validation.requires_confirmation) {
    confirmMessage.value = validation.message
    showConfirm.value = true
    return  // wait for user to click confirm
  }

  await doSubmit()
}

const doSubmit = async () => {
  const result = await api.post('/import', { ...formData.value, confirmed: true })
  if (result.success) emit('success', result.data)
}
</script>

<template>
  <ValidationErrorList :errors="errors" />
  <button @click="handleSubmit">Submit</button>
  <ConfirmDialog
    v-if="showConfirm"
    :message="confirmMessage"
    @confirm="doSubmit"
    @cancel="showConfirm = false"
  />
</template>
```

## MsgBox → Vue 3 Dialog / Toast

| VBA | Vue 3 |
|---|---|
| `MsgBox "Error: " & msg` | `<Toast type="error" :message="msg" />` |
| `MsgBox "Warning: ...", vbYesNo` | `<ConfirmDialog :message="msg" @confirm="onConfirm" />` |
| Multiple MsgBox in loop | `<ValidationErrorList :errors="errors" />` |
| `MsgBox "Done!"` | `<Toast type="success" message="完成" />` |

## Vue Router (replaces ThisWorkbook CommandBars)

```vue
<!-- AppNavbar.vue -->
<script setup lang="ts">
import { RouterLink } from 'vue-router'
</script>

<template>
  <nav>
    <RouterLink to="/query">裝箱單查詢</RouterLink>
    <RouterLink to="/import">匯入</RouterLink>
    <RouterLink to="/edi">EDI</RouterLink>
  </nav>
  <RouterView />
</template>
```

```ts
// router/index.ts
import { createRouter, createWebHistory } from 'vue-router'
const router = createRouter({
  history: createWebHistory(),
  routes: [
    { path: '/query',  component: () => import('@/views/QueryView.vue') },
    { path: '/import', component: () => import('@/views/ImportView.vue') },
    { path: '/edi',    component: () => import('@/views/EdiView.vue') },
  ]
})
```

## Worksheet_Change → Vue 3 Reactive

| VBA | Vue 3 |
|---|---|
| `Worksheet_Change(Target)` | `watch(field, handler)` |
| `Worksheet_SelectionChange` | `@focus` on input element |
| `Workbook_Open` | `onMounted(() => loadSettings())` |
