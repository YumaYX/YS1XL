# LastUsedRow

- Purpose: Retrieves the row number of the last non-empty cell in a specified column of a worksheet.

# Inputs

| Argument Name | Type | Description |
|---------------|------|-------------|
| ws            | Worksheet | The worksheet object to check for the last row. |
| col           | Long | The column index to check (e.g., 1 for Column A, 2 for Column B). Defaults to 1. |

# Output

- Type: Long
- Content: The row number of the last used data. Returns 0 if the specified column is entirely empty.
