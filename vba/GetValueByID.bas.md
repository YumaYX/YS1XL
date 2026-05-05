# GetValueByID

- Purpose: Retrieves a specific value from a worksheet by matching a provided ID value to the designated ID column.

# Parameters

| Parameter Name | Type | Description |
|-----------------|-----|-------------|
| ws | Worksheet | The worksheet object where the search will be performed. |
| idHeader | String | The header name (column title) of the column containing the IDs. |
| idValue | Variant | The specific ID value that is being searched for. |
| targetHeader | String | The header name (column title) of the column from which the desired value should be retrieved. |
| headerRow | Long (Optional) | The row number where the headers are located (defaults to row 1). |

# Return Value

- Type: Variant
- Content: The cell value from the target column corresponding to the found ID. Returns an empty string ("") if no match is found or if the required header columns cannot be located.
