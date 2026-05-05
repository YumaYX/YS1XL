# GetValueByID_Hash

- Purpose: Retrieves a specific value from a designated column by matching a provided ID value in another specified column.

# Inputs

| Parameter Name | Type | Description |
|-----------------|------|-------------|
| ws | Worksheet | The worksheet object containing the data. |
| idHeader | String | The header name of the column that contains the unique IDs (the search column). |
| idValue | Variant | The specific ID value that needs to be searched for. |
| targetHeader | String | The header name of the column whose value should be retrieved. |
| headerRow | Long (Optional) | The row number where the column headers are located (defaults to 1). |

# Output

- Type: Variant
- Content: The value found in the target column corresponding to the provided ID. Returns an empty string ("") if the ID or required headers are not found.
