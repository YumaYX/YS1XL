# GetValueByID_Hash

- Purpose: Retrieves a specific value from a target column by searching for a matching ID in a designated ID column.

# Inputs

| Argument Name | Type | Description |
|---|---|---|
| ws | Worksheet | The target worksheet containing the data. |
| idHeader | String | The header name of the column containing the IDs. |
| idValue | Variant | The specific ID value that needs to be searched for. |
| targetHeader | String | The header name of the column from which the value should be retrieved. |
| headerRow | Long | (Optional) The row number where the headers are located (defaults to 1). |

# Output

- Type: Variant
- Content: The value found in the target column corresponding to the matching ID; returns an empty string if no match is found.
