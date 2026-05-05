# ExportColumnToFile

- Purpose: Exports the values of a specified column from a worksheet into a plain text file.

# Inputs

| ArgumentName | Type | Description |
|---------------|-----|-------------|
| ws            | Worksheet | The worksheet containing the data to be exported. |
| filePath      | String | The full path and filename where the exported text file will be saved. |
| colNum        | Long | (Optional) The column number to export. Defaults to 1 (the first column). |
| delimiter     | String | (Optional) The string used to separate values when writing to the file. Defaults to vbCrLf (new line). |

# Output

- Type: None (Subroutine)
- Content: Writes the concatenated values of the specified column to the file specified by filePath.
