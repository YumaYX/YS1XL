# SearchDataLocation

- Purpose: Reads a CSV file, identifies a specified key column, and returns a dictionary mapping unique key values to their first row index in the dataset.

# Inputs

| 引数名 | 型 | 説明 |
|--------|----|------|
| csvFilePath | String | The path to the CSV file to be processed. Defaults to "sample.csv". |
| targetKey | String | The header name of the key column to use for indexing. Defaults to "id". |

# Output

- Type: Object (Scripting.Dictionary)
- Content: A dictionary where keys are the unique values found in the `targetKey` column, and the corresponding values are the 1-based row index of the first occurrence of that key in the data.
