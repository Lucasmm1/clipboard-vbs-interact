# clipboard-vbs-interact
<small>This is a repository to help developers handle the clipboard in VBS</small>

This repository provides a set of utilities to assist developers in handling the clipboard using VBScript.

## Features

- **PutInClipboard Function**: Easily write data to the clipboard.
- **Clipboard Interaction**: VBScript code to interact with the clipboard efficiently.
- **Temporary File Management**: Creation and deletion of temporary text files for clipboard operations.

## Usage Example

```vbscript
' Uncomment the following lines for a single data example
' data = "StringOfCharacters"
' PutInClipboard data

' Uncomment the following lines for a sequence of lines example
' Assuming "result" is your database result set
' Do While Not result.EOF
'     data = data & result.Fields("ColumnName").Value & vbCrLf
'     result.MoveNext
' Loop
' PutInClipboard data
```

## Contribution
Contributions are welcome! Feel free to open issues or pull requests to enhance the functionality.
