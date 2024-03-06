Function PutInClipboard(data)
    ' Create a temporary text file
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    tempFileName = objFSO.GetSpecialFolder(2) & "\tempClipboard.txt"

    ' Write the data to the file
    Set tempFile = objFSO.CreateTextFile(tempFileName, True)
    tempFile.Write data
    tempFile.Close

    ' Read the text file into the clipboard
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c type " & tempFileName & " | clip", 2, True

    ' Delete the temporary file
    objFSO.DeleteFile tempFileName

    WScript.Echo "Data placed in the clipboard"
End Function

Function main
    ' data = "StringOfCharacters"
    ' PutInClipboard data

    data = ""

    For i = 1 To 3
        data = data & "This is the data: " & i & vbCrLf
    Next

    PutInClipboard data
End Function

main()