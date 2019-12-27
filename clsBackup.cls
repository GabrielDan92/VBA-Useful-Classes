Option Explicit
'The objective of this class is to create a copy of the current excel file, in case something wrong happens with the original workbook while the script runs.
'Class created by Gabriel Pintoiu - December 2019
'the class can be called by using: *Dim backUp As New clsBackup* and *backUp.backupFunction*

Sub backupFunction()

    Dim path As String
    path = ThisWorkbook.path & "\backupFolder"

    'creates a folder called "backupFolder" in case it doesn't exist
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

    'saves a copy of the existing file
    ThisWorkbook.SaveCopyAs path & "\backupCopy.xlsm"

End Sub
