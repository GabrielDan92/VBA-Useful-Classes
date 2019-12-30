'Microsoft Scripting Routine and Microsoft Outlook 16.0 references are needed for this class to work properly
'The objective of this class is to generate a text log file for the errors encountered during the runtime, if any, and send an email to the developer with the error details
'Class created by Gabriel Pintoiu - December 2019
'The class can be called by using: *Dim errorLog As New clsError* and *Call errorLog.errorHandler(errDesciption, errNumber, errSource)*

Option Explicit

Public Sub errorHandler(errDescription As String, errNumber As Integer, errSource As String)

    Dim logFilePath As String
    Dim FSO As Scripting.FileSystemObject
    Dim errorLog As Scripting.TextStream
    Dim errorText As String
    
    
    'the file path's starting point is the same path as the workbook's path
    logFilePath = ThisWorkbook.path & "\errorLog"
    
    
    'create the folder "errorLog" in case it doesn't exist
    If Dir(logFilePath, vbDirectory) = "" Then
        MkDir logFilePath
    End If
    
    
    'generate the new file
    Set FSO = New Scripting.FileSystemObject
    
    
    'declare the file name
    Set errorLog = FSO.CreateTextFile(logFilePath & "\errorLog_" & Month(Date) & "." & Day(Date) & ".txt")
    
    
    errorText = Now & ": Error Description: " & errDescription & " // Error Number: " & errNumber & " // Error Source: " & errSource
    errorLog.WriteLine errorText
    errorLog.Close
    
    
    Set errorLog = Nothing
    Set FSO = Nothing


    Call sendingEmail(errDescription, errNumber, errSource)
    
End Sub

Function sendingEmail(errDescription As String, errNumber As Integer, errSource As String)

    Dim outlookApp As outlook.Application
    Set outlookApp = New outlook.Application
    
    Dim newEmail As outlook.MailItem
    Set newEmail = outlookApp.CreateItem(olMailItem)
    
    
    'creating the email
    newEmail.To = ""
    newEmail.CC = ""
    newEmail.SentOnBehalfOfName = ""
    newEmail.Subject = "Error encountered in " & errSource
    
    
    'email's body
    newEmail.HTMLBody = "Hello Gabriel, <br><p>" _
    & "Please be advised that the script failed at " & Now & "." _
    & " The error description is: " & errDescription & ". The error number is: [" & errNumber & "], with the error source in the project: " & errSource & "<br><p>" _
    & " <br>" _
    & "Best regards, <p>" _
    & "Gabriel <br>" _
    & "Smart Automation Team"
    
    
    'sending the email
    newEmail.send
   
    
    'cleaning
    Set outlookApp = Nothing

End Function

