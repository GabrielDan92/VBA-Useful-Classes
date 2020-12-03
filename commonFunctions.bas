'=============================================
'SQL open and close functions
Public Function OpenConnection(ByVal workbookName As String)

    Dim conn_str As String
    Set Connection = CreateObject("ADODB.Connection")
    Set Recordset = CreateObject("ADODB.Recordset")
    conn_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & workbookName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    Connection.Open conn_str
    
End Function

Public Function CloseConnection(ByVal workbookName As String)

    Connection.Close
    Set Connection = Nothing
    Set Recordset = Nothing
    
End Function
'=============================================

                
'=============================================
'refresh an existing PowerQuery connection and wait until all the data has been retrieved before step into the next line of code
Function refreshPowerQuery(ByVal connection As String)
    Dim boolRefresh As Boolean
    With ThisWorkbook.Connections(connection).OLEDBConnection
        boolRefresh = .BackgroundQuery
        .BackgroundQuery = False
        .Refresh
        .BackgroundQuery = boolRefresh
        Sleep 2000
    End With
End Function
'=============================================

        
'=============================================
'find a string within an existing array. Returns a boolean value
Function IsInArray(ByVal stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
'=============================================


'=============================================
'update pivot's source data and refresh it
Dim pTable As PivotTable
Dim pCache as PivotCache

For Each pTable In Sheet1.PivotTables
    Set pCache = ThisWorkbook.PivotCaches.Create(xlDatabase, Sheet2.Cells(1, 1).CurrentRegion.Address)
    'Set pTable = pCache.CreatePivotTable (TableDestination:=.Sheets("Pivot").Cells(2, 2), TableName:="Pivot")
    pTable.ChangePivotCache pCache
    pTable.RefreshTable
Next
'=============================================


'=============================================
'general settings declared at the beginning of the sub routine, in order to speed up the overall runtime
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayStatusBar = False
Application.EnableEvents = False

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
'=============================================


'=============================================
'run a macro from another workbook (workBookName must be declared as an 'Workbook' object and initialized)
'Application.Run ("'" & workBookName & "'!subRoutineName")
'=============================================


'=============================================
'change the security mode to enable all macros when opening a workbook that already has existing macro files
    Dim security As MsoAutomationSecurity
    security = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    
'set the automation security back to its original setting
    Application.AutomationSecurity = security
'=============================================


'=============================================
'VBS script that calls a sub routine from an opened Excel file. Useful to run in parallel with another sub routine.
WScript.Sleep(1000)
Set objExcel = GetObject("C:\Automations\fileName.xlsm")
objExcel.Application.Run "'C:\Automations\fileName.xlsm'!ModuleName.subRoutineName"
Set objExcel = Nothing

'calling the VB Script:
Shell "wscript path\script.vbs", vbNormalFocus
'=============================================


'=============================================
'click on the HTML element, if it exists; (works with className,Id, tagName, or the regular querySelector method)
    If Not ie.document.getElementsByClassName("className")(0) Is Nothing Then
        ie.document.getElementsByClassName("className")(0).Click
        Call ieBusy(ie)
    End If
'=============================================


'=============================================
'dinamic delay for web pages; waits until the HTML element is loaded (works with className,Id, tagName, or the regular querySelector method)
On Error Resume Next
    Do While ie.document.getElementById("idName") Is Nothing
        DoEvents
        Sleep 1000
        secondsCounter = secondsCounter + 1
        If secondsCounter = 20 Then
        Exit Do
        End If
    Loop
    secondsCounter = 0
On Error GoTo 0
'=============================================


'=============================================
'standard delay for web pages
Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
End Sub
'=============================================


'=============================================
'returns the digits from a string containing digits & letters
Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function
'=============================================
