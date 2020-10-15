Attribute VB_Name = "create db"
Option Compare Database
Option Explicit
Public Function Counts()
Access.DoCmd.RunSQL "UPDATE [Count] SET Field = 'Missing' WHERE Field is null"
Access.DoCmd.RunSQL "SELECT LEFT(Log, LEN(Log)-3) AS Type, Field, Count INTO CountByType FROM [Count]"
Access.DoCmd.RunSQL "SELECT Type, Field, Sum(Count) AS [Sum] INTO SumByType FROM CountByType GROUP BY Type, Field"
Access.DoCmd.RunSQL "SELECT Type, Count(Field) AS Distinct_Errors INTO SumByError FROM SumByType GROUP BY Type"
CurrentDb.Execute "populate_discrepancies", dbFailOnError
End Function
Public Function ClearDBs()
Access.DoCmd.RunSQL "DELETE * FROM Count"
Access.DoCmd.RunSQL "DELETE * FROM Processed"
Access.DoCmd.RunSQL "DELETE * FROM Discrepancies"
End Function

'make sure you have no dots in the log file name
Public Function ImportCSV()

 Dim strPathFile As String
 Dim strFile As String
 Dim strPath As String
 Dim strTable As String
 Dim blnHasFieldNames As Boolean
 Dim cSource As String
 Dim cDest As String
 Dim fs As Object
 Dim fDest As String
 Dim compact As String
 Dim File As String
 
 Set fs = CreateObject("Scripting.FileSystemObject")
 strPath = GetFolder()
 strFile = Dir(strPath & "\*.csv")
 cSource = Application.CurrentProject.FullName
 ' iterate through files with .csv extension in the selected folder
 ' select your csv source folder first, then the access target folder

 Do While Len(strFile) > 0
       strPathFile = strPath & "\" & strFile
       File = Left(strFile, InStr(strFile, ".") - 1)
       ' import csv using a pre-existing setup to the phlex_auditlogs table.
       ' To ensure that macro works, do not edit the setup or the table!
       DoCmd.TransferText acImportDelim, "meta", File, strPathFile, True
       'run primary and secondary queries for pdf generator into the new db
       Access.DoCmd.RunSQL "INSERT INTO Count Select Distinct [{targetAttr}] as Field, '" & File & "'as Log, COUNT([{sourceKey}]) as Count FROM [" & File & "] GROUP BY [{targetAttr}]"
       Access.DoCmd.DeleteObject acTable, File
 ' Uncomment out the next code step if you want to delete the
 ' EXCEL file after it's been imported
 '       Kill strPathFile

       strFile = Dir()

 Loop
       Access.DoCmd.RunSQL "INSERT INTO Discrepancies SELECT c.[Log] AS Logs, Count(*) AS Discrepancies FROM [Count] AS c GROUP BY c.[Log] ORDER BY [Log]"
MsgBox ("Done")
End Function
'make sure you have no dots in the log file name
Public Function ImportLogs()

 Dim strPathFile As String
 Dim strFile As String
 Dim strPath As String
 Dim strTable As String
 Dim blnHasFieldNames As Boolean
 Dim cSource As String
 Dim cDest As String
 Dim fs As Object
 Dim fDest As String
 Dim compact As String
 Dim File As String
 Dim fsoLogIn As TextStream
 Dim strSearch As String
 Dim strNumber As String
 Dim Line As String
 Dim strDiscrepancies As String
 Dim strNotFound As String
 Dim strDiscr As String
 Dim strNF As String
 Dim strRunDate As String
 Dim strCompleted As String
 Dim strRD As String
 Dim strCom As String
 
 Set fs = CreateObject("Scripting.FileSystemObject")
 strSearch = "records processed"
 strDiscrepancies = "discrepancies found"
 strNotFound = "target records not found"
 strRunDate = "Run date:"
 strCompleted = "Process completed at"
 strPath = GetFolder()
 strFile = Dir(strPath & "\*.log")
 cSource = Application.CurrentProject.FullName
 ' iterate through files with .csv extension in the selected folder
 ' select your csv source folder first, then the access target folder

 Do While Len(strFile) > 0
       strPathFile = strPath & "\" & strFile
       Set fsoLogIn = fs.OpenTextFile(strPathFile, ForReading)
        Do Until fsoLogIn.AtEndOfStream
            Line = fsoLogIn.ReadLine
            If InStr(1, Line, strSearch) > 0 Then
                strNumber = Left(Line, InStr(Line, " "))
            ElseIf InStr(1, Line, strDiscrepancies) > 0 Then
                strDiscr = Left(Line, InStr(Line, " "))
            ElseIf InStr(1, Line, strNotFound) > 0 Then
                strNF = Left(Line, InStr(Line, " "))
            ElseIf InStr(1, Line, strRunDate) > 0 Then
                strRD = Right(Line, Len(Line) - InStr(8, Line, " "))
            ElseIf InStr(1, Line, strCompleted) > 0 Then
                strCom = Right(Line, Len(Line) - InStr(20, Line, " "))

            End If
        Loop
       ' import csv using a pre-existing setup to the auditlogs table.
       ' To ensure that macro works, do not edit the setup or the table!
    File = Left(strFile, InStr(strFile, ".") - 1)
    Access.DoCmd.RunSQL "INSERT INTO Processed(Log, Processed, Discrepancies, Not_Found, Start, End) VALUES ('" & File & "','" & strNumber & "','" & strDiscr & "','" & strNF & "','" & strRD & "','" & strCom & "')"
 ' Uncomment out the next code step if you want to delete the
 ' EXCEL file after it's been imported
 '       Kill strPathFile
       strFile = Dir()
 Loop
       Access.DoCmd.RunSQL "UPDATE Processed SET Run_time = format(DATEDIFF('s', Start, End)/86400, 'hh:nn:ss')"
       
MsgBox ("Done")
End Function
Public Function GetFolder() As String
Dim fldr As FileDialog

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Show
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    GetFolder = .SelectedItems(1)
End With

End Function
