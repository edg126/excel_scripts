Option Explicit




Sub SQL_Server_Create_inserts()
' v1.0 created 11/9/2015 by Ed Gallagher
' edg126@gmail.com
' blog: http://edgallagher.blogspot.com/
' this sub will append to the end of a dataset the create table and insert statements for a data set
'example: these columns and rows in an excel file:
'A B C
'1 2 3
'4 5 6
'7 8 9

'Would Produce:


'Create table #tmpTable ([A] varchar(max) ,[B] varchar(max) ,[C] varchar(max) )
'Insert into #tmpTable Values('1','2','3')
'Insert into #tmpTable Values('4','5','6')
'Insert into #tmpTable Values('7','8','9')
'SELECT TOP 200 * FROM #tmpTable

'To test, paste all the code into excel as a macro, then run the macro SQL_Server_Create_inserts.
'if you like it, here are the options to always have it appear in excel as an additional tab


'Excel 2016:
'1) open excel->developer tab -> visual basic
'2) paste the code, save as a microsoft add-in as 'data_utilities.xlam
'3) exit excel
'4) open excel->File->Options->Add-ins, highlight 'Data Utilities' and select 'Go'
'   Make sure data_utilities is selected
'5) Goto File -> Options -> Customize Ribbon
'   Then in addins, select data_utilities,
'   On the right, customize the Ribbon : should say 'Main Tabs'
'   Select New Tab, and rename as 'Personal Macros'
'   Select New Group, and rename as, 'SQL Tools'
'   On the Left, select Choose from Popular Commands: Macros
'   Highlight SQL_Server_Create_inserts
'   highlight on the right SQL Tools (custom)
'   Select Add >>
'   Select ok
'A tab should now appear with this macro.


'To remove all traces
'1) Goto File -> Options -> Customize Ribbon
'   On the Right, highlight 'Personal Macros (Custom), then select <<Remove
'   Select OK
'2) Goto File -> Options -> Add-ins, highlight 'Data_Utilities', then select Go
'   Uncheck Data_Utilities
'3) Goto C:\Users\<username>\AppData\Roaming\Microsoft\AddIns, and delete data_utilities.xlam
'4) Restart excel

'vb -> insert class module



Dim sqlTable As String
Dim lastCol As Long
Dim lastRow As Long
Dim currentRow As Long   'keeps track of the current row value while looping to build the insert statements
Dim currentColumn As Long
Dim headerCheck As Long  'answer to the prompt of if there is a header
Dim maxColumnLength() As Long
Dim columnLength As Long


lastCol = GetLastColumn(ActiveWorkbook.Name, ActiveSheet.Name)



sqlTable = InputBox("Enter Table Name (Prefix # for temporary table)", "SQLServer table prompt", "#tmpTable")
   If sqlTable = "" Then
      Exit Sub
   End If



headerCheck = MsgBox("Does file contain a header?", 3)
   If headerCheck = vbNo Then
      Call BuildHeaderColumns(lastCol)
   ElseIf headerCheck = 2 Then '2 = cancel
      Exit Sub
   End If


'we want to get the last row after the headerCheck, since that class inserts a new row
lastRow = GetLastRow(ActiveWorkbook.Name, ActiveSheet.Name)
ReDim maxColumnLength(1 To lastCol) As Long


 

'Check to see if there are a lot of records for this process to handle
Call CheckForLargeInserts(lastRow, lastCol)



'build the insert statements
For currentRow = 2 To lastRow
     For currentColumn = 1 To lastCol
         columnLength = 0
         columnLength = Len(ActiveWorkbook.ActiveSheet.Cells(currentRow, currentColumn).Value)
         
          If columnLength > maxColumnLength(currentColumn) Then
             maxColumnLength(currentColumn) = columnLength
          End If
          
     Next currentColumn
     
   ActiveSheet.Cells(currentRow, lastCol + 1).Value = buildInsertString(sqlTable, lastCol, currentRow, ActiveWorkbook.Name, ActiveSheet.Name)
   
Next currentRow

'build the header statement
Call buildCreateStatement(sqlTable, lastCol, ActiveWorkbook.Name, ActiveSheet.Name, maxColumnLength)

'build the last line that will show the table output
ActiveSheet.Cells(lastRow + 1, lastCol + 1).Value = "SELECT TOP 200 * FROM " + sqlTable


'copy the sql code to the clipboard
ActiveSheet.Range(Cells(1, lastCol + 1), Cells(lastRow + 1, lastCol + 1)).Copy
MsgBox "Insert statements have been copied to the clipboard, paste into Sql Server"

   

End Sub
Private Sub CheckForLargeInserts(lastRow As Long, lastCol As Long)
'While this function will work on any record size, excel can have issues with large volumes, and there might be constraints in the database with the volume of records.
'If you are using significant volumes of data, you should probably use a better tool than this.


Dim largeRecordCheck As Long 'answer to the prompt on if you want to continue if there is a large row size


If lastRow * lastCol > 1000000 Then
   largeRecordCheck = MsgBox("The row count, " + CStr(lastRow) + " is large, are you sure you want to continue?  It's strongly recommmended you use another method of loading data", vbYesNo)
   If largeRecordCheck = vbNo Then
      End ' exit program
   End If
End If

End Sub
 
Function GetLastColumn(wbName As String, wsName As String)
'This function returns the last column that has populated values in it, regardless of the amount of blank lines in between

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lCol As Long
    
    
    'Set ws = ThisWorkbook.Sheets(wsName)
    Set ws = Workbooks(wbName).Sheets(wsName)
     
    With ws
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            lCol = .Cells.Find(What:="*", _
                   After:=.Range("A1"), _
                   Lookat:=xlPart, _
                   LookIn:=xlValues, _
                   SearchOrder:=xlByColumns, _
                   SearchDirection:=xlPrevious, _
                   MatchCase:=False).Column
        Else
            lCol = 1
        End If
    End With
     
  

GetLastColumn = lCol


End Function
Private Sub BuildHeaderColumns(colNum As Long)
'Insert a row
'Populate a header column with format COLUMN_<columnnumber>

Rows(1).EntireRow.Insert

Dim headerString As String
Dim i As Long
Dim headerText As String

For i = 1 To colNum
  
headerText = "COLUMN_" + CStr(i)
  ActiveSheet.Cells(1, i).Value = headerText
Next i


End Sub

Function GetLastRow(wbName As String, wsName As String)
'This function returns the last row that has populated values in it, regardless of the amount of blank lines in between

    
    Dim ws As Worksheet
    Dim lRow As Long
    
    Set ws = Workbooks(wbName).Sheets(wsName)
         
    With ws
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            lRow = .Cells.Find(What:="*", _
                   After:=.Range("A1"), _
                   Lookat:=xlPart, _
                   LookIn:=xlValues, _
                   SearchOrder:=xlByRows, _
                   SearchDirection:=xlPrevious, _
                   MatchCase:=False).Row
        Else
            lRow = 1
        End If
    End With
    
    GetLastRow = lRow
End Function
Private Sub buildCreateStatement(sqlTable As String, colNum As Long, wbName As String, wsName As String, maxColumnLength() As Long)
'builds the create table output string based off of the header
  
  Dim headerString As String
  Dim i As Long
  Dim ws As Worksheet
  
  headerString = "Create table " + sqlTable + " ("

  For i = 1 To colNum
     If i <> 1 Then
        headerString = headerString + ","
     End If
'     headerString = headerString + "[" + ActiveSheet.Cells(1, i).Text + "] varchar(max) "
     headerString = headerString + "[" + Replace(Workbooks(wbName).Sheets(wsName).Cells(1, i).Text, "'", "''") + "] varchar(" + CStr(maxColumnLength(i)) + ") "
  Next i

  headerString = headerString + ")"

Workbooks(wbName).Sheets(wsName).Cells(1, colNum + 1) = headerString
  
End Sub


Function buildInsertString(sqlTable As String, colNum As Long, currentRow As Long, wbName As String, wsName As String)
'build the insert into statement from values in the sheet

   Dim insertString As String
   Dim i As Long
   Dim insertCell As String
   
   insertString = "Insert into " + sqlTable + " Values("

   For i = 1 To colNum
      If i <> 1 Then
         insertString = insertString + ","
      End If
      
            
      insertString = insertString + "'" + Replace(Workbooks(wbName).Sheets(wsName).Cells(currentRow, i).Text, "'", "''") + "'"
   Next i

   insertString = insertString + ")"

   buildInsertString = insertString
End Function




