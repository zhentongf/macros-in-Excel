Sub DisplaySelectionType()
    'Worksheets("Sheet1").Activate
    'MsgBox "The selection object type is " & TypeName(Selection)
    Debug.Print ("The selection object type is " & TypeName(Selection))
    'UserForm1.Show
End Sub
Sub SelectEmptyCells()
    ' select all blank cells. xlCellTypeBlanks = 4
    ' Application.Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.Selection.SpecialCells(4).Select
    
End Sub
' https://learn.microsoft.com/en-us/office/vba/api/office.filedialog
Sub FileDialog_Test()
 
 'Declare a variable as a FileDialog object.
 Dim fd As FileDialog
 
 'Create a FileDialog object as a File Picker dialog box.
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
 'Declare a variable to contain the path
 'of each selected item. Even though the path is aString,
 'the variable must be a Variant because For Each...Next
 'routines only work with Variants and Objects.
 Dim vrtSelectedItem As Variant
 
 'Use a With...End With block to reference the FileDialog object.
 With fd
 
 'Use the Show method to display the File Picker dialog box and return the user's action.
 'The user pressed the button.
 If .Show = -1 Then
 
 'Step through each string in the FileDialogSelectedItems collection.
 For Each vrtSelectedItem In .SelectedItems
 
 'vrtSelectedItem is aString that contains the path of each selected item.
 'Use any file I/O functions that you want to work with this path.
 'This example displays the path in a message box.
 'MsgBox "The path is: " & vrtSelectedItem
 MsgBox ("The path is: " & vrtSelectedItem)
 
 Next vrtSelectedItem
 'The user pressed Cancel.
 Else
 End If
 End With
 
 'Set the object variable to Nothing.
 Set fd = Nothing
 
End Sub
Sub Copy_and_Export_Sheets()
    Application.ScreenUpdating = False

    Dim wb, wb_D, wb_N, wb_B, wb_L, wb_R As Workbook
    Dim ws As Worksheet
    Dim strColName, str2
    Dim rowID, colID, counter1, counter2 As Integer
    Dim xRow, xColumn As Long
    Dim startTime, endTime, elapsedTime
    
    startTime = Time
    Set wb_D = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B2").Value, Editable:=True)
    Set wb_N = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B3").Value, Editable:=True)
    Set wb_B = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B4").Value, Editable:=True)
    Set wb_L = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B5").Value, Editable:=True)
    Set wb_R = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B6").Value, Editable:=True)
    If wb_D.Worksheets(1).AutoFilterMode Then wb_D.Worksheets(1).AutoFilter.ShowAllData
    If wb_N.Worksheets(1).AutoFilterMode Then wb_N.Worksheets(1).AutoFilter.ShowAllData
    If wb_B.Worksheets(1).AutoFilterMode Then wb_B.Worksheets(1).AutoFilter.ShowAllData
    If wb_L.Worksheets("DDS order").AutoFilterMode Then wb_L.Worksheets("DDS order").AutoFilter.ShowAllData
    If wb_L.Worksheets("NL order").AutoFilterMode Then wb_L.Worksheets("NL order").AutoFilter.ShowAllData
    If wb_L.Worksheets("Back order").AutoFilterMode Then wb_L.Worksheets("Back order").AutoFilter.ShowAllData
    If wb_R.Worksheets(1).AutoFilterMode Then wb_R.Worksheets(1).AutoFilter.ShowAllData
    For counter1 = 1 To 4
        ThisWorkbook.Worksheets(counter1).Activate
        If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
        With Cells
            .NumberFormat = "General"
            .ColumnWidth = 8.43
            .HorizontalAlignment = xlGeneral
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        If counter1 < 3 Then Range("J:J,Q:Q").NumberFormat = "m/d/yyyy"
        If counter1 = 3 Then Range("I:I,P:P").NumberFormat = "m/d/yyyy"
    Next
    
    'paste column headings
    ThisWorkbook.Worksheets("Sheet2").Activate
    Range("A11:R11").Copy
    ThisWorkbook.Worksheets("DDS order").Activate
    Range("A1").PasteSpecial
    Range("A:R").ColumnWidth = 10
    ThisWorkbook.Worksheets("Sheet2").Activate
    Range("A13:S13").Copy
    ThisWorkbook.Worksheets("NL order").Activate
    Range("A1").PasteSpecial
    Range("A:S").ColumnWidth = 10
    ThisWorkbook.Worksheets("Sheet2").Activate
    Range("A15:R15").Copy
    ThisWorkbook.Worksheets("Back order").Activate
    Range("A1").PasteSpecial
    Range("A:R").ColumnWidth = 10

    ' copy DDS sheet
    ThisWorkbook.Worksheets("DDS order").Activate
    Range("A1:H1, R1").Select
    For Each strColName In Selection
        colID = Find_Col_ID(1, wb_D, wb_D.Worksheets("Sheet1"), strColName)
        If colID = 0 Then
            MsgBox "( " & strColName & " ): Cannot find the column."
            Exit Sub
        Else
            Columns(colID).Copy
            ThisWorkbook.Worksheets("DDS order").Activate
            Columns(strColName.Column).PasteSpecial (xlPasteValues)
        End If
    Next

    ' copy NL sheet
    ThisWorkbook.Worksheets("NL order").Activate
    Range("A1:H1, R1:S1").Select
    For Each strColName In Selection
        colID = Find_Col_ID(1, wb_N, wb_N.Worksheets("Sheet1"), strColName)
        If colID = 0 Then
            MsgBox "( " & strColName & " ): Cannot find the column."
            Exit Sub
        Else
            Columns(colID).Copy
            ThisWorkbook.Worksheets("NL order").Activate
            Columns(strColName.Column).PasteSpecial (xlPasteValues)
        End If
    Next

    ' copy BO sheet
    ThisWorkbook.Worksheets("Back order").Activate
    Range("A1:G1, Q1:R1").Select
    For Each strColName In Selection
        colID = Find_Col_ID(1, wb_B, wb_B.Worksheets("Sheet1"), strColName)
        If colID = 0 Then
            MsgBox "( " & strColName & " ): Cannot find the column."
            Exit Sub
        Else
            Columns(colID).Copy
            ThisWorkbook.Worksheets("Back order").Activate
            Columns(strColName.Column).PasteSpecial (xlPasteValues)
        End If
    Next

    ' copy Last Week sheet
    wb_L.Worksheets("DDS order").Activate
    Range("I:J").Copy
    ThisWorkbook.Worksheets("Sheet1").Activate
    Range("L:M").PasteSpecial (xlPasteValues)
    wb_L.Worksheets("NL order").Activate
    Range("I:J").Copy
    ThisWorkbook.Worksheets("Sheet1").Activate
    Range("O:P").PasteSpecial (xlPasteValues)
    wb_L.Worksheets("Back order").Activate
    Range("H:I").Copy
    ThisWorkbook.Worksheets("Sheet1").Activate
    Range("R:S").PasteSpecial (xlPasteValues)

    ' copy RPA result sheet
    ThisWorkbook.Worksheets("Sheet1").Activate
    Range("A1:B1, D1:I1").Select
    For Each strColName In Selection
        colID = Find_Col_ID(1, wb_R, wb_R.Worksheets("result"), strColName)
        If colID = 0 Then
            MsgBox "( " & strColName & " ): Cannot find the column."
            Exit Sub
        Else
            Columns(colID).Copy
            ThisWorkbook.Worksheets("Sheet1").Activate
            Columns(strColName.Column).PasteSpecial (xlPasteValues)
        End If
    Next
    ThisWorkbook.Worksheets("Sheet1").Activate
    Range("C:C").Clear
    Range("C:C").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("C1").Value = "combined"
    
    ' fill the blank cells
    ThisWorkbook.Worksheets("DDS order").Activate
    Range("A:H").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    ThisWorkbook.Worksheets("NL order").Activate
    Range("A:D, R:R").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    ThisWorkbook.Worksheets("Back order").Activate
    Range("A:D").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    
    ' paste the formulas.
    For counter1 = 1 To 3
        ThisWorkbook.Worksheets(counter1).Activate
        If counter1 = 1 Then
            Range("I2").Formula = "=B2&E2"
            Range("J2").Formula = "=VLOOKUP(I2,Sheet1!L:M,2,0)"
            str2 = "I2"
            colID = 9
        End If
        If counter1 = 2 Then
            Range("I2").Formula = "=B2&E2"
            Range("J2").Formula = "=VLOOKUP(I2,Sheet1!O:P,2,0)"
            str2 = "I2"
            colID = 9
        End If
        If counter1 = 3 Then
            Range("H2").Formula = "=B2&E2"
            Range("I2").Formula = "=VLOOKUP(H2,Sheet1!R:S,2,0)"
            str2 = "H2"
            colID = 8
        End If
        For counter2 = 2 To 7
            Cells(2, colID + counter2).Formula = "=VLOOKUP(" & str2 & ",Sheet1!C:I," & counter2 & ",0)"
        Next
    Next
    
    'Cells(Rows.Count, 1).End(xlUp).Row   This is a great way to find the last row.
    ThisWorkbook.Worksheets("DDS order").Activate
    xRow = Cells(Rows.Count, "E").End(xlUp).Row
    Range("I2:P2").AutoFill Destination:=Range("I2:P2").Resize(xRow - 1)
    ThisWorkbook.Worksheets("NL order").Activate
    xRow = Cells(Rows.Count, "E").End(xlUp).Row
    Range("I2:P2").AutoFill Destination:=Range("I2:P2").Resize(xRow - 1)
    ThisWorkbook.Worksheets("Back order").Activate
    xRow = Cells(Rows.Count, "E").End(xlUp).Row
    Range("H2:O2").AutoFill Destination:=Range("H2:O2").Resize(xRow - 1)
    
    'ThisWorkbook.Worksheets("DDS order").Activate
    'Range("A1:G1").ClearContents
    'ThisWorkbook.Worksheets("NL order").Activate
    'Range("A1:G1").ClearContents
    'ThisWorkbook.Worksheets("Back order").Activate
    'Range("A1:G1").ClearContents
    'ThisWorkbook.Worksheets("Sheet2").Activate
    'Range("B1:H1").Copy
    'ThisWorkbook.Worksheets("DDS order").Activate
    'Range("A1:G1").PasteSpecial (xlPasteValues)
    'ThisWorkbook.Worksheets("NL order").Activate
    'Range("A1:G1").PasteSpecial (xlPasteValues)
    'ThisWorkbook.Worksheets("Back order").Activate
    'Range("A1:G1").PasteSpecial (xlPasteValues)
    
    ' remove formulas and paste values.
    ThisWorkbook.Worksheets("DDS order").Activate
    Cells.Copy
    Cells.PasteSpecial (xlPasteValues)
    Cells.Font.Name = "Calibri"
    Cells.Font.Size = 11
    ThisWorkbook.Worksheets("NL order").Activate
    Cells.Copy
    Cells.PasteSpecial (xlPasteValues)
    Cells.Font.Name = "Calibri"
    Cells.Font.Size = 11
    ThisWorkbook.Worksheets("Back order").Activate
    Cells.Copy
    Cells.PasteSpecial (xlPasteValues)
    Cells.Font.Name = "Calibri"
    Cells.Font.Size = 11
    
    ' copy to a new workbook
    ThisWorkbook.Worksheets(Array("DDS order", "NL order", "Back order")).Copy
    str2 = ActiveWorkbook.Name
    For counter1 = 1 To 3
        ActiveWorkbook.Worksheets(counter1).Activate
        Cells.Font.Name = "Calibri"
        Cells.Font.Size = 11
    Next
    For counter1 = 1 To 4
        ThisWorkbook.Worksheets(counter1).Activate
        counter2 = Cells(Rows.Count, "B").End(xlUp).Row
        Range("2:2", Cells(counter2, 2)).Delete
    Next
    
    'myFileName = "working-ETA--" & Year(Date) & "-" & Month(Date) & "-" & Day(Date)
    'If Dir(ThisWorkbook.Path & "\" & myFileName & ".xlsx") <> "" Then
    '    'file does exist
    '    For counter1 = 1 To 100
    '        If counter1 = 100 Then
    '            MsgBox "There are too many files with this name."
    '            Exit Sub
    '        End If
    '        myFileName = myFileName & " (" & counter1 & ")"
    '        If Dir(ThisWorkbook.Path & "\" & myFileName & ".xlsx") = "" Then
    '            'file does not exist
    '            Exit For
    '        End If
    '    Next
    'End If
    'ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & myFileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    
    wb_D.Close SaveChanges:=False
    wb_N.Close SaveChanges:=False
    wb_B.Close SaveChanges:=False
    wb_L.Close SaveChanges:=False
    wb_R.Close SaveChanges:=False
    
    Workbooks(str2).Activate
    Debug.Print str2

    endTime = Time
    elapsedTime = (endTime - startTime) * 24 * 60 * 60
    Debug.Print "startTime: " & startTime & Chr(13) & _
    "endTime: " & endTime & Chr(13) & "duration: " & _
    elapsedTime & " seconds"
    Application.ScreenUpdating = True
'With
    'wb_D.Worksheets(1).Range("A:G").Select
'End With
'wb_D.Worksheets("Sheet1").Columns("A:G").Select
'Set ws = wb.Worksheets("Sheet1")
'ws.Range("B2") = 3
'ThisWorkbook.ActiveSheet.Activate
'Must clear the filter first, before copy the column.
'wb.Worksheets("Sheet1").Columns(2).Copy _
    Destination:=Workbooks("ref.xlsx").Worksheets("Sheet2").Columns(2)
'wb.Close
'clear buffer
End Sub

'Looking for column index.
Public Function Find_Col_ID(ByVal rowID, ByVal objWorkBook, ByVal objWorkSheet, ByVal strColName) As Integer
'https://www.cnblogs.com/sitongyan/p/16168727.html
    objWorkBook.Activate
    objWorkSheet.Select
    'objWorkSheet.Cells(1, 1).Select
    Rows(rowID).Select
    'Rows(rowID).Find(what:=strColName, after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select
    If TypeName(Rows(rowID).Find(what:=strColName, after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)) <> "Nothing" Then
        Rows(rowID).Find(what:=strColName, after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select
        Find_Col_ID = Selection.Column
    Else
        Find_Col_ID = 0
    End If
 
End Function
Sub formatting()
    Application.ScreenUpdating = False
    
    Set wb_U = Workbooks.Open(Filename:=ThisWorkbook.Worksheets("control").Range("B15").Value, Editable:=True)
    
    ' processing DDS sheet
    wb_U.Worksheets("DDS order").Activate
    If wb_U.Worksheets("DDS order").AutoFilterMode Then wb_U.Worksheets("DDS order").AutoFilter.ShowAllData
    Cells.Copy
    ThisWorkbook.Worksheets("DDS order").Activate
    Cells.PasteSpecial (xlPasteAll)
    Range("A:A").Insert
    Range("A:A").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=IF(RC[+1]<10000,RC[+1]+10000,RC[+1])"
    Range("A1").Value = "Terr_2"
    ' if AutoFilterMode is true, then remove AutoFilter.
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Range("A:S").AutoFilter
    Range("A:S").AutoFilter Field:=16, Criteria1:="Invoiced"
    Range("A:S").AutoFilter Field:=11, Criteria1:="=Invoiced", Operator:=xlOr, Criteria2:="=shipped"
    Range("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("L:S").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:G").Select
    Columns("A:G").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    Columns("H:H").EntireColumn.ColumnWidth = 10#
    Columns("H:H").NumberFormat = "m/d/yyyy"
    
    ' processing NL sheet
    wb_U.Worksheets("NL order").Activate
    If wb_U.Worksheets("NL order").AutoFilterMode Then wb_U.Worksheets("NL order").AutoFilter.ShowAllData
    Cells.Copy
    ThisWorkbook.Worksheets("NL order").Activate
    Cells.PasteSpecial (xlPasteAll)
    Range("A:A").Insert
    Range("A:A").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=IF(RC[+1]<10000,RC[+1]+10000,RC[+1])"
    Range("A1").Value = "Terr_2"
    ' if AutoFilterMode is true, then remove AutoFilter.
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Range("A:T").AutoFilter
    Range("A:T").AutoFilter Field:=16, Criteria1:="Invoiced"
    Range("A:T").AutoFilter Field:=11, Criteria1:="=Invoiced", Operator:=xlOr, Criteria2:="=shipped"
    Range("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("L:T").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:G").Select
    Columns("A:G").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    Columns("H:H").EntireColumn.ColumnWidth = 10#
    Columns("H:H").NumberFormat = "m/d/yyyy"
    
    ' processing BO sheet
    wb_U.Worksheets("Back order").Activate
    If wb_U.Worksheets("Back order").AutoFilterMode Then wb_U.Worksheets("Back order").AutoFilter.ShowAllData
    Cells.Copy
    ThisWorkbook.Worksheets("Back order").Activate
    Cells.PasteSpecial (xlPasteAll)
    Range("A:A").Insert
    Range("A:A").SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=IF(RC[+1]<10000,RC[+1]+10000,RC[+1])"
    Range("A1").Value = "Terr_2"
    ' if AutoFilterMode is true, then remove AutoFilter.
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Range("A:S").AutoFilter
    Range("A:S").AutoFilter Field:=15, Criteria1:="Invoiced"
    Range("A:S").AutoFilter Field:=10, Criteria1:="=Invoiced", Operator:=xlOr, Criteria2:="=shipped"
    Range("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("K:S").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("G:G").Select
    Columns("G:G").EntireColumn.ColumnWidth = 10#
    Columns("G:G").NumberFormat = "m/d/yyyy"
    
    Sheets("Sheet2").Select
    Range("B1:I1").Select
    Selection.Copy
    Sheets("DDS order").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("NL order").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Back order").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("G1").Clear
    Range("H1").Cut
    Range("G1").Select
    ActiveSheet.Paste
    
    ThisWorkbook.Worksheets(Array("DDS order", "NL order", "Back order")).Copy
    str2 = ActiveWorkbook.Name
    For counter1 = 1 To 3
        ActiveWorkbook.Worksheets(counter1).Activate
        Cells.Font.Name = "Calibri"
        Cells.Font.Size = 11
        If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Next
    For counter1 = 1 To 3
        ThisWorkbook.Worksheets(counter1).Activate
        counter2 = Cells(Rows.Count, "B").End(xlUp).Row
        Range("2:2", Cells(counter2, 2)).Delete
        If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Next
    
    wb_U.Close SaveChanges:=False
    
    Workbooks(str2).Activate
    
    Application.ScreenUpdating = True
End Sub

Sub TestShowAllData()
ActiveSheet.ShowAllData
 'Worksheets("Sheet1").Range("A1:J10").Activate
If Worksheets("Sheet1").FilterMode Then
Worksheets("Sheet1").ShowAllData
MsgBox "filter"
End If
End Sub
Sub Button1_Click()
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 Dim vrtSelectedItem As Variant
 With fd
  'Use the Show method to display the File Picker dialog box
  If .Show = -1 Then
   If .SelectedItems.Count <> 1 Then
    MsgBox ("You can only select one file!")
    'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
    Exit Sub
   End If
   Range("B2").Value = .SelectedItems.Item(1)
  End If
 End With
End Sub
Sub Button2_Click()
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 Dim vrtSelectedItem As Variant
 With fd
  'Use the Show method to display the File Picker dialog box
  If .Show = -1 Then
   If .SelectedItems.Count <> 1 Then
    MsgBox ("You can only select one file!")
    'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
    Exit Sub
   End If
   Range("B3").Value = .SelectedItems.Item(1)
  End If
 End With
End Sub
Sub Button3_Click()
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 Dim vrtSelectedItem As Variant
 With fd
  'Use the Show method to display the File Picker dialog box
  If .Show = -1 Then
   If .SelectedItems.Count <> 1 Then
    MsgBox ("You can only select one file!")
    'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
    Exit Sub
   End If
   Range("B4").Value = .SelectedItems.Item(1)
  End If
 End With
End Sub
Sub Button4_Click()
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 Dim vrtSelectedItem As Variant
 With fd
  'Use the Show method to display the File Picker dialog box
  If .Show = -1 Then
   If .SelectedItems.Count <> 1 Then
    MsgBox ("You can only select one file!")
    'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
    Exit Sub
   End If
   Range("B5").Value = .SelectedItems.Item(1)
  End If
 End With
End Sub
Sub Button5_Click()
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 Dim vrtSelectedItem As Variant
 With fd
  'Use the Show method to display the File Picker dialog box
  If .Show = -1 Then
   If .SelectedItems.Count <> 1 Then
    MsgBox ("You can only select one file!")
    'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
    Exit Sub
   End If
   Range("B6").Value = .SelectedItems.Item(1)
  End If
 End With
End Sub
Sub Button6_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim vrtSelectedItem As Variant
    With fd
    'Use the Show method to display the File Picker dialog box
        If .Show = -1 Then
            If .SelectedItems.Count <> 1 Then
                MsgBox ("You can only select one file!")
                'learn from this website. https://www.cnblogs.com/Young-shi/p/11690393.html
                Exit Sub
            End If
            Range("B15").Value = .SelectedItems.Item(1)
        End If
    End With
End Sub
