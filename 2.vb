Sub MergeExcelFiles()
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            'clean content from original file
            
             Sheets("Sales Forecast Input").Select
            Cells(1, 1).Select
            Selection.End(xlDown).Select
            x = Selection.End(xlToRight).Column
            
            Z = Selection.Row
            Selection.End(xlDown).Select
            If Z > 10 Then
                Rows(2).Select
                Z = 2
            Else
                Z = Z + 1
                Rows(Z).Select
            End If
       ' Code to clear content if memory dump is presented
       '     y = Selection.End(xlDown).Row
       '     If y > 50000 Then
       '         i = 1
       '         While i <= x
       '         Range(Cells(Z, 1), Cells(y, i)).Clear
       '         i = i + 1
       '         Wend
       '     Else
                Range(Selection, Selection.End(xlDown)).Select
                Application.CutCopyMode = False
                Selection.Clear
       '     End If
                    
                    
       
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
                Application.DisplayAlerts = False
                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
                
                On Error Resume Next
                Sheets("Sales Forecast Input").Select
                Cells(1, 1).Select
                Selection.End(xlDown).Select
                z1 = Selection.Row
                If z1 > 10 Then
                    Rows(2).Select
                    z1 = 2
                Else
                    z1 = z1 + 1
                    Rows(z1).Select
                End If
                Cells(z1, 1).Select
                y1 = Selection.End(xlDown).Row
                If y1 > 1048550 Then
                Rows(z1).Select
                Else
                Rows(z1).Select
                Range(Selection, Selection.End(xlDown)).Select
                End If
                Selection.Copy
                wbkCurBook.Activate
                Sheets("Sales Forecast Input").Select
                Cells(Z, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Cells(Z, 1).Select
                y1 = Selection.End(xlDown).Row
                If y1 > 1048550 Then
                    Z = Z + 1
                Else
                    Selection.End(xlDown).Select
                    Z = Selection.Row + 1
                End If
                Application.DisplayAlerts = False
                
                wbkSrcBook.Close SaveChanges:=False
                Application.DisplayAlerts = True
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
            MsgBox "Processed " & countFiles & " files" & vbCrLf, Title:="Merge Excel files"
        End If
 
    Else
        MsgBox "No files selected", Title:="Merge Excel files"
    End If
    Worksheets("load").Select
End Sub


Sub UpdateFile()
'turn off screenuptade to increase performance
'Application.ScreenUpdating = False
'Variables
Dim pivotfilter As String 'pivotfilter as the first column of the report

Dim location As String
Dim fname As Variant 'raw data input file name
Dim cl As Range 'variable to be used in find procedure
Dim MySheet As Workbook 'spreadsheet name
Dim Mydata As Workbook 'spreadsheet name
Dim multiplefiles As String 'option to slipt in multiple files
Dim accmng As String 'account manager name to put in file path
Dim tab_name As String
Dim ro As Integer


'Set Variables values

'avoit prom

'Define the name of this sheet for future saving purposes

Set MySheet = ActiveWorkbook

Sheets("Update File").Select
'define location path *note that \ must be used by the end of the path ex: c: is wrong c:\ should be used
location = Application.ActiveWorkbook.Path & "\"
'Cells(1, 17).Value = location
multiplefiles = Cells(23, 1).Value
tab_name = Cells(20, 1).Value
deleopt = Cells(26, 1).Value

' UpdateFile Macro
  fname = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose OMP Raw Data Excel file", MultiSelect:=False)
    On Error GoTo procext
    Workbooks.Open Filename:=fname
    Set Mydata = ActiveWorkbook
    Sheets(tab_name).Select
    Cells.Select
    Selection.Copy
    MySheet.Activate
    Sheets(tab_name).Select
    Cells.Select
    ActiveSheet.Paste
 
 
 'find the right column where macro should consider the previous as attributes
With Worksheets(tab_name).Cells
  Set cl = .Find("bucket", after:=.Range("A2"), LookIn:=xlValues)
  If Not cl Is Nothing Then
        cl.Select
    End If
End With
   
      
        'Define the range of cells to consider as attributes cells
    x = Selection.Column
    findresult = Selection.Column
    Range("a1").Select
    
    ActiveCell.Formula = _
        "=DATE(RIGHT(I1,4),MID(I1,2,1),1)"
    'ActiveCell.FormulaR1C1 = _
    '    "=RIGHT(""00""&IF(RIGHT(RIGHT(LEFT(RC[" & x & "],3),2),1)="":"",LEFT(RIGHT(LEFT(RC[" & x & "],3),2),1),RIGHT(LEFT(RC[" & x & "],3),2)),2)&""/""&RIGHT(RC[" & x & "],4)"
    
    Cells(Selection.Row, x + 1).Select
    
    Selection.End(xlDown).Select
    y = Selection.Row
    
    Range(Cells(3, 1), Cells(y - 1, x)).Select
    Application.CutCopyMode = False
    Selection.UnMerge
    If Cells(y, 1).Value = "Total" Then 'verifies if the raw data comes with a total row
    Range(Cells(3, 1), Cells(y - 1, x)).Select
    
    'Fill the blank cells with the upper value. Colum A
    
    Range(Cells(3, 1), Cells(y, 1)).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    
  
    
    'define row before accountMgn
    
    Dim j As Integer
    ActiveSheet.Select
    j = [a1:c100].Find("AccountMgr Name").Row

    linha = j + 1

        
    Do While f <= (x)
   
   f = f + 1
          
    ActiveSheet.Columns(f).Select
      
    Dim h As Integer

    h = (ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row) - 1

    For i = linha To h

    If Sheets(tab_name).Cells(i, f) = "" Then
    Sheets(tab_name).Cells(i, f) = "BLANK"
    If Sheets(tab_name).Cells(i, f - 1).Value = Sheets(tab_name).Cells(i - 1, f - 1) Then
    Sheets(tab_name).Cells(i, f) = Sheets(tab_name).Cells(i - 1, f)

    End If
    End If
    Next i
    
    Loop
        
    
    
    'Copy formula to values in Sales Forecast Input Tab
    Range(Cells(3, 1), Cells(y - 1, x)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Else
      Range(Cells(3, 1), Cells(y, x)).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    
    'Fill the blank cells with the upper value
    Selection.FormulaR1C1 = "=R[-1]C"
    
    'Copy formula to values in Sales Forecast Input Tab
    Range(Cells(3, 1), Cells(y, x)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
   
    
   
  'Sheets("Sales Forecast Input").Clean
    
    
    'copy formulas
Sheets("Formulas").Select
Cells(y, 1).Select
Range(Selection, Selection.End(xlUp)).Select
   
   Range(Rows(1), Rows(Selection.Row)).Select
   Selection.Copy
   Sheets("Sales Forecast Input").Select
   Cells(1, 1).Select
   ActiveSheet.Paste
    
    Selection.EntireRow.Hidden = False

'add formulas to sales forecast input tab
 Sheets("Sales Forecast Input").Select
    Rows("1:2").EntireRow.Hidden = False
    Cells(1, 1).Select
     Selection.End(xlDown).Select
    pivotfilter = Selection.Value 'set the pivotfilter name as the first column in the sales forecast input tab
    Z = Selection.Row
    'add a temp row to formula creation
    If Z > 10 Then
    Rows(2).Select
    Z = 2
    Else
     Z = Z + 1
    Rows(Z).Select
   
    End If
    
   ' Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'clean old values
    Rows(Z + 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    Sheets(tab_name).Select
    Selection.Copy
    Sheets("Sales Forecast Input").Select
    Cells(Z + 1, 1).Select
    ActiveSheet.Paste
    Cells(Z, x).Select
    Selection.End(xlDown).Select
    y = Selection.Row
    Cells(Z - 1, x).Select
    Selection.End(xlToRight).Select
    l = Selection.Column
    Columns(x).Resize(, l - 18).EntireColumn.Select
    Selection.EntireColumn.Hidden = False
    
    
    
    'loop to copy columns and paste values in order to save memory resources
    yloop1 = Z + 1
    
    If yloop2 + 100 < y Then
    yloop2 = yloop1 + 100
    Else
    yloop2 = y
    End If
    
    While yloop1 <= y
    Application.Calculation = xlCalculationAutomatic
    Range(Cells(Z, x + 1), Cells(Z, l)).Select
    Selection.Copy
   Range(Cells(yloop1, x + 1), Cells(yloop2, l)).Select
    ActiveSheet.Paste
    Range(Cells(yloop1, x + 1), Cells(yloop2, l)).Select
    Selection.Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    If yloop1 + 100 < y Then
    yloop1 = yloop1 + 101
    Else
    yloop1 = yloop2 + 1
    End If
    If yloop2 + 100 < y Then
    yloop2 = yloop1 + 100
    Else
    yloop2 = y
    End If
    
    Application.Calculation = xlCalculationManual
    Wend
     
    Rows(Z).Delete
    Rows("1:2").EntireRow.Hidden = True
    
   'copy and past format from formula tab to input tab
    
   Sheets("Formulas").Select
        Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    lastColumn = Selection.Column

    Range(Cells(1, 1), Cells(y - 1, lastColumn)).Select
    Selection.Copy
    Sheets("Sales Forecast Input").Select
    Range(Cells(1, 1), Cells(y - 1, lastColumn)).Select

    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False


    
    
    
    'refresh pivot table
    
        ActiveWorkbook.RefreshAll
        
        Sheets("Pivot Report").Select
        
           ActiveSheet.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sales Forecast Input!R" & Z - 1 & "C1:R" & y & "C" & l, Version:=xlPivotTableVersion15)
        
        
        With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
            .ClearAllFilters
        End With
       On Error Resume Next
    '   With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
          
    '     .PivotItems("CV").Visible = False
    '   On Error GoTo 0
    '    End With
    '    With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
      '       .PivotItems("(blank)").Visible = False
     '   End With
    ActiveSheet.Range("A1").Select
    Mydata.Activate
    Mydata.Close
  
    MySheet.Activate
    ActiveWorkbook.Save
    
    'save different files for each acc managers
    If multiplefiles = "YES" Then
    'copying aacmanagers to do the loop in file saving
    Sheets("Sales Forecast Input").Select
    Range("A" & y - 1).Select
    Range(Selection, Selection.End(xlUp)).Select

    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pivot report").Select
    Range("A1").Select
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
 
    
       ActiveSheet.Range("$XFD$1:$XFD$" & y).RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
    Range("$XFD1").Select
    Selection.End(xlDown).Select
    Z = Selection.Row
    
    i = 1
    
    
    While i < Z
  Aplication.DisplayAlerts = False
    Sheets(Array("Sales Forecast Input", "Pivot Report")).Select
    Sheets(Array("Sales Forecast Input", "Pivot Report")).Copy
    Sheets("pivot report").Select
    
    accmng = Range("$XFD" & Z).Value
    Sheets("Sales Forecast Input").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    r = Selection.Row
    ActiveSheet.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).AutoFilter Field:=1, Criteria1:= _
        "<>" & accmng, Operator:=xlAnd
    Range("A" & r + 1).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
     Selection.ClearContents
    Selection.ClearFormats
    
    Cells.Select
    Range("A" & r + 1).Activate
    Selection.AutoFilter
    
    
    Range("A1000000").Select
    Selection.End(xlUp).Select
    y = Selection.Row
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    x = Selection.Column
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).RemoveDuplicates Columns:=findresult + 1, Header:=xlYes
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).Copy
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).PasteSpecial Paste:=xlPasteValues
    
    'delete first blank line
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
    Rows(Selection.Row + 1).Delete
    Else
     Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
        Rows(Selection.Row + 1).Delete
    End If
    End If
    
     'clear sum zero lines
    If deleopt = "YES" Then
     Range("A1").Select
    Selection.End(xlDown).Select
    y1 = Selection.Row
    ActiveSheet.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).AutoFilter Field:=51, Criteria1:="0", Operator:=xlAnd
    Range("A1000000").Select
    Selection.End(xlUp).Select
    y = Selection.Row
    Range("A" & r + 1).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
     
    If y > y1 Then
    Selection.ClearContents
    Selection.ClearFormats
    Else
    End If
    Cells.Select
    Range("A" & r + 1).Activate
    Selection.AutoFilter
    Range("A1000000").Select
    Selection.End(xlUp).Select
    y = Selection.Row
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    x = Selection.Column
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).RemoveDuplicates Columns:=findresult + 1, Header:=xlYes
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).Copy
    ActiveSheet.Range(Cells(r, 1), Cells(y, x)).PasteSpecial Paste:=xlPasteValues
    'delete first blank line
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
    Rows(Selection.Row + 1).Delete
   Else
    Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
        Rows(Selection.Row + 1).Delete
    End If
   End If
   
   'copy format from formula tab
   
   Set mysheet2 = ActiveWorkbook
   
   Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    y = Selection.Row
   MySheet.Activate
   
   Sheets("Formulas").Select
   Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    lastColumn = Selection.Column
    Range(Cells(1, 1), Cells(y, lastColumn)).Select
    Selection.Copy
    mysheet2.Activate
    Sheets("Sales Forecast Input").Select
    Range(Cells(1, 1), Cells(y, lastColumn)).Select

    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    End If
 
    Selection.End(xlToLeft).Select
    
    
    
    Sheets("pivot report").Select
    
    Columns("XFD:XFD").Select
    Selection.ClearContents
    Selection.End(xlToLeft).Select
    Z = Z - 1
    'refresh pivot to just this account manager
    
    
 Sheets("Pivot Report").Select
     
     ActiveSheet.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sales Forecast Input!R" & r & "C1:R" & y & "C" & l, Version:=xlPivotTableVersion15)
        
        Sheets("Pivot Report").Select
        With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
            .ClearAllFilters
        End With
       On Error Resume Next
       With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
         .PivotItems("CV").Visible = False
       On Error GoTo 0
        End With
       On Error Resume Next
        With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
            .PivotItems("(blank)").Visible = False
        End With
        
        
    
    ChDir location
    
accmng = Replace(Replace(accmng, "/", "-"), "-", "-")
accmng = Replace(Replace(accmng, "\", "-"), "-", "-")
    On Error Resume Next
    
   
      
     
    
    'collapse columns
    Sheets("Sales Forecast InpuT").Select
    
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    
        
    ActiveWorkbook.SaveAs Filename:= _
        location & "SalesForecastCollection_" & Format(Date, "mm-yy -") & accmng & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Sheets("Pivot Report").Select
    Cells(1, 1).Select
    ActiveSheet.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sales Forecast Input!R" & r & "C1:R" & y & "C" & l, Version:=xlPivotTableVersion15)
    ActiveWorkbook.Save
        
            
    Aplication.DisplayAlerts = True
    ActiveWindow.Close
    MySheet.Activate
    Sheets("Update File").Select
    Wend
 Else
    Sheets(Array("Sales Forecast Input", "Pivot Report")).Select
    Sheets(Array("Sales Forecast Input", "Pivot Report")).Copy
    Sheets("Sales Forecast Input").Select
    
    
       'clear sum zero lines
    If deleopt = "YES" Then
     Range("A1").Select
    Selection.End(xlDown).Select
    y1 = Selection.Row
    ActiveSheet.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).AutoFilter Field:=47, Criteria1:="0", Operator:=xlAnd
    Range("A1000000").Select
    Selection.End(xlUp).Select
    y = Selection.Row
    Range("A" & y1 + 1).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
     
    If y > y1 Then
    Selection.ClearContents
    Selection.ClearFormats
    Else
    End If
    Cells.Select
    Range("A" & y1 + 1).Activate
    Selection.AutoFilter
    Range("A1000000").Select
    Selection.End(xlUp).Select
    y = Selection.Row
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    x = Selection.Column
    ActiveSheet.Range(Cells(y1, 1), Cells(y, x)).RemoveDuplicates Columns:=findresult + 1, Header:=xlYes
    ActiveSheet.Range(Cells(y1, 1), Cells(y, x)).Copy
    ActiveSheet.Range(Cells(y1, 1), Cells(y, x)).PasteSpecial Paste:=xlPasteValues
    'delete first blank line
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
    Rows(Selection.Row + 1).Delete
   Else
    Selection.End(xlDown).Select
    If Cells(Selection.Row + 1, findresult + 1).Value = "" Then
        Rows(Selection.Row + 1).Delete
    End If
   End If
   
   'copy format from formula tab
   
   Set mysheet2 = ActiveWorkbook
   
   Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    y = Selection.Row
   MySheet.Activate
   
   Sheets("Formulas").Select
   Cells(1, 1).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    lastColumn = Selection.Column
    Range(Cells(1, 1), Cells(y, lastColumn)).Select
    Selection.Copy
    mysheet2.Activate
    Sheets("Sales Forecast Input").Select
    Range(Cells(1, 1), Cells(y, lastColumn)).Select

    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    End If
 'collapse columns
    Sheets("Sales Forecast InpuT").Select
    
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    

     'refresh pivot table
    
        ActiveWorkbook.RefreshAll
        
        Sheets("Pivot Report").Select
        
           ActiveSheet.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sales Forecast Input!R" & Z - 1 & "C1:R" & y & "C" & l, Version:=xlPivotTableVersion15)
        
        
        With ActiveSheet.PivotTables("PivotTable1").PivotFields(pivotfilter)
            .ClearAllFilters
        End With
       On Error Resume Next
   
    
    
    ChDir location
   On Error Resume Next
    ActiveWorkbook.SaveAs Filename:= _
        location & "SalesForecastCollection_" & Format(Date, "mm-yy -") & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    MySheet.Activate
    
 
 End If
 
  Sheets("pivot report").Select
    
    Columns("XFD:XFD").Select
    Selection.ClearContents
    Selection.End(xlToLeft).Select
    Sheets("Update File").Select

    
    'turn on screenupdate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Worksheets("Update File").Select
    
procext:
    Exit Sub
End Sub
