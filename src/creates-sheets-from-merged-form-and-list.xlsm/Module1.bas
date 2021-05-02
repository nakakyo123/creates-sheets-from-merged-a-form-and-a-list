Attribute VB_Name = "Module1"
Sub SheetMerge_Click()


    Const sheetNameCol As String = "A"
    
    Dim dataStartRow, dataLastRow As Long
    Dim dataRows As String
    Dim cellZahyoRow As Long
    
    Application.ScreenUpdating = False

    Worksheets("DATABASE").Activate 'Activate Database

    dataStartRow = 5   'Row # of database start
        
    cellZahyoRow = 3   'Row # of  Form pointer
    
    dataLastRow = Cells(Rows.Count, 1).End(xlUp).Row 'Last row of database
    
    For i = dataStartRow To dataLastRow ' i is the row seeking
        
        Sheets("FORM").Copy Before:=Sheets("FORM") 'create a new sheet
        
        Dim newSheetName As String
        newSheetName = Worksheets("DATABASE").Range(sheetNameCol & i)
        
        newSheetName = Left(newSheetName, 30)

        ActiveSheet.Name = newSheetName 'Change the new sheet name


        Worksheets("DATABASE").Activate 'Activate database sheet

        'Get last col of form pointer row
        Dim endCol As Long
        endCol = Cells(cellZahyoRow, Columns.Count).End(xlToLeft).Column
        
        For j = 1 To endCol
        
            'On Error GoTo ERR_LABEL
            If Cells(cellZahyoRow, j) Like "[A-Z]#" Then
                Sheets(newSheetName).Range(Cells(cellZahyoRow, j)).Value = Sheets("DATABASE").Cells(i, j).Value
          
            Else
            
            End If
            
        Next
    
    Next


'FOR DEBUG
'ERR_LABEL:
    '// 2  Error number & Error discrittiopn
    'Debug.Print Err.Number & " "; Err.Description
    '// 3. Skip Error
    'On Error GoTo 0
    '// 4. Print 0
    'Debug.Print Err.Number & " "; Err.Description
    '// 5. Resume next
    'Resume Next

    
End Sub
