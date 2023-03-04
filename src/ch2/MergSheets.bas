//MergSheets.bas
//
Sub MergeSheets()

    Dim folderPath As String
    Dim selectedSheetName As String
    Dim newWorkbook As Workbook
    Dim sheetCounter As Integer
    
    ' フォルダのパスを指定
    folderPath = "C:\Users\user\Documents\Excel Files\"
    
    ' コピーするシートの名前を指定
    selectedSheetName = "Sheet1"
    
    ' 新しいブックを作成
    Set newWorkbook = Workbooks.Add
    
    ' シートのカウンターを初期化
    sheetCounter = 0
    
    ' フォルダ内のすべてのExcelブックに対して繰り返し処理
    Dim filename As String
    filename = Dir(folderPath & "*.xls*")
    Do While filename <> ""
        ' Excelブックを開く
        Dim workbookPath As String
        workbookPath = folderPath & filename
        Dim sourceWorkbook As Workbook
        Set sourceWorkbook = Workbooks.Open(workbookPath)
        
        ' コピーするシートを取得
        Dim sourceSheet As Worksheet
        On Error Resume Next
        Set sourceSheet = sourceWorkbook.Worksheets(selectedSheetName)
        On Error GoTo 0
        
        If Not sourceSheet Is Nothing Then
            ' コピーするシートが存在する場合、新しいブックに追加
            sheetCounter = sheetCounter + 1
            sourceSheet.Copy after:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
            newWorkbook.Sheets(sheetCounter).Name = sourceWorkbook.Name & " - " & sourceSheet.Name
        End If
        
        ' Excelブックを閉じる
        sourceWorkbook.Close SaveChanges:=False
        
        ' 次のExcelブックを処理するためにファイル名を取得
        filename = Dir()
    Loop
    
    ' 最初のシートを削除
    Application.DisplayAlerts = False
    newWorkbook.Sheets(1).Delete
    Application.DisplayAlerts = True
    
End Sub
