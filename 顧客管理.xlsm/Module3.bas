Attribute VB_Name = "Module3"
Option Explicit
Sub makeAscendingSheet()
' 整列作業シート生成 Macro

    Dim ws As Worksheet

    '顧客昇順シートがあれば削除する（シート作成時エラー回避のため）
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name = "顧客昇順" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
        
    '顧客昇順シートを作成（シート追加して、入力シートから丸々コピペする）
    Dim inputSheet As Worksheet
    Set inputSheet = Sheets("入力シート")
    Dim ascendingSheet As Worksheet
    Set ascendingSheet = Worksheets.Add ' 新しいシートを追加して変数に代入
    ascendingSheet.Name = "顧客昇順" ' シート名を設定
    
    '最終行、最終列を特定
    Dim inputDateLastRow As Long
    inputDateLastRow = inputSheet.Cells(Rows.Count, 3).End(xlUp).Row '入力シートの最終行を定義 '入力シート、お客様名列の最終行を取得
    Dim inputSheetLastColumn As Long
    inputSheetLastColumn = inputSheet.Cells(1, Columns.Count).End(xlToLeft).Column '入力シート、1行目ヘッダの最終列を取得
    
    '入力シートの内容をコピー
    inputSheet.Activate
    inputSheet.Range(Cells(1, 1), Cells(inputDateLastRow, inputSheetLastColumn)).Copy
    
    '顧客昇順シートへ値のみペースト
    ascendingSheet.Activate
    ascendingSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    '並び替え処理（顧客名で昇順）
    ascendingSheet.Sort.SortFields.Clear
    ascendingSheet.Sort.SortFields.Add2 _
        Key:=Range("H1"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    ascendingSheet.Sort.SortFields.Add2 _
        Key:=Range("I1"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ascendingSheet.Sort
        .SetRange Range("A1", Cells(inputDateLastRow, inputSheetLastColumn))
        .Header = xlYes
        .Apply
    End With
    
End Sub
Function inputSheetFalseCheck()
''''''''''''''''''''''''''''''''''''''''
''このFunctionは顧客管理ブックの入力シート専用のプロシージャです
''Falseがあるかないかをcheckします
''Callする前に入力シートのアクティベートが必須です
''例:inputSheet.Activate
''''''''''''''''''''''''''''''''''''''''
    Dim inputSheet As Worksheet
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")
    Dim booleanArr As Variant
    Dim falseRowList As Variant
    Dim falseRowStr As String
    ReDim falseRowList(0)
    Dim i As Long
    Dim j As Long
    
    inputSheet.Activate
    
    'シートの保護解除
    ActiveSheet.Unprotect Password:="042595"
    
    booleanArr = Range("X2:AA9999").Value
    
    For i = LBound(booleanArr, 1) To UBound(booleanArr, 1)
        For j = LBound(booleanArr, 2) To UBound(booleanArr, 2)
            If booleanArr(i, j) = "FALSE" Then
                falseRowList(UBound(falseRowList)) = i + 1
                ReDim Preserve falseRowList(UBound(falseRowList) + 1)
                Exit For
            End If
        Next
    Next
    
    'シートの保護
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
    '文字列化
    If UBound(falseRowList) = 1 Then
        falseRowStr = falseRowList(0)
    Else
        For i = LBound(falseRowList) To UBound(falseRowList) - 1
            If i <> UBound(falseRowList) - 1 Then
                falseRowStr = falseRowStr & falseRowList(i) & ","
            Else
                falseRowStr = falseRowStr & falseRowList(i)
            End If
        Next
    End If
    
    If UBound(falseRowList) > 0 Then
    
        MsgBox "誤入力があるかもしれません。" & vbCrLf & _
        falseRowStr & "行目を確認してください。" & vbCrLf & _
        "この問題が解決するまで締め処理や顧客情報の更新は行いません。" & vbCrLf & _
        "解決できない場合はシステム管理者にお問い合わせください。", vbExclamation
        
        '上書き保存(セーブ)処理
        On Error Resume Next
        ActiveWorkbook.Save
        If Err.Number > 0 Then
            MsgBox "保存されませんでした"
        End If
        inputSheetFalseCheck = 1
    Else
        inputSheetFalseCheck = 0
    End If
    
End Function

