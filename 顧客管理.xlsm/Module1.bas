Attribute VB_Name = "Module1"
Option Explicit
Sub updateCustomerInfomationButton()
    
    '最初に入力シートの整合性チェックを行う
    If inputSheetFalseCheck() = 1 Then 'inputSheetFalseCheckはModule3
        Exit Sub
    End If
    
    Call makeAscendingSheet 'Module3
    Call updateCustomerInfomation
    
End Sub
Sub updateCustomerInfomation()
'
' 内容の更新 Macro
'

'・・・変数定義・・・
    Dim inputSheet As Worksheet '入力シートを定義
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")
    Dim ascendSheet As Worksheet '顧客昇順シートを定義
    Set ascendSheet = ActiveWorkbook.Worksheets("顧客昇順")
    Dim customerSheet As Worksheet '会員一覧シートを定義
    Set customerSheet = ActiveWorkbook.Worksheets("会員一覧")
    
    Dim i As Long 'Forのカウンタ'
    Dim j As Integer '列インデックス用カウンタ'
    Dim k As Integer '回数分の行調整カウンタ　利用回数分'
    Dim bastInfoArr '会員一覧の出力用_配列 ヘッダ=回数,会員名,電話番号,NG情報,備考'
    Dim historyInfoArr '会員の履歴の出力用_配列 ヘッダ=1回,2回,...'
    Dim ascendCustomerDataArr '顧客昇順表の格納用_配列
    
    ascendSheet.Activate
    
    Dim customerLastRow As Long
    customerLastRow = ascendSheet.Cells(Rows.Count, 3).End(xlUp).Row '顧客昇順シートの最終行を定義
    Dim customerLastColumn As Long
    customerLastColumn = Cells(1, Columns.Count).End(xlToLeft).Column '顧客昇順シートの最終列を定義、1行目ヘッダから
    ascendCustomerDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '顧客昇順シートの内容を全取得'
    
    Dim CountRow '新規数をカウントしてbastInfoArr,historyInfoArrの行数を再定義する用の変数
    CountRow = WorksheetFunction.CountIf(Range(Cells(2, 2), Cells(customerLastRow, 2)), "1")
    Dim CountColumn '顧客中、最高の回数を割り出してhistoryInfoArrの列数を再定義する用の変数
    CountColumn = WorksheetFunction.Max(Range(Cells(2, 2), Cells(customerLastRow, 2)))

'・・・処理・・・
    ReDim bastInfoArr(1 To CountRow, 1 To 5)
    ReDim historyInfoArr(1 To CountRow, 1 To CountColumn)
    
    k = 0
    
    '会員一覧へ出力用の配列整形
    For i = LBound(ascendCustomerDataArr, 1) To UBound(ascendCustomerDataArr, 1)  'ascendCustomerDataArrの行数分ループする
        j = 1 '利用回数カウンタ
        
'        '速度計測用テストコード
'        Dim testLong As Long
'        testLong = speedtest(Val(ascendCustomerDataArr(i, 3)))
'        Debug.Print i & testLong
'        '''''▲End Test
        
        If ascendCustomerDataArr(i, 2) = j Then
            '顧客一覧格納
            bastInfoArr(i - k, 1) = WorksheetFunction.CountIf(Range(Cells(2, 9), Cells(customerLastRow, 9)), ascendCustomerDataArr(i, 9))
            bastInfoArr(i - k, 2) = ascendCustomerDataArr(i, 8) '客名格納
            bastInfoArr(i - k, 3) = ascendCustomerDataArr(i, 9) '客番格納
            If ascendCustomerDataArr(i, 11) <> "" And ascendCustomerDataArr(i, 11) <> 0 Then
                bastInfoArr(i - k, 5) = ascendCustomerDataArr(i, 2) & "," & ascendCustomerDataArr(i, 11) '初回の客備考格納
            End If
            If ascendCustomerDataArr(i, 10) <> "" And ascendCustomerDataArr(i, 10) <> 0 Then
                bastInfoArr(i - k, 4) = ascendCustomerDataArr(i, 10) 'NG情報格納
            End If
            '利用履歴格納 内容=日付3,女子名7,ホテル12,コス13,時間14
            historyInfoArr(i - k, j) = ascendCustomerDataArr(i, 3) & "," & ascendCustomerDataArr(i, 7) & vbLf & ascendCustomerDataArr(i, 12) & "," & ascendCustomerDataArr(i, 13) & "," & ascendCustomerDataArr(i, 14) 'ここでは視認性のために日付は敢えて年yyを省く
        ElseIf ascendCustomerDataArr(i, 2) > j Then '利用2回目以降の格納処理
            k = k + 1
            '2回目以降、備考とNG情報格納
            If ascendCustomerDataArr(i, 11) <> "" And ascendCustomerDataArr(i, 11) <> 0 Then
                bastInfoArr(i - k, 5) = bastInfoArr(i - k, 5) & vbLf & ascendCustomerDataArr(i, 2) & "," & ascendCustomerDataArr(i, 11) '2回目以降の客備考格納
            End If
            If ascendCustomerDataArr(i, 10) <> "" And ascendCustomerDataArr(i, 10) <> 0 Then
                Select Case bastInfoArr(i - k, 4) '改行
                    Case ""
                        bastInfoArr(i - k, 4) = ascendCustomerDataArr(i, 10) 'NG情報格納
                    Case Is <> ""
                        bastInfoArr(i - k, 4) = bastInfoArr(i - k, 4) & vbLf & ascendCustomerDataArr(i, 10) '改行つきNG情報格納
                End Select
            End If
            j = ascendCustomerDataArr(i, 2) '利用回数分、利用履歴の入力列を右にずらす
            '利用履歴格納 内容=日付3,女子名7,ホテル12,コス13,時間14
            historyInfoArr(i - k, j) = ascendCustomerDataArr(i, 3) & "," & ascendCustomerDataArr(i, 7) & vbLf & ascendCustomerDataArr(i, 12) & "," & ascendCustomerDataArr(i, 13) & "," & ascendCustomerDataArr(i, 14) 'ここでは視認性のために日付は敢えて年yyを省く
        Else
            Exit For
        End If
    Next i

    customerSheet.Activate
    
    'フィルターをクリア
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    'セル範囲をbastInfoArr,historyInfoArrの大きさで選択して書き込み
    Range("A3").Resize(UBound(bastInfoArr, 1), UBound(bastInfoArr, 2)) = bastInfoArr
    Range("F3").Resize(UBound(historyInfoArr, 1), UBound(historyInfoArr, 2)) = historyInfoArr
    
    '行の高さ、列の幅を調整
    Range(Cells(1, 1), Cells(UBound(bastInfoArr, 1), UBound(bastInfoArr, 2) - 1)).EntireColumn.AutoFit '顧客備考と利用履歴を除く会員一覧の列を自動調整する
    
    '顧客昇順を削除
    Application.DisplayAlerts = False
        ascendSheet.Delete
    Application.DisplayAlerts = True
    
    '上書き保存(セーブ)処理
    On Error Resume Next
    ActiveWorkbook.Save
    If Err.Number > 0 Then
        MsgBox "保存されませんでした"
    End If

End Sub
