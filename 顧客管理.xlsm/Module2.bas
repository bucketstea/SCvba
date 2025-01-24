Attribute VB_Name = "Module2"
Option Explicit
Sub updateCustomerStatisticsButton()
    Call makeAscendingSheet 'Module3
    Call updateCustomerStatistics
End Sub
Sub updateCustomerStatistics()
'・・・変数定義・・・
    Dim ascendSheet As Worksheet '顧客昇順シートを定義
    Set ascendSheet = ActiveWorkbook.Worksheets("顧客昇順")
    Dim customerStatisticsSheet As Worksheet '会員一覧シートを定義
    Set customerStatisticsSheet = ActiveWorkbook.Worksheets("会員別_統計情報")
    
    Dim i As Long '昇順データの行数ループForのカウンタ'
    Dim k As Integer '回数分の行調整カウンタ　利用回数分、2回目以降の利用回数分、昇順全データと顧客毎データに行数がずれるため必要
    Dim baseInfoArr '会員一覧の出力用_配列 'ヘッダ=回数,初来店日,最終来店日,会員名,電話番号,媒体'
    Dim statisticsInfoArr '会員の統計情報格納用_配列 ヘッダ=媒体,累計売上,累計店落,単価,頻度(日/回),離反日数,アンケ率,本指率,体入率'
    Dim ascendCustomerDataArr '顧客昇順表の格納用_配列
    
    ascendSheet.Activate
    
    Dim customerLastRow As Long
    customerLastRow = Range("C1").CurrentRegion(Range("C1").CurrentRegion.Count).Row '顧客昇順シートの最終行を定義
    Dim customerLastColumn As Long
    customerLastColumn = (Cells(1, Columns.Count).End(xlToLeft).Column) '顧客昇順シートの最終列を定義、1行目ヘッダから
    ascendCustomerDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '顧客昇順シートの内容を全取得
    
    Dim CountRow '新規数をカウントしてbaseInfoArr,statisticsInfoArrの行数を再定義する用の変数
    CountRow = WorksheetFunction.CountIf(Range(Cells(2, 2), Cells(customerLastRow, 2)), "1")

'・・・処理・・・
    Debug.Print ("---------------Prc")
    
    ReDim baseInfoArr(1 To CountRow, 1 To 6) 'ヘッダ=回数,初来店日,最終来店日,会員名,電話番号,媒体'
    ReDim statisticsInfoArr(1 To CountRow, 1 To 8) 'ヘッダ=累計売上,累計店落,単価(店落/回),頻度(日/回),離反日数,ｱﾝｹ率(ｱﾝｹ数/回),本指率(本指/回),体入率(本指/回)
'
    Dim customerCountTotal As Long '会員のご利用回数
    Dim qreCt As Long '
    Dim repeatCt As Long
    Dim newCt As Long
    Dim dtToday As Date
    dtToday = Date
    
    '会員一覧へ出力用の配列整形
    k = 0
    For i = LBound(ascendCustomerDataArr, 1) To UBound(ascendCustomerDataArr, 1)  'ascendCustomerDataArrの行数分ループする
        customerCountTotal = WorksheetFunction.CountIf(Range(Cells(2, 9), Cells(customerLastRow, 9)), ascendCustomerDataArr(i, 9))
        
        '初回利用かつ最終利用ではない時の格納処理
        If ascendCustomerDataArr(i, 2) = 1 And ascendCustomerDataArr(i, 2) <> customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            ''''''''''2回目以降の利用回数分を加算するため初回利用時のここでは省略
'            k = k + 1  '回数分の行調整

            '顧客一覧格納
            baseInfoArr(i - k, 1) = customerCountTotal '回数格納
            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '初回利用日格納
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '最終利用日格納
            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '客名格納
            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '客番格納
            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '媒体名格納
            
            '特定条件利用回数等計算
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "本指" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "体入" Then
                newCt = newCt + 1
            End If
            
            '統計情報格納
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '累計売上
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '累計店落
            ''''''''''単価以降は最終利用時に計算するためここでは省略
'            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '単価
'            statisticsInfoArr(i - k, 4) = "once" '頻度
'            statisticsInfoArr(i - k, 5) = "once" '離反日数
'            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal 'アンケ率
'            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '本指率
'            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '体入率

'            '回数等初期化
            ''''''''''特定条件利用回数は最終利用時にリセットするためここでは省略
'            qreCt = 0
'            repeatCt = 0
'            newCt = 0

        '初回利用かつ最終利用時の計算/格納処理
        ElseIf ascendCustomerDataArr(i, 2) = 1 And ascendCustomerDataArr(i, 2) = customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            ''''''''''2回目以降の利用回数分を加算するため初回利用時のここでは省略
'            k = k + 1  '回数分の行調整
                        
            '顧客一覧格納
            baseInfoArr(i - k, 1) = customerCountTotal '回数格納
            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '初回利用日格納
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '最終利用日格納
            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '客名格納
            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '客番格納
            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '媒体名格納
            
            '特定条件利用回数等計算
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "本指" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "体入" Then
                newCt = newCt + 1
            End If
            
            '統計情報格納
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '累計売上
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '累計店落
            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '単価
            statisticsInfoArr(i - k, 4) = "once" '頻度
            statisticsInfoArr(i - k, 5) = "once" '離反日数
            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal 'アンケ率
            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '本指率
            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '体入率
            
            '回数等初期化
            qreCt = 0
            repeatCt = 0
            newCt = 0

        '利用2回目以降かつ最終利用でない時の計算/格納処理
        ElseIf ascendCustomerDataArr(i, 2) > 1 And ascendCustomerDataArr(i, 2) <> customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))

            k = k + 1 '回数分の行調整

'            '顧客一覧格納
            ''''''''顧客一覧データは最終利用日を除き初回利用時に格納するためここでは省略
'            baseInfoArr(i - k, 1) = customerCountTotal '回数格納
'            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '初回利用日格納
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '最終利用日格納
'            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '客名格納
'            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '客番格納
'            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '媒体名格納

            '特定条件利用回数等計算
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "本指" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "体入" Then
                newCt = newCt + 1
            End If
            
            '統計情報格納
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '累計売上
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '累計店落
            ''''''''''単価以降は最終利用時に計算するためここでは省略
'            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '単価
'            statisticsInfoArr(i - k, 4) = (DateDiff("d", baseInfoArr(i - k, 2), dtToday) + 1) / baseInfoArr(i - k, 1) '頻度
'            statisticsInfoArr(i - k, 5) = (DateDiff("d", baseInfoArr(i - k, 3), dtToday)) '離反日数
'            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal 'アンケ率
'            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '本指率
'            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '体入率

'            '回数等初期化
            ''''''''''特定条件利用回数は最終利用時にリセットするためここでは省略
'            qreCt = 0
'            repeatCt = 0
'            newCt = 0

        '2回目以降かつ最終利用時の計算/格納処理
        ElseIf ascendCustomerDataArr(i, 2) > 1 And ascendCustomerDataArr(i, 2) = customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            k = k + 1  '回数分の行調整
            
'            '顧客一覧格納
            ''''''''顧客一覧データは最終利用日を除き初回利用時に格納するためここでは省略
'            baseInfoArr(i - k, 1) = customerCountTotal '回数格納
'            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '初回利用日格納
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '最終利用日格納
'            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '客名格納
'            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '客番格納
'            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '媒体名格納

            '特定条件利用回数等計算
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "本指" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "体入" Then
                newCt = newCt + 1
            End If
            
            '統計情報格納
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '累計売上
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '累計店落
            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '単価
            statisticsInfoArr(i - k, 4) = (DateDiff("d", baseInfoArr(i - k, 2), dtToday) + 1) / baseInfoArr(i - k, 1) '頻度
            statisticsInfoArr(i - k, 5) = (DateDiff("d", baseInfoArr(i - k, 3), dtToday)) '離反日数
            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal 'アンケ率
            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '本指率
            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '体入率
            
            '回数等初期化
            qreCt = 0
            repeatCt = 0
            newCt = 0
        Else
            Debug.Print ("Exit For")
            Exit For
        End If
    Next i

    customerStatisticsSheet.Activate
    
    'セル範囲をbaseInfoArr,statisticsInfoArrの大きさで選択して書き込み
    Range("A3").Resize(UBound(baseInfoArr, 1), UBound(baseInfoArr, 2)) = baseInfoArr
    Range("G3").Resize(UBound(statisticsInfoArr, 1), UBound(statisticsInfoArr, 2)) = statisticsInfoArr
    
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

