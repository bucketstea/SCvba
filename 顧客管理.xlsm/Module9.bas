Attribute VB_Name = "Module9"
Option Explicit
'計算後配列の座標定義（Excel上の座標とは異なる。ExcelのF列を0列目とし、書き込み時にF3を開始位置とする）
''各合計
Const edCountClm = 2 '本数計
Const edRenominationClm = 3 '本指名数
Const edSalesClm = 4 '売上計
Const edPayClm = 5 '女子給計
Const edIncomeClm = 6 '店落計
Const edNewClm = 7 '新規計
Const edRepeaterClm = 8 '会員計
Const ed2ndCtClm = 10 '2回目計
Const edSBvalueClm = 13 'SB額計
''媒体数
Const edZenraClm = 14
Const edHeavenClm = 15
Const edKyokuClm = 16
Const edJapanClm = 17
Const edDXClm = 18
Const edEkiClm = 19
Const edPureClm = 20
Const edHimechClm = 21
Const edGoogleClm = 22
Const edHPClm = 23
Const edOtherClm = 24
Const edBillClm = 25
Const edT1Clm = 26
''割合、平均数
Const edAvgCountClm = 0 '数/日
Const edAvgIncomeClm = 1 '落/日
Const edRePercentClm = 9
Const ed2ndPercentClm = 11
Const edIncomePerCtClm = 12

'Cast出勤数の座標定義
Const edYearClm = 1
Const edMonthClm = 2
Const edMonthlyLastDateClm = 3
Const edCountPerCastClm = 4
Const edCastCountClm = 5

'オープンの年月（不変の値）
Const startYY = 21 '営業開始年
Const startMM = 10 '営業開始月-1
Sub monthlyStatisticsButton()
    Call updateMonthlyStatistics
End Sub
Sub updateMonthlyStatistics()
'・・・変数定義・・・
    Dim inputSheet As Worksheet '入力シートを定義
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")
    Dim monthlyStatisticsSheet As Worksheet '月別シートを定義
    Set monthlyStatisticsSheet = ActiveWorkbook.Worksheets("月別_統計情報")
    
    
    '''''''''''''''''''''''''''''
    '各シートの最終行までの内容を取得する。
    '''''''''''''''''''''''''''''
    '入力シートの配列定義
    inputSheet.Activate
    Dim inputDateLastRow As Long
    inputDateLastRow = Cells(Rows.Count, 3).End(xlUp).Row '入力シート、お客様名列の最終行を取得
    Dim inputSheetLastColumn As Long
    inputSheetLastColumn = Cells(1, Columns.Count).End(xlToLeft).Column '入力シート、1行目ヘッダの最終列を取得
    Dim inputArray As Variant
    inputArray = Range(Cells(2, 1), Cells(inputDateLastRow, inputSheetLastColumn)).Value '入力シートの内容を配列として取得

    '月別統計データの配列の大きさを定義する(Cast出勤数が絡む列を除く)
    monthlyStatisticsSheet.Activate
    Dim calculatedArray As Variant '書き出し用の配列
    Dim lastDate As Long '最終受付日
    lastDate = inputArray(inputDateLastRow - 1, 3)
    Dim lastYY As Integer '入力シートの対象レコードの年
    lastYY = Val(Left(lastDate, 2))
    Dim lastMM As Integer '入力シートの対象レコードの月
    lastMM = Val(Right(Left(lastDate, 4), 2))
    Dim calculatedArrayLastRow As Long
    calculatedArrayLastRow = ((lastYY - startYY) * 12) + (lastMM - startMM) - 1
    Dim calculatedArrayLastColumn As Long
    calculatedArrayLastColumn = (Cells(2, Columns.Count).End(xlToLeft).Column) - 6 'エクセル上の最終列から年、月、最終日、本出レシオ、総出の5列分引き、0始まりのため1引く、計6引く
    ReDim calculatedArray(calculatedArrayLastRow, calculatedArrayLastColumn)
    
    'Cast出勤数が絡む列を格納する配列を格納する
    Dim castCountArr As Variant
    castCountArr = Range(Cells(3, 1), Cells(calculatedArrayLastRow + 3, 5)) '開始行が3行目からなので開始と最終は+3
    
    '基本データ作成（本数計算、総売上計算、女子給計算、店落計算、新規数計算、会員数計算、各媒体本数、2回目数）
    calculatedArray = monthlyFoundationData(inputArray, calculatedArray)
    
    '計算データ作成（割合、平均等）
    calculatedArray = monthlyRatioAverage(calculatedArray, lastDate) 'lastDateは最新月の平均を割り出すために必要
    
    'cast出勤数平均取得、本数/出勤数の計算処理
    castCountArr = monthlyCastCount(castCountArr)
    
    '書き込み処理
    monthlyStatisticsSheet.Activate
    
    'シートの保護解除
    ActiveSheet.Unprotect Password:="042595"
    
    Range(Cells(3, 6), Cells(calculatedArrayLastRow + 3, calculatedArrayLastColumn + 4)) = calculatedArray '月別統計データの配列(Cast出勤数が絡む列を除く)を書き込み
    Range(Cells(3, 1), Cells(calculatedArrayLastRow + 3, 5)) = castCountArr 'Cast出勤数が絡む列を書き込み
    
    'シートの保護
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
    monthlyStatisticsSheet.Activate

End Sub
Function monthlyFoundationData(ByVal inputArray As Variant, ByVal calculatedArray As Variant)
    Dim i As Long
    Dim j As Long
    
    Dim targetYY As Integer '入力シートの対象レコードの年
    Dim targetMM As Integer '入力シートの対象レコードの月
    Dim monthRow As Long
        
    '本数計算、総売上計算、女子給計算、店落計算、新規数計算、会員数計算、各媒体本数、2回目数
    For i = LBound(inputArray, 1) To UBound(inputArray, 1)
        targetYY = Left(inputArray(i, 3), 2)
        targetMM = Right(Left(inputArray(i, 3), 4), 2)
        monthRow = ((targetYY - startYY) * 12) + (targetMM - startMM) - 1 'calculatedArrayは0始まりの配列であるため、Row -1する。
        
        '本数 同一の対象行が続くと+1加算されていく
        calculatedArray(monthRow, edCountClm) = calculatedArray(monthRow, edCountClm) + 1
        '総売上
        calculatedArray(monthRow, edSalesClm) = calculatedArray(monthRow, edSalesClm) + inputArray(i, 18)
        '女子給
        calculatedArray(monthRow, edPayClm) = calculatedArray(monthRow, edPayClm) + inputArray(i, 19)
        '店落
        calculatedArray(monthRow, edIncomeClm) = calculatedArray(monthRow, edIncomeClm) + inputArray(i, 20)
        
        If inputArray(i, 6) = "本指" Then
        '本指名数
            calculatedArray(monthRow, edRenominationClm) = calculatedArray(monthRow, edRenominationClm) + 1
        End If
        If inputArray(i, 5) = "R" Then
        '会員数
            calculatedArray(monthRow, edRepeaterClm) = calculatedArray(monthRow, edRepeaterClm) + 1
        Else
        '新規数
            calculatedArray(monthRow, edNewClm) = calculatedArray(monthRow, edNewClm) + 1
            '各媒体数
            Select Case inputArray(i, 5)
            Case "隣"
                calculatedArray(monthRow, edZenraClm) = calculatedArray(monthRow, edZenraClm) + 1
            Case "ヘブン"
                calculatedArray(monthRow, edHeavenClm) = calculatedArray(monthRow, edHeavenClm) + 1
            Case "情報局"
                calculatedArray(monthRow, edKyokuClm) = calculatedArray(monthRow, edKyokuClm) + 1
            Case "風俗ジャパン"
                calculatedArray(monthRow, edJapanClm) = calculatedArray(monthRow, edJapanClm) + 1
            Case "DX"
                calculatedArray(monthRow, edDXClm) = calculatedArray(monthRow, edDXClm) + 1
            Case "駅ちか"
                calculatedArray(monthRow, edEkiClm) = calculatedArray(monthRow, edEkiClm) + 1
            Case "ぴゅあらば"
                calculatedArray(monthRow, edPureClm) = calculatedArray(monthRow, edPureClm) + 1
            Case "ヒメチャン"
                calculatedArray(monthRow, edHimechClm) = calculatedArray(monthRow, edHimechClm) + 1
            Case "グーグル"
                calculatedArray(monthRow, edGoogleClm) = calculatedArray(monthRow, edGoogleClm) + 1
            Case "HP"
                calculatedArray(monthRow, edHPClm) = calculatedArray(monthRow, edHPClm) + 1
            Case "その他"
                calculatedArray(monthRow, edOtherClm) = calculatedArray(monthRow, edOtherClm) + 1
            Case "ビル"
                calculatedArray(monthRow, edBillClm) = calculatedArray(monthRow, edBillClm) + 1
            Case "T-1"
                calculatedArray(monthRow, edT1Clm) = calculatedArray(monthRow, edT1Clm) + 1
            End Select
        End If
        '2回目本数
        If inputArray(i, 2) = 2 Then
            calculatedArray(monthRow, ed2ndCtClm) = calculatedArray(monthRow, ed2ndCtClm) + 1
        End If
        
        'SB計
        calculatedArray(monthRow, edSBvalueClm) = calculatedArray(monthRow, edSBvalueClm) + (inputArray(i, 19) * (inputArray(i, 22) / 100))
        
    Next i
    
    monthlyFoundationData = calculatedArray
    
End Function
Function monthlyRatioAverage(ByVal calculatedArray As Variant, ByVal lastDate As String)

    Dim i As Long
    Dim j As Long
    
    Dim monthlyLastDate As Long
    
    '割合計算、平均計算
    For i = LBound(calculatedArray, 1) To UBound(calculatedArray, 1)
        'ゼロパディング
        For j = LBound(calculatedArray, 2) To UBound(calculatedArray, 2)
            If calculatedArray(i, j) = "" Then
                calculatedArray(i, j) = 0
            End If
        Next j
        monthlyLastDate = Cells(i + 3, 3).Value
        '''アベレージ本数、アベレージ店落を計算
        '当月判定
        If i = UBound(calculatedArray, 1) Then
        '▼当日の日付を取得して除算する (問題:日変わり直後の精度が悪い)
'            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / Format(Date, "dd")
'            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / Format(Date, "dd")
        '▼最終入力日の日付で除算する (問題:入力0=本数0の日の直後の精度が悪い？発生頻度が低いので、暫定的にこちらを採用)
            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / Right(lastDate, 2)
            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / Right(lastDate, 2)

        '過去月判定
        Else
            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / monthlyLastDate
            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / monthlyLastDate
        End If
        '''会員率
        calculatedArray(i, edRePercentClm) = calculatedArray(i, edRepeaterClm) / (calculatedArray(i, edRepeaterClm) + calculatedArray(i, edNewClm))
        '''2回目率
        calculatedArray(i, ed2ndPercentClm) = calculatedArray(i, ed2ndCtClm) / calculatedArray(i, edNewClm)
        '''単価計算
        calculatedArray(i, edIncomePerCtClm) = calculatedArray(i, edIncomeClm) / calculatedArray(i, edCountClm)
    Next i
    
    monthlyRatioAverage = calculatedArray
    
End Function
Function monthlyCastCount(ByVal castCountArr As Variant)
    Dim i As Long
    Dim targetyyyymm As String '目的の月のファイル名を指定するための文字列を格納する
    Dim targetyyyymmArr As Variant
    Dim managementBook As Workbook '管理表ブックオブジェクト
    Dim managementBookPath As Variant '管理表ブックのパス
    Dim managementSheet1 As Worksheet
    
    '各行の取得状況を判定して、未取得なら取得する。最新行なら新たに取得する。
    For i = LBound(castCountArr, 1) To UBound(castCountArr, 1)
        '取得できていない月を判定する
        If castCountArr(i, 5) = Empty Or i = UBound(castCountArr, 1) Then
            targetyyyymm = "20" & Format(castCountArr(i, edYearClm), "00") & Format(castCountArr(i, edMonthClm), "00")
            
            'その月の管理表ブックを開いて､キャスト出勤数を取得する
            '本番環境用
'           managementBookPath = "E:\管理表\管理表" & targetyyyymm & ".xlsx"
            'テスト環境用
            managementBookPath = "D:\usb_20241230\管理表\管理表" & targetyyyymm & ".xlsx"
            If Dir(managementBookPath) <> "" Then
                Set managementBook = Workbooks.Open(managementBookPath)
                Set managementSheet1 = managementBook.Worksheets("Z収")
                managementSheet1.Activate
                castCountArr(i, edCastCountClm) = managementSheet1.Range("AD36")
                castCountArr(i, edCountPerCastClm) = "=F" & i + 2 & "/E" & i + 2
                managementBook.Close SaveChanges:=False
            Else
                castCountArr(i, 5) = "Err"
            End If
        End If
    Next
    
    monthlyCastCount = castCountArr
End Function
