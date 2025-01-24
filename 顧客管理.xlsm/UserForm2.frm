VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "女子給表示"
   ClientHeight    =   9090.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17730
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'表示用文字列
Public dateJpDisp As String
Public castDisp As String
Public contentsDisp As String
Public IncomeAndQre As String

'入力日付
Public ansDateIs As Integer

'総売上、店落ち、アンケ計算
Public totalSales As Long '1日の総売上額
Public totalPay As Long '1日の総女子給額
Public totalIncome As Long '1日の店落総額
Public totalQre As Long '1日のアンケ総額

Sub UserForm_Initialize()
    
    castBackOnDay
    
    With Worksheets("入力シート")
        Label5.Caption = dateJpDisp
        Label6.Caption = IncomeAndQre
        Label1.Caption = castDisp
        Label2.Caption = contentsDisp
    End With
    
End Sub
Sub castBackOnDay()
    Dim inputSheet As Worksheet
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")

'//////////////////////////////////
'//日付入力用ダイアログを表示する
'//////////////////////////////////
    Dim ansDate As String
    Do While Not (ansDate Like "######" Or ansDate Like "#")
        ansDate = InputBox("対象の日付を入力してください。" & vbCr & vbCr & _
        "当日分は[0]を入力してもOKです!!" & vbCr & vbCr & _
        "[0]→当日" & vbCr & _
        "[1]→昨日（1日前）" & vbCr & _
        "[2]→一昨日（2日前）" & vbCr & _
        "のように出力されます。（1～9を入力時）", "日付入力", "(例:2090年3月17日なら[900317]と入力)")
        
        ansDateIs = 1
        If StrPtr(ansDate) = 0 Then
            ansDateIs = 0
            Exit Sub
        End If
    Loop
    
'///////////////////////////////////
'//該当する日付の行を特定して取得
'///////////////////////////////////
    'シートの保護解除
    ActiveSheet.Unprotect Password:="042595"
    
    'シートの最終行を定義、取得範囲
    Dim dateColumnArr() As Variant
    Dim dateLastRow As Long
    dateLastRow = Range("C1").CurrentRegion(Range("C1").CurrentRegion.Count).Row
    dateColumnArr = inputSheet.Range(Cells(1, 3), Cells(dateLastRow, 3))
    
    Dim objDate As Variant
    
    '日付入力が"#"だった場合に当日日付に変換する
    If ansDate Like "#" Then
        objDate = Val(Right(Left(Date - ansDate, 4), 2) & Right(Left(Date - ansDate, 7), 2) & Right(Date - ansDate, 2))
    Else
        objDate = Val(ansDate)
    End If
    
    '指定した日が何行目に存在するかを特定
    Dim dateRows() As Variant
    ReDim dateRows(1)
    
    Dim i As Long
    Dim j As Long
    
    i = 0
    
    For i = LBound(dateColumnArr, 1) To UBound(dateColumnArr, 1)
        If dateColumnArr(i, 1) = objDate Then
            dateRows(UBound(dateRows)) = i
            ReDim Preserve dateRows(UBound(dateRows) + 1)
        End If
    Next

    '当日レコードがなければ終了
    If dateRows(1) = "" Then
        MsgBox ("その日は0本みたいです。" & vbCrLf & vbCrLf & _
        "・入力シートに本日の受付履歴を入力していないかも…？" & vbCrLf & _
        "・受付履歴、または今ここに入力した日付が間違っているかも…？" & vbCrLf & vbCrLf & _
        "by_Usui")
        Exit Sub
    End If

    Dim objRowL As Long
    Dim objRowU As Long
    
    objRowL = dateRows(LBound(dateRows) + 1)
    objRowU = dateRows(UBound(dateRows) - 1)
    
    '特定の日の受付データを配列として取得する
    Dim targetArr() As Variant '受付票の配列の格納先
    targetArr = inputSheet.Range(Cells(objRowL, 3), Cells(objRowU, 22)).Value '特定の座標を配列にする

    '女子リスト作成(重複あり)
    Dim tempCastList As Variant
    ReDim tempCastList(1)
    
    i = 0
    For i = LBound(targetArr, 1) To UBound(targetArr, 1)
        tempCastList(i) = targetArr(i, 5)
        ReDim Preserve tempCastList(UBound(tempCastList) + 1)
    Next
    ReDim Preserve tempCastList(UBound(tempCastList) - 1)
    
    '女子リスト重複削除処理
    Dim castList As Variant
    ReDim castList(1)
    
    castList = Call_Array_Dictionary(tempCastList)
    
    'cast毎に本数を格納するため女子リストを2次元配列化
    Dim castArr As Variant
    ReDim castArr(UBound(castList), 2)
    Dim workCt As Integer '本数

    i = 0
    j = 0
    For i = LBound(castList) To UBound(castList)
        workCt = 0
        castArr(i, 1) = castList(i)
        For j = LBound(targetArr, 1) To UBound(targetArr, 1)
            If targetArr(j, 5) = castList(i) Then
                workCt = workCt + 1
            End If
        Next
        castArr(i, 2) = workCt
    Next

    i = 0
    For i = LBound(targetArr, 1) To UBound(targetArr, 1)
        totalSales = totalSales + targetArr(i, 16)
        totalPay = totalPay + targetArr(i, 17)
        totalIncome = totalIncome + targetArr(i, 18)
        totalQre = totalQre + targetArr(i, 19)
    Next

    '女子毎の合計女子給の計算処理、何分入ったかを格納
    Dim castPayArr As Variant '女子毎の給与合計額
    ReDim castPayArr(UBound(castList))
    Dim castMinArr As Variant '分数(羅列)
    ReDim castMinArr(UBound(castList))
    
    i = 0
    j = 0
    For i = LBound(castList, 1) To UBound(castList, 1)
        workCt = 0
        For j = LBound(targetArr, 1) To UBound(targetArr, 1)
            If targetArr(j, 5) = castList(i) Then
                workCt = workCt + 1
                '給与合計額リストへの格納処理
                castPayArr(i) = castPayArr(i) + targetArr(j, 17)

                '分数(羅列)リストへの格納処理
                '最後の一本は文末のカンマ(,)不要
                '本指は"本指XX分"表示
                If workCt <> castArr(i, 2) Then
                    If targetArr(j, 4) = "本指" Then
                        castMinArr(i) = castMinArr(i) & targetArr(j, 4) & targetArr(j, 12) & "分, "
                    Else
                        castMinArr(i) = castMinArr(i) & targetArr(j, 12) & "分, "
                    End If
                Else
                    If targetArr(j, 4) = "本指" Then
                        castMinArr(i) = castMinArr(i) & targetArr(j, 4) & targetArr(j, 12) & "分"
                    Else
                        castMinArr(i) = castMinArr(i) & targetArr(j, 12) & "分"
                    End If
                End If
            End If
        Next
    Next

    '女子給表示用の文字列生成処理
    Dim slashSeparatedDate As Variant
    Dim dateInJp As String
    Dim castNameStr As String
    Dim perCastStr As String

    '日付文字列(Dateへ出力)
    slashSeparatedDate = "20" & Left(objDate, 2) & "/" & Right(Left(objDate, 4), 2) & "/" & Right(objDate, 2)
    dateInJp = "20" & Left(objDate, 2) & "年" & Right(Left(objDate, 4), 2) & "月" & Right(objDate, 2) & "日 (" & WeekdayName(Weekday(slashSeparatedDate), True) & ")"

    '女子名文字列(Nameへ出力)
    i = 0
    For i = LBound(castList, 1) + 1 To UBound(castList, 1)
        castNameStr = castNameStr & "・" & castList(i) & vbLf
    Next

    '内容文字列(Contentsへ出力)
    i = 0
    For i = LBound(castList, 1) + 1 To UBound(castList, 1)
        perCastStr = perCastStr & "…　" & castPayArr(i) & "円　(" & castMinArr(i) & ")" & vbLf
    Next

    'MsgBoxで表示(廃止)
'    MsgBox perCastStr

    'ユーザーフォームで表示する用
    'パブリック変数へ代入する
    dateJpDisp = dateInJp
    castDisp = castNameStr
    contentsDisp = perCastStr
    IncomeAndQre = " 総売上 " & totalSales & "円,　女子給総額 " & totalPay & "円,　店落ち計 " & totalIncome & "円,　" & "アンケ計 " & totalQre & "円"
    
    'シートの保護
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
End Sub

'■Dictionary(連想配列)で重複データを削除する
'▼必要な前準備▼
'ツール>参照設定>Microsoft scripting runtimeにチェック入れる
Public Function Call_Array_Dictionary(arr As Variant)
    Dim dic As New Dictionary
    Dim i As Long

    '■On Error Resume Nextでエラーを無視する
    On Error Resume Next
    For i = 0 To UBound(arr)
        dic.Add arr(i), arr(i)
    Next
    Call_Array_Dictionary = dic.Keys

End Function


Sub CommandButton1_Click()
    Unload UserForm2
End Sub

Sub CommandButton2_Click()
    Dim ansCastPay As Variant
    Dim ansMEx As Variant
    Dim ansDiaryBonus As Variant
    Dim ansCardIncome As Variant
    
    ansCastPay = InputBox("清算済みの女子給額を入力してください。", "清算済み女子給額", 0)
    If StrPtr(ansCastPay) = 0 Then
        Exit Sub
    End If
    
    ansMEx = InputBox("清算済みの雑費額を入力してください。", "清算済み雑費額", 0)
    If StrPtr(ansMEx) = 0 Then
        Exit Sub
    End If
    
    ansDiaryBonus = InputBox("清算済みの日記ボーナス額を入力してください。", "清算済み日記ボーナス額", 0)
    If StrPtr(ansDiaryBonus) = 0 Then
        Exit Sub
    End If

    ansCardIncome = InputBox("本日のカード売上額(カード手数料を除いた税込額)を入力してください。", "カード売上額", 0)
    If StrPtr(ansCardIncome) = 0 Then
        Exit Sub
    End If
    
    MsgBox ("現在の現金総額は、" & totalSales - totalQre - ansCastPay + ansMEx - ansDiaryBonus - ansCardIncome & "円です。" & vbCrLf & vbCrLf & _
        "")
End Sub


