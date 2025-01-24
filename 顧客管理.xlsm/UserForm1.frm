VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "受付料金計算"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20070
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    UserForm1.Caption = "受付料金計算"

    '受付料金計算
    OptionButtonA.Caption = "年末年始"
    OptionButtonA.GroupName = "G0"
    OptionButtonB.Caption = "夜這い特価"
    OptionButtonB.GroupName = "G0"
    OptionButtonB.Value = True

    OptionButton1.Caption = 60
    OptionButton1.GroupName = "G1"
    OptionButton2.Caption = 80
    OptionButton2.GroupName = "G1"
    OptionButton3.Caption = 100
    OptionButton3.GroupName = "G1"
    OptionButton4.Caption = 120
    OptionButton4.GroupName = "G1"
    OptionButton5.Caption = 140
    OptionButton5.GroupName = "G1"
    OptionButton6.Caption = 160
    OptionButton6.GroupName = "G1"
    OptionButton7.Caption = 180
    OptionButton7.GroupName = "G1"
    OptionButton8.Caption = 200
    OptionButton8.GroupName = "G1"
    OptionButton9.Caption = 220
    OptionButton9.GroupName = "G1"
    OptionButton10.Caption = 240
    OptionButton10.GroupName = "G1"
    
    OptionButton11.Caption = 260
    OptionButton11.GroupName = "G1"
    OptionButton12.Caption = 280
    OptionButton12.GroupName = "G1"
    OptionButton13.Caption = 300
    OptionButton13.GroupName = "G1"
    OptionButton14.Caption = 320
    OptionButton14.GroupName = "G1"
    OptionButton15.Caption = 340
    OptionButton15.GroupName = "G1"
    OptionButton16.Caption = 360
    OptionButton16.GroupName = "G1"
    OptionButton17.Caption = 380
    OptionButton17.GroupName = "G1"
    OptionButton18.Caption = 400
    OptionButton18.GroupName = "G1"
    OptionButton19.Caption = 420
    OptionButton19.GroupName = "G1"
    OptionButton20.Caption = 440
    OptionButton20.GroupName = "G1"
    OptionButton1.Value = True
    
    OptionButton101.Caption = "普通 +0"
    OptionButton101.GroupName = "G3"
    OptionButton102.Caption = "スター +1000"
    OptionButton102.GroupName = "G3"
    OptionButton103.Caption = "プラチナ +2000"
    OptionButton103.GroupName = "G3"
    OptionButton101.Value = True
    
    
    OptionButton201.Caption = "新規無料 +0"
    OptionButton201.GroupName = "G4"
    OptionButton202.Caption = "フリー +0"
    OptionButton202.GroupName = "G4"
    OptionButton203.Caption = "会員写指 +1000"
    OptionButton203.GroupName = "G4"
    OptionButton204.Caption = "本指 +2000"
    OptionButton204.GroupName = "G4"
    OptionButton205.Caption = "入会フリー +2000"
    OptionButton205.GroupName = "G4"
    OptionButton201.Value = True
    
    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False
    CheckBox11.Value = False
    
    '女子給計算
    OptionButton1001.Caption = 6000
    OptionButton1002.Caption = 7000
    OptionButton1003.Caption = 8000
    OptionButton1004.Caption = 9000
    OptionButton1005.Caption = 10000
    OptionButton1006.Caption = 11000
    
    OptionButton1003.Value = True
    
    OptionButton1001.GroupName = "G11"
    OptionButton1002.GroupName = "G11"
    OptionButton1003.GroupName = "G11"
    OptionButton1004.GroupName = "G11"
    OptionButton1005.GroupName = "G11"
    OptionButton1006.GroupName = "G11"
    
    itemEnabled
    calculateTotal
    calculateCastBack
    
End Sub
Private Sub itemEnabled()
'年末年始orElseコース
    If OptionButtonB = True Then
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
        OptionButton3.Enabled = True
        OptionButton4.Enabled = True
        OptionButton5.Enabled = True
        OptionButton6.Enabled = True
        OptionButton7.Enabled = True
        OptionButton8.Enabled = True
        OptionButton9.Enabled = True
        OptionButton10.Enabled = True
        OptionButton11.Enabled = True
        OptionButton12.Enabled = True
        OptionButton13.Enabled = True
        OptionButton14.Enabled = True
        OptionButton15.Enabled = True
        OptionButton16.Enabled = True
        OptionButton17.Enabled = True
        OptionButton18.Enabled = True
        OptionButton19.Enabled = True
        OptionButton20.Enabled = True
    ElseIf OptionButtonA = True Then
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
        OptionButton3.Enabled = True
        OptionButton4.Enabled = True
        OptionButton5.Enabled = True
        OptionButton6.Enabled = True
        OptionButton7.Enabled = True
        OptionButton8.Enabled = True
        OptionButton9.Enabled = True
        OptionButton10.Enabled = True
        OptionButton11.Enabled = True
        OptionButton12.Enabled = True
        OptionButton13.Enabled = True
        OptionButton14.Enabled = True
        OptionButton15.Enabled = True
        OptionButton16.Enabled = True
        OptionButton17.Enabled = True
        OptionButton18.Enabled = True
        OptionButton19.Enabled = True
        OptionButton20.Enabled = True
    End If
    
    
    '【年末年始処理】指名周りの有効/無効、caption変更
    If OptionButtonA.Value = True Then
        OptionButton201.Caption = "入会指名 +4000"
        OptionButton202.Caption = "会員フリー +0"
        OptionButton203.Caption = "写指 +2000"
        OptionButton205.Enabled = True
    Else
        OptionButton201.Caption = "新規無料 +0"
        OptionButton202.Caption = "フリー +0円"
        OptionButton203.Caption = "写指 +1000"
        OptionButton205.Enabled = False
    End If
    
'本指名
    If OptionButton204 = True Then
        CheckBox1.Enabled = False
    ElseIf OptionButton204 = False Then
        CheckBox1.Enabled = True
    End If
    
'体入割or口コミ割
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False Then '体入割と口コミ割がチェック無の場合、全EnabledをTrue
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        OptionButton102.Enabled = True
        OptionButton103.Enabled = True
        OptionButton203.Enabled = True
        OptionButton204.Enabled = True
    ElseIf CheckBox1.Value = True Then '体入割がチェック有の場合、
        CheckBox1.Enabled = True
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        OptionButton203.Enabled = False
        OptionButton204.Enabled = False
        OptionButton201.Value = True
    ElseIf CheckBox2.Value = True Then 'メルマガ割がチェック有の場合、
        CheckBox1.Enabled = False
        CheckBox2.Enabled = True
        CheckBox3.Enabled = False
        OptionButton102.Enabled = True
        OptionButton103.Enabled = True
        OptionButton203.Enabled = True
        OptionButton204.Enabled = True
    ElseIf CheckBox3.Value = True Then '口コミ割がチェック有の場合、
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = True
        OptionButton102.Enabled = True
        OptionButton103.Enabled = True
        OptionButton203.Enabled = True
        OptionButton204.Enabled = True
    End If
End Sub
Private Sub calculateTotal()
    Dim courseValue As Long
    Dim plusValue1 As Long
    Dim plusValue2 As Long
    Dim sumPlusValue As Long
    Dim discountValue As Long
    Dim TotalValue As Long
    
    '基本料金設定
    If OptionButtonA.Value = True Then
        Select Case True
            Case OptionButton1.Value = True
                courseValue = 18000
            Case OptionButton2.Value = True
                courseValue = 23000
            Case OptionButton3.Value = True
                courseValue = 28000
            Case OptionButton4.Value = True
                courseValue = 33000
            Case OptionButton5.Value = True
                courseValue = 38000
            Case OptionButton6.Value = True
                courseValue = 43000
            Case OptionButton7.Value = True
                courseValue = 48000
            Case OptionButton8.Value = True
                courseValue = 53000
            Case OptionButton9.Value = True
                courseValue = 58000
            Case OptionButton10.Value = True
                courseValue = 63000
            Case OptionButton11.Value = True
                courseValue = 68000
            Case OptionButton12.Value = True
                courseValue = 73000
            Case OptionButton13.Value = True
                courseValue = 78000
            Case OptionButton14.Value = True
                courseValue = 83000
            Case OptionButton15.Value = True
                courseValue = 88000
            Case OptionButton16.Value = True
                courseValue = 93000
            Case OptionButton17.Value = True
                courseValue = 98000
            Case OptionButton18.Value = True
                courseValue = 103000
            Case OptionButton19.Value = True
                courseValue = 108000
            Case OptionButton20.Value = True
                courseValue = 113000
            Case Else
                courseValue = 0
        End Select
    ElseIf OptionButtonB.Value = True Then
        Select Case True
            Case OptionButton1.Value = True
                courseValue = 15000
            Case OptionButton2.Value = True
                courseValue = 20000
            Case OptionButton3.Value = True
                courseValue = 25000
            Case OptionButton4.Value = True
                courseValue = 30000
            Case OptionButton5.Value = True
                courseValue = 35000
            Case OptionButton6.Value = True
                courseValue = 40000
            Case OptionButton7.Value = True
                courseValue = 45000
            Case OptionButton8.Value = True
                courseValue = 50000
            Case OptionButton9.Value = True
                courseValue = 55000
            Case OptionButton10.Value = True
                courseValue = 60000
            Case OptionButton11.Value = True
                courseValue = 65000
            Case OptionButton12.Value = True
                courseValue = 70000
            Case OptionButton13.Value = True
                courseValue = 75000
            Case OptionButton14.Value = True
                courseValue = 80000
            Case OptionButton15.Value = True
                courseValue = 85000
            Case OptionButton16.Value = True
                courseValue = 90000
            Case OptionButton17.Value = True
                courseValue = 95000
            Case OptionButton18.Value = True
                courseValue = 100000
            Case OptionButton19.Value = True
                courseValue = 105000
            Case OptionButton20.Value = True
                courseValue = 110000
            Case Else
                courseValue = 0
        End Select
    End If
    
    'クラス料プラス
    Select Case True
        Case OptionButton101.Value = True
            plusValue1 = plusValue1 + 0
        Case OptionButton102.Value = True
            plusValue1 = plusValue1 + 1000
        Case OptionButton103.Value = True
            plusValue1 = plusValue1 + 2000
        Case Else
            plusValue1 = 0
    End Select
    
    '指名料プラス
    Select Case True
        Case OptionButton201.Value = True
            If OptionButtonA.Value = True Then
                plusValue2 = plusValue2 + 4000
            Else
                plusValue2 = plusValue2 + 0
            End If
        Case OptionButton202.Value = True
            plusValue2 = plusValue2 + 0
        Case OptionButton203.Value = True
            If OptionButtonA.Value = True Then
                plusValue2 = plusValue2 + 2000
            Else
                plusValue2 = plusValue2 + 1000
            End If
        Case OptionButton204.Value = True
            plusValue2 = plusValue2 + 2000
        Case OptionButton205.Value = True
            plusValue2 = plusValue2 + 2000
        Case Else
            plusValue2 = 0
    End Select
    
    '割引マイナス
    If CheckBox1.Value = True Then
        discountValue = discountValue + 2000
    End If
    If CheckBox2.Value = True Then
        discountValue = discountValue + 1000
    End If
    If CheckBox3.Value = True Then
        discountValue = discountValue + 3000
    End If
    
    'その他割引マイナス
    If TextBox1.Value <> "" Then
        discountValue = discountValue + TextBox1.Value * 1000
    End If
    
    sumPlusValue = plusValue1 + plusValue2
    TotalValue = courseValue + sumPlusValue - discountValue
    
    'OP等プラス料金
    If TextBox2.Value <> "" Then
        sumPlusValue = sumPlusValue + TextBox2.Value * 1000
    End If
    
    TotalValue = courseValue + sumPlusValue - discountValue
    
    'カード払なら総額に10%乗せ
    '税込み版
    If CheckBox11.Value = True Then
        Label26.Caption = "税込み額(税抜き額)"
        Label38.Caption = "カード手数料込み支払総額"
        TotalValue = TotalValue * 1.1
        Label11.Caption = TotalValue & "(" & TotalValue / 1.1 & ")"
        Label37.Caption = (TotalValue / 1.1) * 1.2
    ElseIf CheckBox11.Value = False Then
        Label26.Caption = "税抜き額"
        Label38.Caption = "お客様支払総額(税10％)"
        TotalValue = TotalValue
        Label11.Caption = TotalValue
        Label37.Caption = TotalValue * 1.1
    End If
    
End Sub
Private Sub calculateCastBack()
    Dim castUnitBack As Long
    Dim courseBack As Long
    Dim courseFullBack As Long
    Dim fullBackDiff As Long
    Dim discountBack As Long
    Dim nomBack As Long
    Dim plusBack As Long
    Dim castBack As Long
    Dim classValue As Long
    Dim profitsValue As Long
    
    '単価決定
    Select Case True
        Case OptionButton1001.Value = True
            castUnitBack = OptionButton1001.Caption '6000
        Case OptionButton1002.Value = True
            castUnitBack = OptionButton1002.Caption '7000
        Case OptionButton1003.Value = True
            castUnitBack = OptionButton1003.Caption '8000 デフォルト値
        Case OptionButton1004.Value = True
            castUnitBack = OptionButton1004.Caption '9000
        Case OptionButton1005.Value = True
            castUnitBack = OptionButton1005.Caption '10000
        Case OptionButton1006.Value = True
            castUnitBack = OptionButton1006.Caption '11000
        Case Else
            castUnitBack = 0
    End Select
    
    'コースフルバック算出
    Select Case True
        Case OptionButton1.Value = True '60min
            courseFullBack = 11000
            courseBack = 0
        Case OptionButton2.Value = True '80min
            courseFullBack = 16000
            courseBack = 1
        Case OptionButton3.Value = True '100min
            courseFullBack = 21000
            courseBack = 2
        Case OptionButton4.Value = True '120min
            courseFullBack = 26000
            courseBack = 3
        Case OptionButton5.Value = True '140min
            courseFullBack = 31000
            courseBack = 4
        Case OptionButton6.Value = True '160min
            courseFullBack = 36000
            courseBack = 5
        Case OptionButton7.Value = True '180min
            courseFullBack = 41000
            courseBack = 6
        Case OptionButton8.Value = True '200min
            courseFullBack = 46000
            courseBack = 7
        Case OptionButton9.Value = True '220min
            courseFullBack = 51000
            courseBack = 8
        Case OptionButton10.Value = True '240min
            courseFullBack = 56000
            courseBack = 9
        Case OptionButton11.Value = True '260min
            courseFullBack = 61000
            courseBack = 10
        Case OptionButton12.Value = True '280min
            courseFullBack = 66000
            courseBack = 11
        Case OptionButton13.Value = True '300min
            courseFullBack = 71000
            courseBack = 12
        Case OptionButton14.Value = True '320min
            courseFullBack = 76000
            courseBack = 13
        Case OptionButton15.Value = True '340min
            courseFullBack = 81000
            courseBack = 14
        Case OptionButton16.Value = True '360min
            courseFullBack = 86000
            courseBack = 15
        Case OptionButton17.Value = True '380min
            courseFullBack = 91000
            courseBack = 16
        Case OptionButton18.Value = True '400min
            courseFullBack = 96000
            courseBack = 17
        Case OptionButton19.Value = True '420min
            courseFullBack = 101000
            courseBack = 18
        Case OptionButton20.Value = True '440min
            courseFullBack = 106000
            courseBack = 19
        Case Else
            courseFullBack = 0
            courseBack = 0
    End Select
    
    courseBack = 2500 * courseBack
    
    '指名料プラス
    Select Case True
        Case OptionButton204.Value = True
            nomBack = 2000
        Case Else
            nomBack = 0
    End Select
    
    'ゲリラ割等減額
    If TextBox3.Value <> "" Then
        discountBack = TextBox3.Value * 1000
    End If
    
    'OPバック増額
    If TextBox4.Value <> "" Then
        plusBack = TextBox4.Value * 1000
    End If
    
    '女子給合計処理
    'キャスト単価＋コース料金＋本指-ゲリラ割＋OPバック
    castBack = castUnitBack + courseBack + nomBack - discountBack + plusBack
    Label22.Caption = castBack
    
    'フルバック差額算出
    'コースフルバック額＋クラス料-(女子給-本指-OPバック)(本指名料とOPバックは通常バックに含んでいるため差額から除外)
    If CheckBox101.Value = True Then
        Select Case True
            Case OptionButton101.Value = True
                classValue = 0
            Case OptionButton102.Value = True
                classValue = 1000
            Case OptionButton103.Value = True
                classValue = 2000
            Case Else
                classValue = 0
        End Select

        fullBackDiff = courseFullBack + classValue - (castBack - nomBack - plusBack)
        Label30.Caption = fullBackDiff
    ElseIf CheckBox101.Value = False Then
        Label30.Caption = 0
    End If
    
    '総女子給算出
    '女子給+フルバック差額
    Label36.Caption = castBack + fullBackDiff
    
    '店落ち算出
    '受付支払い額-(女子給+フルバック差額)
    '税抜き版
'    If CheckBox11.Value = True Then
'        profitsValue = (Label11.Caption / 1.1) - (castBack + fullBackDiff)
'    ElseIf CheckBox11.Value = False Then
'        profitsValue = Label11.Caption - (castBack + fullBackDiff)
'    End If
    
    '税込み版
    If CheckBox11.Value = True Then
        profitsValue = (Label37.Caption / 1.2) * 1.1 - (castBack + fullBackDiff)
    ElseIf CheckBox11.Value = False Then
        profitsValue = Label37.Caption - (castBack + fullBackDiff)
    End If
    
    Label33.Caption = profitsValue
End Sub
Private Sub writeAccount() '最終行への書き込み処理
    Dim ws As Worksheet
    Set ws = Worksheets("入力シート")
    
    Dim lRowSale As Long
    Dim lRowBack As Long
    Dim lRowProfits As Long
    
    '18,19,20列目の最終行を割り出し
    lRowSale = ws.Cells(Rows.Count, 18).End(xlUp).Row
    lRowBack = ws.Cells(Rows.Count, 19).End(xlUp).Row
    lRowProfits = ws.Cells(Rows.Count, 20).End(xlUp).Row
    
    '18,19,20列目の最終行へ書き出し
    '税抜き版 総売上書き出し
'    If CheckBox11.Value = True Then
'        ws.Cells(lRowSale + 1, 18) = Label11.Caption / 1.1
'    ElseIf CheckBox11.Value = False Then
'        ws.Cells(lRowSale + 1, 18) = Label11.Caption
'    End If
    
'    '税込み版 総売上書き出し
    If CheckBox11.Value = True Then
        ws.Cells(lRowSale + 1, 18) = (Label37.Caption / 1.2) * 1.1
    ElseIf CheckBox11.Value = False Then
        ws.Cells(lRowSale + 1, 18) = Label37.Caption
    End If
    
    ws.Cells(lRowBack + 1, 19) = Label36.Caption
    ws.Cells(lRowProfits + 1, 20) = Label33.Caption
    
End Sub
Private Sub OptionButtonA_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButtonB_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton1_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton2_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton3_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton4_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton5_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton6_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton7_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton8_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton9_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton10_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton11_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton12_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton13_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton14_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton15_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton16_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton17_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton18_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton19_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton20_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton101_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton102_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton103_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton201_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton202_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton203_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton204_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton205_Click()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub CheckBox1_Change()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub CheckBox2_Change()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub CheckBox3_Change()
    itemEnabled
    calculateTotal
    calculateCastBack
End Sub
Private Sub CheckBox11_Click()
    calculateTotal
    calculateCastBack
End Sub
Private Sub TextBox1_Change()
    calculateTotal
    calculateCastBack
End Sub
Private Sub TextBox2_Change()
    calculateTotal
    calculateCastBack
End Sub
Private Sub OptionButton1001_Click()
    calculateCastBack
End Sub
Private Sub OptionButton1002_Click()
    calculateCastBack
End Sub
Private Sub OptionButton1003_Click()
    calculateCastBack
End Sub
Private Sub OptionButton1004_Click()
    calculateCastBack
End Sub
Private Sub OptionButton1005_Click()
    calculateCastBack
End Sub
Private Sub OptionButton1006_Click()
    calculateCastBack
End Sub
Private Sub TextBox3_Change()
    calculateCastBack
End Sub
Private Sub TextBox4_Change()
    calculateCastBack
End Sub
Private Sub CheckBox101_Click()
    calculateCastBack
End Sub
Private Sub CommandButton1_Click()
    writeAccount
End Sub
