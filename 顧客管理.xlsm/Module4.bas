Attribute VB_Name = "Module4"
Option Explicit
'受付用の料金計算を行うユーザーフォーム
Sub calFrontValue()
    UserForm1.Show vbModeless
End Sub
'CASTへの支払い金額を計算するユーザーフォーム
Sub showCastPayment()
    UserForm2.Show vbModal
End Sub
'行挿入する機能
Sub insertRow()
    Dim targetRow As String
    
    targetRow = InputBox("何行目に挿入したいですか？", "行の挿入", "(例:999)")
    If StrPtr(targetRow) = 0 Or targetRow = 0 Then
            Exit Sub
    End If
    
    'シートの保護解除
    ActiveSheet.Unprotect Password:="042595"
    
    ActiveWorkbook.Worksheets("入力シート").Rows(targetRow - 1).Copy
    ActiveWorkbook.Worksheets("入力シート").Rows(targetRow).Insert
    ActiveWorkbook.Worksheets("入力シート").Rows(targetRow).PasteSpecial (xlPasteAll)
    ActiveWorkbook.Worksheets("入力シート").Cells(targetRow, 3).ClearContents
    ActiveWorkbook.Worksheets("入力シート").Range(Cells(targetRow, 5), Cells(targetRow, 22)).ClearContents
    
    'シートの保護
    ActiveSheet.Protect Password:="042595"
End Sub
'顧客の電話番号を自動で入力する機能
Sub enterCustomerNumber()
    '''''''''''''''''''''''''''''
    '''''''''''''''''ファイル定義
    '''''''''''''''''''''''''''''
'    Dim customerTablePath As Variant 'ファイルパス格納用変数
'    customerTablePath = "E:\顧客管理.xlsm"
'    customerTablePath = "C:\Users\seifu\OneDrive\ドキュメント\売上データ_230628\顧客管理.xlsm"
'    Dim customerTable As Workbook 'ブックを定義
'    Set customerTable = Workbooks.Open(customerTablePath)
    Dim inputSheet As Worksheet '入力シートを定義
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")
    Dim customerSheet As Worksheet '会員一覧シートを定義
    Set customerSheet = ActiveWorkbook.Worksheets("会員一覧")
    
    '''''''''''''''''''''''''''''
    '''''''''''''''''''各座標定義
    '''''''''''''''''''''''''''''
    Dim inputMediaColumn As Integer '入力シート、媒体の列番号定義
    inputMediaColumn = 5
    Dim inputNameColumn As Integer '入力シート、お客様名の列番号定義
    inputNameColumn = 8
    Dim inputPhoneNumColumn As Integer '入力シート、電話番号の列番号定義
    inputPhoneNumColumn = 9
    
    Dim customerNameColumn As Integer '会員一覧シート、会員名の列番号定義
    customerNameColumn = 2
    Dim customerPhoneNumColumn As Integer '会員一覧シート、電話番号の列番号定義
    customerPhoneNumColumn = 3
    
    inputSheet.Activate
    'シートの保護解除
    ActiveSheet.Unprotect Password:="042595"
    
    
    '''''''''''''''''''''''''''''
    '各シートの最終行までの内容を取得する。列範囲は上で定義した列番号を使う
    '''''''''''''''''''''''''''''
    inputSheet.Activate
    Dim inputNameLastRow As Long
    inputNameLastRow = Range("H1").CurrentRegion(Range("H1").CurrentRegion.Count).Row '入力シート、お客様名列の最終行を取得
    Dim inputArray
    inputArray = Range(Cells(3, 1), Cells(inputNameLastRow, inputPhoneNumColumn)).Value '入力シートの内容を配列として取得
    
    customerSheet.Activate
    Dim customerNameLastRow As Long
    customerNameLastRow = Range("B1").CurrentRegion(Range("B1").CurrentRegion.Count).Row '会員一覧シート、お客様名列の最終行を取得
    Dim customerArray
    customerArray = Range(Cells(2, 1), Cells(customerNameLastRow, customerPhoneNumColumn)).Value '会員一覧シートの内容を配列として取得
    
    '''''''''''''''''''''''''''''
    '電話番号検索/書き込み処理
    '''''''''''''''''''''''''''''
    Dim i As Long '入力シート,行カウンタ
    Dim j As Long '会員一覧シート,行カウンタ
    Dim k As Long '同名リスト用のForカウンタ
    Dim nonNumName As Variant '電話番号が無い会員の名前格納用変数
    Dim sameNameFlag As Integer '同名存在フラグ
    Dim sameNameNumberList As Variant '同名会員の番号リスト
    Dim sameNameNumberListStr As String '同名会員の下4桁リストを文字列化
    Dim underNum4 As Long
    Dim underNum4SuccessFlag As Integer '下4桁入力が成功したフラグ
    
    inputSheet.Activate
    For i = LBound(inputArray, 1) To UBound(inputArray, 1) '入力シートの全行を調べる
        If inputArray(i, inputMediaColumn) = "R" And inputArray(i, inputPhoneNumColumn) = "" Then '媒体がR、かつ電話番号がない
            Debug.Print (inputArray(i, inputNameColumn))
            sameNameFlag = 0
            ReDim sameNameNumberList(0)
            nonNumName = inputArray(i, inputNameColumn) '電話番号のない会員の名前を格納
            For j = LBound(customerArray, 1) To UBound(customerArray, 1) '会員一覧シートの全行を調べる
                If customerArray(j, customerNameColumn) = nonNumName Then '一致する会員名かどうか
                    Debug.Print (customerArray(j, customerPhoneNumColumn))
                    Cells(i + 2, inputPhoneNumColumn).Value = customerArray(j, customerPhoneNumColumn) '会員の電話番号をセルに入力する
                    sameNameNumberList(sameNameFlag) = customerArray(j, customerPhoneNumColumn)
                    sameNameFlag = sameNameFlag + 1
                    ReDim Preserve sameNameNumberList(sameNameFlag)
                End If
            Next j
            '同名存在時の処理
            If sameNameFlag > 1 Then
                
                '下4桁入力時の入力値が成功するまで繰り返す
                underNum4SuccessFlag = 0
                Do While (underNum4SuccessFlag = 0)
                    sameNameNumberListStr = nonNumName + "様は複数います。正しい下4桁を入力してください。"
                    '同名顧客の下4桁を並べた文字列を形成する処理
                    k = 0
                    For k = LBound(sameNameNumberList) To UBound(sameNameNumberList) - 1
                        sameNameNumberListStr = sameNameNumberListStr + vbCrLf + "・" + Right(sameNameNumberList(k), 4)
                    Next k
                    
                    '入力ボックス表示して、入力値を受け取る
                    underNum4 = Application.InputBox(Prompt:=sameNameNumberListStr, Title:="同名の顧客が複数います。", Default:="0000")
                    
                    '入力された4桁が同名顧客の番号リストにあればセルに書き込む処理
                    k = 0
                    For k = LBound(sameNameNumberList) To UBound(sameNameNumberList) - 1
                        If underNum4 = Right(sameNameNumberList(k), 4) Then
                            Cells(i + 2, inputPhoneNumColumn).Value = sameNameNumberList(k)
                            underNum4SuccessFlag = 1 '入力値が合っていたためフラグを立てる
                        End If
                    Next k
                    If underNum4SuccessFlag = 0 Then
                        MsgBox "その入力値は合っていますか…？(注:全角はNGです)", vbExclamation
                    End If
                    If StrPtr(underNum4) = 0 Then
                        MsgBox "キャンセルがクリックされたため" + nonNumName + "様の番号は省略します", vbExclamation
                        Cells(i + 2, inputPhoneNumColumn).ClearContents
                        Exit Do
                    End If
                Loop
            End If
        ElseIf inputArray(i, 8) = "" Then
            Exit For
        End If
    Next i
    
    '最終行のセルを選択する(スクリプト実行後に表示が乱れるバグ対策)
    Cells(i, 3).Select
    
    'シートの保護
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
End Sub
'特定顧客の特定CASTを案内した履歴を簡易表示する機能
Sub showHistory()
    '・・・変数定義・・・
    Dim inputSheet As Worksheet '入力シートを定義
    Set inputSheet = ActiveWorkbook.Worksheets("入力シート")
    Dim customerSheet As Worksheet
    Set customerSheet = ActiveWorkbook.Worksheets("会員一覧")
    
    inputSheet.Activate
    
    '入力シート格納配列の定義
    Dim customerLastRow As Long
    customerLastRow = Cells(Rows.Count, 3).End(xlUp).Row '入力シートの最終行を定義
    Dim customerLastColumn As Long
    customerLastColumn = (Cells(1, Columns.Count).End(xlToLeft).Column) '入力シートの最終列を定義、1行目ヘッダから
    
    Dim inputDataArr As Variant
    inputDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '入力シートの内容を全取得
    
    '履歴格納配列の定義
    Dim castHistoryOnCustomerHistory As Variant
    ReDim castHistoryOnCustomerHistory(0)
    
    customerSheet.Activate
    
    '顧客番号の変数宣言、InputBox呼び出し
    Dim customerNumber As String
    customerNumber = customerNumberInput
    
    'キャスト名の変数宣言、InputBox呼び出し
    Dim castName As String
    castName = castNameInput
    
    '受付データに顧客名とCast名が一致したらcastHistoryOnCustomerHistoryに代入する
    Dim i As Long
    i = 0
    Dim x As Long
    x = 0
    
    For i = LBound(inputDataArr, 1) To UBound(inputDataArr, 1)
        If inputDataArr(i, 9) = customerNumber And inputDataArr(i, 7) = castName Then
            castHistoryOnCustomerHistory(x) = inputDataArr(i, 3)
            x = x + 1
            ReDim Preserve castHistoryOnCustomerHistory(x)
        End If
    Next i
    
    '履歴を整形して文字列化（読点の挿入）
    Dim stringHistory As String
    
    For i = LBound(castHistoryOnCustomerHistory) To UBound(castHistoryOnCustomerHistory)
        If i = LBound(castHistoryOnCustomerHistory) Then
            stringHistory = castHistoryOnCustomerHistory(i)
        ElseIf castHistoryOnCustomerHistory(i) <> Empty Then
            stringHistory = stringHistory & "、" & castHistoryOnCustomerHistory(i)
        End If
    Next i
    
    customerSheet.Activate
    
    MsgBox "会員番号【" & customerNumber & " 】の会員様、【" & castName & "】さんでの受付は、" & vbCrLf & "合計で【" & x & "】回です。" & vbCrLf & vbCrLf & "指名日付" & vbCrLf & stringHistory
End Sub
'電話番号入力ダイアログ_showHistory()用のfunctionプロシージャ
Function customerNumberInput()
    customerNumberInput = InputBox("顧客の会員番号（登録電話番号）を入力してください。", "会員番号入力", "08012345678")
End Function
'源氏名入力ダイアログ_showHistory()用のfunctionプロシージャ
Function castNameInput()
    castNameInput = InputBox("女の子の源氏名を入力してください。", "CAST名入力", "あつし")
End Function
