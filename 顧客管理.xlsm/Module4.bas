Attribute VB_Name = "Module4"
Option Explicit
'��t�p�̗����v�Z���s�����[�U�[�t�H�[��
Sub calFrontValue()
    UserForm1.Show vbModeless
End Sub
'CAST�ւ̎x�������z���v�Z���郆�[�U�[�t�H�[��
Sub showCastPayment()
    UserForm2.Show vbModal
End Sub
'�s�}������@�\
Sub insertRow()
    Dim targetRow As String
    
    targetRow = InputBox("���s�ڂɑ}���������ł����H", "�s�̑}��", "(��:999)")
    If StrPtr(targetRow) = 0 Or targetRow = 0 Then
            Exit Sub
    End If
    
    '�V�[�g�̕ی����
    ActiveSheet.Unprotect Password:="042595"
    
    ActiveWorkbook.Worksheets("���̓V�[�g").Rows(targetRow - 1).Copy
    ActiveWorkbook.Worksheets("���̓V�[�g").Rows(targetRow).Insert
    ActiveWorkbook.Worksheets("���̓V�[�g").Rows(targetRow).PasteSpecial (xlPasteAll)
    ActiveWorkbook.Worksheets("���̓V�[�g").Cells(targetRow, 3).ClearContents
    ActiveWorkbook.Worksheets("���̓V�[�g").Range(Cells(targetRow, 5), Cells(targetRow, 22)).ClearContents
    
    '�V�[�g�̕ی�
    ActiveSheet.Protect Password:="042595"
End Sub
'�ڋq�̓d�b�ԍ��������œ��͂���@�\
Sub enterCustomerNumber()
    '''''''''''''''''''''''''''''
    '''''''''''''''''�t�@�C����`
    '''''''''''''''''''''''''''''
'    Dim customerTablePath As Variant '�t�@�C���p�X�i�[�p�ϐ�
'    customerTablePath = "E:\�ڋq�Ǘ�.xlsm"
'    customerTablePath = "C:\Users\seifu\OneDrive\�h�L�������g\����f�[�^_230628\�ڋq�Ǘ�.xlsm"
'    Dim customerTable As Workbook '�u�b�N���`
'    Set customerTable = Workbooks.Open(customerTablePath)
    Dim inputSheet As Worksheet '���̓V�[�g���`
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")
    Dim customerSheet As Worksheet '����ꗗ�V�[�g���`
    Set customerSheet = ActiveWorkbook.Worksheets("����ꗗ")
    
    '''''''''''''''''''''''''''''
    '''''''''''''''''''�e���W��`
    '''''''''''''''''''''''''''''
    Dim inputMediaColumn As Integer '���̓V�[�g�A�}�̗̂�ԍ���`
    inputMediaColumn = 5
    Dim inputNameColumn As Integer '���̓V�[�g�A���q�l���̗�ԍ���`
    inputNameColumn = 8
    Dim inputPhoneNumColumn As Integer '���̓V�[�g�A�d�b�ԍ��̗�ԍ���`
    inputPhoneNumColumn = 9
    
    Dim customerNameColumn As Integer '����ꗗ�V�[�g�A������̗�ԍ���`
    customerNameColumn = 2
    Dim customerPhoneNumColumn As Integer '����ꗗ�V�[�g�A�d�b�ԍ��̗�ԍ���`
    customerPhoneNumColumn = 3
    
    inputSheet.Activate
    '�V�[�g�̕ی����
    ActiveSheet.Unprotect Password:="042595"
    
    
    '''''''''''''''''''''''''''''
    '�e�V�[�g�̍ŏI�s�܂ł̓��e���擾����B��͈͂͏�Œ�`������ԍ����g��
    '''''''''''''''''''''''''''''
    inputSheet.Activate
    Dim inputNameLastRow As Long
    inputNameLastRow = Range("H1").CurrentRegion(Range("H1").CurrentRegion.Count).Row '���̓V�[�g�A���q�l����̍ŏI�s���擾
    Dim inputArray
    inputArray = Range(Cells(3, 1), Cells(inputNameLastRow, inputPhoneNumColumn)).Value '���̓V�[�g�̓��e��z��Ƃ��Ď擾
    
    customerSheet.Activate
    Dim customerNameLastRow As Long
    customerNameLastRow = Range("B1").CurrentRegion(Range("B1").CurrentRegion.Count).Row '����ꗗ�V�[�g�A���q�l����̍ŏI�s���擾
    Dim customerArray
    customerArray = Range(Cells(2, 1), Cells(customerNameLastRow, customerPhoneNumColumn)).Value '����ꗗ�V�[�g�̓��e��z��Ƃ��Ď擾
    
    '''''''''''''''''''''''''''''
    '�d�b�ԍ�����/�������ݏ���
    '''''''''''''''''''''''''''''
    Dim i As Long '���̓V�[�g,�s�J�E���^
    Dim j As Long '����ꗗ�V�[�g,�s�J�E���^
    Dim k As Long '�������X�g�p��For�J�E���^
    Dim nonNumName As Variant '�d�b�ԍ�����������̖��O�i�[�p�ϐ�
    Dim sameNameFlag As Integer '�������݃t���O
    Dim sameNameNumberList As Variant '��������̔ԍ����X�g
    Dim sameNameNumberListStr As String '��������̉�4�����X�g�𕶎���
    Dim underNum4 As Long
    Dim underNum4SuccessFlag As Integer '��4�����͂����������t���O
    
    inputSheet.Activate
    For i = LBound(inputArray, 1) To UBound(inputArray, 1) '���̓V�[�g�̑S�s�𒲂ׂ�
        If inputArray(i, inputMediaColumn) = "R" And inputArray(i, inputPhoneNumColumn) = "" Then '�}�̂�R�A���d�b�ԍ����Ȃ�
            Debug.Print (inputArray(i, inputNameColumn))
            sameNameFlag = 0
            ReDim sameNameNumberList(0)
            nonNumName = inputArray(i, inputNameColumn) '�d�b�ԍ��̂Ȃ�����̖��O���i�[
            For j = LBound(customerArray, 1) To UBound(customerArray, 1) '����ꗗ�V�[�g�̑S�s�𒲂ׂ�
                If customerArray(j, customerNameColumn) = nonNumName Then '��v�����������ǂ���
                    Debug.Print (customerArray(j, customerPhoneNumColumn))
                    Cells(i + 2, inputPhoneNumColumn).Value = customerArray(j, customerPhoneNumColumn) '����̓d�b�ԍ����Z���ɓ��͂���
                    sameNameNumberList(sameNameFlag) = customerArray(j, customerPhoneNumColumn)
                    sameNameFlag = sameNameFlag + 1
                    ReDim Preserve sameNameNumberList(sameNameFlag)
                End If
            Next j
            '�������ݎ��̏���
            If sameNameFlag > 1 Then
                
                '��4�����͎��̓��͒l����������܂ŌJ��Ԃ�
                underNum4SuccessFlag = 0
                Do While (underNum4SuccessFlag = 0)
                    sameNameNumberListStr = nonNumName + "�l�͕������܂��B��������4������͂��Ă��������B"
                    '�����ڋq�̉�4������ׂ���������`�����鏈��
                    k = 0
                    For k = LBound(sameNameNumberList) To UBound(sameNameNumberList) - 1
                        sameNameNumberListStr = sameNameNumberListStr + vbCrLf + "�E" + Right(sameNameNumberList(k), 4)
                    Next k
                    
                    '���̓{�b�N�X�\�����āA���͒l���󂯎��
                    underNum4 = Application.InputBox(Prompt:=sameNameNumberListStr, Title:="�����̌ڋq���������܂��B", Default:="0000")
                    
                    '���͂��ꂽ4���������ڋq�̔ԍ����X�g�ɂ���΃Z���ɏ������ޏ���
                    k = 0
                    For k = LBound(sameNameNumberList) To UBound(sameNameNumberList) - 1
                        If underNum4 = Right(sameNameNumberList(k), 4) Then
                            Cells(i + 2, inputPhoneNumColumn).Value = sameNameNumberList(k)
                            underNum4SuccessFlag = 1 '���͒l�������Ă������߃t���O�𗧂Ă�
                        End If
                    Next k
                    If underNum4SuccessFlag = 0 Then
                        MsgBox "���̓��͒l�͍����Ă��܂����c�H(��:�S�p��NG�ł�)", vbExclamation
                    End If
                    If StrPtr(underNum4) = 0 Then
                        MsgBox "�L�����Z�����N���b�N���ꂽ����" + nonNumName + "�l�̔ԍ��͏ȗ����܂�", vbExclamation
                        Cells(i + 2, inputPhoneNumColumn).ClearContents
                        Exit Do
                    End If
                Loop
            End If
        ElseIf inputArray(i, 8) = "" Then
            Exit For
        End If
    Next i
    
    '�ŏI�s�̃Z����I������(�X�N���v�g���s��ɕ\���������o�O�΍�)
    Cells(i, 3).Select
    
    '�V�[�g�̕ی�
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
End Sub
'����ڋq�̓���CAST���ē������������ȈՕ\������@�\
Sub showHistory()
    '�E�E�E�ϐ���`�E�E�E
    Dim inputSheet As Worksheet '���̓V�[�g���`
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")
    Dim customerSheet As Worksheet
    Set customerSheet = ActiveWorkbook.Worksheets("����ꗗ")
    
    inputSheet.Activate
    
    '���̓V�[�g�i�[�z��̒�`
    Dim customerLastRow As Long
    customerLastRow = Cells(Rows.Count, 3).End(xlUp).Row '���̓V�[�g�̍ŏI�s���`
    Dim customerLastColumn As Long
    customerLastColumn = (Cells(1, Columns.Count).End(xlToLeft).Column) '���̓V�[�g�̍ŏI����`�A1�s�ڃw�b�_����
    
    Dim inputDataArr As Variant
    inputDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '���̓V�[�g�̓��e��S�擾
    
    '�����i�[�z��̒�`
    Dim castHistoryOnCustomerHistory As Variant
    ReDim castHistoryOnCustomerHistory(0)
    
    customerSheet.Activate
    
    '�ڋq�ԍ��̕ϐ��錾�AInputBox�Ăяo��
    Dim customerNumber As String
    customerNumber = customerNumberInput
    
    '�L���X�g���̕ϐ��錾�AInputBox�Ăяo��
    Dim castName As String
    castName = castNameInput
    
    '��t�f�[�^�Ɍڋq����Cast������v������castHistoryOnCustomerHistory�ɑ������
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
    
    '�����𐮌`���ĕ����񉻁i�Ǔ_�̑}���j
    Dim stringHistory As String
    
    For i = LBound(castHistoryOnCustomerHistory) To UBound(castHistoryOnCustomerHistory)
        If i = LBound(castHistoryOnCustomerHistory) Then
            stringHistory = castHistoryOnCustomerHistory(i)
        ElseIf castHistoryOnCustomerHistory(i) <> Empty Then
            stringHistory = stringHistory & "�A" & castHistoryOnCustomerHistory(i)
        End If
    Next i
    
    customerSheet.Activate
    
    MsgBox "����ԍ��y" & customerNumber & " �z�̉���l�A�y" & castName & "�z����ł̎�t�́A" & vbCrLf & "���v�Ły" & x & "�z��ł��B" & vbCrLf & vbCrLf & "�w�����t" & vbCrLf & stringHistory
End Sub
'�d�b�ԍ����̓_�C�A���O_showHistory()�p��function�v���V�[�W��
Function customerNumberInput()
    customerNumberInput = InputBox("�ڋq�̉���ԍ��i�o�^�d�b�ԍ��j����͂��Ă��������B", "����ԍ�����", "08012345678")
End Function
'���������̓_�C�A���O_showHistory()�p��function�v���V�[�W��
Function castNameInput()
    castNameInput = InputBox("���̎q�̌���������͂��Ă��������B", "CAST������", "����")
End Function
