Attribute VB_Name = "Module1"
Option Explicit
Sub updateCustomerInfomationButton()
    
    '�ŏ��ɓ��̓V�[�g�̐������`�F�b�N���s��
    If inputSheetFalseCheck() = 1 Then 'inputSheetFalseCheck��Module3
        Exit Sub
    End If
    
    Call makeAscendingSheet 'Module3
    Call updateCustomerInfomation
    
End Sub
Sub updateCustomerInfomation()
'
' ���e�̍X�V Macro
'

'�E�E�E�ϐ���`�E�E�E
    Dim inputSheet As Worksheet '���̓V�[�g���`
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")
    Dim ascendSheet As Worksheet '�ڋq�����V�[�g���`
    Set ascendSheet = ActiveWorkbook.Worksheets("�ڋq����")
    Dim customerSheet As Worksheet '����ꗗ�V�[�g���`
    Set customerSheet = ActiveWorkbook.Worksheets("����ꗗ")
    
    Dim i As Long 'For�̃J�E���^'
    Dim j As Integer '��C���f�b�N�X�p�J�E���^'
    Dim k As Integer '�񐔕��̍s�����J�E���^�@���p�񐔕�'
    Dim bastInfoArr '����ꗗ�̏o�͗p_�z�� �w�b�_=��,�����,�d�b�ԍ�,NG���,���l'
    Dim historyInfoArr '����̗����̏o�͗p_�z�� �w�b�_=1��,2��,...'
    Dim ascendCustomerDataArr '�ڋq�����\�̊i�[�p_�z��
    
    ascendSheet.Activate
    
    Dim customerLastRow As Long
    customerLastRow = ascendSheet.Cells(Rows.Count, 3).End(xlUp).Row '�ڋq�����V�[�g�̍ŏI�s���`
    Dim customerLastColumn As Long
    customerLastColumn = Cells(1, Columns.Count).End(xlToLeft).Column '�ڋq�����V�[�g�̍ŏI����`�A1�s�ڃw�b�_����
    ascendCustomerDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '�ڋq�����V�[�g�̓��e��S�擾'
    
    Dim CountRow '�V�K�����J�E���g����bastInfoArr,historyInfoArr�̍s�����Ē�`����p�̕ϐ�
    CountRow = WorksheetFunction.CountIf(Range(Cells(2, 2), Cells(customerLastRow, 2)), "1")
    Dim CountColumn '�ڋq���A�ō��̉񐔂�����o����historyInfoArr�̗񐔂��Ē�`����p�̕ϐ�
    CountColumn = WorksheetFunction.Max(Range(Cells(2, 2), Cells(customerLastRow, 2)))

'�E�E�E�����E�E�E
    ReDim bastInfoArr(1 To CountRow, 1 To 5)
    ReDim historyInfoArr(1 To CountRow, 1 To CountColumn)
    
    k = 0
    
    '����ꗗ�֏o�͗p�̔z�񐮌`
    For i = LBound(ascendCustomerDataArr, 1) To UBound(ascendCustomerDataArr, 1)  'ascendCustomerDataArr�̍s�������[�v����
        j = 1 '���p�񐔃J�E���^
        
'        '���x�v���p�e�X�g�R�[�h
'        Dim testLong As Long
'        testLong = speedtest(Val(ascendCustomerDataArr(i, 3)))
'        Debug.Print i & testLong
'        '''''��End Test
        
        If ascendCustomerDataArr(i, 2) = j Then
            '�ڋq�ꗗ�i�[
            bastInfoArr(i - k, 1) = WorksheetFunction.CountIf(Range(Cells(2, 9), Cells(customerLastRow, 9)), ascendCustomerDataArr(i, 9))
            bastInfoArr(i - k, 2) = ascendCustomerDataArr(i, 8) '�q���i�[
            bastInfoArr(i - k, 3) = ascendCustomerDataArr(i, 9) '�q�Ԋi�[
            If ascendCustomerDataArr(i, 11) <> "" And ascendCustomerDataArr(i, 11) <> 0 Then
                bastInfoArr(i - k, 5) = ascendCustomerDataArr(i, 2) & "," & ascendCustomerDataArr(i, 11) '����̋q���l�i�[
            End If
            If ascendCustomerDataArr(i, 10) <> "" And ascendCustomerDataArr(i, 10) <> 0 Then
                bastInfoArr(i - k, 4) = ascendCustomerDataArr(i, 10) 'NG���i�[
            End If
            '���p�����i�[ ���e=���t3,���q��7,�z�e��12,�R�X13,����14
            historyInfoArr(i - k, j) = ascendCustomerDataArr(i, 3) & "," & ascendCustomerDataArr(i, 7) & vbLf & ascendCustomerDataArr(i, 12) & "," & ascendCustomerDataArr(i, 13) & "," & ascendCustomerDataArr(i, 14) '�����ł͎��F���̂��߂ɓ��t�͊����ĔNyy���Ȃ�
        ElseIf ascendCustomerDataArr(i, 2) > j Then '���p2��ڈȍ~�̊i�[����
            k = k + 1
            '2��ڈȍ~�A���l��NG���i�[
            If ascendCustomerDataArr(i, 11) <> "" And ascendCustomerDataArr(i, 11) <> 0 Then
                bastInfoArr(i - k, 5) = bastInfoArr(i - k, 5) & vbLf & ascendCustomerDataArr(i, 2) & "," & ascendCustomerDataArr(i, 11) '2��ڈȍ~�̋q���l�i�[
            End If
            If ascendCustomerDataArr(i, 10) <> "" And ascendCustomerDataArr(i, 10) <> 0 Then
                Select Case bastInfoArr(i - k, 4) '���s
                    Case ""
                        bastInfoArr(i - k, 4) = ascendCustomerDataArr(i, 10) 'NG���i�[
                    Case Is <> ""
                        bastInfoArr(i - k, 4) = bastInfoArr(i - k, 4) & vbLf & ascendCustomerDataArr(i, 10) '���s��NG���i�[
                End Select
            End If
            j = ascendCustomerDataArr(i, 2) '���p�񐔕��A���p�����̓��͗���E�ɂ��炷
            '���p�����i�[ ���e=���t3,���q��7,�z�e��12,�R�X13,����14
            historyInfoArr(i - k, j) = ascendCustomerDataArr(i, 3) & "," & ascendCustomerDataArr(i, 7) & vbLf & ascendCustomerDataArr(i, 12) & "," & ascendCustomerDataArr(i, 13) & "," & ascendCustomerDataArr(i, 14) '�����ł͎��F���̂��߂ɓ��t�͊����ĔNyy���Ȃ�
        Else
            Exit For
        End If
    Next i

    customerSheet.Activate
    
    '�t�B���^�[���N���A
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
    
    '�Z���͈͂�bastInfoArr,historyInfoArr�̑傫���őI�����ď�������
    Range("A3").Resize(UBound(bastInfoArr, 1), UBound(bastInfoArr, 2)) = bastInfoArr
    Range("F3").Resize(UBound(historyInfoArr, 1), UBound(historyInfoArr, 2)) = historyInfoArr
    
    '�s�̍����A��̕��𒲐�
    Range(Cells(1, 1), Cells(UBound(bastInfoArr, 1), UBound(bastInfoArr, 2) - 1)).EntireColumn.AutoFit '�ڋq���l�Ɨ��p��������������ꗗ�̗��������������
    
    '�ڋq�������폜
    Application.DisplayAlerts = False
        ascendSheet.Delete
    Application.DisplayAlerts = True
    
    '�㏑���ۑ�(�Z�[�u)����
    On Error Resume Next
    ActiveWorkbook.Save
    If Err.Number > 0 Then
        MsgBox "�ۑ�����܂���ł���"
    End If

End Sub
