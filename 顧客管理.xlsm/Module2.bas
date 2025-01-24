Attribute VB_Name = "Module2"
Option Explicit
Sub updateCustomerStatisticsButton()
    Call makeAscendingSheet 'Module3
    Call updateCustomerStatistics
End Sub
Sub updateCustomerStatistics()
'�E�E�E�ϐ���`�E�E�E
    Dim ascendSheet As Worksheet '�ڋq�����V�[�g���`
    Set ascendSheet = ActiveWorkbook.Worksheets("�ڋq����")
    Dim customerStatisticsSheet As Worksheet '����ꗗ�V�[�g���`
    Set customerStatisticsSheet = ActiveWorkbook.Worksheets("�����_���v���")
    
    Dim i As Long '�����f�[�^�̍s�����[�vFor�̃J�E���^'
    Dim k As Integer '�񐔕��̍s�����J�E���^�@���p�񐔕��A2��ڈȍ~�̗��p�񐔕��A�����S�f�[�^�ƌڋq���f�[�^�ɍs��������邽�ߕK�v
    Dim baseInfoArr '����ꗗ�̏o�͗p_�z�� '�w�b�_=��,�����X��,�ŏI���X��,�����,�d�b�ԍ�,�}��'
    Dim statisticsInfoArr '����̓��v���i�[�p_�z�� �w�b�_=�}��,�݌v����,�݌v�X��,�P��,�p�x(��/��),��������,�A���P��,�{�w��,�̓���'
    Dim ascendCustomerDataArr '�ڋq�����\�̊i�[�p_�z��
    
    ascendSheet.Activate
    
    Dim customerLastRow As Long
    customerLastRow = Range("C1").CurrentRegion(Range("C1").CurrentRegion.Count).Row '�ڋq�����V�[�g�̍ŏI�s���`
    Dim customerLastColumn As Long
    customerLastColumn = (Cells(1, Columns.Count).End(xlToLeft).Column) '�ڋq�����V�[�g�̍ŏI����`�A1�s�ڃw�b�_����
    ascendCustomerDataArr = Range(Cells(2, 1), Cells(customerLastRow, customerLastColumn)).Value '�ڋq�����V�[�g�̓��e��S�擾
    
    Dim CountRow '�V�K�����J�E���g����baseInfoArr,statisticsInfoArr�̍s�����Ē�`����p�̕ϐ�
    CountRow = WorksheetFunction.CountIf(Range(Cells(2, 2), Cells(customerLastRow, 2)), "1")

'�E�E�E�����E�E�E
    Debug.Print ("---------------Prc")
    
    ReDim baseInfoArr(1 To CountRow, 1 To 6) '�w�b�_=��,�����X��,�ŏI���X��,�����,�d�b�ԍ�,�}��'
    ReDim statisticsInfoArr(1 To CountRow, 1 To 8) '�w�b�_=�݌v����,�݌v�X��,�P��(�X��/��),�p�x(��/��),��������,�ݹ��(�ݹ��/��),�{�w��(�{�w/��),�̓���(�{�w/��)
'
    Dim customerCountTotal As Long '����̂����p��
    Dim qreCt As Long '
    Dim repeatCt As Long
    Dim newCt As Long
    Dim dtToday As Date
    dtToday = Date
    
    '����ꗗ�֏o�͗p�̔z�񐮌`
    k = 0
    For i = LBound(ascendCustomerDataArr, 1) To UBound(ascendCustomerDataArr, 1)  'ascendCustomerDataArr�̍s�������[�v����
        customerCountTotal = WorksheetFunction.CountIf(Range(Cells(2, 9), Cells(customerLastRow, 9)), ascendCustomerDataArr(i, 9))
        
        '���񗘗p���ŏI���p�ł͂Ȃ����̊i�[����
        If ascendCustomerDataArr(i, 2) = 1 And ascendCustomerDataArr(i, 2) <> customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            ''''''''''2��ڈȍ~�̗��p�񐔕������Z���邽�ߏ��񗘗p���̂����ł͏ȗ�
'            k = k + 1  '�񐔕��̍s����

            '�ڋq�ꗗ�i�[
            baseInfoArr(i - k, 1) = customerCountTotal '�񐔊i�[
            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '���񗘗p���i�[
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '�ŏI���p���i�[
            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '�q���i�[
            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '�q�Ԋi�[
            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '�}�̖��i�[
            
            '����������p�񐔓��v�Z
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "�{�w" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "�̓�" Then
                newCt = newCt + 1
            End If
            
            '���v���i�[
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '�݌v����
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '�݌v�X��
            ''''''''''�P���ȍ~�͍ŏI���p���Ɍv�Z���邽�߂����ł͏ȗ�
'            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '�P��
'            statisticsInfoArr(i - k, 4) = "once" '�p�x
'            statisticsInfoArr(i - k, 5) = "once" '��������
'            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal '�A���P��
'            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '�{�w��
'            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '�̓���

'            '�񐔓�������
            ''''''''''����������p�񐔂͍ŏI���p���Ƀ��Z�b�g���邽�߂����ł͏ȗ�
'            qreCt = 0
'            repeatCt = 0
'            newCt = 0

        '���񗘗p���ŏI���p���̌v�Z/�i�[����
        ElseIf ascendCustomerDataArr(i, 2) = 1 And ascendCustomerDataArr(i, 2) = customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            ''''''''''2��ڈȍ~�̗��p�񐔕������Z���邽�ߏ��񗘗p���̂����ł͏ȗ�
'            k = k + 1  '�񐔕��̍s����
                        
            '�ڋq�ꗗ�i�[
            baseInfoArr(i - k, 1) = customerCountTotal '�񐔊i�[
            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '���񗘗p���i�[
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '�ŏI���p���i�[
            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '�q���i�[
            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '�q�Ԋi�[
            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '�}�̖��i�[
            
            '����������p�񐔓��v�Z
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "�{�w" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "�̓�" Then
                newCt = newCt + 1
            End If
            
            '���v���i�[
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '�݌v����
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '�݌v�X��
            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '�P��
            statisticsInfoArr(i - k, 4) = "once" '�p�x
            statisticsInfoArr(i - k, 5) = "once" '��������
            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal '�A���P��
            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '�{�w��
            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '�̓���
            
            '�񐔓�������
            qreCt = 0
            repeatCt = 0
            newCt = 0

        '���p2��ڈȍ~���ŏI���p�łȂ����̌v�Z/�i�[����
        ElseIf ascendCustomerDataArr(i, 2) > 1 And ascendCustomerDataArr(i, 2) <> customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))

            k = k + 1 '�񐔕��̍s����

'            '�ڋq�ꗗ�i�[
            ''''''''�ڋq�ꗗ�f�[�^�͍ŏI���p�����������񗘗p���Ɋi�[���邽�߂����ł͏ȗ�
'            baseInfoArr(i - k, 1) = customerCountTotal '�񐔊i�[
'            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '���񗘗p���i�[
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '�ŏI���p���i�[
'            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '�q���i�[
'            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '�q�Ԋi�[
'            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '�}�̖��i�[

            '����������p�񐔓��v�Z
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "�{�w" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "�̓�" Then
                newCt = newCt + 1
            End If
            
            '���v���i�[
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '�݌v����
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '�݌v�X��
            ''''''''''�P���ȍ~�͍ŏI���p���Ɍv�Z���邽�߂����ł͏ȗ�
'            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '�P��
'            statisticsInfoArr(i - k, 4) = (DateDiff("d", baseInfoArr(i - k, 2), dtToday) + 1) / baseInfoArr(i - k, 1) '�p�x
'            statisticsInfoArr(i - k, 5) = (DateDiff("d", baseInfoArr(i - k, 3), dtToday)) '��������
'            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal '�A���P��
'            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '�{�w��
'            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '�̓���

'            '�񐔓�������
            ''''''''''����������p�񐔂͍ŏI���p���Ƀ��Z�b�g���邽�߂����ł͏ȗ�
'            qreCt = 0
'            repeatCt = 0
'            newCt = 0

        '2��ڈȍ~���ŏI���p���̌v�Z/�i�[����
        ElseIf ascendCustomerDataArr(i, 2) > 1 And ascendCustomerDataArr(i, 2) = customerCountTotal Then
            Debug.Print (ascendCustomerDataArr(i, 2))
            
            k = k + 1  '�񐔕��̍s����
            
'            '�ڋq�ꗗ�i�[
            ''''''''�ڋq�ꗗ�f�[�^�͍ŏI���p�����������񗘗p���Ɋi�[���邽�߂����ł͏ȗ�
'            baseInfoArr(i - k, 1) = customerCountTotal '�񐔊i�[
'            baseInfoArr(i - k, 2) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '���񗘗p���i�[
            baseInfoArr(i - k, 3) = CDate("20" & Left(ascendCustomerDataArr(i, 3), 2) & "/" & Right(Left(ascendCustomerDataArr(i, 3), 4), 2) & "/" & Right(ascendCustomerDataArr(i, 3), 2)) '�ŏI���p���i�[
'            baseInfoArr(i - k, 4) = ascendCustomerDataArr(i, 8) '�q���i�[
'            baseInfoArr(i - k, 5) = ascendCustomerDataArr(i, 9) '�q�Ԋi�[
'            baseInfoArr(i - k, 6) = ascendCustomerDataArr(i, 5) '�}�̖��i�[

            '����������p�񐔓��v�Z
            If ascendCustomerDataArr(i, 21) > 0 Then
                qreCt = qreCt + 1
            End If
            If ascendCustomerDataArr(i, 6) = "�{�w" Then
                repeatCt = repeatCt + 1
            ElseIf ascendCustomerDataArr(i, 6) = "�̓�" Then
                newCt = newCt + 1
            End If
            
            '���v���i�[
            statisticsInfoArr(i - k, 1) = statisticsInfoArr(i - k, 1) + ascendCustomerDataArr(i, 18) '�݌v����
            statisticsInfoArr(i - k, 2) = statisticsInfoArr(i - k, 2) + ascendCustomerDataArr(i, 20) '�݌v�X��
            statisticsInfoArr(i - k, 3) = statisticsInfoArr(i - k, 2) / customerCountTotal '�P��
            statisticsInfoArr(i - k, 4) = (DateDiff("d", baseInfoArr(i - k, 2), dtToday) + 1) / baseInfoArr(i - k, 1) '�p�x
            statisticsInfoArr(i - k, 5) = (DateDiff("d", baseInfoArr(i - k, 3), dtToday)) '��������
            statisticsInfoArr(i - k, 6) = qreCt / customerCountTotal '�A���P��
            statisticsInfoArr(i - k, 7) = repeatCt / customerCountTotal '�{�w��
            statisticsInfoArr(i - k, 8) = newCt / customerCountTotal '�̓���
            
            '�񐔓�������
            qreCt = 0
            repeatCt = 0
            newCt = 0
        Else
            Debug.Print ("Exit For")
            Exit For
        End If
    Next i

    customerStatisticsSheet.Activate
    
    '�Z���͈͂�baseInfoArr,statisticsInfoArr�̑傫���őI�����ď�������
    Range("A3").Resize(UBound(baseInfoArr, 1), UBound(baseInfoArr, 2)) = baseInfoArr
    Range("G3").Resize(UBound(statisticsInfoArr, 1), UBound(statisticsInfoArr, 2)) = statisticsInfoArr
    
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

