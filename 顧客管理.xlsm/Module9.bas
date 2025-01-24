Attribute VB_Name = "Module9"
Option Explicit
'�v�Z��z��̍��W��`�iExcel��̍��W�Ƃ͈قȂ�BExcel��F���0��ڂƂ��A�������ݎ���F3���J�n�ʒu�Ƃ���j
''�e���v
Const edCountClm = 2 '�{���v
Const edRenominationClm = 3 '�{�w����
Const edSalesClm = 4 '����v
Const edPayClm = 5 '���q���v
Const edIncomeClm = 6 '�X���v
Const edNewClm = 7 '�V�K�v
Const edRepeaterClm = 8 '����v
Const ed2ndCtClm = 10 '2��ڌv
Const edSBvalueClm = 13 'SB�z�v
''�}�̐�
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
''�����A���ϐ�
Const edAvgCountClm = 0 '��/��
Const edAvgIncomeClm = 1 '��/��
Const edRePercentClm = 9
Const ed2ndPercentClm = 11
Const edIncomePerCtClm = 12

'Cast�o�ΐ��̍��W��`
Const edYearClm = 1
Const edMonthClm = 2
Const edMonthlyLastDateClm = 3
Const edCountPerCastClm = 4
Const edCastCountClm = 5

'�I�[�v���̔N���i�s�ς̒l�j
Const startYY = 21 '�c�ƊJ�n�N
Const startMM = 10 '�c�ƊJ�n��-1
Sub monthlyStatisticsButton()
    Call updateMonthlyStatistics
End Sub
Sub updateMonthlyStatistics()
'�E�E�E�ϐ���`�E�E�E
    Dim inputSheet As Worksheet '���̓V�[�g���`
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")
    Dim monthlyStatisticsSheet As Worksheet '���ʃV�[�g���`
    Set monthlyStatisticsSheet = ActiveWorkbook.Worksheets("����_���v���")
    
    
    '''''''''''''''''''''''''''''
    '�e�V�[�g�̍ŏI�s�܂ł̓��e���擾����B
    '''''''''''''''''''''''''''''
    '���̓V�[�g�̔z���`
    inputSheet.Activate
    Dim inputDateLastRow As Long
    inputDateLastRow = Cells(Rows.Count, 3).End(xlUp).Row '���̓V�[�g�A���q�l����̍ŏI�s���擾
    Dim inputSheetLastColumn As Long
    inputSheetLastColumn = Cells(1, Columns.Count).End(xlToLeft).Column '���̓V�[�g�A1�s�ڃw�b�_�̍ŏI����擾
    Dim inputArray As Variant
    inputArray = Range(Cells(2, 1), Cells(inputDateLastRow, inputSheetLastColumn)).Value '���̓V�[�g�̓��e��z��Ƃ��Ď擾

    '���ʓ��v�f�[�^�̔z��̑傫�����`����(Cast�o�ΐ������ޗ������)
    monthlyStatisticsSheet.Activate
    Dim calculatedArray As Variant '�����o���p�̔z��
    Dim lastDate As Long '�ŏI��t��
    lastDate = inputArray(inputDateLastRow - 1, 3)
    Dim lastYY As Integer '���̓V�[�g�̑Ώۃ��R�[�h�̔N
    lastYY = Val(Left(lastDate, 2))
    Dim lastMM As Integer '���̓V�[�g�̑Ώۃ��R�[�h�̌�
    lastMM = Val(Right(Left(lastDate, 4), 2))
    Dim calculatedArrayLastRow As Long
    calculatedArrayLastRow = ((lastYY - startYY) * 12) + (lastMM - startMM) - 1
    Dim calculatedArrayLastColumn As Long
    calculatedArrayLastColumn = (Cells(2, Columns.Count).End(xlToLeft).Column) - 6 '�G�N�Z����̍ŏI�񂩂�N�A���A�ŏI���A�{�o���V�I�A���o��5�񕪈����A0�n�܂�̂���1�����A�v6����
    ReDim calculatedArray(calculatedArrayLastRow, calculatedArrayLastColumn)
    
    'Cast�o�ΐ������ޗ���i�[����z����i�[����
    Dim castCountArr As Variant
    castCountArr = Range(Cells(3, 1), Cells(calculatedArrayLastRow + 3, 5)) '�J�n�s��3�s�ڂ���Ȃ̂ŊJ�n�ƍŏI��+3
    
    '��{�f�[�^�쐬�i�{���v�Z�A������v�Z�A���q���v�Z�A�X���v�Z�A�V�K���v�Z�A������v�Z�A�e�}�̖{���A2��ڐ��j
    calculatedArray = monthlyFoundationData(inputArray, calculatedArray)
    
    '�v�Z�f�[�^�쐬�i�����A���ϓ��j
    calculatedArray = monthlyRatioAverage(calculatedArray, lastDate) 'lastDate�͍ŐV���̕��ς�����o�����߂ɕK�v
    
    'cast�o�ΐ����ώ擾�A�{��/�o�ΐ��̌v�Z����
    castCountArr = monthlyCastCount(castCountArr)
    
    '�������ݏ���
    monthlyStatisticsSheet.Activate
    
    '�V�[�g�̕ی����
    ActiveSheet.Unprotect Password:="042595"
    
    Range(Cells(3, 6), Cells(calculatedArrayLastRow + 3, calculatedArrayLastColumn + 4)) = calculatedArray '���ʓ��v�f�[�^�̔z��(Cast�o�ΐ������ޗ������)����������
    Range(Cells(3, 1), Cells(calculatedArrayLastRow + 3, 5)) = castCountArr 'Cast�o�ΐ������ޗ����������
    
    '�V�[�g�̕ی�
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
    monthlyStatisticsSheet.Activate

End Sub
Function monthlyFoundationData(ByVal inputArray As Variant, ByVal calculatedArray As Variant)
    Dim i As Long
    Dim j As Long
    
    Dim targetYY As Integer '���̓V�[�g�̑Ώۃ��R�[�h�̔N
    Dim targetMM As Integer '���̓V�[�g�̑Ώۃ��R�[�h�̌�
    Dim monthRow As Long
        
    '�{���v�Z�A������v�Z�A���q���v�Z�A�X���v�Z�A�V�K���v�Z�A������v�Z�A�e�}�̖{���A2��ڐ�
    For i = LBound(inputArray, 1) To UBound(inputArray, 1)
        targetYY = Left(inputArray(i, 3), 2)
        targetMM = Right(Left(inputArray(i, 3), 4), 2)
        monthRow = ((targetYY - startYY) * 12) + (targetMM - startMM) - 1 'calculatedArray��0�n�܂�̔z��ł��邽�߁ARow -1����B
        
        '�{�� ����̑Ώۍs��������+1���Z����Ă���
        calculatedArray(monthRow, edCountClm) = calculatedArray(monthRow, edCountClm) + 1
        '������
        calculatedArray(monthRow, edSalesClm) = calculatedArray(monthRow, edSalesClm) + inputArray(i, 18)
        '���q��
        calculatedArray(monthRow, edPayClm) = calculatedArray(monthRow, edPayClm) + inputArray(i, 19)
        '�X��
        calculatedArray(monthRow, edIncomeClm) = calculatedArray(monthRow, edIncomeClm) + inputArray(i, 20)
        
        If inputArray(i, 6) = "�{�w" Then
        '�{�w����
            calculatedArray(monthRow, edRenominationClm) = calculatedArray(monthRow, edRenominationClm) + 1
        End If
        If inputArray(i, 5) = "R" Then
        '�����
            calculatedArray(monthRow, edRepeaterClm) = calculatedArray(monthRow, edRepeaterClm) + 1
        Else
        '�V�K��
            calculatedArray(monthRow, edNewClm) = calculatedArray(monthRow, edNewClm) + 1
            '�e�}�̐�
            Select Case inputArray(i, 5)
            Case "��"
                calculatedArray(monthRow, edZenraClm) = calculatedArray(monthRow, edZenraClm) + 1
            Case "�w�u��"
                calculatedArray(monthRow, edHeavenClm) = calculatedArray(monthRow, edHeavenClm) + 1
            Case "����"
                calculatedArray(monthRow, edKyokuClm) = calculatedArray(monthRow, edKyokuClm) + 1
            Case "�����W���p��"
                calculatedArray(monthRow, edJapanClm) = calculatedArray(monthRow, edJapanClm) + 1
            Case "DX"
                calculatedArray(monthRow, edDXClm) = calculatedArray(monthRow, edDXClm) + 1
            Case "�w����"
                calculatedArray(monthRow, edEkiClm) = calculatedArray(monthRow, edEkiClm) + 1
            Case "�҂゠���"
                calculatedArray(monthRow, edPureClm) = calculatedArray(monthRow, edPureClm) + 1
            Case "�q���`����"
                calculatedArray(monthRow, edHimechClm) = calculatedArray(monthRow, edHimechClm) + 1
            Case "�O�[�O��"
                calculatedArray(monthRow, edGoogleClm) = calculatedArray(monthRow, edGoogleClm) + 1
            Case "HP"
                calculatedArray(monthRow, edHPClm) = calculatedArray(monthRow, edHPClm) + 1
            Case "���̑�"
                calculatedArray(monthRow, edOtherClm) = calculatedArray(monthRow, edOtherClm) + 1
            Case "�r��"
                calculatedArray(monthRow, edBillClm) = calculatedArray(monthRow, edBillClm) + 1
            Case "T-1"
                calculatedArray(monthRow, edT1Clm) = calculatedArray(monthRow, edT1Clm) + 1
            End Select
        End If
        '2��ږ{��
        If inputArray(i, 2) = 2 Then
            calculatedArray(monthRow, ed2ndCtClm) = calculatedArray(monthRow, ed2ndCtClm) + 1
        End If
        
        'SB�v
        calculatedArray(monthRow, edSBvalueClm) = calculatedArray(monthRow, edSBvalueClm) + (inputArray(i, 19) * (inputArray(i, 22) / 100))
        
    Next i
    
    monthlyFoundationData = calculatedArray
    
End Function
Function monthlyRatioAverage(ByVal calculatedArray As Variant, ByVal lastDate As String)

    Dim i As Long
    Dim j As Long
    
    Dim monthlyLastDate As Long
    
    '�����v�Z�A���όv�Z
    For i = LBound(calculatedArray, 1) To UBound(calculatedArray, 1)
        '�[���p�f�B���O
        For j = LBound(calculatedArray, 2) To UBound(calculatedArray, 2)
            If calculatedArray(i, j) = "" Then
                calculatedArray(i, j) = 0
            End If
        Next j
        monthlyLastDate = Cells(i + 3, 3).Value
        '''�A�x���[�W�{���A�A�x���[�W�X�����v�Z
        '��������
        If i = UBound(calculatedArray, 1) Then
        '�������̓��t���擾���ď��Z���� (���:���ς�蒼��̐��x������)
'            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / Format(Date, "dd")
'            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / Format(Date, "dd")
        '���ŏI���͓��̓��t�ŏ��Z���� (���:����0=�{��0�̓��̒���̐��x�������H�����p�x���Ⴂ�̂ŁA�b��I�ɂ�������̗p)
            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / Right(lastDate, 2)
            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / Right(lastDate, 2)

        '�ߋ�������
        Else
            calculatedArray(i, edAvgCountClm) = calculatedArray(i, edCountClm) / monthlyLastDate
            calculatedArray(i, edAvgIncomeClm) = calculatedArray(i, edIncomeClm) / monthlyLastDate
        End If
        '''�����
        calculatedArray(i, edRePercentClm) = calculatedArray(i, edRepeaterClm) / (calculatedArray(i, edRepeaterClm) + calculatedArray(i, edNewClm))
        '''2��ڗ�
        calculatedArray(i, ed2ndPercentClm) = calculatedArray(i, ed2ndCtClm) / calculatedArray(i, edNewClm)
        '''�P���v�Z
        calculatedArray(i, edIncomePerCtClm) = calculatedArray(i, edIncomeClm) / calculatedArray(i, edCountClm)
    Next i
    
    monthlyRatioAverage = calculatedArray
    
End Function
Function monthlyCastCount(ByVal castCountArr As Variant)
    Dim i As Long
    Dim targetyyyymm As String '�ړI�̌��̃t�@�C�������w�肷�邽�߂̕�������i�[����
    Dim targetyyyymmArr As Variant
    Dim managementBook As Workbook '�Ǘ��\�u�b�N�I�u�W�F�N�g
    Dim managementBookPath As Variant '�Ǘ��\�u�b�N�̃p�X
    Dim managementSheet1 As Worksheet
    
    '�e�s�̎擾�󋵂𔻒肵�āA���擾�Ȃ�擾����B�ŐV�s�Ȃ�V���Ɏ擾����B
    For i = LBound(castCountArr, 1) To UBound(castCountArr, 1)
        '�擾�ł��Ă��Ȃ����𔻒肷��
        If castCountArr(i, 5) = Empty Or i = UBound(castCountArr, 1) Then
            targetyyyymm = "20" & Format(castCountArr(i, edYearClm), "00") & Format(castCountArr(i, edMonthClm), "00")
            
            '���̌��̊Ǘ��\�u�b�N���J���Ĥ�L���X�g�o�ΐ����擾����
            '�{�Ԋ��p
'           managementBookPath = "E:\�Ǘ��\\�Ǘ��\" & targetyyyymm & ".xlsx"
            '�e�X�g���p
            managementBookPath = "D:\usb_20241230\�Ǘ��\\�Ǘ��\" & targetyyyymm & ".xlsx"
            If Dir(managementBookPath) <> "" Then
                Set managementBook = Workbooks.Open(managementBookPath)
                Set managementSheet1 = managementBook.Worksheets("Z��")
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
