Attribute VB_Name = "Module3"
Option Explicit
Sub makeAscendingSheet()
' �����ƃV�[�g���� Macro

    Dim ws As Worksheet

    '�ڋq�����V�[�g������΍폜����i�V�[�g�쐬���G���[����̂��߁j
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name = "�ڋq����" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
        
    '�ڋq�����V�[�g���쐬�i�V�[�g�ǉ����āA���̓V�[�g����ہX�R�s�y����j
    Dim inputSheet As Worksheet
    Set inputSheet = Sheets("���̓V�[�g")
    Dim ascendingSheet As Worksheet
    Set ascendingSheet = Worksheets.Add ' �V�����V�[�g��ǉ����ĕϐ��ɑ��
    ascendingSheet.Name = "�ڋq����" ' �V�[�g����ݒ�
    
    '�ŏI�s�A�ŏI������
    Dim inputDateLastRow As Long
    inputDateLastRow = inputSheet.Cells(Rows.Count, 3).End(xlUp).Row '���̓V�[�g�̍ŏI�s���` '���̓V�[�g�A���q�l����̍ŏI�s���擾
    Dim inputSheetLastColumn As Long
    inputSheetLastColumn = inputSheet.Cells(1, Columns.Count).End(xlToLeft).Column '���̓V�[�g�A1�s�ڃw�b�_�̍ŏI����擾
    
    '���̓V�[�g�̓��e���R�s�[
    inputSheet.Activate
    inputSheet.Range(Cells(1, 1), Cells(inputDateLastRow, inputSheetLastColumn)).Copy
    
    '�ڋq�����V�[�g�֒l�̂݃y�[�X�g
    ascendingSheet.Activate
    ascendingSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    '���ёւ������i�ڋq���ŏ����j
    ascendingSheet.Sort.SortFields.Clear
    ascendingSheet.Sort.SortFields.Add2 _
        Key:=Range("H1"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    ascendingSheet.Sort.SortFields.Add2 _
        Key:=Range("I1"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ascendingSheet.Sort
        .SetRange Range("A1", Cells(inputDateLastRow, inputSheetLastColumn))
        .Header = xlYes
        .Apply
    End With
    
End Sub
Function inputSheetFalseCheck()
''''''''''''''''''''''''''''''''''''''''
''����Function�͌ڋq�Ǘ��u�b�N�̓��̓V�[�g��p�̃v���V�[�W���ł�
''False�����邩�Ȃ�����check���܂�
''Call����O�ɓ��̓V�[�g�̃A�N�e�B�x�[�g���K�{�ł�
''��:inputSheet.Activate
''''''''''''''''''''''''''''''''''''''''
    Dim inputSheet As Worksheet
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")
    Dim booleanArr As Variant
    Dim falseRowList As Variant
    Dim falseRowStr As String
    ReDim falseRowList(0)
    Dim i As Long
    Dim j As Long
    
    inputSheet.Activate
    
    '�V�[�g�̕ی����
    ActiveSheet.Unprotect Password:="042595"
    
    booleanArr = Range("X2:AA9999").Value
    
    For i = LBound(booleanArr, 1) To UBound(booleanArr, 1)
        For j = LBound(booleanArr, 2) To UBound(booleanArr, 2)
            If booleanArr(i, j) = "FALSE" Then
                falseRowList(UBound(falseRowList)) = i + 1
                ReDim Preserve falseRowList(UBound(falseRowList) + 1)
                Exit For
            End If
        Next
    Next
    
    '�V�[�g�̕ی�
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
    '������
    If UBound(falseRowList) = 1 Then
        falseRowStr = falseRowList(0)
    Else
        For i = LBound(falseRowList) To UBound(falseRowList) - 1
            If i <> UBound(falseRowList) - 1 Then
                falseRowStr = falseRowStr & falseRowList(i) & ","
            Else
                falseRowStr = falseRowStr & falseRowList(i)
            End If
        Next
    End If
    
    If UBound(falseRowList) > 0 Then
    
        MsgBox "����͂����邩������܂���B" & vbCrLf & _
        falseRowStr & "�s�ڂ��m�F���Ă��������B" & vbCrLf & _
        "���̖�肪��������܂Œ��ߏ�����ڋq���̍X�V�͍s���܂���B" & vbCrLf & _
        "�����ł��Ȃ��ꍇ�̓V�X�e���Ǘ��҂ɂ��₢���킹���������B", vbExclamation
        
        '�㏑���ۑ�(�Z�[�u)����
        On Error Resume Next
        ActiveWorkbook.Save
        If Err.Number > 0 Then
            MsgBox "�ۑ�����܂���ł���"
        End If
        inputSheetFalseCheck = 1
    Else
        inputSheetFalseCheck = 0
    End If
    
End Function

