VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "���q���\��"
   ClientHeight    =   9090.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17730
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�\���p������
Public dateJpDisp As String
Public castDisp As String
Public contentsDisp As String
Public IncomeAndQre As String

'���͓��t
Public ansDateIs As Integer

'������A�X�����A�A���P�v�Z
Public totalSales As Long '1���̑�����z
Public totalPay As Long '1���̑����q���z
Public totalIncome As Long '1���̓X�����z
Public totalQre As Long '1���̃A���P���z

Sub UserForm_Initialize()
    
    castBackOnDay
    
    With Worksheets("���̓V�[�g")
        Label5.Caption = dateJpDisp
        Label6.Caption = IncomeAndQre
        Label1.Caption = castDisp
        Label2.Caption = contentsDisp
    End With
    
End Sub
Sub castBackOnDay()
    Dim inputSheet As Worksheet
    Set inputSheet = ActiveWorkbook.Worksheets("���̓V�[�g")

'//////////////////////////////////
'//���t���͗p�_�C�A���O��\������
'//////////////////////////////////
    Dim ansDate As String
    Do While Not (ansDate Like "######" Or ansDate Like "#")
        ansDate = InputBox("�Ώۂ̓��t����͂��Ă��������B" & vbCr & vbCr & _
        "��������[0]����͂��Ă�OK�ł�!!" & vbCr & vbCr & _
        "[0]������" & vbCr & _
        "[1]������i1���O�j" & vbCr & _
        "[2]�������i2���O�j" & vbCr & _
        "�̂悤�ɏo�͂���܂��B�i1�`9����͎��j", "���t����", "(��:2090�N3��17���Ȃ�[900317]�Ɠ���)")
        
        ansDateIs = 1
        If StrPtr(ansDate) = 0 Then
            ansDateIs = 0
            Exit Sub
        End If
    Loop
    
'///////////////////////////////////
'//�Y��������t�̍s����肵�Ď擾
'///////////////////////////////////
    '�V�[�g�̕ی����
    ActiveSheet.Unprotect Password:="042595"
    
    '�V�[�g�̍ŏI�s���`�A�擾�͈�
    Dim dateColumnArr() As Variant
    Dim dateLastRow As Long
    dateLastRow = Range("C1").CurrentRegion(Range("C1").CurrentRegion.Count).Row
    dateColumnArr = inputSheet.Range(Cells(1, 3), Cells(dateLastRow, 3))
    
    Dim objDate As Variant
    
    '���t���͂�"#"�������ꍇ�ɓ������t�ɕϊ�����
    If ansDate Like "#" Then
        objDate = Val(Right(Left(Date - ansDate, 4), 2) & Right(Left(Date - ansDate, 7), 2) & Right(Date - ansDate, 2))
    Else
        objDate = Val(ansDate)
    End If
    
    '�w�肵���������s�ڂɑ��݂��邩�����
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

    '�������R�[�h���Ȃ���ΏI��
    If dateRows(1) = "" Then
        MsgBox ("���̓���0�{�݂����ł��B" & vbCrLf & vbCrLf & _
        "�E���̓V�[�g�ɖ{���̎�t��������͂��Ă��Ȃ������c�H" & vbCrLf & _
        "�E��t�����A�܂��͍������ɓ��͂������t���Ԉ���Ă��邩���c�H" & vbCrLf & vbCrLf & _
        "by_Usui")
        Exit Sub
    End If

    Dim objRowL As Long
    Dim objRowU As Long
    
    objRowL = dateRows(LBound(dateRows) + 1)
    objRowU = dateRows(UBound(dateRows) - 1)
    
    '����̓��̎�t�f�[�^��z��Ƃ��Ď擾����
    Dim targetArr() As Variant '��t�[�̔z��̊i�[��
    targetArr = inputSheet.Range(Cells(objRowL, 3), Cells(objRowU, 22)).Value '����̍��W��z��ɂ���

    '���q���X�g�쐬(�d������)
    Dim tempCastList As Variant
    ReDim tempCastList(1)
    
    i = 0
    For i = LBound(targetArr, 1) To UBound(targetArr, 1)
        tempCastList(i) = targetArr(i, 5)
        ReDim Preserve tempCastList(UBound(tempCastList) + 1)
    Next
    ReDim Preserve tempCastList(UBound(tempCastList) - 1)
    
    '���q���X�g�d���폜����
    Dim castList As Variant
    ReDim castList(1)
    
    castList = Call_Array_Dictionary(tempCastList)
    
    'cast���ɖ{�����i�[���邽�ߏ��q���X�g��2�����z��
    Dim castArr As Variant
    ReDim castArr(UBound(castList), 2)
    Dim workCt As Integer '�{��

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

    '���q���̍��v���q���̌v�Z�����A���������������i�[
    Dim castPayArr As Variant '���q���̋��^���v�z
    ReDim castPayArr(UBound(castList))
    Dim castMinArr As Variant '����(����)
    ReDim castMinArr(UBound(castList))
    
    i = 0
    j = 0
    For i = LBound(castList, 1) To UBound(castList, 1)
        workCt = 0
        For j = LBound(targetArr, 1) To UBound(targetArr, 1)
            If targetArr(j, 5) = castList(i) Then
                workCt = workCt + 1
                '���^���v�z���X�g�ւ̊i�[����
                castPayArr(i) = castPayArr(i) + targetArr(j, 17)

                '����(����)���X�g�ւ̊i�[����
                '�Ō�̈�{�͕����̃J���}(,)�s�v
                '�{�w��"�{�wXX��"�\��
                If workCt <> castArr(i, 2) Then
                    If targetArr(j, 4) = "�{�w" Then
                        castMinArr(i) = castMinArr(i) & targetArr(j, 4) & targetArr(j, 12) & "��, "
                    Else
                        castMinArr(i) = castMinArr(i) & targetArr(j, 12) & "��, "
                    End If
                Else
                    If targetArr(j, 4) = "�{�w" Then
                        castMinArr(i) = castMinArr(i) & targetArr(j, 4) & targetArr(j, 12) & "��"
                    Else
                        castMinArr(i) = castMinArr(i) & targetArr(j, 12) & "��"
                    End If
                End If
            End If
        Next
    Next

    '���q���\���p�̕����񐶐�����
    Dim slashSeparatedDate As Variant
    Dim dateInJp As String
    Dim castNameStr As String
    Dim perCastStr As String

    '���t������(Date�֏o��)
    slashSeparatedDate = "20" & Left(objDate, 2) & "/" & Right(Left(objDate, 4), 2) & "/" & Right(objDate, 2)
    dateInJp = "20" & Left(objDate, 2) & "�N" & Right(Left(objDate, 4), 2) & "��" & Right(objDate, 2) & "�� (" & WeekdayName(Weekday(slashSeparatedDate), True) & ")"

    '���q��������(Name�֏o��)
    i = 0
    For i = LBound(castList, 1) + 1 To UBound(castList, 1)
        castNameStr = castNameStr & "�E" & castList(i) & vbLf
    Next

    '���e������(Contents�֏o��)
    i = 0
    For i = LBound(castList, 1) + 1 To UBound(castList, 1)
        perCastStr = perCastStr & "�c�@" & castPayArr(i) & "�~�@(" & castMinArr(i) & ")" & vbLf
    Next

    'MsgBox�ŕ\��(�p�~)
'    MsgBox perCastStr

    '���[�U�[�t�H�[���ŕ\������p
    '�p�u���b�N�ϐ��֑������
    dateJpDisp = dateInJp
    castDisp = castNameStr
    contentsDisp = perCastStr
    IncomeAndQre = " ������ " & totalSales & "�~,�@���q�����z " & totalPay & "�~,�@�X�����v " & totalIncome & "�~,�@" & "�A���P�v " & totalQre & "�~"
    
    '�V�[�g�̕ی�
    ActiveSheet.Protect Password:="042595", AllowFiltering:=True
    
End Sub

'��Dictionary(�A�z�z��)�ŏd���f�[�^���폜����
'���K�v�ȑO������
'�c�[��>�Q�Ɛݒ�>Microsoft scripting runtime�Ƀ`�F�b�N�����
Public Function Call_Array_Dictionary(arr As Variant)
    Dim dic As New Dictionary
    Dim i As Long

    '��On Error Resume Next�ŃG���[�𖳎�����
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
    
    ansCastPay = InputBox("���Z�ς݂̏��q���z����͂��Ă��������B", "���Z�ςݏ��q���z", 0)
    If StrPtr(ansCastPay) = 0 Then
        Exit Sub
    End If
    
    ansMEx = InputBox("���Z�ς݂̎G��z����͂��Ă��������B", "���Z�ςݎG��z", 0)
    If StrPtr(ansMEx) = 0 Then
        Exit Sub
    End If
    
    ansDiaryBonus = InputBox("���Z�ς݂̓��L�{�[�i�X�z����͂��Ă��������B", "���Z�ςݓ��L�{�[�i�X�z", 0)
    If StrPtr(ansDiaryBonus) = 0 Then
        Exit Sub
    End If

    ansCardIncome = InputBox("�{���̃J�[�h����z(�J�[�h�萔�����������ō��z)����͂��Ă��������B", "�J�[�h����z", 0)
    If StrPtr(ansCardIncome) = 0 Then
        Exit Sub
    End If
    
    MsgBox ("���݂̌������z�́A" & totalSales - totalQre - ansCastPay + ansMEx - ansDiaryBonus - ansCardIncome & "�~�ł��B" & vbCrLf & vbCrLf & _
        "")
End Sub


