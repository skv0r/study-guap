VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12825
   OleObjectBlob   =   "����� ���������������.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
    With ThisWorkbook.Sheets("������������")
        .Range("B3").Value = Worker.Value           ' ���
        .Range("B5").Value = Org.Value              ' �������� �����������
        .Range("B7").Value = Spec.Value             ' �������������
        .Range("E7").Value = Prof.Value             ' ���������
        .Range("B9").Value = Gorod.Value            ' ����� (������, �����)
        .Range("E9").Value = OrgTo.Value            ' ����������� ����������
        .Range("B11").Value = Cel.Value             ' ���� ������������
        .Range("B13").Value = Dney.Value            ' ����������� ���
        .Range("B15").Value = S.Value               ' ���� ������
        .Range("D15").Value = Po.Value              ' ���� ���������
        .Range("C17").Value = Doc.Value             ' �������� ���������
        .Range("F17").Value = DocNumber.Value       ' �����, �����
        .Range("B19").Value = RucD.Value            ' ��������� ������������
        .Range("E19").Value = RucPod.Value          ' ������� (�����������)
    End With

    MsgBox "������ ������� ��������� � ��������������� �������������!", vbInformation
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Worker.Value = "������ ���� ��������"
    Org.Value = "��� ""�����������"""
    Spec.Value = "�������-�����������"
    Prof.Value = "����������� ��"
    Gorod.Value = "������, ������"
    OrgTo.Value = "�� ""��-�������"""
    Cel.Value = "������� � ��������� �� ������� �������"
    Dney.Value = "3"
    S.Value = "12.05.2025"
    Po.Value = "14.05.2025"
    Doc.Value = "������� ���������� ��"
    DocNumber.Value = "1234 �567890"
    RucD.Value = "����������� ��������"
    RucPod.Value = "������ �.�."
End Sub

