Attribute VB_Name = "Module1"
Sub �����������������������()
    Dim ws������ As Worksheet, ws�������� As Worksheet, ws����������� As Worksheet, ws�������� As Worksheet
    Dim ��������������� As Integer, �������������� As Integer
    Dim ������������ As Integer, ����������� As Integer
    Dim i As Integer
    
    
    Set ws������ = Worksheets("������")
    Set ws�������� = Worksheets("�������� � �����������")
    Set ws����������� = Worksheets("����������� �����")
    Set ws�������� = Worksheets("�������� �����")
    
    
    ws������.Cells.Clear
    
    
    ��������������� = Application.WorksheetFunction.Max(ws�����������.Range("I3:I22"))
    �������������� = Application.WorksheetFunction.Min(ws�����������.Range("I3:I22"))
    ������������ = Application.WorksheetFunction.Max(ws��������.Range("I3:I22"))
    ����������� = Application.WorksheetFunction.Min(ws��������.Range("I3:I22"))
    
    
    ws������.Range("A1").Value = "����� �� ������ �����������"
    ws������.Range("A3").Value = "1. ���������� � ������������� ������������ ������� (" & ��������������� & "):"
    ws������.Range("A10").Value = "2. ���������� � ������������ ������������ ������� (" & �������������� & "):"
    ws������.Range("A17").Value = "3. ���������� � ������������� ��������� ������� (" & ������������ & "):"
    ws������.Range("A24").Value = "4. ���������� � ������������ ��������� ������� (" & ����������� & "):"
    
    
    Call ��������������(ws������, ws��������, ws�����������, "I", ���������������, 4)
    Call ��������������(ws������, ws��������, ws�����������, "I", ��������������, 11)
    Call ��������������(ws������, ws��������, ws��������, "I", ������������, 18)
    Call ��������������(ws������, ws��������, ws��������, "I", �����������, 25)
    
    
    ws������.Columns("A:K").AutoFit
    MsgBox "����� �����������!", vbInformation
End Sub

Sub ��������������(ws������ As Worksheet, ws�������� As Worksheet, ws����� As Worksheet, ������������� As String, �������� As Integer, ����������� As Integer)
    Dim i As Integer, j As Integer
    Dim found As Boolean
    
    found = False
    For i = 3 To 22
        If ws�����.Range(������������� & i).Value = �������� Then
            
            If Not found Then
                For j = 1 To 10
                    ws������.Cells(�����������, j).Value = ws��������.Cells(2, j).Value
                Next j
                ����������� = ����������� + 1
                found = True
