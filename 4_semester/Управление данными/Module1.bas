Attribute VB_Name = "Module1"
Sub СформироватьПолныйОтчет()
    Dim wsОтчеты As Worksheet, wsСведения As Worksheet, wsПремиальные As Worksheet, wsШтрафные As Worksheet
    Dim максПремиальные As Integer, минПремиальные As Integer
    Dim максШтрафные As Integer, минШтрафные As Integer
    Dim i As Integer
    
    
    Set wsОтчеты = Worksheets("Отчеты")
    Set wsСведения = Worksheets("Сведенья о сотрудниках")
    Set wsПремиальные = Worksheets("Премиальные баллы")
    Set wsШтрафные = Worksheets("Штрафные баллы")
    
    
    wsОтчеты.Cells.Clear
    
    
    максПремиальные = Application.WorksheetFunction.Max(wsПремиальные.Range("I3:I22"))
    минПремиальные = Application.WorksheetFunction.Min(wsПремиальные.Range("I3:I22"))
    максШтрафные = Application.WorksheetFunction.Max(wsШтрафные.Range("I3:I22"))
    минШтрафные = Application.WorksheetFunction.Min(wsШтрафные.Range("I3:I22"))
    
    
    wsОтчеты.Range("A1").Value = "Отчет по баллам сотрудников"
    wsОтчеты.Range("A3").Value = "1. Сотрудники с МАКСИМАЛЬНЫМИ премиальными баллами (" & максПремиальные & "):"
    wsОтчеты.Range("A10").Value = "2. Сотрудники с МИНИМАЛЬНЫМИ премиальными баллами (" & минПремиальные & "):"
    wsОтчеты.Range("A17").Value = "3. Сотрудники с МАКСИМАЛЬНЫМИ штрафными баллами (" & максШтрафные & "):"
    wsОтчеты.Range("A24").Value = "4. Сотрудники с МИНИМАЛЬНЫМИ штрафными баллами (" & минШтрафные & "):"
    
    
    Call ДобавитьДанные(wsОтчеты, wsСведения, wsПремиальные, "I", максПремиальные, 4)
    Call ДобавитьДанные(wsОтчеты, wsСведения, wsПремиальные, "I", минПремиальные, 11)
    Call ДобавитьДанные(wsОтчеты, wsСведения, wsШтрафные, "I", максШтрафные, 18)
    Call ДобавитьДанные(wsОтчеты, wsСведения, wsШтрафные, "I", минШтрафные, 25)
    
    
    wsОтчеты.Columns("A:K").AutoFit
    MsgBox "Отчет сформирован!", vbInformation
End Sub

Sub ДобавитьДанные(wsОтчеты As Worksheet, wsСведения As Worksheet, wsБаллы As Worksheet, столбецБаллов As String, значение As Integer, стартСтрока As Integer)
    Dim i As Integer, j As Integer
    Dim found As Boolean
    
    found = False
    For i = 3 To 22
        If wsБаллы.Range(столбецБаллов & i).Value = значение Then
            
            If Not found Then
                For j = 1 To 10
                    wsОтчеты.Cells(стартСтрока, j).Value = wsСведения.Cells(2, j).Value
                Next j
                стартСтрока = стартСтрока + 1
                found = True
