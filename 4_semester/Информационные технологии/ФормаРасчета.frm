VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9660.001
   OleObjectBlob   =   "ФормаРасчета.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCalculate_Click()
    Dim transportCost As Long
    Dim hotelCostPerDay As Long
    Dim days As Long
    Dim extraCost As Long
    Dim total As Long


    If OptPlane.Value Then
        transportCost = 5000
    ElseIf OptTrain.Value Then
        transportCost = 2500
    ElseIf OptBus.Value Then
        transportCost = 1500
    End If

    If optHotel3.Value Then
        hotelCostPerDay = 1000
    ElseIf optHotel4.Value Then
        hotelCostPerDay = 2000
    ElseIf optHotel5.Value Then
        hotelCostPerDay = 3500
    End If


    If optDays3.Value Then
        days = 3
    ElseIf optDays5.Value Then
        days = 5
    ElseIf optDays7.Value Then
        days = 7
    End If


    extraCost = 0
    If chkExcursion.Value Then extraCost = extraCost + 1000
    If chkTheater.Value Then extraCost = extraCost + 1500
    If chkMuseum.Value Then extraCost = extraCost + 800


    total = transportCost + hotelCostPerDay * days + extraCost

    lblResult.Caption = "Итоговая стоимость: " & total & " рублей."
End Sub

Private Sub Label1_Click()

End Sub
