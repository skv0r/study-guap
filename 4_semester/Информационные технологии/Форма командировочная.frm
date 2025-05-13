VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12825
   OleObjectBlob   =   "Форма командировочная.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
    With ThisWorkbook.Sheets("Командировка")
        .Range("B3").Value = Worker.Value           ' ФИО
        .Range("B5").Value = Org.Value              ' Название организации
        .Range("B7").Value = Spec.Value             ' Специальность
        .Range("E7").Value = Prof.Value             ' Профессия
        .Range("B9").Value = Gorod.Value            ' Город (страна, город)
        .Range("E9").Value = OrgTo.Value            ' Организация назначения
        .Range("B11").Value = Cel.Value             ' Цель командировки
        .Range("B13").Value = Dney.Value            ' Календарные дни
        .Range("B15").Value = S.Value               ' Дата начала
        .Range("D15").Value = Po.Value              ' Дата окончания
        .Range("C17").Value = Doc.Value             ' Название документа
        .Range("F17").Value = DocNumber.Value       ' Номер, серия
        .Range("B19").Value = RucD.Value            ' Должность руководителя
        .Range("E19").Value = RucPod.Value          ' Подпись (расшифровка)
    End With

    MsgBox "Данные успешно сохранены в командировочное удостоверение!", vbInformation
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Worker.Value = "Иванов Иван Иванович"
    Org.Value = "ООО ""АльфаПроект"""
    Spec.Value = "Инженер-программист"
    Prof.Value = "Разработчик ПО"
    Gorod.Value = "Казань, Россия"
    OrgTo.Value = "АО ""ИТ-Системы"""
    Cel.Value = "Участие в совещании по запуску проекта"
    Dney.Value = "3"
    S.Value = "12.05.2025"
    Po.Value = "14.05.2025"
    Doc.Value = "Паспорт гражданина РФ"
    DocNumber.Value = "1234 №567890"
    RucD.Value = "Генеральный директор"
    RucPod.Value = "Петров С.С."
End Sub

