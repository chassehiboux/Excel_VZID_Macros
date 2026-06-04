Attribute VB_Name = "CommandStubs"
Option Explicit

Public Sub Doc_Packets()
    MsgBox "Каркас команды 'Пакеты' работает." & vbCrLf & _
           "Дальше сюда будет импортирована реальная логика.", vbInformation, "VZID"
End Sub

Public Sub Make_Cover_Letters()
    MsgBox "Каркас команды 'Сопроводительные' для Кургана работает.", vbInformation, "VZID"
End Sub

Public Sub Make_Cover_Letters_ByRecipient()
    MsgBox "Каркас команды 'Сопроводительные по получателю' для Кургана работает.", vbInformation, "VZID"
End Sub

Public Sub TMN_Make_Cover_Letters_likeKGN()
    MsgBox "Каркас команды 'Сопроводительные' для Тюмени работает.", vbInformation, "VZID"
End Sub

Public Sub TMN_Make_Cover_Letters_ByRecipient_likeKGN()
    MsgBox "Каркас команды 'Сопроводительные по получателю' для Тюмени работает.", vbInformation, "VZID"
End Sub

Public Sub RIC_Zakaznye_Create()
    MsgBox "Каркас команды 'Заказные' для Екатеринбурга работает.", vbInformation, "VZID"
End Sub
