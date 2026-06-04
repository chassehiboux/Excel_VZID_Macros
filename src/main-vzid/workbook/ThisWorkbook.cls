Option Explicit

Private Sub Workbook_Open()
    On Error GoTo done
    MainEntryPoint_Initialize
done:
    ' Ошибка инициализации не должна блокировать Excel.
End Sub
