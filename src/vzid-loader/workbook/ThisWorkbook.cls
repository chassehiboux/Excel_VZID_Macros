Option Explicit

Private Sub Workbook_Open()
    On Error GoTo done
    LoaderEntryPoint_HandleWorkbookOpen
done:
    ' Ошибка загрузчика не должна ломать запуск Excel.
End Sub
