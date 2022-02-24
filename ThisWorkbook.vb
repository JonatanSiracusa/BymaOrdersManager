'************************************************************************************************************************
'************************************************************************************************************************
'************************* THIS MODULE IS USED FOR PROCEDURES EXECUTED AT OPENING THE FILE ******************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' LinkedIn: https://www.linkedin.com/in/ajsiracusa/
' GitHub: https://github.com/JonatanSiracusa
' Instagram: @JonaSiracusa
'
'

Private Sub Workbook_Open()
Dim contador, sistema As Integer
Dim libro, hoja, Hoja2 As String
Dim ahora As Date

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Hoja4.Cells(2, 46) = ThisWorkbook.Name
Hoja4.Cells(2, 47) = Hoja4.Name


hoja = "log"
Hoja2 = "Datos"

'******** insertar el nombre del libro en una hoja, y que de ahi lo tome el resto de las variables.
Worksheets(Hoja2).Cells(7, 3) = ThisWorkbook.Name
libro = Worksheets(Hoja2).Cells(7, 3)


Workbooks(libro).Worksheets(hoja).Cells(1, 16) = Now()
ahora = Workbooks(libro).Worksheets(hoja).Cells(1, 16)
sistema = Workbooks(libro).Worksheets(Hoja2).Cells(9, 14)


Application.OnKey "^{Enter}", "asignarCargarOrdenes"
Application.OnKey "^ ", "asignarCancelarOrdenes"

Call insertarNombreHoja
Call asignarIngresarSistema


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

End Sub
