'************************************************************************************************************************
'************************************************************************************************************************
'********* THIS MODULE REDIRECTS TO THE DIFFERENT MODULES SYSTEMS: RODI, GALLO1, GALLO2 OR MODULES USING PYTHON *********
'************************************ THE SYSTEM USED IS CHOSEN BY THE USER *********************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' LinkedIn: https://www.linkedin.com/in/ajsiracusa/
' GitHub: https://github.com/JonatanSiracusa
' Instagram: @JonaSiracusa
'
'

Sub asignarCargarOrdenes()
Dim sistema As Integer
Dim Hoja2 As String


Hoja2 = "Datos"
libro = Worksheets(Hoja2).Cells(7, 3)

sistema = Workbooks(libro).Worksheets(Hoja2).Cells(9, 14)

Select Case sistema
'    Case 1
'        Call cargarOrdenes
'    Case 2
'        Call cargarOrdenes
'    Case 3
'        Call cargarOrdenesGallo2
    Case 4
        Call cargarOrdenesPython
End Select

End Sub
Sub asignarCancelarOrdenes()
Dim sistema As Integer
Dim Hoja2 As String

'libro = ThisWorkbook.Name
Hoja2 = "Datos"
libro = Worksheets(Hoja2).Cells(7, 3)

sistema = Workbooks(libro).Worksheets(Hoja2).Cells(9, 14)

Select Case sistema
'    Case 1
'        Call cancelarOrden
'    Case 2
'        Call cancelarOrden
'    Case 3
'        Call cancelarOrdenGallo2
    Case 4
        Call cancelarOrdenPython
End Select

End Sub
Sub asignarIngresarSistema()
Dim sistema As Integer
Dim Hoja2 As String

'libro = ThisWorkbook.Name
Hoja2 = "Datos"
libro = Worksheets(Hoja2).Cells(7, 3)

sistema = Workbooks(libro).Worksheets(Hoja2).Cells(9, 14)

Select Case sistema
'    Case 1
'        Call ingresarGalloRodi(sistema)
'    Case 2
'        Call ingresarGalloRodi(sistema)
'    Case 3
'        Call ingresarGallo2
    Case 4
        Call ingresarPython
End Select

End Sub
Sub asignarRefreshPeriodico()
Dim sistema As Integer
Dim Hoja2 As String

'libro = ThisWorkbook.Name
Hoja2 = "Datos"
libro = Worksheets(Hoja2).Cells(7, 3)

sistema = Workbooks(libro).Worksheets(Hoja2).Cells(9, 14)

Select Case sistema
    Case 1
        Call refreshPeriodicoSistemas(libro)
    Case 2
        Call refreshPeriodicoSistemas(libro)
    Case 3
        Call refreshPeriodicoSistemas(libro)
End Select

End Sub

