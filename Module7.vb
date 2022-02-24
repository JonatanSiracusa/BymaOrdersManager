'************************************************************************************************************************
'************************************************************************************************************************
'**************************** THIS MODULE IS USED FOR CODING GENERAL USE PROCEDURES *************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' LinkedIn: https://www.linkedin.com/in/ajsiracusa/
' GitHub: https://github.com/JonatanSiracusa
' Instagram: @JonaSiracusa
'
'


Sub generarLog(ByRef log() As String)
Dim m As Integer
Dim libroOrden, hoja, Hoja2 As String
'Dim rang As range


Hoja2 = "Datos"
libroOrden = Worksheets(Hoja2).Cells(7, 3)
hoja = "log"


'REGISTRO DE ACCIONES
m = 1

While log(m, 1) <> ""
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1).EntireRow.Insert Shift:=xlDown
    
    If Workbooks(libroOrden).Worksheets(hoja).Cells(5, 1) = "" Then
        Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1) = 1
    Else
        Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1) = "=+A5+1"
    End If
    
    'oper    activo  precio  estado  comentarios
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 2) = CDate(log(m, 14))
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 3) = CDate(log(m, 15))
    
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 4) = log(m, 1)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 5) = log(m, 2)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 6) = CDbl(log(m, 3))
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 7) = CDbl(log(m, 4))
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 8) = log(m, 5)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 9) = log(m, 6)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 10) = log(m, 7)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 11) = log(m, 8)
    
    m = m + 1
Wend


End Sub
Sub insertarNombreHoja()
Dim nombre, hojaDatos As String
Dim a, b, i, j As Integer

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

hojaDatos = "Datos"
libroOrden = Worksheets(hojaDatos).Cells(7, 3)
i = 10

While Worksheets(hojaDatos).Cells(i, 3) <> ""
    nombre = Worksheets(hojaDatos).Cells(i, 3)
    Worksheets(nombre).Cells(1, 244) = nombre
    
    i = i + 1
Wend

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

End Sub
Sub limpiarPrueba()
Dim libroOrden, hojaOrden As String
Dim i, j, cont As Integer

libroOrden = ActiveWorkbook.Name
hojaOrden = ActiveSheet.Name

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


i = 54
j = 146
cont = 0

While i < 303
    While cont < 4
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 0) = ""
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 1) = ""
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 2) = ""
        
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 20) = ""
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 21) = ""
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 22) = ""
        i = i + 1
        cont = cont + 1
    Wend
    i = i + 2
    cont = 0
Wend


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

MsgBox "Se limpiaron todas las celdas de prueba.", vbInformation, "Limpiar celdas de Pruebas"

End Sub
Sub copiarPosActual()
Dim libroOrden, hojaOrden As String
Dim i, j, cont, m As Integer
Dim posActualBruto(1 To 50, 1 To 10) As String

libroOrden = ActiveWorkbook.Name
hojaOrden = ActiveSheet.Name

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


i = 355
j = 133
cont = 0
m = 1

'tomo bases, cantidades y precios de la posición actual
While cont < 42
    posActualBruto(m, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 0)
    posActualBruto(m, 2) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 1)
    posActualBruto(m, 3) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 3)
    posActualBruto(m, 4) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 4)
    posActualBruto(m, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 6)
    
    i = i + 1
    m = m + 1
    cont = cont + 1
Wend



i = 54
j = 146
cont = 0
m = 1

'inserto el Resultado Acumulado Real en FE50
Workbooks(libroOrden).Worksheets(hojaOrden).Cells(50, 161) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(308, 133 + 19)
'Workbooks(libroOrden).Worksheets(hojaOrden).Cells(50, 161) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(6, 94)

While cont < 42
    If posActualBruto(m, 2) > 0 Then
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 0) = posActualBruto(m, 2)
    ElseIf posActualBruto(m, 2) < 0 Then
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 1) = -posActualBruto(m, 2)
    End If
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 2) = CDbl(posActualBruto(m, 3))
    
    If posActualBruto(m, 4) > 0 Then
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 20) = posActualBruto(m, 4)
    ElseIf posActualBruto(m, 4) < 0 Then
        Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 21) = -posActualBruto(m, 4)
    End If
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 22) = CDbl(posActualBruto(m, 5))
            
    i = i + 6
    m = m + 1
    cont = cont + 1
Wend


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

MsgBox "Se copio la posición actual a la planilla de prueba. Al cambiar la posición actual, la posición de prueba NO cambia.", vbInformation, "Copiar posición actual"

End Sub

Sub irAPosicion()

libro = Worksheets("Datos").Cells(7, 3)
hoja = "GGAL"

Sheets(hoja).Select
Cells(20, 77).Select

End Sub

Sub irABoletos()

libro = Worksheets("Datos").Cells(7, 3)
hoja = "boletos"

Sheets(hoja).Select
Cells(3, 2).Select

End Sub

