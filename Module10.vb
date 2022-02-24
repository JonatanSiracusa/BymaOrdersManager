'************************************************************************************************************************
'************************************************************************************************************************
'********************** THIS MODULE CONTAINS PROCEDURES FOR TAKING THE NECESSARY ORDER'S PARAMETERS *********************
'**************** FROM A M. EXCEL SHEET FOR SENDING, CANCELLING AND WRITING IN THE LOG EVERY ORDER TAKEN ****************
'************************************************************************************************************************
'************************************************************************************************************************
'
' LinkedIn: https://www.linkedin.com/in/ajsiracusa/
' GitHub: https://github.com/JonatanSiracusa
' Instagram: @JonaSiracusa
'
'

Public Sub ingresarPython()
Dim MyHTML_Element As IHTMLElement
Dim winID As Long
Dim MyURL, hojaDOM, userGallo, passGallo, userRodi, passRodi, dni, cuenta, fileName As String
Dim libro, hoja, comitente As String
Dim log(1 To 30, 1 To 15) As String


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


libro = ThisWorkbook.Name
hoja = "log"

dni = Workbooks(libro).Worksheets("Datos").Cells(20, 8)
userGallo = Workbooks(libro).Worksheets("Datos").Cells(21, 8)
passGallo = Workbooks(libro).Worksheets("Datos").Cells(22, 8)
comitente = Workbooks(libro).Worksheets("Datos").Cells(19, 8)
fileName = Workbooks(libro).Worksheets("Datos").Cells(7, 3)

'INGRESO A GALLO 2 a traves de Python
RunPython ("import sniper_v1; sniper_v1.primer_login('" & dni & "', '" & userGallo & "', '" & passGallo & "', '" & comitente & "', '" & fileName & "')")

'Genero e imprimo el Log
log(1, 1) = "Login"
log(1, 2) = "HomeBroker"
log(1, 3) = "0"
log(1, 4) = "0"
log(1, 5) = ""
log(1, 6) = comitente
log(1, 7) = "ok"
log(1, 8) = "ingresarPython"
log(1, 14) = Date
log(1, 15) = Time

Call generarLog(log())


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

MsgBox "Se hizo el Login a GALLO 2 (Home Broker) usando Python.", vbInformation, "INGRESO A SISTEMAS"


End Sub
Sub cargarOrdenesPython()
Dim arrParamOrden(1 To 20, 1 To 6) As String
Dim arrInforOrden(1 To 20, 1 To 15) As String
'Dim arrEnlaceOrden(1 To 20) As String
Dim sistema As Integer
Dim user, pass As String
Dim oldStatusBar As String

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


sistema = 0
user = ""
pass = ""


oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Cargando ordenes.."

Call generarParamOrdenGallo2Python(arrParamOrden(), arrInforOrden())
Call cargarOrdenGallo2Python(arrParamOrden(), arrInforOrden())

Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

End Sub
Sub generarParamOrdenGallo2Python(ByRef arrParamOrden() As String, ByRef arrInforOrden() As String)
Dim comitente As String
Dim hojaOrden, libroOrden, cuenta As String
Dim i, j, a, b, c, d, m, n, p, q As Integer
Dim esPut As Boolean

' EN ESTE PROCREDIMIENTO SE TOMAN LOS PARAMETROS PARA CARGAR LAS ORDENES, SE INSERTAN EN UN ARREGLO DE VBA Y SE IMPRIMEN EN
' EL RANGO EC402:EI441, PARA QUE LUEGO PUEDAN SER TOMADAS POR PYTHON COMO UN ARREGLO.


libroOrden = ActiveWorkbook.Name
hojaOrden = ActiveSheet.Name

'Determino si es Put o Call, según lo que seleccionado
esPut = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(1, 73)

'La cuenta que se esta operando
cuenta = Workbooks(libroOrden).Worksheets("Datos").Cells(9, 16)

'Determino el nro. de cuenta comitente
comitente = Workbooks(libroOrden).Worksheets("Datos").Cells(19, 8)


'Para recorrer las celdas
i = 7   ' limpiarCantidades() usa este mismo i. Si se cambia, tmb cambiar el i de limpiarCantidades()
j = 82  ' limpiarCantidades() usa este mismo j. Si se cambia, tmb cambiar el j de limpiarCantidades()

'Para recorrer el arreglo
a = 1
b = 1


'1) SE GENERAN LOS PARAMETROS PARA EMITIR LAS ORDENES Y SE ACUMULAN DENTRO DEL ARREGLO.

While Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 56) <> ""
    
    'COMPRA
    If Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 5) > 0 Then     'Solo genera la fila si hay CANTIDAD.
        If esPut = False Then
            arrParamOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 68)    'ESPECIE/INSTRUMENO
            arrInforOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 68)
        Else
            arrParamOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 66)    'ESPECIE/INSTRUMENTO
            arrInforOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 66)
        End If
        
        arrParamOrden(a, 2) = 1             'COMPRA
        arrInforOrden(a, 2) = "COMPRA"
        arrParamOrden(a, 3) = "2"            'PLAZO
        arrInforOrden(a, 3) = "24hs."
        arrParamOrden(a, 4) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 5)      'CANTIDAD
        arrInforOrden(a, 4) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 5)
        
        If Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 4) > 0 Then                           'PRECIO
            arrParamOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 4)  'Si hay precio manual, toma el manual.
            arrInforOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 4)
        Else
            arrParamOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 2)  'Si NO hay precio manual, toma la PUNTA VENDEDORA.
            arrInforOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 2)
        End If
        
        arrParamOrden(a, 6) = comitente               'COMITENTE
        arrInforOrden(a, 6) = comitente
        
        'arrParamOrden(a, 13) = "Ejecutado por 3.08"     'COMENTARIO
        
        
        a = a + 1
    End If
    
    'VENTA
    If Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 5) > 0 Then           'Solo genera la fila si hay CANTIDAD.
        If esPut = False Then
            arrParamOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 68)    'ESPECIE
            arrInforOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 68)
        Else
            arrParamOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 66)    'ESPECIE
            arrInforOrden(a, 1) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 66)
        End If
        
        arrParamOrden(a, 2) = 2             'VENTA
        arrInforOrden(a, 2) = "VENTA"
        arrParamOrden(a, 3) = "2"            'PLAZO
        arrInforOrden(a, 3) = "24hs."
        arrParamOrden(a, 4) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 5)      'CANTIDAD
        arrInforOrden(a, 4) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 5)
        
        If Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 4) > 0 Then                           'PRECIO
            arrParamOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 4)  'Si hay precio manual, toma el manual.
            arrInforOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 4)
        Else
            arrParamOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 2)   'Si NO hay precio manual, toma la PUNTA COMPRADORA.
            arrInforOrden(a, 5) = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 2)
        End If
        
        arrParamOrden(a, 6) = comitente               'COMITENTE
        arrInforOrden(a, 6) = comitente
        
        'arrParamOrden(a, 13) = "Ejecutado por 3.08"    'COMENTARIO
        
        a = a + 1
    End If
    
    i = i + 1
Wend

c = 1
m = 402
n = 133
While arrParamOrden(c, 1) <> ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n) = arrParamOrden(c, 1)       'ESPECIE
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n + 1) = CInt(arrParamOrden(c, 4))   'CANTIDAD
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n + 3) = CDbl(arrParamOrden(c, 5))   'PRECIO
    'Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n + 3) = arrParamOrden(c, 5)   'PRECIO
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n + 4) = arrParamOrden(c, 2)   'TIPO 1=COMPRA; 2=VENTA
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(m, n + 6) = arrParamOrden(c, 3)   'PLAZO 2=24HS
    c = c + 1
    m = m + 1
Wend


End Sub
Sub cargarOrdenGallo2Python(ByRef arrParamOrden() As String, arrInforOrden() As String)
Dim log(1 To 30, 1 To 15) As String
Dim rngParamOrden, cookie, fileName As String
Dim a, contadorAux As Integer

contadorAux = 0

a = 1
libroOrden = ActiveWorkbook.Name
hojaOrden = "Datos"
hoja = ActiveSheet.Name


'ENVIO LA CARGA DE LA ORDEN A TRAVES DE PYTHON
'1) Asigno a un String los valores del rango donde se encuetran los parametros de las ordenes
'rngParamOrden = "EC402:EI441"
rngParamOrden = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(7, 2)
cookie = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(18, 8)
fileName = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(7, 3)

'2) Hago el llamado a la Fjunction de Python para cargar las ordenes
RunPython ("import sniper_v1; sniper_v1.cargar_ordenes('" & rngParamOrden & "','" & hoja & "','" & cookie & "', '" & fileName & "')")


' EL LOG SE GENERA DESDE PYTHON.


End Sub
Sub limpiarCantidadesGallo2()
Dim hojaOrden, libroOrden As String
Dim i, j As Integer

libroOrden = ActiveWorkbook.Name
hojaOrden = ActiveSheet.Name

i = 7
j = 82
While Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 56) <> ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j - 5) = ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 5) = ""
    i = i + 1
Wend

i = 402
j = 133
While Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j) <> ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j) = ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 1) = ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 3) = ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 4) = ""
    Workbooks(libroOrden).Worksheets(hojaOrden).Cells(i, j + 6) = ""
    i = i + 1
Wend

End Sub
Sub generarMensajePython(ByRef arr_resultado() As String)
Dim a, b, i, j As Integer
Dim mensaje As String

a = 1
b = 1
mensaje = ""

While arr_resultado(a, 1) <> ""
    mensaje = mensaje & "Orden nro. " & arr_resultado(a, 7) & ":  " & arr_resultado(a, 1) & "  "
    mensaje = mensaje & arr_resultado(a, 2) & Chr(13) '& "Plazo: " & arrInforOrden(a, 3) & Chr(13)
    mensaje = mensaje & "Lotes:       " & arr_resultado(a, 3) & Chr(13) & "Precio:  $ " & arr_resultado(a, 4) & Chr(13) & Chr(13)
    
    a = a + 1
Wend

'4) SE VISUALIZA UN MENSAJE CON LAS ORDENES CARGADAS.

MsgBox "Se cargaron las siguientes ordenes: " & Chr(13) & Chr(13) & mensaje, vbInformation, "Ordenes cargadas en GALLO 2 (Home Broker) por PYTHON"

End Sub
Sub generarLogPython(arr_resultado_carga_orden)
Dim arr_resultado(1 To 40, 1 To 8) As String
Dim a, b, c, d, m, n, i, j, total_caracteres, terminar, pos, pos2, pos3, pos4, pos_final As Integer
Dim libroOrden, hoja, Hoja2, cadena, st_item_respuesta As String
'Dim arr_resultado_carga_orden, cadena As String


'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'Application.EnableEvents = False
'ActiveSheet.DisplayPageBreaks = False


Hoja2 = "Datos"
libroOrden = Worksheets(Hoja2).Cells(7, 3)
hoja = "log"

cadena = arr_resultado_carga_orden
i = 0
j = 0
terminar = 0

While terminar <> 1
    cadena = Trim(cadena)
    If Left(cadena, 1) = "<" Then
        cadena = Trim(Right(cadena, Len(cadena) - 1))
        i = i + 1
        j = 1
    End If
    pos_final = InStr(cadena, ">")
    pos = InStr(cadena, ";")
    
    If pos <> 0 And pos < pos_final Then
        st_item_respuesta = Left(cadena, pos - 1)
        arr_resultado(i, j) = st_item_respuesta
        cadena = Right(cadena, Len(cadena) - pos)
    'End If
    ElseIf ((pos <> 0) And (pos > pos_final)) Then
        pos2 = InStr(cadena, ">")
        st_item_respuesta = Left(cadena, pos2)
        arr_resultado(i, j) = st_item_respuesta
        cadena = Right(cadena, Len(cadena) - pos2)
    'End If
    ElseIf (pos = 0) Then
        pos3 = InStr(cadena, ">")
        st_item_respuesta = Left(cadena, pos3)
        arr_resultado(i, j) = st_item_respuesta
        cadena = Right(cadena, Len(cadena) - pos3)
        terminar = 1
    End If
    
    If j = 7 Then
        If Len(arr_resultado(i, j)) > 5 Then
            pos4 = InStr(arr_resultado(i, j), "Nro")
            arr_resultado(i, j) = Mid(arr_resultado(i, j), pos4 + 4, 6)
        Else
            arr_resultado(i, j) = "ERROR"
        End If
    End If
    
    j = j + 1
Wend


m = 1
While arr_resultado(m, 1) <> ""
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1).EntireRow.Insert Shift:=xlDown
    
    If Workbooks(libroOrden).Worksheets(hoja).Cells(5, 1) = "" Then
        Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1) = 1
    Else
        Workbooks(libroOrden).Worksheets(hoja).Cells(4, 1) = "=+A5+1"
    End If
    
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 2) = CDate(Date)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 3) = CDate(Time)
    
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 4) = arr_resultado(m, 1)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 5) = arr_resultado(m, 2)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 6) = CDbl(arr_resultado(m, 3))
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 7) = CDbl(arr_resultado(m, 4))
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 8) = arr_resultado(m, 5)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 9) = arr_resultado(m, 6)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 10) = arr_resultado(m, 7)
    Workbooks(libroOrden).Worksheets(hoja).Cells(4, 11) = "cargarOrdenesPython"
    
    m = m + 1
Wend

Call limpiarCantidadesGallo2
Call generarMensajePython(arr_resultado())


End Sub
Sub cancelarOrdenPython()
Dim MyHTML_Element As IHTMLElement
Dim MyURL, libro, comitente, cuenta As String
Dim log(1 To 30, 1 To 15) As String
Dim a, cont, lim, contadorAux, hecho As Integer
Dim hayCancelar, primero, hayOrdenes As Boolean

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "CANCELANDO ordenes.."


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


a = 1
libroOrden = ActiveWorkbook.Name
hojaOrden = "Datos"
hoja = ActiveSheet.Name
comitente = Workbooks(libroOrden).Worksheets("Datos").Cells(19, 8)


'ENVIO LA CARGA DE LA ORDEN A TRAVES DE PYTHON
'1) Asigno a un String los valores del rango donde se encuetran los parametros de las ordene
'rngParamOrden = "EC402:EI441"
cookie = Workbooks(libroOrden).Worksheets(hojaOrden).Cells(18, 8)

'2) Hago el llamado a la Fjunction de Python para cargar las ordenes
RunPython ("import sniper_v1; sniper_v1.cancelar_todas_ordenes('" & cookie & "')")


log(a, 1) = "Cancelar"
log(a, 2) = "gallo2"
log(a, 3) = "0"
log(a, 4) = "0"
log(a, 5) = ""
log(a, 6) = comitente
log(a, 7) = "ok"
log(a, 8) = "cancelarOrdenPython"
log(a, 14) = Date
log(a, 15) = Time


Call generarLog(log())


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar


MsgBox "Se ANULARON TODAS las ordenes cargadas.", vbInformation, "Ordenes ANULADAS en Gallo 2 (Home Broker)" '


End Sub

