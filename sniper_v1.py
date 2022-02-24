"""
************************************************************************************************************************
************************************************************************************************************************
***************** THIS PROGRAM IS USED TO SEND AND CANCEL ByMA STOCK MARKET ORDERS *************************************
************************************************************************************************************************
************************************************************************************************************************

LinkedIn: https://www.linkedin.com/in/ajsiracusa
GitHub: https://github.com/JonatanSiracusa
Instagram: @JonaSiracusa

"""

import requests
import asyncio
import json
from selenium import webdriver
from seleniumwire import webdriver
import xlwings as xw
import time
import numpy as np



datosExcel = {
    'hoja_datos': 'Datos',
    'hoja_log': 'log',
}

datosRequest = {
    'ip_address': '',
    'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36',
    'cookie': '',
    'url_login': 'https://clientes.tuagentedebolsa.com.ar/',
    'url_cargar_orden1': 'https://clientes.tuagentedebolsa.com.ar/Order/ValidarCargaOrdenAsync',
    'url_cargar_orden2': 'https://clientes.tuagentedebolsa.com.ar/Order/EnviarOrdenConfirmadaAsyc',
    'url_consulta_ordenes': 'https://clientes.tuagentedebolsa.com.ar/Consultas/GetConsulta',
    'url_cancelar_orden1': 'https://clientes.tuagentedebolsa.com.ar/Order/EnviarCancelacionAsyc',
    'url_cancelar_orden2': 'https://clientes.tuagentedebolsa.com.ar/Order/EnviarOrdenCanceladaAsyc',
    'comitente': '',
    'proceso': '121',
}

dataCancelar = {
    'comitente': '',
    'proceso': '121',
}

datosLogin = {
    'ip_address': '',
    'url_login': 'https://clientes.tuagentedebolsa.com.ar/',
}


def cantidadOrden(cantidad):
    """Metodo que convierte cualquier numero string con decimales a un string sin decimales"""
    cantidad = str(cantidad)
    return cantidad


def precioOrden(precio):
    """Metodo que convierte cualquier precio al formato necesario"""
    precio_str = str(precio)
    precio = precio_str.replace(".", ",")
    return precio


def optionTipoOrden(option_tipo):
    """Metodo que convierte cualquier numero string con decimales a un string sin decimales"""
    option_tipo = str(option_tipo)
    return option_tipo


def optionTipoPlazoOrden(option_tipo_plazo):
    """Metodo que convierte cualquier numero string con decimales a un string sin decimales"""
    option_tipo_plazo = str(option_tipo_plazo)
    return option_tipo_plazo


def dateValid():
    """Metodo que establece la fecha valida de una orden"""
    fecha = time.strftime("%d/%m/%Y")
    return fecha


class Orden():
    """En esta clase van a estar los parametros de una orden. Asimismo, se definen los Metodos para cargar una orden."""

    def __init__(self):
        self.nombre_especie = '',  # 'GFGC200.AB',
        self.cantidad = '',  # '5',
        self.precio = '',  # '31,20',
        self.date_valid = dateValid(),  # '24/02/2022',  # la validez es UN DIA MAS
        self.option_tipo = '',  # '1': Compra; '2': Venta
        self.option_tipo_plazo = '2',  # '2' 24hs

    def cargar_orden(self, nombre_especie, cantidad, precio, option_tipo, option_tipo_plazo):

        def resultado_carga_orden():
            """Tomo el resultado de la carga de la orden: AcceptMessage (nro de orden) o ErrorMessage"""

            if data['OptionTipo'] == '1':
                info_orden = "<COMPRA; "
            else:
                info_orden = "<VENTA; "

            info_orden = info_orden + data['NombreEspecie'] + '; '
            info_orden = info_orden + data['Cantidad'] + '; '
            info_orden = info_orden + data['Precio'] + '; '

            if data['OptionTipoPlazo'] == '2':
                info_orden = info_orden + '24hs; '
            else:
                info_orden = info_orden + ' ; '

            info_orden = info_orden + datosRequest['comitente'] + '; '

            resultado_orden_dict = json.loads(r.text)
            accept_message = (resultado_orden_dict["Result"]["ResponseOrden"]["AcceptMessage"])
            error_message = (resultado_orden_dict["Result"]["ResponseOrden"]["ErrorMessage"])
            if accept_message is None:
                info_orden = info_orden + error_message + '> '
            else:
                info_orden = info_orden + accept_message + '> '

            return info_orden

        url = datosRequest['url_cargar_orden1']
        headers = datosRequest
        data = {
            'NombreEspecie': nombre_especie,
            'Cantidad': cantidadOrden(cantidad),
            'Precio': precioOrden(precio),
            'Importe': '',
            'DateValid': self.date_valid,  # la validez es UN DIA MAS
            'OptionTipo': optionTipoOrden(option_tipo),  # 1: Compra; 2: Venta
            'OptionTipoPlazo': optionTipoPlazoOrden(option_tipo_plazo),
        }
        r = requests.post(url=url, data=data, headers=headers)

        url = datosRequest['url_cargar_orden2']
        headers = datosRequest
        r = requests.post(url=url, headers=headers)

        return resultado_carga_orden()


def cargar_ordenes(rng_params_orden, hoja_excel, cookie, file_name):
    """Este procedimiento se encarga de hacer la carga de las ordenes que envía la planilla de Excel"""

    # El arreglo se pasa como un rango en String ("A1:C4") y luego se convierte en un arreglo
    # usando: arr = xw.Range(arr_params_orden).options(np.array, ndim=2).value

    datosRequest['cookie'] = cookie
    wb = xw.Book(file_name)
    sheet = wb.sheets[hoja_excel]

    arr_params_orden_bruto = xw.Range(rng_params_orden).options(np.array, ndim=2).value
    print("arr_params_orden_bruto:", arr_params_orden_bruto)

    # Determino cuántas ordenes hay en el arreglo
    i = 0
    q_ordenes = 0
    while arr_params_orden_bruto[i, 0] != "nan":
        q_ordenes = q_ordenes + 1
        i = i + 1

    # Genero un arreglo nuevo solo con los parametros necesarios para cargar las ordenes: arr_params_orden
    a = 0
    arr_params_orden = []
    item_arr_params_orden = []

    while a < q_ordenes:
        item_arr_params_orden.append(arr_params_orden_bruto[a, 0])
        item_arr_params_orden.append(int(float(arr_params_orden_bruto[a, 1])))
        item_arr_params_orden.append(arr_params_orden_bruto[a, 3])
        item_arr_params_orden.append(int(float(arr_params_orden_bruto[a, 4])))
        item_arr_params_orden.append(int(float(arr_params_orden_bruto[a, 6])))

        arr_params_orden.append(item_arr_params_orden[:])
        item_arr_params_orden.clear()
        a += 1

    b = 0
    arr_resultado_carga_orden = ''
    while b < q_ordenes:
        nombre_especie = arr_params_orden[b][0]
        cantidad = arr_params_orden[b][1]
        precio = arr_params_orden[b][2]
        option_tipo = arr_params_orden[b][3]
        option_tipo_plazo = arr_params_orden[b][4]

        orden = Orden()
        resultado_orden = orden.cargar_orden(nombre_especie, cantidad, precio, option_tipo, option_tipo_plazo)
        arr_resultado_carga_orden = arr_resultado_carga_orden + resultado_orden
        b += 1

    macro_log_excel = wb.macro('generarLogPython')
    macro_log_excel(arr_resultado_carga_orden)


def cancelar_todas_ordenes(cookie):
    def hacer_request(url, data, headers):
        return requests.post(url, data=data, headers=headers)

    def cantidad_de_ordenes():
        """ Esta Function indica la cantidad de ordenes que estan en cualquier "Estado". """
        a = 0
        orden = 0
        CantOrdenes = 0

        vuelta1 = len(estadoOrdenes_dict["Result"])

        while a < vuelta1:
            vuelta2 = len(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"])

            while orden < vuelta2:
                CantOrdenes += 1
                # print(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["NUME"])
                orden += 1

            a += 1
            orden = 0

        print("Cant de Ordenes:", CantOrdenes)
        return CantOrdenes


    # 1) Consulto el ESTADO DE ORDENES
    # print("\n\n1) HAGO UNA CONSULTA AL ESTADO DE ORDENES (Consultas/GetConsulta): ")
    print("\n\nSe estan cancelando las ordenes pendientes:")

    datosRequest['cookie'] = cookie

    url = datosRequest['url_consulta_ordenes']
    headers = datosRequest
    data = dataCancelar
    r = requests.post(url=url, data=data, headers=headers)

    # print("\nRespuesta del Request a IniciarOrdenAsync: \n", r.text)
    # print("\nStatus Code: ", r.status_code)

    estadoOrdenes_dict = json.loads(r.text)
    # print(estadoOrdenes_dict, "\n")

    qOrdenes = cantidad_de_ordenes()
    print("qOrdenes: ", qOrdenes)

    ordenesRecibidas = []
    itemOrdenesRecibidas = []

    # Empieza
    a = 0
    orden = 0
    CantOrdenes = 0

    vuelta1 = len(estadoOrdenes_dict["Result"])
    # print("Vuelta1 de Result:", vuelta1)

    while a < vuelta1:
        vuelta2 = len(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"])

        while orden < vuelta2:
            CantOrdenes += 1
            estado = estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["ESTA"]

            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["CESP"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["TICK"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["CANT"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["PCIO"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["TIPO"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["PLAZ"])
            itemOrdenesRecibidas.append(estadoOrdenes_dict["Result"][a]["listaDetalleTiker"][0]["ORDE"][orden]["NUME"])
            if (estado == "Recibida") or (estado == "Pendiente") or (estado == "Parcial"):
                ordenesRecibidas.append(itemOrdenesRecibidas[:])
            itemOrdenesRecibidas.clear()

            orden += 1
        a += 1
        orden = 0
    # termina

    qOrdenes = len(ordenesRecibidas)
    i = 0
    print(ordenesRecibidas)

    for each in ordenesRecibidas:
        # 1) Intento llamar a Order/EnviarCancelacionAsyc y cancelar una orden
        # print("\n\n1) 1ER PASO EN LA CANCELACION DE ORDEN (EnviarCancelacionAsyc): ")
        url = datosRequest['url_cancelar_orden1']
        headers = datosRequest

        if i < qOrdenes:
            data = {
                'especie': ordenesRecibidas[i][0],
                'Ticker': ordenesRecibidas[i][1],
                'Cantidad': ordenesRecibidas[i][2],
                'Precio': ordenesRecibidas[i][3],
                'OptionTipo': ordenesRecibidas[i][4],
                'OptionTipoPlazo': ordenesRecibidas[i][5],
                'Numero': ordenesRecibidas[i][6],
            }
            i += 1

        async def cancelar_todas_ordenes_async(data):
            # prueba async
            url1 = datosRequest['url_cancelar_orden1']
            headers1 = datosRequest
            url2 = datosRequest['url_cancelar_orden2']
            headers2 = datosRequest
            loop = asyncio.get_event_loop()

            future1 = loop.run_in_executor(None, hacer_request, url1, data, headers1)
            response1 = await future1
            # print(response1, url1, data)

            future2 = loop.run_in_executor(None, hacer_request, url2, data, headers2)
            response2 = await future2
            # print(response2, url2, data)
            print("Se cancelo la siguiente orden: ", data)

        loop = asyncio.get_event_loop()
        loop.run_until_complete(cancelar_todas_ordenes_async(data))

    print("Se finalizo la cancelación de las ordenes.\n")


def primer_login(dni_login, usuario_login, pass_login, comitente, file_name):
    """Metodo que realiza el Login a la web utilizando Selenium y permite generar la cookie."""

    def guardar_cookie():
        """Este procedimiento guarda la cookie en una celda del Excel"""
        wb = xw.Book(file_name)
        sheet = wb.sheets[datosExcel['hoja_datos']]
        sheet[17, 7].value = cookie

    # Ingreso a la URL para hacer Login
    browser = webdriver.Chrome('C:\@ MyPython\chromedriver')
    browser.get(datosLogin['url_login'])

    # Ubico los elementos HTML
    dni = browser.find_element_by_name('Dni')
    usuario = browser.find_element_by_id('usuario')
    clave = browser.find_element_by_id('passwd')

    # Inserto los datos correspondientes en los elementos HTML para el Login
    dni.send_keys(dni_login)
    usuario.send_keys(usuario_login)
    clave.send_keys(pass_login)

    submit = browser.find_element_by_id('loginButton')
    submit.click()

    # Busco los datos necesarios para generar la cookie
    cookies_ahora = browser.get_cookies()
    cookies_dict = []
    cookies_dict = cookies_ahora

    # Genero la cookie que se va a almacenar en datosConfig
    cookie_generada = 'ASP.NET_SessionId=' + cookies_dict[1]['value'] + '; '
    cookie_generada = cookie_generada + '.ASPXAUTH=' + cookies_dict[0]['value'] + ""
    cookie = cookie_generada

    # Inserto en Excel la cookie, para tenerla disponible.
    guardar_cookie()

    # Guardo el nro. de comitente
    datosRequest['comitente'] = comitente
    dataCancelar['comitente'] = comitente

    browser.quit()
    print('Se hizo el LogIn correctamente.\n')
    return cookie


"""
************************************************************************************************************************
***************** THE NEXT CODE IS COMMENTED DUE TO BE USEFUL ONLY WHILE DEBUGGING BEFORE MAKING ***********************
*************************** THE CONNECTION TO MICROSOFT EXCEL USING VISUAL BASIC CODE **********************************
************************************************************************************************************************
"""


'''
def mje_inicio():
    """Metodo que emite un mensaje al iniciar el programa"""
    print("\n**********  GESTION DE ORDENES by @JonaSiracusa  **********\n\n")
    print("Ingresando al sistema.")


def indicacion_usuario():
    """Metodo que toma lo que el usuario desea realizar. Se usa momentáneamente."""
    rta = input(
        '\nPulse "O" para cargar una orden.\nPulse "C" para cancerlar todas las ordenes.\nPulse "S" para finalizar.\n¿Qué se desea hacer? ')
    rta = rta.lower()
    return rta


def main():
    """Programa principal"""

    # Hago el LogIn, tomo los datos necesarios para generar la cookie y la asigno a una variable.
    mje_inicio()
    #cookie_login = primer_login()
    cookie_login = primer_login()
    datosRequest['cookie'] = cookie_login
    print("\nCookie del Dic: ", datosRequest['cookie'], "\n")

    hacer = indicacion_usuario()

    # ------------- Bucle principal del programa ---------------
    while hacer != "s":
        if hacer == "o":
            orden = Orden()
            respuesta = orden.cargar_orden('GFGC10681O', '5', '1,02', '1', '2')
            print("\nRespuesta: ", respuesta, "\n")
            

        if hacer == "c":
            # cancelar_todas_ordenes()
            cancelar_todas_ordenes(datosRequest['cookie'])

        hacer = indicacion_usuario()


if __name__ == '__main__':
    main()
'''