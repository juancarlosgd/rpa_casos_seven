from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchWindowException
from selenium.webdriver.remote.command import Command
import urllib3
import pywinauto as pywin
from pywinauto.timings import TimeoutError
from pywinauto.findwindows import WindowNotFoundError
import time
import xlrd as excel
import ctypes
import configparser
import warnings
import os
from pathlib import Path
import datetime

def status_driver(driver):
    try:
        handle=driver.current_window_handle
        if handle != None:
            status=True
        else:
            status=False  
    except urllib3.exceptions.MaxRetryError as e:
        status=False
    except NoSuchWindowException as e:
        status=False
    return status

def parse_cadena(archivo,hoja,hojaCasos,filaCadena,colCadena,colRangoi,colRangof,ColDuplicado):
    wb = excel.open_workbook(filename=archivo)
    hojaCadena=wb.sheet_by_name(hoja)
    hojaDatos=wb.sheet_by_name(hojaCasos)
    cadena=hojaCadena.cell_value(filaCadena-1,colCadena-1)
    lista= cadena.split(sep=',')
    tamLista=len(lista)
    # Creamos listas
    filas = []
    varCasoprevio=''
    for fila in range(0,tamLista+1):
        columnas = []
        if varCasoprevio != hojaDatos.cell_value(fila,ColDuplicado): #colRangoi-1):
            for columna in range(colRangoi-1,colRangof):
                columnas.append(hojaDatos.cell_value(fila,columna))
            filas.append(columnas)
        varCasoprevio = hojaDatos.cell_value(fila,ColDuplicado) #colRangoi-1)
    #Elimina la fila de titulos
    filas.pop(0)
    lista=list(dict.fromkeys(lista))
    tamLista=len(lista)
    return lista, tamLista,filas
if __name__ == "__main__":
    tic= time.perf_counter()
    dir_actual= Path(__file__).parent
    print (dir_actual)
    warnings.simplefilter('ignore', category=UserWarning)
    #Lee archivo de configuracion
    print('Leyendo archivo de configuracion...')
    configuracion= configparser.ConfigParser()
    configuracion.read(str(dir_actual)+os.sep+'configuracion.ini',encoding='utf-8-sig')
    configuracion.sections()
    tipoProceso= int(configuracion['Proceso']['Tipo'])
    archivoOrigen= configuracion['Archivos']['Excel']
    hojaDatos=configuracion['Archivos']['Hoja_origen']
    hojaLista= configuracion['Archivos']['Hoja_destino']
    filaDatos=configuracion['Archivos']['Fila_origen']
    columnaDatos=configuracion['Archivos']['Col_origen']
    colListai=configuracion['Archivos']['Col_Destino_i']
    colListaf=configuracion['Archivos']['Col_Destino_f']
    colIndice=configuracion['Archivos']['Col_Indice']
    colEjecutor=configuracion['Archivos']['Col_Ejecutor']
    colTexto = configuracion['Archivos']['Col_Texto']
    colDuplicidad=configuracion['Archivos']['Col_Duplicidad']
    textoAenviar=configuracion['Elementos']['Texto_a_enviar']
    siguienteEjecutor=configuracion['Elementos']['Siguiente_ejecutor']
    autorizaOrden=configuracion['Proceso']['Autoriza_orden']
    print('Archivo de configuración cargado al robot')
    print('Leyendo archivo de origen con los casos a procesar...')
    lista_casos,num_casos,fila_casos= parse_cadena(archivoOrigen,hojaDatos,hojaLista,int(filaDatos),int(columnaDatos),int(colListai),int(colListaf),int(colDuplicidad))
    #print (lista_casos, num_casos,fila_casos)
    print('Archivo de origen cargado')
    key=0
    for key in range(0,num_casos):
        driver= webdriver.Ie('IEDriverServer.exe')
        if tipoProceso == 0:
            varCasoweb=0
        elif tipoProceso == 1:
            varCasoweb=1
        elif tipoProceso == 2:
            varCasoweb=0
        else:
            varCasoweb=0
        #Abre el navegador con la cadena del caso
        cadenaWeb=configuracion['Web_casos']['Url_1']+ str(int(fila_casos[key][varCasoweb])) + '&' + str(int(fila_casos[key][int(colIndice)-int(colListai)])) + configuracion['Web_casos']['Url_2']
        wndSeven= driver.get(cadenaWeb)
        driver.maximize_window()
        try:
            elemento=WebDriverWait(driver,int(configuracion['Elementos']['Espera'])).until(EC.element_to_be_clickable((By.ID, "CWfSvrcnAcx")))
            #obtiene el control del objeto activex para hacer clic en el  panel de opciones y el tabcontrol
            wndapp= pywin.Application(backend="win32").connect(title=configuracion['Elementos']['Texto_titulo'],visible_only=True)
            sevenErp=wndapp.window(title_re=configuracion['Elementos']['Texto_titulo'])
            #Codigo para automatizar la aprobación en el modulo SCMAUOCO
            if autorizaOrden=='Si':
                elemento2=WebDriverWait(driver,int(configuracion['Elementos']['Espera'])).until(EC.element_to_be_clickable((By.ID, "SCMAUOCO")))
                activex=sevenErp.child_window(class_name="TCCmAuocoAcX")
                btn= activex.child_window(class_name='TBitBtn',ctrl_index=2)
                btnAutoriza = pywin.controls.win32_controls.ButtonWrapper(btn.wrapper_object())
                if btn.is_enabled():
                    btnAutoriza.click()
                    try:
                        pywin.timings.wait_until_passes(20, .5, lambda: (pywin.findwindows.find_window(title='Information', class_name="TMessageForm")),(WindowNotFoundError))
                        messBox= pywin.Application(backend="win32").connect(title='Information', class_name="TMessageForm")
                        dlgFrm=messBox.window(title_re='Information')
                        dlgFrm.wait('ready',timeout=10, retry_interval=0.5)
                        pywin.controls.win32_controls.ButtonWrapper(dlgFrm.child_window(class_name='TButton').wrapper_object()).click()
                        print ('Caso '+ str(int(fila_casos[key][varCasoweb])) +' - orden de compra autorizada')                
                    except TimeoutError as e:
                        print("Se excedió el timepo de espera del cuadro de aceptación de la aprobación de la orden.")
                        driver.close()
                        toc= time.perf_counter()
                        ctypes.windll.user32.MessageBoxW(0, "El robot ha finalizado anticipadamente por demora excesiva en la respuesta de SEVEN ERP. Se procesaron " + str(key+1) + " casos en " + str(datetime.timedelta(seconds=toc-tic)) + "el caso " + str(int(fila_casos[key][varCasoweb])) + " no fue procesado.", "Robot Terminado", 0)
                else:
                    print ('El caso ' + str(int(fila_casos[key][varCasoweb])) + ' ya fue autorizado previamente')
            #Fin del codigo SCMAUOCO
            activex=sevenErp.child_window(title="CWfSvrcnAcx", class_name="TCWfSvrcnAcx")
            ophelia=sevenErp.child_window(class_name='Internet Explorer_Server')
            ophelia_ctrl= ophelia.wrapper_object()
            page_control = activex.child_window(class_name="TPageControl")
            tab_control = pywin.controls.common_controls.TabControlWrapper(page_control.wrapper_object())
            tab_control.select(1)
            tabPage=page_control.child_window(title_re="A enviar",class_name="TTabSheet")
            memoField= tabPage.child_window(class_name="TMemo")
            if textoAenviar.lower() == 'variable':
                memoField.type_keys(str(fila_casos[key][int(colTexto)-int(colListai)]), with_spaces = True)
            else:               
                memoField.type_keys(textoAenviar, with_spaces = True)
            tpanelAcciones= activex.child_window(title_re="PnlAcciones",class_name="TPanel")
            checkList=tpanelAcciones.child_window(class_name="TCheckListBox")
            if tipoProceso < 3:
                listItem = pywin.controls.win32_controls.ListBoxWrapper(checkList.wrapper_object())
                for x, res in enumerate(listItem.item_texts()):
                    if res.strip()==configuracion['Elementos']['Texto_aprobacion']:
                        varAccion=x
                varx = listItem.item_rect(varAccion).left + int(configuracion['Elementos']['Offsetx'])
                vary = listItem.item_rect(varAccion).top + int(configuracion['Elementos']['Offsety'])
                listItem.click(coords=(varx,vary))
            #Codigo para capturar el link enviar
            vinculo= driver.find_element_by_xpath('//table/tbody/tr[7]/td[3]/a')
            ubicacion_vinculo=vinculo.location
            x=ophelia_ctrl.rectangle().left + int(ubicacion_vinculo['x'])+int(configuracion['Elementos']['Offset_enviarx'])
            y=ophelia_ctrl.rectangle().top + int(ubicacion_vinculo['y'])+10
            pywin.mouse.click(button='left',coords=(x,y))
            time.sleep(1)
            # Codigo para el popup de terminar seguimiento
            if tipoProceso == 2 or tipoProceso == 3:
                time.sleep(3)
                continue
            #Código para capturar el cuadro de dialogo de siguiente ejecución
            try:
                pywin.timings.wait_until_passes(20, .5, lambda: (pywin.findwindows.find_window(class_name="TCWfGetExecDataFrm")),(WindowNotFoundError))
                wndejecucion= pywin.Application(backend="win32").connect(class_name="TCWfGetExecDataFrm")
                dlgFrm=wndejecucion.window(title_re="Selecciones en ejecución")
                dlgFrm.wait('ready',timeout=10, retry_interval=0.5)
                tpanelEjecutor= dlgFrm.child_window(class_name="TPanel", ctrl_index=2)
                ejecutorList=tpanelEjecutor.child_window(class_name="TListBox")
                ejecutorList.wait('ready',timeout=10, retry_interval=0.5)
                listItem = pywin.controls.win32_controls.ListBoxWrapper(ejecutorList.wrapper_object())
                if siguienteEjecutor.lower() == 'variable':
                    for y, res in enumerate(listItem.item_texts()):
                        if str(fila_casos[key][int(colEjecutor)-int(colListai)]).lower() in res.lower():
                            varAccion= y
                else:
                    for y, res in enumerate(listItem.item_texts()):
                        if configuracion['Elementos']['Siguiente_ejecutor'].lower() in res.lower():
                            varAccion= y
                # time.sleep(1)
                listItem.select(varAccion)
                dlgFrm.Accept.click()
                time.sleep(3)
                try:
                    wndapp.wait_cpu_usage_lower(threshold=2.5,timeout=10,usage_interval=1)
                    wndapp.wait_for_process_exit(timeout=10,retry_interval=0.5)
                except TimeoutError as err:
                    wndapp.kill()
                    # toc= time.perf_counter()
                    print('Proceso terminado anticipadamente en la etapa siguiente ejecutor. No fue posible cerrar la ventana de IE.\nLa tarea se finalizará en el Sistema Operativo')
                    # ctypes.windll.user32.MessageBoxW(0, "El robot ha finalizado anticipadamente por demora excesiva en la respuesta de SEVEN ERP. Se procesaron " + str(key+1) + " casos en " + str(datetime.timedelta(seconds=toc-tic)) + "el caso " + str(int(fila_casos[key][varCasoweb])) + " no fue procesado.", "Robot Terminado", 0)
                print('Caso ' + str(int(fila_casos[key][varCasoweb])) + ' procesado exitosamente')
            except TimeoutError as error:
                print("Se excedió el timepo de espera del cuadro de siguiente ejecutor.")
                driver.close()
                toc= time.perf_counter()
                ctypes.windll.user32.MessageBoxW(0, "El robot ha finalizado anticipadamente por demora excesiva en la respuesta de SEVEN ERP. Se procesaron " + str(key+1) + " casos en " + str(datetime.timedelta(seconds=toc-tic)) + "el caso " + str(int(fila_casos[key][varCasoweb])) + " no fue procesado.", "Robot Terminado", 0)
                exit()
            #driver.close()
            #alertIe=driver.switch_to_alert()
            #alertIe.dismiss()
            #time.sleep(1)
            #driver.close()
            #time.sleep(2)
            #break
        except TimeoutException as ex:
            print ("El caso No. " + str(int(fila_casos[key][varCasoweb]))+" ya fue realizado")
            driver.close()
    toc= time.perf_counter()
    if status_driver(driver):
        driver.quit()
    ctypes.windll.user32.MessageBoxW(0, "El robot ha finalizado. Se procesaron " + str(num_casos) + " casos en " + str(datetime.timedelta(seconds=toc-tic)), "Robot Terminado", 0)