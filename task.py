"""
proyecto busca facturas de tigo
Template robot with Python.
"""
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import time

lib = Files()
browser = Selenium()

cc=None
correo=None



def Consulta():
#leemos datos a consultar
   lib.open_workbook("DtTigo.xlsx")        #ubicacion del libro
   lib.read_worksheet("Cliente")           #nombre de la hoja
   lib.read_worksheet(header=True)         #activamos las cabezeras 
   lista=lib.read_worksheet_as_table(name="Cliente",header=True, start=1).data
   for x in lista:
       
       cc=x[0]
       print("Consultamos cc ")
       print(cc)
       correo=x[1]
       print("Consultamos correo ")
       print(correo)
#abrimos navegador y hacemos la consulta    
       browser.open_available_browser("https://transacciones.tigo.com.co/servicios/facturas")
       time.sleep(3)
       browser.maximize_browser_window()
       print("abrimos pagina web de tigo → https://transacciones.tigo.com.co/servicios/facturas")
       time.sleep(20)

       print("La pagina esta funcionado")

       browser.click_element("//LABEL[@for='edit-radios-0'][text()='Por documento']/self::LABEL")
       print("selecionamos radio button de (por documento)")
       time.sleep(2)
       browser.click_element("//SELECT[@id='edit-document-type']/self::SELECT")
       print("Seleccionamos → "+browser.get_text("//*[@id='edit-document-type']/option[1]"))
       browser.click_element("//*[@id='edit-document-type']/option[1]")
       time.sleep(2)
       browser.input_text("//INPUT[@id='edit-document']/self::INPUT",cc)
       print("introduccimos el numero de documento")
       time.sleep(2)
       browser.input_text("//INPUT[@id='edit-email-home']/self::INPUT",correo)
       print("introduccimos el email")
       time.sleep(10)
       browser.click_element_if_visible("//INPUT[@id='edit-consult-home']/self::INPUT")
       print("esperando que termine de procesar la consulta")
       time.sleep(25)
       try:
        
        CapturaValor=browser.get_text("//*[@id='content']/section")
        print("Capturamos datos")
        print(CapturaValor)
       except:
        print("error en la captura de datos de factura")
        CapturaValor="error en la captura de datos de factura"
        
       
       print("Registramos valores consultados ")
       
   
   
   browser.close_all_browsers
   lib.save_workbook
   lib.close_workbook


def minimal_task():
    print("Done.")


if __name__ == "__main__":
   Consulta()
   minimal_task()