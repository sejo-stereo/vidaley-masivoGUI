import tkinter as tk
import customtkinter
from tkinter import filedialog

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException

import pandas as pd

from PIL import Image
import requests
from io import BytesIO

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("dark-blue")

url_vida_ley = "https://api-seguridad.sunat.gob.pe/v1/clientessol/b3639111-1546-4d06-b74f-de2c40629748/oauth2/login?originalUrl=https://apps.trabajo.gob.pe/si.segurovida/index.jsp&state=m1ntr4"
chrome_options = Options()
chrome_options.add_experimental_option("detach",True)
driver = webdriver.Chrome(options=chrome_options)
driver.implicitly_wait(3)
driver.get(url_vida_ley)


def seleccionar_excel():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=(("Archivos de texto", "*.xlsx"), ("Todos los archivos", "*.*")))
    if archivo:
        print(f"seleccion: {archivo}")
        excel_label.configure(text=archivo)
    
def eliminar_trabajadores():
    ceses_df = pd.read_excel(excel_label.cget("text"),sheet_name="CESES",dtype={"NRO_DOC":str})
    for i in range(len(ceses_df)):
        dni = ceses_df.loc[i,"NRO_DOC"]
        driver.find_element(By.NAME,"v_codtrabus").clear()
        driver.find_element(By.NAME,"v_codtrabus").send_keys(dni)
        driver.find_element(By.XPATH,'//a[@href="javascript:buscarTraLisBD();"]').click()
        try:
            result_table = driver.find_element(By.XPATH,'//table[@id="lstPolizaTrabajador"]//tr/td[1]/input')
            result_table.click()
            driver.find_element(By.XPATH,'//img[@src="/si.segurovida/util/images/botones/eliminar.gif"]').click()
            driver.switch_to.alert.accept()
        except NoSuchElementException:
            print(f"{dni}: No se encontro trabajador.")
            with open("log_ceses.txt","a+") as f:
                f.write(f"{dni}: No se encontro trabajador.\n")

def agregar_trabajadores():
    ingresos_df = pd.read_excel(excel_label.cget("text"),sheet_name="INGRESOS",dtype={"NRO_DOC":str})
    for i in range(len(ingresos_df)):
        tipo_doc = ingresos_df.loc[i,"TIPO_DOC"]
        nro_doc = ingresos_df.loc[i,"NRO_DOC"]
        reingreso = ingresos_df.loc[i,"REINGRESO"]
        ex_trabajador = ingresos_df.loc[i,"SEGURO_EX_TRABAJADOR"]
        fecha_aseguramiento = ingresos_df.loc[i,"FECHA_ASEGURAMIENTO"].strftime("%d%m%Y")
        tipo_moneda = ingresos_df.loc[i,"TIPO_MONEDA"]
        monto_asegurable = str(ingresos_df.loc[i,"MONTO_REM_ASEGURABLE"])
        
        Select(driver.find_element(By.ID,"v_codtdocide")).select_by_visible_text(tipo_doc)
        driver.find_element(By.ID,"v_codtra").clear()
        driver.find_element(By.ID,"v_codtra").send_keys(nro_doc,Keys.ENTER)
        driver.find_element(By.XPATH, f"//input[@type='radio' and @name='v_flgreing' and @value='{reingreso}']").click()
        driver.find_element(By.XPATH, f"//input[@type='radio' and @name='v_flgcontseg' and @value='{ex_trabajador}']").click()
        driver.find_element(By.ID,"d_fecasetra").clear()
        driver.find_element(By.ID,"d_fecasetra").send_keys(fecha_aseguramiento)
        Select(driver.find_element(By.NAME,"v_codtmon")).select_by_visible_text(tipo_moneda)
        driver.find_element(By.NAME,"n_monrem").clear()
        driver.find_element(By.NAME,"n_monrem").send_keys(monto_asegurable)
        
        # try:
        #     driver.find_element(By.ID,"v_aceptaingreso").click()
        # except NoSuchElementException:
        #     pass
            
        driver.find_element(By.XPATH,"//a[@href='javascript:grabarPolizaxTra();']").click()

        try:
            alert_text = driver.switch_to.alert.text
            with open("log_ingresos.txt","a") as f:
                f.write(f"{nro_doc}: {alert_text}\n") 
            driver.switch_to.alert.accept()
        except NoAlertPresentException:
            pass

def cerrar_chrome():
    driver.quit()

def get_plin_image():
    r = requests.get("https://i.ibb.co/fQX3QWj/img-plin.jpg")
    if r.status_code == 200:
        image = Image.open(BytesIO(r.content))
        return image
    else:
        return None

app = customtkinter.CTk()
app.geometry("560x700")
app.title("MasivoVidaLey-v0.1.0")
app.resizable(width=False,height=False)

title = customtkinter.CTkLabel(app,text="REGISTRO MASIVO DE TRABAJADORES VIDA LEY\n Creado por José Melgarejo",font=('Helvetica', 15))
title.pack(padx=10,pady=10)

excel_label = customtkinter.CTkLabel(app,text="",width=400,wraplength=300)
excel_label.pack(pady=10)

select_file_button = customtkinter.CTkButton(app,text="1. Seleccionar archivo Excel",command=seleccionar_excel)
select_file_button.pack(padx=10,pady=10)

borrar_trabajador_button = customtkinter.CTkButton(app,text="2. Eliminar Trabajadores",command=eliminar_trabajadores)
borrar_trabajador_button.pack(pady=10)

agregar_trabajador_button = customtkinter.CTkButton(app,text="3. Agregar Trabajadores",command=agregar_trabajadores)
agregar_trabajador_button.pack(pady=10)

cerrar_chrome_button = customtkinter.CTkButton(app,text="Cerrar Chrome",command=cerrar_chrome)
cerrar_chrome_button.pack(pady=10)

plin_label = customtkinter.CTkLabel(app,text="Es una herramienta gratuita pero cualquier apoyo sera bienvenido :D")
plin_label.pack(pady=10)

# plin_img = customtkinter.CTkImage(light_image=Image.open("img_plin.jpeg"),dark_image=Image.open("img_plin.jpeg"),size=(200,235))
plin_img = customtkinter.CTkImage(light_image=get_plin_image(),dark_image=get_plin_image(),size=(200,235))
image_label = customtkinter.CTkLabel(app,image=plin_img,text="")
image_label.pack(pady=10)

info_label = customtkinter.CTkLabel(app,text="Cualquier duda o sugerencia pueden escribir por:\n Linkedin: https://www.linkedin.com/in/jose-melgarejo/ \n Correo: josemelgarejo88@gmail.com ")
info_label.pack(pady=10)

app.mainloop()