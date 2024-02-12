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
import sys
import os
from PIL import Image
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("dark-blue")

driver = None
apoyo_window = None

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def start_chrome():
    url_vida_ley = "https://api-seguridad.sunat.gob.pe/v1/clientessol/b3639111-1546-4d06-b74f-de2c40629748/oauth2/login?originalUrl=https://apps.trabajo.gob.pe/si.segurovida/index.jsp&state=m1ntr4"
    chrome_options = Options()
    chrome_options.add_experimental_option("detach",True)
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(5)
    driver.get(url_vida_ley)
    return driver

def seleccionar_excel():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=(("Archivos de texto", "*.xlsx"), ("Todos los archivos", "*.*")))
    if archivo:
        # print(f"seleccion: {archivo}")
        excel_path.configure(text=archivo)
        log_eliminados_label.configure(text="")
        log_ingresos_label.configure(text="")
    
def eliminar_trabajador(dni):
    try:
        dni_input = driver.find_element(By.NAME,"v_codtrabus")
        dni_input.clear()
        dni_input.send_keys(dni)

        driver.find_element(By.XPATH,'//a[@href="javascript:buscarTraLisBD();"]').click()

        result_table = driver.find_element(By.XPATH,'//table[@id="lstPolizaTrabajador"]//tr/td[1]/input')
        result_table.click()

        eliminar_button = driver.find_element(By.XPATH,'//img[@src="/si.segurovida/util/images/botones/eliminar.gif"]')
        eliminar_button.click()

        driver.switch_to.alert.accept()
        
        return None
    except NoSuchElementException:
        return(f"{dni}: No se encontró trabajador.")

def eliminar_trabajadores():
    ceses_df = pd.read_excel(excel_path.cget("text"), sheet_name="CESES", dtype={"NRO_DOC": str})
    eliminar_errores = []

    for _, row in ceses_df.iterrows():
        dni = row["NRO_DOC"]
        error = eliminar_trabajador(dni)
        if error is not None:
            eliminar_errores.append(error)

    if eliminar_errores:
        with open("log_errores_ceses.txt", "w") as f:
            f.write("\n".join(eliminar_errores))

    total_eliminar = len(ceses_df)
    total_eliminar_errores = len(eliminar_errores)
    total_eliminados = total_eliminar - total_eliminar_errores

    log_eliminados_label.configure(text=f"FINALIZADO: Total: {total_eliminar} - Eliminados: {total_eliminados} - Errores: {total_eliminar_errores}")

def agregar_trabajador(tipo_doc,nro_doc,reingreso,ex_trabajador,fecha_aseguramiento,tipo_moneda,monto_asegurable):
    final_log = None
    try:
        Select(driver.find_element(By.ID,"v_codtdocide")).select_by_visible_text(tipo_doc)
        driver.find_element(By.ID,"v_codtra").clear()
        driver.find_element(By.ID,"v_codtra").send_keys(nro_doc,Keys.ENTER)

        try:
            dni_alert_text = driver.switch_to.alert.text
            driver.switch_to.alert.accept()
            final_log = dni_alert_text
            
        except NoAlertPresentException:
            driver.find_element(By.XPATH, f"//input[@type='radio' and @name='v_flgreing' and @value='{reingreso}']").click()
            driver.find_element(By.XPATH, f"//input[@type='radio' and @name='v_flgcontseg' and @value='{ex_trabajador}']").click()
            driver.find_element(By.ID,"d_fecasetra").clear()
            driver.find_element(By.ID,"d_fecasetra").send_keys(fecha_aseguramiento)
            Select(driver.find_element(By.NAME,"v_codtmon")).select_by_visible_text(tipo_moneda)
            driver.find_element(By.NAME,"n_monrem").clear()
            driver.find_element(By.NAME,"n_monrem").send_keys(monto_asegurable)
            driver.find_element(By.XPATH,"//a[@href='javascript:grabarPolizaxTra();']").click()

            try:
                alert_grabar = driver.switch_to.alert
                grabar_msg = alert_grabar.text
                alert_grabar.accept()
                final_log = grabar_msg
            except NoAlertPresentException:
                pass
                # final_log = f"Registro exitoso."
    except:
        final_log = None
    return final_log

def agregar_trabajadores():
    ingresos_df = pd.read_excel(excel_path.cget("text"),sheet_name="INGRESOS",dtype={"NRO_DOC":str})
    log_ingresos = []

    for i in range(len(ingresos_df)):
        tipo_doc = ingresos_df.loc[i,"TIPO_DOC"]
        nro_doc = ingresos_df.loc[i,"NRO_DOC"]
        reingreso = ingresos_df.loc[i,"REINGRESO"]
        ex_trabajador = ingresos_df.loc[i,"SEGURO_EX_TRABAJADOR"]
        fecha_aseguramiento = ingresos_df.loc[i,"FECHA_ASEGURAMIENTO"].strftime("%d%m%Y")
        tipo_moneda = ingresos_df.loc[i,"TIPO_MONEDA"]
        monto_asegurable = str(ingresos_df.loc[i,"MONTO_REM_ASEGURABLE"])
        
        log = agregar_trabajador(tipo_doc,nro_doc,reingreso,ex_trabajador,fecha_aseguramiento,tipo_moneda,monto_asegurable)
        
        if log:
            log_ingresos.append(f"{nro_doc}: {log}")

    with open("log_errores_ingresos.txt","w") as f:
        f.write("\n".join(log_ingresos))

    total_agregar = len(ingresos_df)
    total_agregar_errores = len(log_ingresos)
    total_agregados = total_agregar - total_agregar_errores

    log_ingresos_label.configure(text=f"FINALIZADO: Total: {total_agregar} - Registrados: {total_agregados} - Errores: {total_agregar_errores}")

def cerrar_chrome():
    global driver
    if driver is not None:
        driver.quit()

def login_ruc():
    ruc_file = "ruc_data.txt"

    if ruc_file in os.listdir(os.getcwd()):
    
        with open(ruc_file,"r") as f:
            accesos = f.readlines()
        
        ruc,usuario,contrasena = accesos[0],accesos[1],accesos[2]

        driver.find_element(By.ID,"txtRuc").send_keys(ruc)
        driver.find_element(By.ID,"txtUsuario").send_keys(usuario)
        driver.find_element(By.ID,"txtContrasena").send_keys(contrasena)

        driver.find_element(By.ID,"btnAceptar").click()

def open_apoyo_window():

    global apoyo_window

    if apoyo_window is None or not apoyo_window.winfo_exists():
        apoyo_window = customtkinter.CTkToplevel()
        apoyo_window.geometry("300x350")
        apoyo_window.title("Apoya el proyecto")
        apoyo_window.resizable(width=False,height=False)

        apoyo_label = customtkinter.CTkLabel(apoyo_window,text="Gracias por apoyar al mantenimiento\ny actualización del proyecto.")
        apoyo_label.pack(padx=10,pady=10)

        plin_img = customtkinter.CTkImage(light_image=Image.open(resource_path("plin_img.jpg")),size=(200,276))
        support_label = customtkinter.CTkLabel(apoyo_window,image=plin_img,text="")
        support_label.pack(pady=10,padx=10)

        apoyo_window.grab_set()
    
    else:
        apoyo_window.focus()

driver = start_chrome()

login_ruc()

app = customtkinter.CTk()
app.geometry("450")
app.title("MasivoVidaLey 1.0")
app.resizable(width=False,height=False)

app_title = customtkinter.CTkLabel(app,text="REGISTRO MASIVO DE TRABAJADORES VIDA LEY",font=('Helvetica', 15,"bold"))
app_title.pack(padx=10)

developer_label = customtkinter.CTkLabel(app,text="Desarrollado por José Melgarejo",font=('Helvetica', 12))
developer_label.pack()

frame_excel = customtkinter.CTkFrame(app,width=350,height=50,border_color=("black","white"))
frame_excel.pack(padx=10,pady=5)

excel_path = customtkinter.CTkLabel(frame_excel,text="",width=350,wraplength=300)
excel_path.pack(pady=5)

select_file_button = customtkinter.CTkButton(app,text="1. Seleccionar archivo Excel",command=seleccionar_excel,width=200)
select_file_button.pack(pady=5)

borrar_trabajador_button = customtkinter.CTkButton(app,text="2. Eliminar Trabajadores",command=eliminar_trabajadores,fg_color="firebrick3",hover_color="firebrick2",width=200)
borrar_trabajador_button.pack(pady=5)

log_eliminados_label = customtkinter.CTkLabel(app,text="")
log_eliminados_label.pack(pady=5)

agregar_trabajador_button = customtkinter.CTkButton(app,text="3. Agregar Trabajadores",command=agregar_trabajadores,fg_color="SpringGreen4",hover_color="SpringGreen3",width=200)
agregar_trabajador_button.pack(pady=5)

log_ingresos_label = customtkinter.CTkLabel(app,text="")
log_ingresos_label.pack(pady=5)

cerrar_chrome_button = customtkinter.CTkButton(app,text="Cerrar Chrome",command=cerrar_chrome,width=200,fg_color="gray63",hover="gray62")
cerrar_chrome_button.pack(pady=5)

info_label = customtkinter.CTkLabel(app,text="Cualquier duda o sugerencia pueden escribir a:\nLinkedin: in/jose-melgarejo\nCorreo: josemelgarejo88@gmail.com")
info_label.pack(pady=10,padx=10)

apoyar_button = customtkinter.CTkButton(app,text="Apoya el Proyecto",command=open_apoyo_window)
apoyar_button.pack(pady=10,padx=10,)

app.mainloop()



