# vidaley-masivoGUI
La aplicación **MasivoVidaLey** permite dar de alta y baja de trabajadores masivamente en el portal de REGISTRO OBLIGATORIO DE TRABAJADORES DEL SEGURO DE VIDA LEY.
La aplicación cuenta con una interfaz gráfica para facilitar su uso.

Video Explicación: https://www.youtube.com/watch?v=efqmRvPqcTo

![gui_image_1 0](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/460f8f2c-96cf-46b8-a11b-0397236e33ef)

## Funciones

Antes de agregar o eliminar trabajadores debe completar el archivo Excel *excel_demo.xlsx*. Solo registrar a trabajadores cuyo documento de identidad sea DNI. La opción de registrar personal con otro documento de identidad será agregada en una próxima actualización.

### 1. Agregar trabajadores

- La aplicación registra en el portal a los trabajadores registrados en la pestaña *INGRESOS* del archivo excel.<br>
- Si se presenta algún error en el registro de trabajadores se creará un archivo ***log_errores_ingresos.txt*** en la misma carpeta donde se encuentra la aplicación.<br>
- Dentro del archivo .txt encontrará el DNI y el error correspondiente.

![gui_ingresos](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/679bb801-8e3c-4ee5-96e7-8f84c81cc1d9)

### 2. Eliminar trabajadores

- La aplicación elimina en el portal a los trabajadores registrados en la pestaña *CESES* del archivo excel.<br>
- Si se presenta algún error en la eliminación de trabajadores se creará un archivo ***log_errores_ceses.txt*** en la misma carpeta donde se encuentra la aplicación.<br>
- Dentro del archivo .txt encontrará el DNI y el error correspondiente.

![gui ceses](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/18ba4a0d-02bb-4812-9641-871e5c3f9d01)

### 3. Complementarias

#### Login autómatico de Clave Sol Sunat

Si desea que la aplicación ingrese automáticamente con sus accesos a Clave Sol Sunat debe crear un archivo ***ruc_data.txt*** donde indiqué RUC, USUARIO y CONTRASEÑA.
No debe registrar más líneas.

![gui_ruc_data](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/bc9b1fdf-0814-4d79-ad75-06a0b7c32435)

## Instrucciones de uso
1. Descargar desde ***Releases*** el lanzamiento más reciente.
2. Descargar y completar el archivo excel ***excel_demo.xlsx***.
   - INGRESOS: registrar todos los trabajadores que se darán de ALTA en la póliza.
   - CESES: registrar todos los trabajadores a ELIMINAR de la póliza
3. Abrir la aplicación descargada en el 1er paso.
4. Al abrir la aplicación se abriran hasta 3 ventanas.
   - Simbolo del sistema (ventana negra) **NO CERRAR**
   - Google Chrome
   - Programa MasivoVidaLey-0.1.0
5. En la ventana de Google Chrome, ingresar hasta la opción que desee (REGISTRAR NUEVA PÓLIZA o RENOVAR PÓLIZA). Complete todos los campos necesarios y solo deje la parte de registrar o eliminar trabajadores.<br>
   Aseguresé de marcar la siguiente opción.

   ![image](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/a7f79a59-8c57-4676-a2cd-739e2f5fbbd0)

6. Si cumplio hasta el paso 5, ya puede usar la aplicación en el orden que se indica.
   1. Seleccionar archivo excel
   2. Agregar Trabajadores
   3. Eliminar trabajadores
  
   La aplicación puede que deje de responder durante su ejecución pero seguirá con el registro en segundo plano. No apagar el equipo y aseguré buena conexión a internet.

7. Si se proceso correctamente la alta/baja se mostrara la cantidad de trabajadores procesados y la cantidad de errores.

![image](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/7192213d-7502-4921-a37d-3cf295608c85)

8. De presentar algún error no contemplado, el mensaje de error se mostrará en la ventana negra. Para reportar error envié una captura y una descripción a mi correo o Linkedin. Trataré de resolverlo en una siguiente actualización. 

![image](https://github.com/sejo-stereo/vidaley-masivoGUI/assets/51570964/64d16b50-4cc7-49d3-bf45-712a6282281a)

## PRÓXIMAS ACTUALIZACIONES

- Registro de Beneficiarios por trabajador
- Se actualizará semanalmente para la solución de errores.
- Para más información pueden contactarme por LinkedIn.

## VIDEO DEMOSTRACIÓN

### PRESENTACIÓN
https://www.youtube.com/watch?v=3J9sW6vPZmA

### DEMOSTRACIÓN CORTA

https://youtu.be/VO_5TTiJLCs






