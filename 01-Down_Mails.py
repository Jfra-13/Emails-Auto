import imaplib
import email
from email.header import decode_header
import os
import tkinter as tk
from tkinter import filedialog

# Función para seleccionar la carpeta de descarga
def seleccionar_carpeta():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de tkinter
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde deseas guardar los archivos")
    return carpeta

# Seleccionar la carpeta de destino
carpeta_destino = seleccionar_carpeta()

# Ingresar el asunto que se desea buscar
asunto_buscado = input("Ingrese el asunto del correo que desea buscar: ").strip().lower()

# Contadores para los tipos de archivos
contador_pdf = 0
contador_excel = 0
contador_word = 0

PASSWORD = "n h r j r v c bf c n r k n h w"
USER = "product.maxima.23@gmail.com"

# Credenciales de Usuario
username = USER
password = PASSWORD

# Creación de conexión
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Inicio de sesión
imap.login(username, password)
status, mensajes = imap.select("INBOX")

# Cantidad total de correos
mensajes = int(mensajes[0])
for i in range(mensajes, 0, -1):
    try:
        res, mensaje = imap.fetch(str(i), "(RFC822)")
    except:
        break
    for respuesta in mensaje:
        if isinstance(respuesta, tuple):
            # Obtener contenidos
            mensaje = email.message_from_bytes(respuesta[1])
            # Decodificar contenido
            subject = decode_header(mensaje["Subject"])[0][0]
            if isinstance(subject, bytes):
                # Convertir a string
                subject = subject.decode()
            # Verificar si el asunto coincide con el buscado
            if asunto_buscado in subject.lower():
                # Obtener origen del correo
                from_ = mensaje.get("From")
                
                print(">[Subject]: ", subject)
                print(">[From]: ", from_)
                print(">[MENSAJE OBTENIDO]: Exitoso")
                print("{-------------------------------------------}")
                # Si el correo es un HTML
                if mensaje.is_multipart():
                    # Recorremos las partes del correo
                    for part in mensaje.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        
                        # Si hay un adjunto
                        if "attachment" in content_disposition:
                            nombre_archivo = part.get_filename()
                            if nombre_archivo:
                                # Crear subcarpeta para el asunto si no existe
                                ruta_asunto = os.path.join(carpeta_destino, subject)
                                if not os.path.isdir(ruta_asunto):
                                    os.mkdir(ruta_asunto)
                                
                                # Clasificar archivos por tipo
                                extension = os.path.splitext(nombre_archivo)[1].lower()
                                if extension == ".pdf":
                                    tipo_carpeta = "pdfs"
                                    contador_pdf += 1
                                elif extension in [".xls", ".xlsx"]:
                                    tipo_carpeta = "excels"
                                    contador_excel += 1
                                elif extension in [".doc", ".docx"]:
                                    tipo_carpeta = "words"
                                    contador_word += 1
                                else:
                                    tipo_carpeta = "otros"
                                
                                # Crear subcarpeta para el tipo de archivo si no existe
                                ruta_subcarpeta = os.path.join(ruta_asunto, tipo_carpeta)
                                if not os.path.isdir(ruta_subcarpeta):
                                    os.mkdir(ruta_subcarpeta)
                                
                                # Ruta completa del archivo
                                ruta_archivo = os.path.join(ruta_subcarpeta, nombre_archivo)
                                
                                # Guardar archivo
                                open(ruta_archivo, "wb").write(part.get_payload(decode=True))

# Cierre de conexión
imap.close()
imap.logout()

# Mostrar resultados
print(f"\nDescarga completada:")
print(f"Archivos PDF descargados: {contador_pdf}")
print(f"Archivos Excel descargados: {contador_excel}")
print(f"Archivos Word descargados: {contador_word}")

