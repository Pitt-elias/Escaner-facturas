import os              # Sistema operativo: para verificar si el archivo Excel ya existe en el disco duro y leer variables de entorno.
import json            # Traductor: para interpretar y convertir los datos de texto en la estructura JSON que devuelve la IA.
import pandas as pd        # Pandas: librería avanzada para manejar tablas de datos, filas y columnas, y escribir en Excel.
from PIL import Image      # Pillow: librería para abrir, leer y manipular la foto de la factura recibida.
from groq import Groq      # SDK oficial de Groq para conectar nuestro código con el modelo de visión Qwen en la nube.
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks # FastAPI para el servidor web y BackgroundTasks para evitar bloqueos por tiempo de espera (Timeout 504).
import io                  # IO: maneja flujos de bytes en la memoria RAM del PC para leer imágenes sin guardarlas en disco.
import base64              # Base64: codificador que transforma los bytes de la foto en una cadena gigante de texto seguro para enviarla por internet.
import socket              # Socket: librería de red nativa de Python para consultar la dirección IP local del PC.
import difflib             # Difflib: librería nativa para comparar textos de forma inteligente (coincidencia aproximada o fuzzy matching).


# ==============================================================================
# 1. INICIALIZACIÓN DEL SERVIDOR WEB Y DE LA INTELIGENCIA ARTIFICIAL (GROQ - QWEN)
# ==============================================================================

app = FastAPI(title="Motoline Backend Comercial")  # Inicializamos el servidor web creando la aplicación principal 'app' mediante FastAPI.

# Inicializamos el cliente de Groq buscando de forma segura la API key en las variables de entorno del sistema operativo.
client = Groq(api_key=os.environ.get("GROQ_API_KEY", "v"))

# Definimos el nombre fijo del archivo Excel oficial del taller donde se acumularán todos los repuestos.
ARCHIVO_EXCEL_JEFE = "inventario_motoline.xlsx"  


# ==============================================================================
# 2. CATÁLOGO MAESTRO INTELIGENTE (Estandarización de Nombres con Pandas y Difflib)
# ==============================================================================
def obtener_catalogo_maestro():
    """Lee el Excel del almacén mediante Pandas y extrae una lista única con todos los nombres 
       de repuestos oficiales que ya existen, para evitar duplicados o variaciones de la IA."""
    if os.path.exists(ARCHIVO_EXCEL_JEFE) and os.path.getsize(ARCHIVO_EXCEL_JEFE) > 0:
        try:
            # Leemos el archivo Excel oficial del taller especificando el motor openpyxl.
            df = pd.read_excel(ARCHIVO_EXCEL_JEFE, engine='openpyxl')
            if "nombre" in df.columns:
                # .dropna() quita nulos, .astype(str) pasa a texto, .str.strip() quita espacios, 
                # .unique() deja solo valores únicos y .tolist() lo vuelve una lista de Python.
                nombres = df["nombre"].dropna().astype(str).str.strip().unique().tolist() 

                # Recorremos la lista temporal 'n' asegurándonos de omitir celdas vacías o con texto "nan".
                return [n for n in nombres if n.lower() != "nan" and n != ""] 
        except Exception as e: 
            print(f"⚠️ No se pudo leer el catálogo maestro del Excel: {e}")
    return []  # Si el Excel no existe o está vacío, retornamos una lista vacía.

def estandarizar_nombre_repuesto(nombre_ia):
    """Compara el nombre que trajo la IA con el catálogo maestro del almacén utilizando difflib.
       Si encuentra un repuesto similar (ej. 'Disco clutch' vs 'Discos de clutch'),
       lo corrige automáticamente al nombre oficial registrado en el Excel."""

    # 'catalogo' almacena la lista limpia devuelta por nuestra función de arriba.
    catalogo = obtener_catalogo_maestro()
    if not catalogo: 
        return nombre_ia  # Si el catálogo está vacío, conservamos el nombre original devuelto por la IA.

    # get_close_matches busca coincidencias aproximadas. n=1 trae la mejor opción, 
    # y cutoff=0.5 exige al menos un 50% de similitud en los caracteres del texto.
    coincidencias = difflib.get_close_matches(nombre_ia, catalogo, n=1, cutoff=0.5)

    if coincidencias:
        # 'coincidencias' es una lista de resultados; [0] accede al primer elemento (la mejor coincidencia hallada).
        print(f"🔍 [Estandarización] La IA leyó '{nombre_ia}' -> Mapeado al catálogo oficial: '{coincidencias[0]}'") 
        return coincidencias[0]  # Retornamos el nombre oficial y estandarizado del almacén.

    return nombre_ia  # Si el repuesto es totalmente nuevo y no está en el catálogo, lo dejamos pasar.


# ==============================================================================
# 3. CAPA DE BLINDAJE E INTEGRIDAD DE DATOS (Validación de Stock y Precios)
# ==============================================================================
def validar_y_limpiar_repuestos(datos_crudos):    # Función que recibe los datos sin filtrar devueltos por la IA y los sanitiza.
    """Filtra y asegura la integridad de los datos para evitar errores numéricos 
       o interpretaciones erróneas (como confundir números o dejar campos vacíos)."""
    repuestos_validados = []  # Lista vacía donde iremos guardando los diccionarios limpios.

    for item in datos_crudos:  # Recorremos cada repuesto extraído de la factura uno por uno.
        # Asegurarnos de que el repuesto tenga un nombre válido.
        # item.get("nombre", "") busca la clave 'nombre'; si no está, usa un texto vacío "" por defecto.
        # .strip() elimina los espacios vacíos sobrantes al inicio y final de la cadena.
        nombre_crudo = str(item.get("nombre", "")).strip()  

        # Si el nombre está totalmente vacío o la IA escribió la palabra "null", descartamos el registro completo.
        if not nombre_crudo or nombre_crudo.lower() == "null":  
            continue  # Salta al siguiente elemento y no lo guarda.

        # Cruzamos el nombre leído con nuestro Catálogo Maestro para mantener la estandarización del almacén.
        nombre_oficial = estandarizar_nombre_repuesto(nombre_crudo)

        # Validación estricta de Precio (debe ser numérico y mayor o igual a 0).
        try:
            precio = float(item.get("precio", 0))  # Intentamos convertirlo a número decimal. Si falla, usa 0.
        except (ValueError, TypeError):  # Si la IA mandó texto ilegible (ej. "Diez mil"), atrapamos el error.
            precio = 0.0  # Valor de respaldo seguro para que el programa no colapse.

        # Validación estricta de Stock (número entero, previniendo que se guarde texto o valores nulos).
        try:
            stock = int(item.get("Stock", 1))  # Intentamos convertirlo a entero. Por defecto asigna 1 si no existe.
            if stock <= 0:    
                stock = 1  # Forzamos un stock mínimo de 1 si el valor interpretado fue cero o negativo.
        except (ValueError, TypeError):  # Si ocurre un error de tipo o valor inválido:
            stock = 1  # Valor seguro por defecto si la IA interpretó mal el número.

        # Estructuramos el diccionario final totalmente limpio y blindado para el Excel.
        repuesto_limpio = {
            "descripcion_original": str(item.get("descripcion_original", "N/D")),
            "nombre": nombre_oficial,  # Usamos el nombre ya estandarizado contra el inventario real.
            "precio": precio,
            "descripción": str(item.get("descripción", "Sin descripción")),
            "Stock": stock,
            "Imagen 1(.jpg)": str(item.get("Imagen 1(.jpg)", "")),
            "Imagen 2(.jpg)": str(item.get("Imagen 2(.jpg)", "")),
            "Imagen 3(.jpg)": str(item.get("Imagen 3(.jpg)", "")),
            "Imagen 4(.jpg)": str(item.get("Imagen 4(.jpg)", "")),
            "Marca": str(item.get("Marca", "Genérica")),
            "Categoría": str(item.get("Categoría", "General")),
            "MarcaVehiculo": str(item.get("MarcaVehiculo", "Universal"))
        }

        # Agregamos el repuesto limpio a nuestra lista utilizando .append()
        repuestos_validados.append(repuesto_limpio)  

    return repuestos_validados  # Devolvemos la lista de repuestos listos para ser ingresados al Excel.


# ==============================================================================
# 4. TAREA EN SEGUNDO PLANO (Con sistema de reintento / doble pasada si falta info)
# ==============================================================================
def tarea_procesar_factura_fondo(contenido_imagen: bytes):    
    """Procesa la imagen en segundo plano, valida la integridad de los campos obligatorios 
       como el stock y actualiza de forma persistente el Excel del taller.""" 
    try:
        print("🔄 Procesando y validando factura en segundo plano con Qwen...")

        # Codificamos los bytes de la foto a Base64 y los decodificamos a texto plano UTF-8 para enviarlos a Groq.
        imagen_base64 = base64.b64encode(contenido_imagen).decode("utf-8")   

        # Este prompt de texto se envía directamente a la IA junto con la imagen para darle la orden exacta de extracción.
        prompt = """Analiza esta factura de repuestos de motos (puede incluir texto impreso o escrito a mano). 
Extrae una lista estricta con los siguientes campos para cada repuesto. TODOS los campos son OBLIGATORIOS (si el stock o el precio no se ven del todo claros, infiere lógicamente o asigna 1 por defecto al stock, nunca dejes vacíos):    
- descripcion_original (el texto exacto que aparece en la factura)
- nombre (nombre comercial del repuesto)
- precio (número exacto)
- descripción (String)
- Stock (Número entero obligatorio, ej. 1, 2, 5...)
- Imagen 1(.jpg), Imagen 2(.jpg), Imagen 3(.jpg), Imagen 4(.jpg)
- Marca (String)
- Categoría (String)
- MarcaVehiculo (String)

Devuelve la respuesta ÚNICAMENTE en formato JSON válido, como una lista de objetos:
[
  {"descripcion_original": "...", "nombre": "...", "precio": 0.0, "descripción": "...", "Stock": 1, "Imagen 1(.jpg)": "", "Imagen 2(.jpg)": "", "Imagen 3(.jpg)": "", "Imagen 4(.jpg)": "", "Marca": "", "Categoría": "", "MarcaVehiculo": ""}
]
"""

        nuevos_datos = []
        intentos = 0
        max_intentos = 2  # Sistema de doble pasada: reintenta una segunda vez si la primera lectura deja vacíos o falla.

        # Bucle que se ejecuta mientras no alcancemos el máximo de intentos y 'nuevos_datos' siga vacío.
        while intentos < max_intentos and not nuevos_datos:  
            intentos += 1  # Incrementamos el contador de intentos.
            try:
                # Petición oficial a la API de Groq utilizando el modelo multimodal de visión Qwen.
                completion = client.chat.completions.create(  
                    model="qwen/qwen3.8-27b", # Modelo de Qwen seleccionado para OCR.
                    messages=[{ 
                        "role": "user",  # Rol de la conversación ('user' indica que nosotros damos la orden).
                        "content": [
                            # Si es el primer intento enviamos el prompt principal; si falla, enviamos un prompt correctivo de refuerzo.
                            {"type": "text", "text": prompt if intentos == 1 else "La extracción anterior dejó campos vacíos o incompletos. Por favor, reintenta asegurándote de rellenar obligatoriamente el stock y los precios de todos los repuestos visibles."},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{imagen_base64}"}} # Adjuntamos la imagen en formato Base64.
                        ],
                    }],
                    temperature=0.2,            # Temperatura baja (0.2) para que la IA sea estricta, matemática y precisa con los números (cero inventos).
                    max_completion_tokens=2048, # Límite máximo de tokens que la IA puede generar en su respuesta.
                    response_format={"type": "json_object"} # Exigimos que la respuesta de la IA venga estructurada obligatoriamente como un JSON.
                )

                # Extraemos el texto puro del mensaje que devolvió la IA dentro de las opciones (choices).
                texto_respuesta = completion.choices[0].message.content 

                # Limpiamos las marcas de formato markdown (como ```json y ```) que suelen venir en las respuestas.
                texto_json = texto_respuesta.replace("```json", "").replace("```", "").strip()

                # Traducimos el texto limpio a una estructura de datos nativa de Python mediante json.loads().
                parsed_data = json.loads(texto_json) 

                # Verificamos si la estructura devuelta es una lista directa o un diccionario contenedor.
                if isinstance(parsed_data, list):  # isinstance() comprueba si el objeto es de tipo lista (list).
                    nuevos_datos = parsed_data
                elif isinstance(parsed_data, dict): # Si es un diccionario (dict), buscamos de forma segura cualquier valor interno que sea una lista.
                    nuevos_datos = next((v for v in parsed_data.values() if isinstance(v, list)), [parsed_data])

                # Aplicamos nuestra capa de validación numérica, de tipos y estandarización por catálogo maestro.
                nuevos_datos = validar_y_limpiar_repuestos(nuevos_datos) 

            except Exception as e:    
                print(f"⚠️ Intento {intentos} falló: {e}") # Capturamos y mostramos en consola cualquier error del intento.

        # Si tras los reintentos la factura sigue siendo totalmente ilegible, generamos un registro de alerta seguro.
        if not nuevos_datos:
            print("❌ La factura no pudo ser interpretada con claridad. Registrando aviso de revisión manual.")
            nuevos_datos = [{
                "descripcion_original": "REVISAR MANUALMENTE - Factura ilegible, borrosa o ambigua",
                "nombre": "Repuesto Pendiente de Verificación",
                "precio": 0.0,
                "descripción": "La IA no pudo extraer los campos con precisión.",
                "Stock": 1,
                "Imagen 1(.jpg)": "", "Imagen 2(.jpg)": "", "Imagen 3(.jpg)": "", "Imagen 4(.jpg)": "",
                "Marca": "Revisar", "Categoría": "Revisar", "MarcaVehiculo": "Revisar"
            }]

        # Conversión de los datos limpios a un DataFrame de Pandas (una tabla virtual de filas y columnas).
        df_nuevos = pd.DataFrame(nuevos_datos)  

        # Verificación y unión acumulativa con el archivo Excel oficial del taller en el disco duro.
        # Validamos dos cosas: que el archivo exista física y lógicamente, y que su tamaño en bytes sea mayor a 0.
        if os.path.exists(ARCHIVO_EXCEL_JEFE) and os.path.getsize(ARCHIVO_EXCEL_JEFE) > 0: 
           try:
               # Leemos el Excel anterior para traer los datos guardados previamente a la memoria RAM.
               # 'engine=openpyxl' es el motor técnico que permite leer y escribir archivos .xlsx modernos.
               df_existente = pd.read_excel(ARCHIVO_EXCEL_JEFE, engine='openpyxl') 

               # Unimos la tabla vieja con la nueva usando pd.concat(). 
               # 'ignore_index=True' reinicia la numeración de filas de forma secuencial (0, 1, 2...) ignorando los índices viejos.
               df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
           except Exception: 
               df_final = df_nuevos # Si ocurre un error al leer el Excel viejo, por seguridad tomamos solo los datos nuevos.
        else:
           df_final = df_nuevos # Si el archivo no existía o estaba vacío, la tabla final es simplemente la nueva.

        # Guardado persistente escribiendo la tabla unida de vuelta en el archivo Excel del disco.
        # index=False evita guardar una columna adicional innecesaria con los números de fila.
        df_final.to_excel(ARCHIVO_EXCEL_JEFE, index=False, engine='openpyxl') 
        print(f"¡Proceso completado! Se guardaron {len(df_nuevos)} registros validados en '{ARCHIVO_EXCEL_JEFE}'.")

    except Exception as e:
        print(f"❌ Error crítico en la tarea de fondo: {str(e)}")


# ==============================================================================
# 5. RUTA (ENDPOINT) DE LA API: RECIBE LA FOTO Y RESPONDE AL INSTANTE
# ==============================================================================

@app.post("/procesar-factura/") # Registramos un endpoint web de tipo POST en la ruta /procesar-factura/.
async def procesar_factura_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):   
    # 'async def' permite manejar conexiones concurrentes eficientemente.
    # 'BackgroundTasks' nos permite delegar tareas pesadas al fondo sin bloquear la red.
    # 'UploadFile' recibe el archivo adjunto (la foto) mandado por la app de Flutter.
    try:
        # await pausa la ejecución el microsegundo exacto que toma descargar el archivo recibido por la red.
        contenido_imagen = await file.read() 

        # Delegamos el trabajo pesado (IA y Excel) a un segundo plano para evitar el error 504 de Timeout en el celular.
        background_tasks.add_task(tarea_procesar_factura_fondo, contenido_imagen) 

        print("¡Foto recibida desde el celular! Respuesta enviada al instante.")

        # Devolvemos una respuesta inmediata y exitosa al teléfono celular.
        return {
            "status": "success",
            "mensaje": "¡Factura recibida con éxito! Analizando integridad y catálogo en segundo plano."
        }

    except Exception as e:
        # Si ocurre un problema al recibir el archivo, elevamos una excepción HTTP 500 informando el fallo de red.
        raise HTTPException(status_code=500, detail=f"Error al recibir la factura: {str(e)}") 


# ==============================================================================
# 6. AUTO-DETECCIÓN DE IP PARA CONFIGURACIÓN MANUAL EN LA TERMINAL AL ARRANCAR
# ==============================================================================
def obtener_ip_local():
    """Crea un socket temporal para descubrir automáticamente la IP local del PC en la red Wi-Fi."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # No llega a enviar nada real, pero fuerza al sistema operativo a revelar la ruta de red local activa.
        s.connect(("10.255.255.255", 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = "127.0.0.1"  # Valor por defecto de respaldo si falla la red.
    finally:
        s.close()
    return IP

if __name__ == "__main__":
    import uvicorn
    ip_local = obtener_ip_local()
    puerto = 8000
    url_servidor = f"http://{ip_local}:{puerto}/procesar-factura/"

    print("\n" + "="*60)
    print(f"  SERVIDOR MOTOLINE LISTO Y FUNCIONANDO")
    print(f"  Dirección IP del Servidor: {ip_local}")
    print(f"  URL de Conexión para la App: {url_servidor}")
    print("="*60 + "\n")
    print("💡 Copie esta URL en el cuadro de configuración manual de la IP en su app de Flutter.")

    # Encendemos el servidor web Uvicorn abriéndolo a cualquier dispositivo en la red local (host='0.0.0.0').
    uvicorn.run(app, host="0.0.0.0", port=puerto)