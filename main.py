import os                  # Sistema operativo: para verificar si el archivo Excel ya existe en el disco
import json                # Traductor: para interpretar los datos en la estructura JSON que devuelve la IA
import pandas as pd        # Pandas: herramienta avanzada para manejar tablas de datos, filas y columnas para escribir en Excel
from PIL import Image      # Pillow: librería para abrir, leer y manipular la foto de la factura
from google import genai   # SDK oficial de Google para conectar nuestro código con la IA de Gemini


# ==============================================================================
# INICIALIZACIÓN Y CARGA DE IMÁGEN
# ==============================================================================

# Almaceno el cliente de la IA en la variable client, buscando la llave secreta en el entorno
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

def procesar_factura_proveedor(ruta_imagen, archivo_excel="inventario_motoline.xlsx"):
    print("Analizando la factura con IA...")
    imagen = Image.open(ruta_imagen) # Abrimos la imagen con Pillow para que la IA pueda "verla"


    # ==========================================================================
    # ENVÍO DE LA ORDEN VISUAL A LA IA Y OBTENCIÓN DE LA RESPUESTA EN JSON
    # ==========================================================================

    prompt = """Analiza esta factura de repuestos de motos. Extrae una lista con los siguientes campos para cada repuesto:
    - descripcion_original (el texto exacto que aparece en la factura)
    - nombre (aplica estas reglas de limpieza: si dice 'campana con discos de clutch' cámbialo a 'centro de clutch'...)
    - precio (número)
    - descripción (String)
    - Stock (Número entero)
    - Imagen 1(.jpg)
    - Imagen 2(.jpg)
    - Imagen 3(.jpg)
    - Imagen 4(.jpg)
    - Marca (String)
    - Categoría (String)
    - MarcaVehiculo (String)
    
    Devuelve la respuesta ÚNICAMENTE en formato JSON válido, como una lista de objetos...
    """
    
    response = client.models.generate_content(        
        model='gemini-2.5-flash',                        # Seleccionamos el modelo de IA más rápido y eficiente
        contents=[imagen, prompt]                        # Empaquetamos en una lista la imagen y las instrucciones (prompt)
    )


    # ==========================================================================
    # TRANSFORMACIÓN DE LOS DATOS Y ALMANECAMIENTO EN EL ARCHIVO DE EXCEL
    # ==========================================================================

    # 1. Limpieza de formato: Quitamos las marcas de texto ```json y ``` que la IA suele agregar por error
    texto_json = response.text.replace("```json", "").replace("```", "").strip()
    
    # 2. Traducción: Convertimos el texto limpio en una estructura de datos que Python entiende
    nuevos_datos = json.loads(texto_json)

    # 3. Conversión a Tabla: Creamos una tabla virtual de filas y columnas con Pandas
    df_nuevos = pd.DataFrame(nuevos_datos)

    # 4. Verificación y Unión: Comprobamos si ya existe un archivo Excel previo
    if os.path.exists(archivo_excel):
       df_existente = pd.read_excel(archivo_excel) # Leemos el Excel anterior para no borrar lo que ya estaba
       
       # Unimos la tabla vieja con la nueva, usando ignore_index=True para reordenar los números de fila secuencialmente (0, 1, 2...)
       df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
    else:
       df_final = df_nuevos # Si el archivo no existía, la tabla final es simplemente la nueva

    # 5. Guardado definitivo: Escribimos el resultado en el archivo Excel (sin guardar los índices numéricos automáticos de pandas)
    df_final.to_excel(archivo_excel, index=False)
    
    print(f"¡Listo! Se agregaron {len(df_nuevos)} repuestos estandarizados al Excel.")