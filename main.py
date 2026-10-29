import os                  # Sistema operativo: para verificar si el archivo Excel ya existe en el disco duro
import json                # Traductor: para interpretar los datos en la estructura JSON que devuelve la IA
import pandas as pd        # Pandas: herramienta avanzada para manejar tablas de datos, filas y columnas para escribir en Excel
from PIL import Image      # Pillow: librería para abrir, leer y manipular la foto de la factura
from groq import Groq      # SDK oficial de Groq para conectar con el modelo de visión Qwen en la nube
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks # FastAPI para el servidor y BackgroundTasks para evitar bloqueos (Timeout 504)
import io                  # IO: maneja flujos de datos y bytes en la memoria RAM del PC
import base64              # Base64: codificador necesario para enviar imágenes a través de la API de Groq


# ==============================================================================
# INICIALIZACIÓN DEL SERVIDOR WEB Y DE LA INTELIGENCIA ARTIFICIAL (GROQ - QWEN)
# ==============================================================================

app = FastAPI(title="Motoline Backend Blindado de Facturas")

# Inicializamos el cliente de Groq (puedes configurar tu API key directamente o vía variable de entorno)
client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

# Definimos el nombre del archivo Excel oficial del taller donde se guardarán los repuestos
ARCHIVO_EXCEL_JEFE = "inventario_motoline.xlsx"


# ==============================================================================
# CAPA DE BLINDAJE E INTEGRIDAD DE DATOS (Validación de Stock y Precios)
# ==============================================================================
def validar_y_limpiar_repuestos(datos_crudos):
    """Filtra y asegura la integridad de los datos para evitar errores numéricos 
       o interpretaciones erróneas (como confundir números o dejar campos vacíos)."""
    repuestos_validados = []
    
    for item in datos_crudos:
        # Asegurar que el repuesto tenga un nombre válido
        nombre = str(item.get("nombre", "")).strip()
        if not nombre or nombre.lower() == "null":
            continue 
            
        # Validación estricta de Precio (numérico y mayor o igual a 0)
        try:
            precio = float(item.get("precio", 0))
        except (ValueError, TypeError):
            precio = 0.0
            
        # Validación estricta de Stock (número entero, previene que se guarde texto o valores nulos)
        try:
            stock = int(item.get("Stock", 1))
            if stock <= 0: 
                stock = 1 
        except (ValueError, TypeError):
            stock = 1  # Valor seguro por defecto si la IA interpretó mal

        repuesto_limpio = {
            "descripcion_original": str(item.get("descripcion_original", "N/D")),
            "nombre": nombre,
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
        repuestos_validados.append(repuesto_limpio)
        
    return repuestos_validados


# ==============================================================================
# TAREA EN SEGUNDO PLANO (Con sistema de reintento / doble pasada si falta info)
# ==============================================================================
def tarea_procesar_factura_fondo(contenido_imagen: bytes):
    """Procesa la imagen en segundo plano, valida la integridad de los campos obligatorios 
       como el stock y actualiza de forma persistente el Excel del taller."""
    try:
        print("🔄 Procesando y validando factura en segundo plano con Qwen...")
        imagen_base64 = base64.b64encode(contenido_imagen).decode("utf-8")

        # Prompt estricto exigiendo campos obligatorios (incluso en facturas escritas a mano)
        prompt = """Analiza esta factura de repuestos de motos (puede incluir texto impreso o escrito a mano). 
Extrae una lista estricta con los siguientes campos para cada repuesto. TODOS los campos son OBLIGATORIOS (si el stock o el precio no se ven del todo claros, infiere lógicamente o asigna 1 por defecto al stock, nunca dejes vacíos):
- descripcion_original (el texto exacto que aparece en la factura)
- nombre (aplica reglas de limpieza: ej. 'campana con discos de clutch' -> 'centro de clutch')
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
        max_intentos = 2  # Doble pasada de seguridad si la primera lectura deja vacíos o falla

        while intentos < max_intentos and not nuevos_datos:
            intentos += 1
            try:
                completion = client.chat.completions.create(
                    model="qwen/qwen3.8-27b",
                    messages=[{
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt if intentos == 1 else "La extracción anterior dejó campos vacíos o incompletos. Por favor, reintenta asegurándote de rellenar obligatoriamente el stock y los precios de todos los repuestos visibles."},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{imagen_base64}"}}
                        ],
                    }],
                    temperature=0.2, # Temperatura baja para garantizar precisión numérica y analítica
                    max_completion_tokens=2048,
                    response_format={"type": "json_object"}
                )

                texto_respuesta = completion.choices[0].message.content
                texto_json = texto_respuesta.replace("```json", "").replace("```", "").strip()
                parsed_data = json.loads(texto_json)
                
                # Extraemos la lista de repuestos de la respuesta JSON
                if isinstance(parsed_data, list):
                    nuevos_datos = parsed_data
                elif isinstance(parsed_data, dict):
                    nuevos_datos = next((v for v in parsed_data.values() if isinstance(v, list)), [parsed_data])
                    
                # Aplicamos el filtro de integridad numérica y tipos de datos
                nuevos_datos = validar_y_limpiar_repuestos(nuevos_datos)

            except Exception as e:
                print(f"⚠️ Intento {intentos} falló: {e}")

        # Si tras los reintentos la factura es totalmente ilegible, generamos un registro de alerta seguro
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

        # Conversión a DataFrame de Pandas
        df_nuevos = pd.DataFrame(nuevos_datos)

        # Verificación y unión acumulativa con el Excel oficial del taller
        if os.path.exists(ARCHIVO_EXCEL_JEFE) and os.path.getsize(ARCHIVO_EXCEL_JEFE) > 0:
           try:
               df_existente = pd.read_excel(ARCHIVO_EXCEL_JEFE, engine='openpyxl')
               df_final = pd.concat([df_existente, df_nuevos], ignore_index=True)
           except Exception:
               df_final = df_nuevos
        else:
           df_final = df_nuevos

        # Guardado persistente especificando openpyxl
        df_final.to_excel(ARCHIVO_EXCEL_JEFE, index=False, engine='openpyxl')
        print(f"¡Proceso completado! Se guardaron {len(df_nuevos)} registros validados en '{ARCHIVO_EXCEL_JEFE}'.")

    except Exception as e:
        print(f"❌ Error crítico en la tarea de fondo: {str(e)}")


# ==============================================================================
# RUTA (ENDPOINT) DE LA API: RECIBE LA FOTO Y RESPONDE AL INSTANTE
# ==============================================================================

@app.post("/procesar-factura/")
async def procesar_factura_endpoint(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    try:
        contenido_imagen = await file.read()
        
        # Delegamos el trabajo pesado y la validación al fondo para evitar el error 504 de Timeout
        background_tasks.add_task(tarea_procesar_factura_fondo, contenido_imagen)
        
        print("¡Foto recibida desde el celular! Respuesta enviada al instante.")

        return {
            "status": "success",
            "mensaje": "¡Factura recibida con éxito! Analizando integridad y campos obligatorios en segundo plano."
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al recibir la factura: {str(e)}")


# ==============================================================================
# INSTRUCCIÓN PARA CORRER EL SERVIDOR LOCALMENTE EN REPLIT
# ==============================================================================
# uvicorn main:app --host 0.0.0.0 --reload