import streamlit as st
import io
import time
from datetime import datetime
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import openpyxl

class TicketGasto(BaseModel):
    fecha: str = Field(description="Fecha del ticket DD/MM/YYYY")
    codigo_postal: str = Field(description="Código postal de 5 dígitos. Búscalo por NIF o localidad si no sale.")
    localidad: str = Field(description="Nombre del municipio o localidad")
    detalle: str = Field(description="Si es bar pon 'desayuno', 'comida' o 'cena'. Si es gasolinera pon combustible.")
    importe: float = Field(description="Importe total final pagado, ej: 13.00")
    concepto: str = Field(description="SÓLO: 'COMIDAS', 'COMBUSTIBLE GASOIL', 'COMBUSTIBLE GASOLINA', 'PEAJES', 'HOTELES', 'OTROS GASTOS DE VIAJES', 'OTROS GASTOS MATERIALES Y EQUIPOS'.")

class ListaTickets(BaseModel):
    tickets: list[TicketGasto]

def convertir_fecha(fecha_str):
    try: return datetime.strptime(fecha_str, "%d/%m/%Y")
    except: return datetime(2099, 12, 31)

st.set_page_config(page_title="Gastos Electren", page_icon="🧾")
st.title("🧾 Automatizador de Notas de Gastos")
st.markdown("Sube las fotos de tus tickets y descarga el Excel listo para copiar y pegar.")

api_key = st.secrets["GEMINI_API_KEY"] if "GEMINI_API_KEY" in st.secrets else ""
fotos = st.file_uploader("Arrastra aquí los tickets", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if st.button("🚀 Procesar Gastos"):
    if not api_key:
        st.error("Error: Falta la clave API en la configuración.")
    elif not fotos:
        st.warning("⚠️ Por favor, sube al menos un ticket.")
    else:
        client = genai.Client(api_key=api_key)
        todos_los_gastos = []
        lote_tamano = 3  # Agrupa de 3 en 3 fotos por petición
        
        barra_progreso = st.progress(0, text="Iniciando procesamiento agrupado...")
        
        for i in range(0, len(fotos), lote_tamano):
            lote = fotos[i:i + lote_tamano]
            nombres_lote = ", ".join([f.name for f in lote])
            
            # Preparar las imágenes comprimidas del bloque
            contents = []
            for f in lote:
                img = Image.open(f)
                img.thumbnail((1024, 1024))
                contents.append(img)
            
            prompt = "Analiza todas estas fotos. Hay varios tickets. Extrae la información de cada uno de ellos por separado. Para concepto usa SÓLO mayúsculas según la lista estricta."
            contents.append(prompt)
            
            exito = False
            for intento in range(3):
                try:
                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=contents,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=ListaTickets,
                            temperature=0.1,
                        ),
                    )
                    if response.parsed and response.parsed.tickets:
                        for t in response.parsed.tickets:
                            desc = f"{t.codigo_postal} - {t.localidad}, {t.detalle}"
                            todos_los_gastos.append({
                                "fecha": t.fecha, "fecha_dt": convertir_fecha(t.fecha),
                                "descripcion": desc, "importe": t.importe, "concepto": t.concepto
                            })
                        exito = True
                        break
                    else:
                        raise Exception("No se detectaron tickets válidos en este grupo.")
                except Exception as e:
                    error_str = str(e).upper()
                    if "503" in error_str or "429" in error_str or "UNAVAILABLE" in error_str:
                        st.toast(f"⏳ Servidor saturado. Reintentando bloque ({nombres_lote}) en 10s...", icon="🔄")
                        time.sleep(10)
                    else:
                        st.error(f"Error procesando {nombres_lote}: {e}")
                        break
            
            porcentaje = min(1.0, (i + lote_tamano) / len(fotos))
            barra_progreso.progress(porcentaje, text=f"Procesando bloque {min(i+lote_tamano, len(fotos))} de {len(fotos)} tickets...")
            time.sleep(1)

        if todos_los_gastos:
            todos_los_gastos.sort(key=lambda x: (x["concepto"], x["fecha_dt"]))
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Datos Extraidos"
            ws.append(["FECHA", "DESCRIPCIÓN", "IMPORTE", "CONCEPTO"])
            for gasto in todos_los_gastos:
                ws.append([gasto["fecha"], gasto["descripcion"], gasto["importe"], gasto["concepto"]])
            for row in range(2, len(todos_los_gastos) + 2):
                ws[f"C{row}"].number_format = '#,##0.00'
            
            output = io.BytesIO()
            wb.save(output)
            
            st.success(f"🎉 ¡Proceso finalizado! Se han extraído {len(todos_los_gastos)} gastos.")
            st.download_button(
                label="📥 Descargar Excel para Copiar y Pegar",
                data=output.getvalue(),
                file_name="TICKETS_LISTOS_PARA_PEGAR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se ha podido extraer ningún gasto. Revisa que las fotos sean legibles o que no se haya agotado el cupo temporal.")
