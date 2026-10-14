import streamlit as st
import io
import time
from datetime import datetime
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import openpyxl

# Estructura exigida por el Excel de la empresa
class TicketGasto(BaseModel):
    fecha: str = Field(description="Fecha del ticket DD/MM/YYYY")
    codigo_postal: str = Field(description="Código postal de 5 dígitos. Búscalo por NIF o localidad si no sale.")
    localidad: str = Field(description="Nombre del municipio o localidad")
    detalle: str = Field(description="Si es bar pon 'desayuno', 'comida' o 'cena'. Si es gasolinera pon combustible.")
    importe: float = Field(description="Importe total final pagado con decimales (sin símbolo €), ej: 13.00")
    concepto: str = Field(description="SÓLO: 'COMIDAS', 'COMBUSTIBLE GASOIL', 'COMBUSTIBLE GASOLINA', 'PEAJES', 'HOTELES', 'OTROS GASTOS DE VIAJES', 'OTROS GASTOS MATERIALES Y EQUIPOS'.")

class ListaTickets(BaseModel):
    tickets: list[TicketGasto]

def convertir_fecha(fecha_str):
    try: return datetime.strptime(fecha_str, "%d/%m/%Y")
    except: return datetime(2099, 12, 31)

# Configuración de Streamlit
st.set_page_config(page_title="Gastos Electren", page_icon="🧾")
st.title("🧾 Automatizador de Notas de Gastos")
st.markdown("Sube las fotos de tus tickets y descarga el Excel listo para copiar y pegar.")

api_key = st.secrets["GEMINI_API_KEY"] if "GEMINI_API_KEY" in st.secrets else ""
fotos = st.file_uploader("Arrastra aquí los tickets (puedes subir todos los que quieras a la vez)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if st.button("🚀 Procesar Gastos"):
    if not api_key:
        st.error("Error: Falta la clave API en Secrets.")
    elif not fotos:
        st.warning("⚠️ Por favor, sube al menos un ticket.")
    else:
        client = genai.Client(api_key=api_key)
        todos_los_gastos = []
        tickets_fallidos = []
        
        barra_progreso = st.progress(0, text="Iniciando lectura ultrarrápida...")
        
        for i, foto in enumerate(fotos):
            img = Image.open(foto)
            img.thumbnail((1024, 1024)) # Optimizar tamaño para subir en milisegundos
            
            prompt = "Extrae los gastos de este ticket. Para concepto usa SÓLO mayúsculas según la lista estricta proporcionada."
            exito = False
            error_msg = ""
            
            # Intento de lectura
            for intento in range(2):
                try:
                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=[img, prompt],
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
                except Exception as e:
                    error_msg = str(e)
                    time.sleep(1)
            
            if not exito:
                tickets_fallidos.append(f"{foto.name} ({error_msg[:100]})")
            
            barra_progreso.progress((i + 1) / len(fotos), text=f"Procesando {i+1} de {len(fotos)}: {foto.name}...")

        # Avisos de resultados
        if tickets_fallidos:
            st.error("⚠️ No se pudieron leer estos tickets:")
            for tf in tickets_fallidos:
                st.write(f"- {tf}")
        
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
            
            st.success(f"🎉 ¡Proceso finalizado! Se han extraído {len(todos_los_gastos)} gastos correctamente.")
            st.download_button(
                label="📥 Descargar Excel para Copiar y Pegar",
                data=output.getvalue(),
                file_name="TICKETS_LISTOS_PARA_PEGAR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
