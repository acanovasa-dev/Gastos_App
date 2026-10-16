import streamlit as st
import io
import time
from datetime import datetime
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import openpyxl

# LISTA OFICIAL Y EXACTA DE CONCEPTOS DE ELECTREN
CONCEPTOS_OFICIALES = """
- FERRETERÍA / HERRAMIENTAS
- REPUESTOS Y REPARACIONES
- OTROS GASTOS MATERIALES Y EQUIPOS
- COMIDAS
- COMBUSTIBLE GASOLINA
- COMBUSTIBLE GASOIL
- COMBUSTIBLE 100% RENOVABLE
- COMBUSTIBLE ELECTRICIDAD
- AD BLUE
- PEAJES
- APARCAMIENTOS
- TAXIS
- VIAJES
- ALQUILER VEHÍCULOS
- HOTELES
- OTROS GASTOS DE VIAJES
- INVITACIONES
"""

class TicketGasto(BaseModel):
    fecha: str = Field(description="Fecha del ticket en formato DD/MM/YYYY")
    descripcion: str = Field(
        description="Descripción exacta según las normas de Electren: "
                    "1. COMIDAS: '[CP] - [Localidad], [desayuno/comida/cena]' (ej: '28012 - Madrid, comida'). "
                    "2. INVITACIONES: '[CP] - [Localidad], [desayuno/comida/cena] ([Nombres comensales])' (ej: '28012 - Madrid, comida (Pablo y Jorge)'). "
                    "3. COMBUSTIBLE: '[CP] - [Localidad], [Tipo combustible]' (ej: '28012 - Madrid, Gasóleo A'). "
                    "4. MATERIALES/AMAZON: '[Comercio] - [Detalle de los artículos comprados]' (ej: 'Amazon - Cables de red'). "
                    "5. TAXIS: '[Origen] a [Destino]'."
    )
    importe: float = Field(description="Importe total final pagado con decimales (sin símbolo €), ej: 13.00")
    concepto: str = Field(description=f"DEBE SER ESTRICTAMENTE UNO DE ESTOS VALORES EXACTOS EN MAYÚSCULAS:\n{CONCEPTOS_OFICIALES}")

class ListaTickets(BaseModel):
    tickets: list[TicketGasto]

def convertir_fecha(fecha_str):
    try: return datetime.strptime(fecha_str, "%d/%m/%Y")
    except: return datetime(2099, 12, 31)

# Configuración Web
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
        
        barra_progreso = st.progress(0, text="Iniciando lectura inteligente...")
        
        prompt = f"""
        Analiza esta imagen y extrae los datos contables siguiendo estrictamente estas reglas de la empresa Electren:
        
        1. REGLA DE INVITACIONES (NOTAS A BOLI):
           Presta mucha atención a anotaciones manuscritas a bolígrafo. Si en el ticket hay escritos dos o más nombres (ej: 'comida pablo y jorge', 'invitación a Juan', etc.), el concepto DEBE ser 'INVITACIONES' (NO 'COMIDAS') y debes incluir los nombres en la descripción.
        
        2. REGLA DE COMIDAS:
           Si es comida individual, pon 'COMIDAS'. Deduce si es desayuno, comida o cena por la hora o productos. Pon siempre el Código Postal (5 dígitos) y Municipio.
           
        3. REGLA DE MATERIALES Y AMAZON:
           Si es una factura de Amazon, ferretería o compra de piezas/cables/herramientas:
           - Concepto: 'FERRETERÍA / HERRAMIENTAS' o 'OTROS GASTOS MATERIALES Y EQUIPOS'.
           - Descripción: Indica el comercio y los materiales concretos comprados (ej: 'Amazon - Cables y conectores').
           
        4. CONCEPTOS VÁLIDOS PERMITIDOS:
        {CONCEPTOS_OFICIALES}
        """

        for i, foto in enumerate(fotos):
            img = Image.open(foto)
            img.thumbnail((1024, 1024))
            exito = False
            error_msg = ""
            
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
                            todos_los_gastos.append({
                                "fecha": t.fecha, "fecha_dt": convertir_fecha(t.fecha),
                                "descripcion": t.descripcion, "importe": t.importe, "concepto": t.concepto
                            })
                        exito = True
                        break
                except Exception as e:
                    error_msg = str(e)
                    time.sleep(1)
            
            if not exito:
                tickets_fallidos.append(f"{foto.name} ({error_msg[:100]})")
            
            barra_progreso.progress((i + 1) / len(fotos), text=f"Procesando {i+1} de {len(fotos)}: {foto.name}...")

        if tickets_fallidos:
            st.error("⚠️ No se pudieron leer estos tickets:")
            for tf in tickets_fallidos:
                st.write(f"- {tf}")
        
        if todos_los_gastos:
            # Ordenar por concepto y fecha
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
