from flask import Flask, render_template, request, send_file, redirect, jsonify
import pandas as pd
import unicodedata
import io
import json
import plotly.express as px
import sys, os

def ruta_archivo(nombre):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nombre)
    return os.path.join(os.path.abspath("."), nombre)

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__, template_folder=ruta_archivo("templates"))
app.secret_key = "1234"
data_filtrada_global = None
login_activo = False

# =========================
# FUNCIONES DE APOYO
# =========================
def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.strip()

def formato_pesos(x):
    try:
        return "${:,.2f}".format(float(x))
    except:
        return "$0.00"


df = pd.read_excel(ruta_archivo("ejemplo base.xlsx"))
df["Municipio_normalizado"] = df["Municipio"].apply(normalizar)
df["Monto en pesos"] = pd.to_numeric(
    df["Monto en pesos"].astype(str)
    .str.replace("$","", regex=False)
    .str.replace(",","", regex=False),
    errors="coerce"
).fillna(0)

# =========================
# USERS (Cargar desde archivo)
# =========================
try:
    with open(ruta_archivo("users.json")) as f:
        USERS = json.load(f)
except:
    USERS = {"admin": {"password": "123"}}

# =========================
# RUTAS (LOGIN Y DASHBOARD)
# =========================
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if USERS.get(request.form["user"], {}).get("password") == request.form["pass"]:
            global login_activo
            login_activo = True
            return redirect("/")
    return render_template("login.html")

@app.route("/logout")
def logout():
    global login_activo
    login_activo = False
    return redirect("/login")

@app.route("/", methods=["GET","POST"])
def inicio():
    global login_activo
    if not login_activo:
        return redirect("/login")

    data = df.copy()
    total = len(data)
    suma = data["Monto en pesos"].sum()

    data_vista = data.copy()
    data_vista["Monto en pesos"] = data_vista["Monto en pesos"].apply(formato_pesos)
    
    # Excluir las columnas "Municipio_normalizado" y "Sexo (catálogo)" de la vista
    columnas_a_excluir = ["Municipio_normalizado", "Sexo (catálogo)"]
    columnas_a_mostrar = [col for col in data_vista.columns if col not in columnas_a_excluir]
    data_vista = data_vista[columnas_a_mostrar]

    global data_filtrada_global
    data_filtrada_global = data

    return render_template(
        "index.html",
        tablas=data_vista.values.tolist(),
        columnas=data_vista.columns,
        total=total,
        suma=formato_pesos(suma),
        municipios=sorted(df["Municipio"].dropna().unique()),
        subprogramas=sorted(df["Subprograma"].dropna().unique())
    )

@app.route("/filtrar", methods=["POST"])
def filtrar():
    data = df.copy()
    req = request.json

    if req.get("busqueda"):
        b = req["busqueda"]
        data = data[data.astype(str).apply(lambda r: r.str.contains(b, case=False).any(), axis=1)]

    if req.get("municipio") and req["municipio"].strip():
        municipio_norm = normalizar(req["municipio"])
        data = data[data["Municipio_normalizado"] == municipio_norm]

    if req.get("subprograma") and req["subprograma"].strip():
        data = data[data["Subprograma"] == req["subprograma"]]

    total = len(data)
    suma = data["Monto en pesos"].sum()
    
    # Validar que suma no sea NaN
    if pd.isna(suma):
        suma = 0

    data_vista = data.copy()
    data_vista["Monto en pesos"] = data_vista["Monto en pesos"].apply(formato_pesos)
    
    # Excluir las columnas "Municipio_normalizado" y "Sexo (catálogo)" de la vista
    columnas_a_excluir = ["Municipio_normalizado", "Sexo (catálogo)"]
    columnas_a_mostrar = [col for col in data_vista.columns if col not in columnas_a_excluir]
    data_vista = data_vista[columnas_a_mostrar]
    
    # Limpiar NaN para que sea serializable a JSON
    tabla_limpia = []
    for fila in data_vista.values.tolist():
        fila_limpia = [None if (isinstance(v, float) and pd.isna(v)) else v for v in fila]
        tabla_limpia.append(fila_limpia)

    global data_filtrada_global
    data_filtrada_global = data

    return jsonify({
        "total": total,
        "suma": formato_pesos(suma),
        "tabla": tabla_limpia,
        "columnas": list(data_vista.columns),
        "conteo": data.groupby("Municipio_normalizado").size().to_dict()
    })

@app.route("/excel")
def excel():
    global data_filtrada_global
    if data_filtrada_global is None:
        return "No hay datos", 400

    data = data_filtrada_global.copy()
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data_excel = data.copy()

        subtotal = data_excel["Monto en pesos"].sum()

        # agrega fila TOTAL como numérico para la columna de montos
        fila_total = {col: "" for col in data_excel.columns}
        # opcional: etiqueta TOTAL en la columna Municipio si existe
        if "Municipio" in data_excel.columns:
            fila_total["Subprograma"] = "TOTAL"
        fila_total["Monto en pesos"] = subtotal

        data_excel = pd.concat([data_excel, pd.DataFrame([fila_total])], ignore_index=True)

        data_excel.to_excel(writer, index=False, sheet_name="Reporte")
        ws = writer.sheets["Reporte"]

        # obtener columna "Monto en pesos"
        monto_col = None
        for idx, cell in enumerate(ws[1], start=1):
            if cell.value == "Monto en pesos":
                monto_col = idx
                break

        if monto_col is not None:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=monto_col).number_format = "$#,##0.00"

    output.seek(0)
    return send_file(output, download_name="reporte.xlsx", as_attachment=True)
    
@app.route("/pdf")
def pdf():

    global data_filtrada_global

    if data_filtrada_global is None:
        return "No hay datos", 400

    data = data_filtrada_global.copy()

    total = len(data)
    suma = data["Monto en pesos"].sum()

    # Excluir columnas innecesarias
    columnas_a_excluir = ["Municipio_normalizado", "Sexo (catálogo)"]
    columnas_a_mostrar = [col for col in data.columns if col not in columnas_a_excluir]
    data = data[columnas_a_mostrar]

    data["Monto en pesos"] = data["Monto en pesos"].apply(formato_pesos)
    # 🔥 FORMATO DE FECHA CORTO
    if "Fecha" in data.columns:
        data["Fecha"] = pd.to_datetime(data["Fecha"]).dt.strftime("%d-%m-%Y")
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm

    estilo = ParagraphStyle(
        name='Normal',
        fontSize=7,
        leading=9,
        alignment=1
    )

    filas = []
    for _, row in data.iterrows():
        fila = []
        for col in data.columns:
            val = str(row[col])[:120]
            if col.lower() in ["apoyo", "subprograma"]:
                fila.append(Paragraph(val, estilo))
            else:
                fila.append(val)
        filas.append(fila)

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=1*cm,
        rightMargin=1*cm,
        topMargin=1*cm,
        bottomMargin=1*cm
    )
    styles = getSampleStyleSheet()
    from datetime import datetime
    fecha = datetime.now().strftime("%d/%m/%Y")

    elementos = []

    elementos.append(Paragraph("REPORTE BENEFICIADOS SADERH 2025", styles['Title']))
    elementos.append(Spacer(1, 10))

    

    # 🔥 AJUSTE PROFESIONAL DE COLUMNAS
    anchos = []

    
 
    tabla = Table(
        [list(data.columns)] + filas,
        repeatRows=1
    )
    
    tabla._argW = [None] * len(data.columns)
    tabla.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#691b31")),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTSIZE',(0,0),(-1,-1),7),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('LEFTPADDING',(0,0),(-1,-1),3),
        ('RIGHTPADDING',(0,0),(-1,-1),3),
        ('TOPPADDING',(0,0),(-1,-1),2),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
        ('GRID',(0,0),(-1,-1),0.25,colors.black)
   ]))
    tabla._argW = [None] * len(data.columns)
    elementos.append(tabla)
    elementos.append(Spacer(1, 20))

    resumen = Table([
        ["TOTAL REGISTROS", str(total)],
        ["MONTO TOTAL", formato_pesos(suma)]
    ])

    resumen.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),colors.HexColor("#a02142")),
        ('TEXTCOLOR',(0,0),(-1,-1),colors.white),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('FONTSIZE',(0,0),(-1,-1),14)
    ]))

    elementos.append(resumen)
    elementos.append(Spacer(1, 6 ))
    elementos.append(Paragraph(f"Fecha de emisión: {fecha}", styles["Normal"]))

    doc.build(elementos)

    buffer.seek(0)

    return send_file(buffer, download_name="reporte.pdf", as_attachment=True)


import webbrowser
import threading

def abrir_navegador():
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == "__main__":
    threading.Timer(1, abrir_navegador).start()
    if __name__ == "__main__":
        app.run(host="0.0.0.0", port=10000)
