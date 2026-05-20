from flask import Flask, render_template, request, send_file, redirect, jsonify, session
import sys, os
import json
import io

# Lazy import de pandas
pd = None
unicodedata = None
reportlab = None

def get_pandas():
    """Importa pandas solo cuando se necesita"""
    global pd, unicodedata
    if pd is None:
        import pandas as pd_temp
        import unicodedata as u
        pd = pd_temp
        unicodedata = u
    return pd

def get_reportlab():
    """Importa reportlab solo cuando se necesita"""
    global reportlab
    if reportlab is None:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter, A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
        reportlab = {
            'colors': colors,
            'pagesizes': {'letter': letter, 'A4': A4},
            'SimpleDocTemplate': SimpleDocTemplate,
            'Table': Table,
            'TableStyle': TableStyle,
            'Paragraph': Paragraph,
            'Spacer': Spacer,
            'getSampleStyleSheet': getSampleStyleSheet,
            'landscape': landscape
        }
    return reportlab

def ruta_archivo(nombre):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nombre)
    return os.path.join(os.path.abspath("."), nombre)

app = Flask(__name__, template_folder=ruta_archivo("templates"))
app.secret_key = "saderh_2025_secret_key_abc123!@#"
data_filtrada_global = None
_df_cache = None

print("[app.py] ✓ Flask inicializado")

def normalizar(texto):
    pd_local = get_pandas()
    if pd_local.isna(texto):
        return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.strip()

def formato_pesos(x):
    try:
        return "${:,.2f}".format(float(x))
    except:
        return "$0.00"

def cargar_dataframe():
    """Carga el DataFrame (lazy loading)"""
    global _df_cache
    if _df_cache is None:
        pd_local = get_pandas()
        print("[app.py] Leyendo Excel...")
        _df_cache = pd_local.read_excel(ruta_archivo("ejemplo base.xlsx"))
        _df_cache["Municipio_normalizado"] = _df_cache["Municipio"].apply(normalizar)
        _df_cache["Monto en pesos"] = pd_local.to_numeric(
            _df_cache["Monto en pesos"].astype(str)
            .str.replace("$","", regex=False)
            .str.replace(",","", regex=False),
            errors="coerce"
        ).fillna(0)
        print("[app.py] ✓ Excel cargado")
    return _df_cache

try:
    with open(ruta_archivo("users.json")) as f:
        USERS = json.load(f)
except:
    USERS = {"admin": {"password": "123"}}

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        user = request.form.get("user", "")
        password = request.form.get("pass", "")
        if USERS.get(user, {}).get("password") == password:
            session["usuario"] = user
            session.permanent = True
            print(f"[LOGIN] ✓ Usuario {user} autenticado")
            return redirect("/")
        else:
            print(f"[LOGIN] ✗ Credenciales inválidas para {user}")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    print("[LOGOUT] ✓ Sesión cerrada")
    return redirect("/login")

@app.route("/", methods=["GET","POST"])
def inicio():
    if "usuario" not in session:
        return redirect("/login")

    df = cargar_dataframe()
    data = df.copy()
    total = len(data)
    suma = data["Monto en pesos"].sum()

    data_vista = data.copy()
    data_vista["Monto en pesos"] = data_vista["Monto en pesos"].apply(formato_pesos)
    
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
    if "usuario" not in session:
        return {"error": "No autenticado"}, 401
    
    pd_local = get_pandas()
    df = cargar_dataframe()
    data = df.copy()
    req = request.json or {}

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
    
    if pd_local.isna(suma):
        suma = 0

    data_vista = data.copy()
    data_vista["Monto en pesos"] = data_vista["Monto en pesos"].apply(formato_pesos)
    
    columnas_a_excluir = ["Municipio_normalizado", "Sexo (catálogo)"]
    columnas_a_mostrar = [col for col in data_vista.columns if col not in columnas_a_excluir]
    data_vista = data_vista[columnas_a_mostrar]
    
    tabla_limpia = []
    for fila in data_vista.values.tolist():
        fila_limpia = [None if (isinstance(v, float) and pd_local.isna(v)) else v for v in fila]
        tabla_limpia.append(fila_limpia)

    global data_filtrada_global
    data_filtrada_global = data

    municipios_monto = (
    data.groupby("Municipio")["Monto en pesos"]
    .sum()
    .sort_values(ascending=False)
    .head(30)   # 👈 SOLO TOP 30 PARA LA PANTALLA
    ) 
    municipios_data = {
        "labels": municipios_monto.index.tolist(),
        "valores": municipios_monto.values.tolist()
    }
    # TODOS LOS MUNICIPIOS (para descarga completa)
    municipios_full = data.groupby("Municipio")["Monto en pesos"].sum().sort_values(ascending=False)

    municipios_full_data = {
        "labels": municipios_full.index.tolist(),
        "valores": municipios_full.values.tolist()
    }   

    subprogramas_monto = data.groupby("Subprograma")["Monto en pesos"].sum()
    subprogramas_monto = subprogramas_monto.sort_values(ascending=False)
    total_monto = subprogramas_monto.sum()
    subprogramas_data = {
        "labels": subprogramas_monto.index.tolist(),
        "valores": subprogramas_monto.values.tolist(),
        "porcentajes": [(v / total_monto * 100) if total_monto > 0 else 0 for v in subprogramas_monto.values.tolist()]
    }

    return jsonify({
        "total": total,
        "suma": formato_pesos(suma),
        "tabla": tabla_limpia,
        "columnas": list(data_vista.columns),
        "conteo": data.groupby("Municipio_normalizado").size().to_dict(),
        "municipios_monto": municipios_data,
        "subprogramas_data": subprogramas_data,
        "municipios_full": municipios_full_data,
    })

@app.route("/excel")
def excel():
    if "usuario" not in session:
        return redirect("/login")
    
    global data_filtrada_global
    if data_filtrada_global is None:
        return "No hay datos", 400

    pd_local = get_pandas()
    data = data_filtrada_global.copy()
    output = io.BytesIO()

    with pd_local.ExcelWriter(output, engine="openpyxl") as writer:
        data_excel = data.copy()
        subtotal = data_excel["Monto en pesos"].sum()
        fila_total = {col: "" for col in data_excel.columns}
        if "Municipio" in data_excel.columns:
            fila_total["Subprograma"] = "TOTAL"
        fila_total["Monto en pesos"] = subtotal

        data_excel = pd_local.concat([data_excel, pd_local.DataFrame([fila_total])], ignore_index=True)
        data_excel.to_excel(writer, index=False, sheet_name="Reporte")
        ws = writer.sheets["Reporte"]

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
    try:
        global data_filtrada_global

        if data_filtrada_global is None:
            return "No hay datos para generar el PDF", 400

        data = data_filtrada_global.copy()

        total = len(data)
        suma = data["Monto en pesos"].sum()

        reportlab_pkg = get_reportlab()
        SimpleDocTemplate = reportlab_pkg['SimpleDocTemplate']
        Table = reportlab_pkg['Table']
        TableStyle = reportlab_pkg['TableStyle']
        ParagraphReport = reportlab_pkg['Paragraph']
        Spacer = reportlab_pkg['Spacer']
        getSampleStyleSheet = reportlab_pkg['getSampleStyleSheet']
        colors = reportlab_pkg['colors']
        letter = reportlab_pkg['pagesizes']['letter']
        landscape = reportlab_pkg['landscape']

        # Excluir columnas innecesarias
        columnas_a_excluir = ["Municipio_normalizado", "Sexo (catálogo)"]
        columnas_a_mostrar = [col for col in data.columns if col not in columnas_a_excluir]
        data = data[columnas_a_mostrar]

        data["Monto en pesos"] = data["Monto en pesos"].apply(formato_pesos)
        # 🔥 FORMATO DE FECHA CORTO
        if "Fecha" in data.columns:
            data["Fecha"] = pd.to_datetime(data["Fecha"]).dt.strftime("%d-%m-%Y")
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm

        def footer(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            page_num = canvas.getPageNumber()
            text = f"Página {page_num}"
            canvas.drawRightString(landscape(letter)[0] - 1*cm, 1*cm, text)
            canvas.restoreState()

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
                    fila.append(ParagraphReport(val, estilo))
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
            bottomMargin=1*cm,
            onPage=footer
        )
        styles = getSampleStyleSheet()
        from datetime import datetime
        fecha = datetime.now().strftime("%d-%m-%Y")

        elementos = []

        elementos.append(ParagraphReport("REPORTE BENEFICIADOS SADERH 2025", styles['Title']))
        elementos.append(Spacer(1, 10))

        # 🔥 AJUSTE PROFESIONAL DE COLUMNAS
        num_cols = len(data.columns)
        ancho_disponible = landscape(letter)[0] - 2*cm  # ancho total menos márgenes
        ancho_min = 40  # mínimo 40 puntos por columna
        ancho_max = 120  # máximo 120 puntos
        if num_cols * ancho_min > ancho_disponible:
            ancho_col = ancho_disponible / num_cols
        else:
            ancho_col = max(ancho_min, ancho_disponible / num_cols)
            ancho_col = min(ancho_col, ancho_max)
        anchos = [ancho_col] * num_cols

        tabla = Table(
            [list(data.columns)] + filas,
            colWidths=anchos,
            repeatRows=1
        )

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
        elementos.append(tabla)
        elementos.append(Spacer(1, 20))

        resumen = Table([
            ["TOTAL BENEFICIADOS", f"{total:,}"],
            ["MONTO TOTAL", formato_pesos(suma)]
        ], colWidths=[200, 200])

        resumen.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),colors.HexColor("#a02142")),
            ('TEXTCOLOR',(0,0),(-1,-1),colors.white),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTSIZE',(0,0),(-1,-1),14),
            ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold')
        ]))

        elementos.append(resumen)
        elementos.append(Spacer(1, 6))
        elementos.append(ParagraphReport(f"Fecha de emisión: {fecha}", styles["Normal"]))

        doc.build(elementos)

        buffer.seek(0)

        return send_file(buffer, download_name="reporte.pdf", as_attachment=True)
    
    except Exception as e:
        print(f"[PDF ERROR] Error generando PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"Error interno del servidor al generar PDF: {str(e)}", 500

print("[app.py] ✓ Rutas configuradas")

# Precargar el DataFrame al iniciar
print("[app.py] Precargando Excel en background...")
try:
    df_temp = cargar_dataframe()
    print(f"[app.py] ✓ Excel precargado ({len(df_temp)} registros)")
except Exception as e:
    print(f"[app.py] ✗ Error precargando: {e}")

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
