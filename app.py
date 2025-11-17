from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import random
from datetime import datetime, timedelta
import xlsxwriter
from io import BytesIO

app = Flask(__name__)
CORS(app)

def generar_datos_ejemplo():
    estatus_opciones = ["Abierto", "En proceso", "Cerrado"]
    ingenieros = ["Marco Antonio", "Carlos Pérez", "Ana López", "Luis García"]
    data = []
    base = datetime.now() - timedelta(days=15)

    for i in range(250):
        data.append({
            "ticket_id": i + 1,
            "date": base + timedelta(hours=random.randint(1, 300)),
            "status": random.choice(estatus_opciones),
            "engineer": random.choice(ingenieros)
        })

    df = pd.DataFrame(data)
    return df

def estilizar_tabla(ws, df, workbook):
    header_format = workbook.add_format({
        "bold": True, "bg_color": "#1F4E79", "font_color": "white",
        "border": 1, "align": "center"
    })

    body_format = workbook.add_format({
        "border": 1, "align": "center"
    })

    zebra_format = workbook.add_format({
        "border": 1, "align": "center", "bg_color": "#D9E1F2"
    })

    for col_num, value in enumerate(df.columns.values):
        ws.write(0, col_num, value, header_format)

    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_format = zebra_format if row % 2 == 0 else body_format
            ws.write(row + 1, col, df.iloc[row, col], cell_format)

    ws.autofilter(0, 0, len(df), len(df.columns)-1)
    for i, col in enumerate(df.columns):
        ws.set_column(i, i, len(col) + 12)

def crear_reporte(df):
    output = BytesIO()

    resumen_estatus = df["status"].value_counts().reset_index()
    resumen_estatus.columns = ["Estatus", "Cantidad"]

    resumen_ing = df.groupby(["engineer", "status"]).size().unstack(fill_value=0)
    resumen_ing = resumen_ing.reindex(columns=["Abierto", "En proceso", "Cerrado"])
    resumen_ing.reset_index(inplace=True)

    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book

    # ===== Hoja 1: Tickets =====
    df.to_excel(writer, sheet_name="Tickets", index=False)
    ws0 = writer.sheets["Tickets"]
    estilizar_tabla(ws0, df, workbook)

    # ===== Hoja 2: Resumen Estatus =====
    resumen_estatus.to_excel(writer, sheet_name="Resumen Estatus", index=False)
    ws1 = writer.sheets["Resumen Estatus"]

    ws1.set_tab_color("#A6A6A6")
    ws1.set_default_row(20)
    ws1.set_column("A:B", 20)

    colores = {
        "Abierto": "#FF9999",
        "En proceso": "#FFE699",
        "Cerrado": "#C6EFCE"
    }

    header_format = workbook.add_format({
        "bold": True, "font_color": "white", "align": "center",
        "border": 1, "bg_color": "#4F81BD"
    })

    title_format = workbook.add_format({
        "bold": True, "font_size": 18, "align": "center",
        "bg_color": "#E6E6E6", "border": 2
    })

    ws1.merge_range("A1:B1", "📌 Resumen por Estatus", title_format)

    ws1.write(1, 0, "Estatus", header_format)
    ws1.write(1, 1, "Cantidad", header_format)

    for r in range(len(resumen_estatus)):
        est = resumen_estatus.iloc[r, 0]
        bg = colores.get(est, "#FFFFFF")

        style = workbook.add_format({
            "align": "center", "border": 1, "bg_color": bg, "bold": True
        })

        ws1.write(r + 2, 0, est, style)
        ws1.write(r + 2, 1, resumen_estatus.iloc[r, 1], style)

    points = [{"fill": {"color": colores.get(est)}} for est in resumen_estatus["Estatus"]]

    chart1 = workbook.add_chart({"type": "pie"})
    chart1.add_series({
        "name": "Estatus",
        "categories": f"=Resumen Estatus!$A$3:$A${2+len(resumen_estatus)}",
        "values": f"=Resumen Estatus!$B$3:$B${2+len(resumen_estatus)}",
        "data_labels": {"percentage": True, "category": True, "value": True},
        "points": points
    })
    chart1.set_title({"name": "Estatus de Tickets"})
    chart1.set_legend({"position": "bottom"})
    chart1.set_style(10)
    ws1.insert_chart("D3", chart1)

    # ===== Hoja 3: Resumen Ingenieros =====
    resumen_ing.to_excel(writer, sheet_name="Resumen Ingenieros", index=False)
    ws2 = writer.sheets["Resumen Ingenieros"]
    estilizar_tabla(ws2, resumen_ing, workbook)

    num_rows = len(resumen_ing)

    chart2 = workbook.add_chart({"type": "column"})
    for i, status in enumerate(["Abierto", "En proceso", "Cerrado"]):
        chart2.add_series({
            "name": status,
            "categories": f"='Resumen Ingenieros'!$A$2:$A${num_rows+1}",
            "values": f"='Resumen Ingenieros'!${chr(66+i)}$2:${chr(66+i)}${num_rows+1}",
            "data_labels": {"value": True}
        })

    chart2.set_title({"name": "Estatus por Ingeniero"})
    ws2.insert_chart("G3", chart2)

    # ===== Hoja 4: Dashboard =====
    ws3 = workbook.add_worksheet("Dashboard")
    title = workbook.add_format({"bold": True, "font_size": 22})
    card = workbook.add_format({"bold": True, "font_size": 18, "align": "center", "border": 2})

    ws3.write("A1", "📊 Dashboard Ejecutivo de Tickets", title)

    total = len(df)
    abiertos = resumen_estatus.loc[resumen_estatus["Estatus"] == "Abierto", "Cantidad"].values[0]
    proceso = resumen_estatus.loc[resumen_estatus["Estatus"] == "En proceso", "Cantidad"].values[0]
    cerrados = resumen_estatus.loc[resumen_estatus["Estatus"] == "Cerrado", "Cantidad"].values[0]

    ws3.merge_range("A3:C5", f"Total: {total}", card)
    ws3.merge_range("E3:G5", f"Abiertos: {abiertos}", workbook.add_format({"bg_color": "#FFC7CE", "align": "center", "bold": True, "border": 2}))
    ws3.merge_range("A7:C9", f"En Proceso: {proceso}", workbook.add_format({"bg_color": "#FFEB9C", "align": "center", "bold": True, "border": 2}))
    ws3.merge_range("E7:G9", f"Cerrados: {cerrados}", workbook.add_format({"bg_color": "#C6EFCE", "align": "center", "bold": True, "border": 2}))

    # ===== Hoja 5: Rating Ingenieros =====
    ws4 = workbook.add_worksheet("Rating Ingenieros")
    puntos = {"Cerrado": 5, "En proceso": 3, "Abierto": 1}
    df["puntos"] = df["status"].map(puntos)

    rating = df.groupby("engineer")["puntos"].mean().reset_index()
    rating.columns = ["Ingeniero", "Rating Promedio"]
    rating["Rating Promedio"] = rating["Rating Promedio"].round(2)
    rating = rating.sort_values(by="Rating Promedio", ascending=False)

    rating.to_excel(writer, sheet_name="Rating Ingenieros", index=False)
    estilizar_tabla(ws4, rating, workbook)

    top = rating.iloc[0]
    ws4.write("E2", "🏆 Mejor Ingeniero:", workbook.add_format({"bold": True}))
    ws4.write("F2", f"{top['Ingeniero']} ({top['Rating Promedio']})")

    chart3 = workbook.add_chart({"type": "column"})
    num_rat = len(rating)
    chart3.add_series({
        "name": "Rating",
        "categories": f"='Rating Ingenieros'!$A$2:$A${num_rat+1}",
        "values": f"='Rating Ingenieros'!$B$2:$B${num_rat+1}",
        "data_labels": {"value": True}
    })
    chart3.set_title({"name": "Rating por Ingeniero"})
    ws4.insert_chart("E5", chart3)

    writer.close()
    output.seek(0)
    return output

@app.route("/generate-excel", methods=["POST"])
def generate_excel():
    data = request.json.get("tickets", [])
    df = pd.DataFrame(data) if data else generar_datos_ejemplo()
    archivo = crear_reporte(df)
    nombre_archivo = f"Reporte_Tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(archivo, as_attachment=True, download_name=nombre_archivo)

@app.route("/")
def home():
    return jsonify({"message": "Servidor de Reportes está funcionando 🚀"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
