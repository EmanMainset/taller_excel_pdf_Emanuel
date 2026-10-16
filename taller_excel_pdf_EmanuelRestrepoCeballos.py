import os

import openpyxl
from tabulate import tabulate
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(CARPETA, "ventas_tecnologia_v2.xlsx")
ARCHIVO_PDF = os.path.join(CARPETA, "reporte_ejecutivo_ventas.pdf")
PRECIO_MINIMO = 100

ENCABEZADOS = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
ALINEACION = ("center", "left", "center", "right", "right")
ANCHOS_COLS = [40, 160, 100, 80, 80]


def cargar_datos(ruta=ARCHIVO_EXCEL):
    """Lee el Excel y devuelve una lista de filas (listas), sin filas vacías."""
    wb = openpyxl.load_workbook(ruta, data_only=True)
    hoja = wb.active
    datos = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if all(celda is None for celda in fila):
            continue  # se omiten filas completamente vacías
        datos.append(list(fila))
    return datos



def filtrar_por_precio(datos, minimo=PRECIO_MINIMO):
    """Reto 1: productos con precio unitario superior a `minimo`."""
    return [fila for fila in datos if fila[3] is not None and fila[3] > minimo]


def calcular_metricas(datos):
    """Reto 2: devuelve (total de unidades vendidas, precio promedio)."""
    if not datos:
        return 0, 0.0
    total_unidades = sum(fila[4] or 0 for fila in datos)
    precio_promedio = sum(fila[3] or 0 for fila in datos) / len(datos)
    return total_unidades, precio_promedio


def buscar_por_id(datos, id_buscado):
    """Reto 3: devuelve las filas cuyo ID coincide (comparación como texto)."""
    return [fila for fila in datos if str(fila[0]).strip() == str(id_buscado).strip()]



def mostrar_tabla(datos, titulo):
    print(f"\n=== {titulo} ({len(datos)} PRODUCTOS) ===")
    print(tabulate(datos, headers=ENCABEZADOS, tablefmt="fancy_grid", colalign=ALINEACION))


def mostrar_metricas(datos):
    total_unidades, precio_promedio = calcular_metricas(datos)
    print("\n--- MÉTRICAS DE RESUMEN ---")
    print(f"Total de unidades vendidas : {total_unidades:,.0f}")
    print(f"Precio promedio            : ${precio_promedio:,.2f}")



def generar_pdf(datos, nombre_pdf=ARCHIVO_PDF, titulo="Informe Ejecutivo de Ventas Tecno"):
    """Genera el informe PDF con la tabla de productos y una fila final de resumen."""
    doc = SimpleDocTemplate(nombre_pdf, pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10,
    )
    estilo_cuerpo = ParagraphStyle("EstiloCuerpo", parent=estilos["Normal"], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        "EstiloEncabezado", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.whitesmoke, alignment=1,
    )

    story = [
        Paragraph(titulo, estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12),
    ]


    st_c = ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)
    st_l = ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)
    st_r = ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in ENCABEZADOS]
    filas_pdf = [
        [
            Paragraph(str(r[0]), st_c),
            Paragraph(str(r[1]), st_l),
            Paragraph(str(r[2]), st_c),
            Paragraph(f"${r[3]:,.2f}", st_r),
            Paragraph(f"{r[4]:,.0f}", st_r),
        ]
        for r in datos
    ]


    total_unidades, precio_promedio = calcular_metricas(datos)
    fila_resumen = [
        Paragraph("", st_c),
        Paragraph("<b>RESUMEN</b>", st_l),
        Paragraph("<b>Prom. precio / Total unid.</b>", st_c),
        Paragraph(f"<b>${precio_promedio:,.2f}</b>", st_r),
        Paragraph(f"<b>{total_unidades:,.0f}</b>", st_r),
    ]

    tabla = Table([encabezados_pdf] + filas_pdf + [fila_resumen],
                  colWidths=ANCHOS_COLS, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),  # fila resumen
    ]))

    story.append(tabla)
    doc.build(story)



def opcion_catalogo(datos):
    mostrar_tabla(datos, "REPORTE DE VENTAS EN CONSOLA")
    mostrar_metricas(datos)


def opcion_filtrar(datos):
    filtrados = filtrar_por_precio(datos)
    if not filtrados:
        print(f"\n[AVISO] No hay productos con precio superior a ${PRECIO_MINIMO}.")
        return
    mostrar_tabla(filtrados, f"PRODUCTOS CON PRECIO > ${PRECIO_MINIMO}")
    mostrar_metricas(filtrados)


def opcion_buscar(datos):
    id_buscado = input("\nIngrese el ID del producto: ").strip()
    if not id_buscado:
        print("[ADVERTENCIA] Debe ingresar un ID.")
        return
    resultado = buscar_por_id(datos, id_buscado)
    if resultado:
        mostrar_tabla(resultado, f"RESULTADO DE BÚSQUEDA (ID {id_buscado})")
    else:
        print(f"[ADVERTENCIA] El ID '{id_buscado}' no existe en el catálogo.")


def opcion_exportar(datos):
    generar_pdf(datos)
    print(f"\n[ÉXITO] Archivo '{os.path.basename(ARCHIVO_PDF)}' generado correctamente.")


def menu(datos):
    while True:
        print("\n========== MENÚ TECNOVENTAS ==========")
        print("1) Ver catálogo completo")
        print(f"2) Filtrar por precio (> ${PRECIO_MINIMO})")
        print("3) Buscar por ID")
        print("4) Exportar informe PDF")
        print("5) Salir")
        opcion = input("Seleccione una opción (1-5): ").strip()

        if opcion == "1":
            opcion_catalogo(datos)
        elif opcion == "2":
            opcion_filtrar(datos)
        elif opcion == "3":
            opcion_buscar(datos)
        elif opcion == "4":
            opcion_exportar(datos)
        elif opcion == "5":
            print("\nHasta pronto.")
            break
        else:
            print("[ADVERTENCIA] Opción no válida. Elija un número del 1 al 5.")


def main():
    try:
        datos = cargar_datos()
    except FileNotFoundError:
        print(f"[ERROR] No se encontró '{ARCHIVO_EXCEL}'. "
              "Coloque el Excel en la misma carpeta que este script.")
        return
    print(f"Hoja cargada. Total de registros leídos en memoria: {len(datos)}")
    menu(datos)


if __name__ == "__main__":
    main()