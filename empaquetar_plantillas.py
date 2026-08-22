"""Empaqueta las plantillas oficiales dentro de la aplicación.

¿Para qué? Cuando la página se abre con doble clic (dirección `file://`), el
navegador prohíbe leer archivos de la carpeta: por eso aparecía «Selecciona la
plantilla oficial». Un archivo .js sí se puede cargar con `<script src=...>`
aunque no haya servidor web, así que las plantillas se guardan ahí en base64 y
la app las encuentra siempre.

Cuando el SENA publique una versión nueva de un formato: reemplaza el .docx o
.xlsx en esta carpeta y vuelve a ejecutar

    python3 empaquetar_plantillas.py

Uso normal (con servidor o GitHub Pages): la app intenta primero leer el archivo
de la carpeta, y solo si no puede usa la copia empaquetada. Así, si cambias una
plantilla y olvidas ejecutar este script, el sitio publicado sigue usando la
versión nueva.
"""
import base64, os, sys

CARPETA = os.path.dirname(os.path.abspath(__file__))

PAQUETES = {
    "plantillas-023.js": [
        "GFPI-F-023_V06.__Formato_de_Planeación_Seguimiento_y_Evaluación_de_Etapa_Productiva_Grado10_Modelo.docx",
        "GFPI-F-023_V06.__Formato_de_Planeación_Seguimiento_y_Evaluación_de_Etapa_Productiva_Grado11_Modelo.docx",
    ],
    "plantillas-193.js": [
        "GFPI-F-193Formato_acta_de_inicio_y_confidencialidad_del_proyecto_productivo_V3_NombreApellidosAprendiz_Modelo.docx",
    ],
    "plantillas-147.js": [
        "GFPI-F-147-Bitacora-#_Modelo.xlsx",
    ],
}

CABECERA = """/* Plantilla(s) oficial(es) empaquetada(s) en base64.
   Generado por empaquetar_plantillas.py — NO editar a mano.
   Sirve para que la aplicación funcione al abrirla con doble clic, sin
   servidor web. Si cambias la plantilla, vuelve a ejecutar el script. */
window.PLANTILLAS_EMBEBIDAS = window.PLANTILLAS_EMBEBIDAS || {};
"""

faltantes, total = [], 0
for salida, archivos in PAQUETES.items():
    partes = [CABECERA]
    for nombre in archivos:
        ruta = os.path.join(CARPETA, nombre)
        if not os.path.exists(ruta):
            faltantes.append(nombre)
            continue
        datos = base64.b64encode(open(ruta, "rb").read()).decode("ascii")
        partes.append(f'window.PLANTILLAS_EMBEBIDAS[{nombre!r}] =\n  "{datos}";\n')
        print(f"  {nombre[:64]:66} {len(datos)/1024:8.0f} KB en base64")
    if len(partes) > 1:
        contenido = "\n".join(partes)
        open(os.path.join(CARPETA, salida), "w", encoding="utf8").write(contenido)
        total += len(contenido)
        print(f"→ {salida} ({len(contenido)/1024:.0f} KB)\n")

if faltantes:
    print("FALTAN estas plantillas en la carpeta:")
    for f in faltantes:
        print("  -", f)
    sys.exit(1)
print(f"Listo. {total/1024/1024:.1f} MB empaquetados.")
