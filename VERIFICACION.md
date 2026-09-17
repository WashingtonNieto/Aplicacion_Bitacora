# Verificación — Formatos de etapa productiva

## Bitácora GFPI-F-147

Resultado de `python3 verificar.py` sobre las copias generadas para el grupo de la bitácora 6
(Danna, Lizeth y Víctor), con firmas cargadas. Se entregan con el prefijo `PRUEBA-` para no pisar
el archivo original que ya estaba en la carpeta.

| # | Criterio | Estado | Detalle |
|---|---|---|---|
| 1 | Nombre del archivo según la fórmula | ✅ Cumple | Calculado desde el formulario, no escrito fijo |
| 2 | Logos, fusiones, bordes y encabezado intactos | ✅ Cumple | 3 → 9 imágenes · 132 → 137 fusiones · **0 fusiones perdidas** · encabezado y bordes idénticos |
| 3 | Hojas `Hoja1` e `Instructivo` presentes e intactas | ✅ Cumple | Comparación celda por celda: sin diferencias |
| 4 | Datos en las celdas del mapa, sin desplazamientos | ✅ Cumple | 15/15 celdas de control correctas |
| 5 | Fechas como fecha real con formato `dd/mm/yyyy` | ✅ Cumple | 9 celdas de fecha verificadas como `datetime`, ninguna celda desbordada |
| 6 | Documentos y teléfonos completos, como texto | ✅ Cumple | 6/6 valores exactos, sin notación científica |
| 7 | Exactamente una `"X"` en la alternativa | ✅ Cumple | Solo `H45` (proyecto productivo) |
| 8 | Rótulos de firma intactos y misma paginación | ✅ Cumple | Ningún rótulo sobrescrito · 2 páginas la plantilla, 2 la salida |
| 9 | Funciona sin servidor (`file://`) | ✅ Cumple | Ofrece selección manual de la plantilla; ExcelJS carga desde la copia local |
| 10 | Firmas insertadas sobre su línea | ✅ Cumple | 9 imágenes: 3 del formato + casilla "Pública" + 5 firmas · anclas correctas para los 3 aprendices, el instructor y el ente co-formador |
| 11 | Una copia por aprendiz, idénticas y con su nombre | ✅ Cumple | 3 archivos · contenido idéntico byte a byte |

**11/11 criterios cumplidos.**

> Los conteos de 9 imágenes y 137 fusiones son mayores que los de la plantilla a propósito: son la
> casilla "Pública" restaurada, las 5 firmas y las fusiones de la cuarta actividad (fila 59), que
> la plantilla trae a medio formatear.

---

## Dónde queda cada firma

| Firmante | Línea en el formato | Ancla de la imagen |
|---|---|---|
| Aprendiz 1 | `C80:E80` | columna C, fila 78, centrada |
| Aprendiz 2 | `H80:I80` | columna H, fila 78, centrada |
| Aprendiz 3 | `C83:E83` | columna C, fila 81, centrada |
| Aprendiz 4 *(si se habilita)* | `H83:I83` | columna H, fila 81 |
| Aprendiz 5 *(si se habilita)* | `C86:E86` | columna C, fila 84 |
| Instructor de seguimiento | `B91:E91` | columna B, fila 89, centrada |
| Ente co-formador | `H91:I91` | columna H, fila 89, centrada |

Cada imagen se escala conservando su proporción, con un tope de 55 px de alto y del 75 % del ancho
de su línea, de modo que nunca invade el espacio de otro firmante.

---

## Pruebas automáticas incluidas

| Archivo | Qué comprueba |
|---|---|
| `probar.js` | El núcleo de escritura, los calendarios de **los dos grados** (9 casos: quincenal de grado 11, mensual de grado 10, salto de diciembre-enero y cierre de febrero el 28), las 5 firmas y que las 3 copias sean idénticas byte a byte. Extrae el código del propio `bitacora.html`, así que no puede desincronizarse |
| `prueba_ui.py` | La app en un navegador real: comprueba que **cambiar de grado** rehaga el calendario (grado 10 mensual, febrero cerrando el 28; grado 11 quincenal), valida el formulario, respeta el tope de 3 aprendices, bloquea fechas fuera del período, comprueba que **elegir el colegio llene dirección, NIT y correo** (y que editarlos a mano devuelva la lista a "Otra entidad"), **exige las firmas de los aprendices**, descarga una copia por aprendiz, comprueba que sean idénticas, restaura el borrador tras recargar, verifica que el **grupo quede guardado solo al generar** y que exportarlo, **limpiar todo el almacenamiento** e importarlo devuelva aprendices, colegio y firmas intactos. Sin errores de consola |
| `verificar.py` | Los 11 criterios de esta tabla, incluida la conversión a PDF para comparar la paginación |
| `prueba_formatos.py` | El menú y los tres formatos de Word (ver la sección final) |
| `probar_seguimiento.js` | El núcleo del Momento 2 fuera del navegador: el cálculo de la fecha del seguimiento, el renglón de cierre con la X, la redacción sugerida a partir de las «M» y la generación de un documento por aprendiz |
| `verificar_seguimiento.py` | El documento del Momento 2 celda por celda: cada S/M en su columna, las filas sobrantes en blanco, las tres firmas, el texto en negro y el bloqueo de solo lectura |
| `huella_referencia.py` | Calcula la huella de contraseña de Word por su cuenta, para contrastar la de `docx.js`. Que dos implementaciones independientes coincidan descarta errores de programación |

---

## Cuatro defectos que encontraron las pruebas

**1. ExcelJS descartaba la casilla de "Pública".** Conserva las 3 imágenes del formato sin
problema, pero la marca roja junto a "Pública" no es una imagen: es una *forma* dibujada
(`<xdr:sp>`), y esas sí se pierden al reescribir el archivo. Se detectó comparando el XML de
dibujo de ambos archivos y convirtiéndolos a PDF para verlos. La app la reinserta como imagen en
la misma posición.

**2. Las fechas salían como texto ISO** (`"2026-07-15T00:00:00.000Z"`). El primer arnés de pruebas
ejecutaba el núcleo en un contexto aislado de Node (`vm`), donde los objetos `Date` pertenecen a
otro *realm* y no pasan el `instanceof Date` de ExcelJS. Era un defecto del arnés, no de la app,
pero de no haberlo corregido la prueba habría estado validando algo distinto de lo que corre en el
navegador. Ahora usa `new Function`, que comparte realm.

**3. La app dependía del CDN.** En una sala de sistemas sin internet (o con el CDN bloqueado por el
filtro del colegio) la app cargaba pero no generaba nada. Se agregó `exceljs.min.js` como copia
local, con el CDN de respaldo.

**4. Al quitar un aprendiz del medio, su firma se le asignaba a otro.** Las firmas se guardan como
`aprendiz0`, `aprendiz1`, `aprendiz2`; si se eliminaba el segundo aprendiz, el tercero heredaba la
firma del segundo. `reindexarFirmasAprendices()` corre los índices junto con la lista.

---

## Observaciones sobre el formato oficial

- **Fuentes grises casi invisibles.** Las celdas `I71:I74` (EPP), `G75` y `H86` (fecha de entrega)
  traen fuente gris `A5A5A5`/`BFBFBF` en la plantilla, y `B21:B23` y `H53` traen azul de
  hipervínculo. Al imprimir, el dato prácticamente no se ve. La app pasa a negro **solo** las
  celdas donde escribe un dato real; no altera ningún otro atributo del formato.
- **El período de la bitácora 6 del ejemplo estaba mal.** El archivo diligenciado dice
  "Desde 1/07/2026 hasta 15/07/2026", pero sus actividades están fechadas entre el 15 y el 30 de
  julio, que es lo que corresponde a la bitácora 6 según la regla de dos bitácoras por mes. La app
  calcula el período y ese error deja de ser posible.
- **El número de bitácora iba en el rótulo.** En el ejemplo quedó escrito como "Bitácora N° 6"
  dentro de `B10`, dejando vacía `B11`, que es la celda de valor. La app escribe en `B11` y deja el
  rótulo intacto.

---

## Formatos de Word (GFPI-F-023 y GFPI-F-193)

`prueba_formatos.py` recorre el menú, entra a cada formulario, importa el grupo guardado desde la
bitácora, diligencia y descarga. Después compara los documentos generados contra la plantilla y
contra los que los aprendices habían llenado a mano:

| Comprobación | Resultado |
|---|---|
| El menú ofrece los tres formatos, en orden | ✅ |
| El grupo guardado en la bitácora se reutiliza en el 023 y en el 193 | ✅ |
| GFPI-F-023: cada grado carga **su propia plantilla** y muestra sus fechas | ✅ |
| GFPI-F-023: plantilla intacta (3 imágenes del formato + 1 firma nueva) | ✅ |
| GFPI-F-023: 22 campos en el párrafo exacto donde los puso la aprendiz | ✅ |
| GFPI-F-023: la firma del aprendiz queda en `5448632A` y las del instructor y el ente co-formador siguen intactas | ✅ |
| GFPI-F-023: la firma es obligatoria; si se quita, bloquea la generación | ✅ |
| GFPI-F-023: cada copia trae los datos de SU aprendiz (es un formato individual) | ✅ |
| GFPI-F-023 grado 10: sale de su plantilla, con sus fechas y su firma | ✅ |
| GFPI-F-023: por defecto genera **solo el del aprendiz elegido**, y «Todo el grupo» genera los de todos | ✅ |
| GFPI-F-023: un aprendiz sin grupo escribe sus datos, elige su colegio y descarga **un solo archivo** | ✅ |
| Elegir el colegio llena también la **ficha** y el **correo de la entidad**, en la bitácora y en el 023 | ✅ |
| **La ficha es obligatoria en los cuatro formatos**: sin ella ninguno genera | ✅ |
| En la bitácora la ficha queda **centrada** en su celda, como el resto de su fila | ✅ |
| Al cambiar de colegio, la ficha que puso la app se reemplaza por la del colegio nuevo | ✅ |
| Una ficha escrita a mano no se pisa al cambiar de colegio | ✅ |
| El acta pone al **instructor técnico** en la viñeta de «instructor proyecto productivo», una sola vez | ✅ |
| Los nombres de persona salen como **nombre propio** en los cuatro formatos, aunque se escriban en mayúscula sostenida | ✅ |
| También en el nombre de los archivos generados | ✅ |
| Las partículas van en minúscula (`Juan de la Cruz`) y las siglas se respetan (`SENA-Washington`) | ✅ |
| El nombre de la entidad y el cargo **no** se tocan | ✅ |
| Los renglones de lista que sobran se **quitan**, no quedan como una viñeta vacía | ✅ |
| Los dos instructores (nombre, cédula, correo, teléfono y firma) se capturan una vez y llegan a los cuatro formatos | ✅ |
| «Aplicar estos instructores a todos mis grupos guardados» actualiza todos los grupos de una vez | ✅ |
| …y reemplaza también la **firma**: la del instructor anterior no queda junto al nombre del nuevo | ✅ |
| El equipo se puede armar **desde el acta**, sin grupo previo, y topa en 3 integrantes | ✅ |
| Ese equipo aparece después en el 023 (3 destinatarios) y en la bitácora (3 aprendices) | ✅ |
| **Grupal**: los 3 integrantes salen en el personal vinculado y en las conclusiones del acta, y las 3 copias son idénticas | ✅ |
| **Individual**: cada copia del 023 trae los datos de su propio aprendiz | ✅ |
| El acta pide los dos correos y los pone en su única columna, institucional arriba y personal debajo | ✅ |
| El acta pide los **objetivos específicos (1.1)** y reemplaza con ellos la instrucción de la plantilla | ✅ |
| Los objetivos sin numerar salen numerados; los que ya vienen numerados se respetan tal cual | ✅ |
| El **objetivo general** del formato (punto 1) no se toca | ✅ |
| Una viñeta por persona: con 3 aprendices el acta muestra a., b. y c., no los tres dentro de la «a.» | ✅ |
| Lo mismo en la lista de instructores (a., b.) | ✅ |
| Los 3 aprendices firman su fila en la columna «FIRMA» de la tabla de aprobación | ✅ |
| El texto que escribe la app queda en **negro**, no en el rojo de los textos de ejemplo | ✅ |
| El acta exige al menos un correo por integrante | ✅ |
| «Grupo nuevo» vacía el formulario (aprendices, colegio y firmas) y **conserva los grupos guardados** | ✅ |
| El menú ofrece los **cuatro** formatos, cada uno con su propio color | ✅ |
| **Momento 2**: los aprendices del grupo aparecen numerados y los 13 factores empiezan en «S» | ✅ |
| **Momento 2**: cada «S»/«M» cae en la columna exacta de su aprendiz, en las dos tablas | ✅ |
| **Momento 2**: las filas que sobran de las dos tablas quedan en blanco | ✅ |
| **Momento 2**: la sugerencia de redacción nombra los factores marcados «M» | ✅ |
| **Momento 2**: un seguimiento virtual sin enlace de grabación queda bloqueado | ✅ |
| **Momento 2**: exige la observación de cada aprendiz | ✅ |
| **Momento 2**: avisa que faltan las firmas de los instructores, sin bloquear | ✅ |
| **Momento 2**: las tres firmas se insertan (la plantilla no trae ninguna) | ✅ |
| **Momento 2**: una copia por aprendiz, con su encabezado y la tabla de todo el equipo | ✅ |
| **Momento 2**: una sola plantilla sirve a los dos grados, con la fecha lectiva de cada uno | ✅ |
| **Momento 2**: el documento sale en solo lectura (`w:edit="readOnly"` + `w:enforcement="1"`) | ✅ |
| **Momento 2**: la huella SHA-512 de la contraseña coincide con una segunda implementación escrita aparte | ✅ |
| **Momento 2**: la contraseña NO queda escrita ni en el documento ni en el navegador | ✅ |
| **Momento 2**: la etiqueta de protección respeta el orden que exige el esquema de Word | ✅ |
| **Momento 2**: sin pedir bloqueo, el documento sigue saliendo editable | ✅ |

---

## Lo que estas pruebas NO pueden demostrar

- **Que Word acepte la contraseña del bloqueo.** En este equipo no hay Word. Lo comprobado es que
  la huella coincide con una segunda implementación del algoritmo publicado, que la etiqueta va en
  el orden que exige el esquema y que **LibreOffice sí abre el documento en solo lectura**
  (`LoadReadonly = true` al convertirlo, frente a `false` en el documento sin bloquear). Conviene
  verificar la contraseña una vez, en Word, con un documento de prueba.
- **Que el bloqueo resista a quien quiera romperlo.** No es su propósito: es un sello. El .docx es
  un .zip y la protección es una línea de XML.

---

## Modo «doble clic» (sin servidor web)

`prueba_sin_servidor.py` abre la app con `file://`, como cuando el aprendiz hace doble clic sobre
`index.html` desde una carpeta o una USB, y genera los tres formatos:

| Comprobación | Resultado |
|---|---|
| El menú abre con doble clic | ✅ |
| Los tres formatos encuentran su plantilla **sin pedirla a mano** | ✅ |
| Cambiar de grado en el 023 tampoco la pide | ✅ |
| El grupo guardado en un formato llega a los otros dos | ✅ |
| GFPI-F-023 generado sin servidor: datos, plantilla de grado 10 y NIT correctos | ✅ |
| GFPI-F-193 generada sin servidor: conserva las 41 partes del .docx | ✅ |
| GFPI-F-147 generada sin servidor: 132 fusiones intactas | ✅ |
| GFPI-F-023 Momento 2 generado sin servidor: valoración, fecha y grado correctos | ✅ |
| Sin errores de JavaScript en ninguna página | ✅ |
| GFPI-F-193: plantilla intacta (4 imágenes, 41 partes) | ✅ |
| GFPI-F-193: acta, personal vinculado, párrafo de conclusiones y tabla de aprobación | ✅ |
| GFPI-F-193: las filas sobrantes de la tabla quedan en blanco | ✅ |
| GFPI-F-193: las copias son idénticas byte a byte | ✅ |
| Validaciones bloquean la generación (horario, fechas, documentos) | ✅ |
| Sin errores de consola en ninguna de las cuatro páginas | ✅ |

La prueba del 023 no se limita a «hay texto»: compara contra los 22 valores exactos del documento
que llenó Danna, `paraId` por `paraId`. Si el mapa se desalineara aunque fuera un párrafo, la
prueba falla.
