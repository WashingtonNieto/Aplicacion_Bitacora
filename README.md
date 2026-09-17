# Formatos de etapa productiva — GFPI-F-023, GFPI-F-193 y GFPI-F-147

Aplicación web que reemplaza el diligenciamiento manual de los formatos del SENA que entregan los
aprendices de Articulación con la Educación Media:

| # | Formato | Qué es | Cómo se diligencia | Salida |
|---|---|---|---|---|
| 1 | **GFPI-F-023 v06** | Planeación, Seguimiento y Evaluación de Etapa Productiva (Momento 1) | **Individual**: cada documento lleva los datos de un solo aprendiz | Un `.docx` por aprendiz, **cada uno distinto** |
| 2 | **GFPI-F-193 v3** | Acta de inicio y confidencialidad del proyecto productivo | **Grupal**: el documento lleva a todo el equipo | Una copia por integrante, **idénticas** |
| 3 | **GFPI-F-147 v05** | Bitácora de seguimiento de etapa productiva | **Grupal**: el documento lleva a todo el equipo | Una copia por integrante, **idénticas** |
| 4 | **GFPI-F-023 v06 · Momento 2** | Seguimiento de la etapa productiva. **Lo diligencia el instructor** | **Mixto**: el encabezado es del aprendiz que recibe la copia; la tabla de valoración lleva a todo el equipo | Un `.docx` por aprendiz, **cada uno distinto** |

En los formatos grupales el contenido es el mismo para todos y lo único que cambia es el nombre del
archivo, para que cada integrante entregue el suyo. En el individual, cada documento trae los datos
de su aprendiz: nombre, documento, teléfono, dirección, correos y su firma.

**El equipo se arma una vez, desde cualquiera de los formatos** («+ Agregar integrante»,
máximo 3) y queda igual en todos. Lo mismo con **los dos instructores**: nombre, cédula, correo,
teléfono y firma se capturan una sola vez en la tarjeta «Instructores» y quedan guardados con el
grupo, listos para los documentos siguientes.

> El **instructor técnico** es la misma persona que la bitácora y el GFPI-F-023 rotulan «ente
> co-formador»: se guarda una vez y cada formato le pone su propio rótulo.

En los cuatro casos la app **abre la plantilla oficial y escribe dentro de ella**: nunca construye
un documento nuevo, así que logos, tablas, bordes y textos legales quedan exactamente como los
publicó el SENA. Y los cuatro comparten el mismo **grupo de trabajo**: los aprendices, el colegio,
el instructor y las firmas se escriben una sola vez.

---

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | El menú con los cuatro formatos |
| `planeacion.html` | Formulario del GFPI-F-023 (Momento 1) |
| `acta-inicio.html` | Formulario del GFPI-F-193 |
| `bitacora.html` | Formulario del GFPI-F-147 |
| `seguimiento.html` | Formulario del GFPI-F-023 **Momento 2** (valoración S/M del instructor) |
| `comun.js` | **CONFIG** (colegios, grados, competencias, factores del Momento 2) y los grupos guardados. Lo comparten las cuatro páginas: es el único sitio donde hay que editar esos datos |
| `docx.js` | Motor que edita los formatos de Word conservando el diseño |
| `exceljs.min.js` · `fflate.min.js` | Librerías locales, para que funcione **sin internet** |
| `GFPI-F-147-Bitacora-#_Modelo.xlsx` | Plantillas oficiales. **Deben estar junto a las páginas** |
| `GFPI-F-023_..._Grado10_Modelo.docx` · `..._Grado11_Modelo.docx` · `..._Momento2_Modelo.docx` · `GFPI-F-193..._Modelo.docx` | |
| `probar.js` · `probar_seguimiento.js` · `prueba_ui.py` · `prueba_formatos.py` · `prueba_sin_servidor.py` · `verificar.py` · `verificar_seguimiento.py` | Pruebas automáticas |
| `huella_referencia.py` | Segunda implementación de la huella de contraseña de Word, para contrastar la de `docx.js` |
| `plantillas-023.js` · `plantillas-023m2.js` · `plantillas-193.js` · `plantillas-147.js` | Las plantillas empaquetadas, para que funcione con doble clic. Las genera `empaquetar_plantillas.py` |
| `firmas_demo/` | Firmas que usan las pruebas |

---

## El proceso, paso a paso

**Si cada aprendiz hace el suyo** (GFPI-F-023):

1. Abre el menú y entra a **Planeación, Seguimiento y Evaluación**.
2. En «¿Para quién es este formato?» deja su propio nombre (es lo que viene por defecto).
3. Escribe sus datos en «Datos del aprendiz», o carga el grupo si ya existe.
4. Elige su **grado**: la app carga la plantilla correcta con las fechas de ese grado.
5. Elige el **colegio** de la lista y completa el horario.
6. Carga **su firma** (las del instructor y el ente co-formador ya vienen en el formato).
7. **Generar formatos** → descarga un solo archivo, el suyo.

**Si el instructor los genera todos:** carga el grupo, elige «Todo el grupo» en el selector y
genera; salen tantos documentos como aprendices, cada uno con sus datos y su firma.

---

## Cómo se editan los formatos de Word

Un `.docx` es un `.zip` que por dentro tiene `word/document.xml` con todo el texto. El motor
(`docx.js`) abre ese XML, reemplaza el contenido de párrafos concretos y vuelve a empaquetar el
archivo. Todo lo demás —imágenes, tablas, estilos, encabezados— se queda intacto.

Para las **firmas** hacen falta tres cosas a la vez: el archivo en `word/media/`, una relación en
`word/_rels/document.xml.rels` que le dé un identificador, y un `<w:drawing>` dentro del párrafo
que apunte a ese identificador. Si falta una, Word abre el documento con el recuadro rojo de
«imagen no encontrada».

Cada párrafo se identifica por su **`w14:paraId`**, un código único que Word le asigna y que no
cambia al escribir dentro de él. Es el equivalente a la referencia de celda del Excel. Los mapas
de los dos formatos (`MAPA_PLANEACION` y `MAPA_ACTA`, dentro de cada página) **no se adivinaron**:
salieron de comparar la plantilla vacía contra los documentos que los aprendices ya habían llenado
a mano, de modo que cada dato cae exactamente donde ellos lo pusieron.

Si algún día el SENA publica una versión nueva de la plantilla, los `paraId` cambian. La app lo
detecta y avisa —«la plantilla no coincide con la esperada»— en vez de entregar un documento a
medias. Para rehacer el mapa basta con volver a comparar plantilla vacía contra documento lleno.

---

## Qué se llena en cada formato

**GFPI-F-023 — Planeación (Momento 1).** Ficha, fecha fin de la etapa lectiva, datos del aprendiz,
alternativa registrada en SofiaPlus, instructor de seguimiento, ente co-formador, fechas y horario
de la etapa productiva. Es individual: **cada aprendiz recibe su propio documento** con sus datos.
El momento 2 tiene su propia opción en el menú (más abajo); el momento 3 (evaluación) todavía se
completa a mano.

**Hay dos plantillas, una por grado**, con valores distintos ya escritos (fechas de la etapa
productiva, registro en SofiaPlus, competencias y resultados de aprendizaje). Al elegir el grado,
la app carga la plantilla que corresponde y muestra en el formulario los valores que esa plantilla
trae, para que el aprendiz los vea y pueda corregirlos:

| | Grado 10 | Grado 11 |
|---|---|---|
| Plantilla | `..._Grado10_Modelo.docx` | `..._Grado11_Modelo.docx` |
| Fecha fin etapa lectiva | 26/11/2027 | 26/11/2026 |
| Etapa productiva | 01/07/26 a 26/11/27 | 01/05/26 a 26/11/26 |
| Nombre del archivo | `..._Grado10_Nombre.docx` | `..._Grado11_Nombre.docx` |

**Un campo vacío no borra lo que la plantilla ya trae.** La app solo escribe los campos
diligenciados; los demás conservan el valor oficial del formato. Por eso, para *vaciar* un dato hay
que hacerlo en Word después de generarlo.

**Cada aprendiz puede hacer el suyo.** Al inicio del formulario hay un selector
**«¿Para quién es este formato?»**:

- Por defecto está en **un solo aprendiz**: aparece la tarjeta «Datos del aprendiz», solo se pide
  su firma y se descarga un único archivo.
- Con un grupo cargado también aparece **«Todo el grupo»**, que genera un documento por aprendiz
  de una sola vez. Es lo que le sirve al instructor.

Sin grupo guardado, el aprendiz escribe sus propios datos y elige su colegio de la lista (la
dirección y el NIT se llenan solos). Al generar, esos datos quedan guardados como grupo y la
próxima vez ya no hay que repetirlos.

**Las firmas sí se insertan aquí.** Las del instructor de seguimiento y el ente co-formador ya
vienen dentro de la plantilla; la app agrega la del aprendiz, en el espacio en blanco sobre el
rótulo «Firma del aprendiz» (párrafo `5448632A`). Si el aprendiz ya cargó su firma en la bitácora,
aparece puesta sin volver a pedirla.

**GFPI-F-193 — Acta de inicio.** Ciudad y fecha, nombre y descripción del proyecto, **los objetivos
específicos (punto 1.1)**, personal vinculado (aprendices e instructores con su documento), el
párrafo de conclusiones —que se arma solo con los nombres— y la tabla de registro de aprobación.
Las copias son idénticas: cambia el nombre del archivo para que cada aprendiz entregue el suyo.

**Los objetivos específicos se escriben uno por renglón** y la app los numera. Si se pegan desde el
plan de negocio ya numerados —o con viñeta, o con guion—, se respetan tal cual: numerar encima
daría «1. 1. Diseñar…». El párrafo del formato viene justificado, así que la app lo pasa a la
izquierda: con saltos de línea, cada renglón se estiraba de margen a margen.

> El **objetivo general** (punto 1) no se pide: el formato ya trae el del SENA, que es el mismo
> para todos los proyectos productivos, y la app no lo toca.

El acta es el primer formato del ciclo, así que **el equipo se puede escribir ahí mismo** con
«+ Agregar integrante», sin necesidad de un grupo previo. Lo que se escriba queda guardado y
aparece luego en la bitácora y en el 023.

De cada integrante se piden nombre, tipo y número de documento, teléfono y **los dos correos**. La
tabla de aprobación del formato tiene una sola columna «CORREO ELECTRÓNICO», así que ahí van los
dos: el institucional arriba y el personal debajo. Se exige al menos uno.

**Una viñeta por persona.** El formato trae dos renglones en «Personal vinculado» (a. y b.). Si el
equipo es de tres, la app **clona el renglón** —con su viñeta y con el espacio en blanco que el
formato deja entre uno y otro— y Word sigue numerando solo: a., b., c. Antes los tres nombres caían
dentro de la «a.» y la «b.» quedaba vacía. Y al revés: los renglones que **sobran se quitan**, no se
dejan en blanco, porque una viñeta «b.» sin nada al lado se lee como un error en un acta oficial.

**Cada instructor en su viñeta.** El «Personal vinculado» tiene tres viñetas y la app pone a cada
quien una sola vez:

| Viñeta del formato | Quién va ahí |
|---|---|
| «…con rol de aprendiz» | los integrantes del equipo |
| «…instructor de seguimiento y/o instructor técnico» | el **instructor de seguimiento** |
| «…facilitador de seguimiento etapa productiva…, **instructor proyecto productivo** o instructor de investigación - SENNOVA» | el **instructor técnico**, que es quien acompaña el proyecto productivo |

> La viñeta del medio dice «y/o», así que admitiría a los dos. Se dejó uno en cada una porque
> nombrar dos veces a la misma persona en un acta de una página se lee como un error.

**GFPI-F-147 — Bitácora.** Lo de siempre: período, actividades, evidencias, ARL y firmas.

**Las firmas del acta** van en la columna «FIRMA» de la tabla de aprobación: una por integrante y
la del instructor técnico. Se piden en el mismo formulario y se reutilizan las que ya estén
cargadas en los otros formatos.

**El texto que escribe la app va en negro.** Los textos de ejemplo de la plantilla
—«<Nombre del proyecto>», «<Ciudad y fecha>», la descripción— están en rojo a propósito: son marcas
para quien diligencia. Si se heredara ese color, el dato real saldría rojo. La app conserva
negrita, tamaño, fuente y alineación, y solo cambia el color.

**GFPI-F-023 Momento 2 — Seguimiento.** Lo diligencia **el instructor**, una vez por visita de
seguimiento. Es un archivo aparte del Momento 1 (`..._Momento2_Modelo.docx`): conserva el mismo
encabezado «Información general» y en lugar de la planeación trae la tabla de valoración.

Por cada aprendiz del equipo se marca **S** (satisfactorio) o **M** (por mejorar) en trece
factores —ocho técnicos y cinco actitudinales— y se escribe una observación con el compromiso de
mejora. **Todos empiezan en S**: solo hay que bajar a M lo que corresponda.

> El botón **«Sugerir redacción a partir de las M»** arma la primera frase con los factores que
> quedaron en M —«Presenta una valoración por mejorar (M) en el factor técnico de mejora continua y
> en el factor actitudinal de organización. Se compromete a …»— y deja el cursor donde sigue el
> compromiso. El compromiso lo escribe el instructor: eso no se puede adivinar.

Tres diferencias con el Momento 1, que conviene tener presentes:

| | Momento 1 | Momento 2 |
|---|---|---|
| Plantillas | Dos, una por grado | **Una sola**: la app escribe la fecha fin de etapa lectiva del grado |
| Firmas en la plantilla | Trae las del instructor y el ente co-formador | **No trae ninguna**: la app inserta las tres |
| Alcance | Todo el documento es de un aprendiz | Encabezado del aprendiz + tabla de todo el equipo |

Se genera **un documento por aprendiz**: el encabezado cambia (es su carpeta), la tabla de
valoración es la misma en las tres copias. La tabla del formato admite **hasta 5 aprendices**
(la de factores técnicos trae 7 renglones y la actitudinal 5; manda la más corta); las filas que
sobran quedan en blanco.

Si el seguimiento fue **virtual**, se exige el enlace de la grabación, y la X del renglón de cierre
—«de forma presencial ___ o virtual ___»— se pone donde corresponda. Las firmas del instructor de
seguimiento y del ente co-formador **avisan si faltan, pero no bloquean**: esa línea puede firmarse
a mano el día de la visita.

**El documento sale bloqueado.** Una vez calificado, no debería poder modificarse: por eso el
Momento 2 se genera con la restricción de edición de Word activada —**solo lectura**— y,
opcionalmente, con una contraseña para quitarla. Se controla en la tarjeta «Bloqueo del documento»;
viene marcado por defecto.

> **Qué protege y qué no.** Es un **sello, no cifrado**. Evita que se cambie una calificación por
> descuido o por las buenas, pero un .docx sigue siendo un .zip y quien tenga conocimiento técnico
> puede borrar esa línea. Para un documento realmente cerrado, ábralo en Word y use
> **Guardar como → PDF**.

Tres consecuencias del bloqueo, que conviene tener presentes:

- **Nadie puede escribir después.** Las observaciones del aprendiz y del ente co-formador tienen
  que quedar escritas en el formulario, o anotarse a mano sobre el impreso.
- **La contraseña no se guarda en ninguna parte** —ni en el navegador ni dentro del documento—:
  hay que escribirla en cada tanda. Si se olvida, el documento se vuelve a generar desde la app.
- Si el navegador no ofrece criptografía, el documento sale bloqueado **pero sin contraseña**, y la
  app lo dice en el mensaje: es peor entregar algo que no protege lo que promete.

Por dentro es la etiqueta `<w:documentProtection w:edit="readOnly" w:enforcement="1">` de
`word/settings.xml`, con la huella SHA-512 de la contraseña calculada como la calcula Word
(MS-OFFCRYPTO 2.3.7.1: `H0 = SHA512(sal + contraseña UTF-16LE)`, y luego
`Hn = SHA512(Hn-1 + número de vuelta)`). El cálculo está comprobado contra una segunda
implementación escrita aparte, en Python (`huella_referencia.py`), y LibreOffice abre el documento
en solo lectura. Lo que no se pudo comprobar aquí es que **Word acepte la contraseña**, porque en
este equipo no hay Word: conviene que lo verifique una vez, con un documento de prueba.

**La plantilla no es coherente consigo misma y la app lo corrige.** En la tabla de valoración, unas
casillas vienen centradas y en negrita y otras alineadas a la izquierda y más pequeñas; respetando
cada celda tal cual, la columna de S/M sale desordenada. La app fuerza el mismo formato en las
trece, y alinea la observación a la izquierda (el formato la trae centrada, y un párrafo largo
centrado se lee mal). Es la única excepción a la regla de conservar el formato oficial.

---

## Grupos guardados (no volver a escribir lo mismo)

Los aprendices de un grupo, sus documentos, teléfonos, correos y direcciones, el colegio, el
instructor, el jefe inmediato y **las firmas** son iguales en todas las bitácoras. La app los
guarda como un **grupo** y los reutiliza: de la segunda bitácora en adelante solo hay que escribir
el período y las actividades.

**Se guarda solo.** Al generar las bitácoras, el grupo queda guardado con el nombre
`Ficha 3192383 — Danna, Lizeth, Victor`. La próxima vez que se abra la app, aparece en la lista
«Grupo guardado» y con elegirlo se llena todo. Si algo cambió (un teléfono, una firma), se corrige
y al generar de nuevo el grupo se actualiza.

**Botones de la tarjeta «Grupo de trabajo»:**

| Botón | Para qué |
|---|---|
| **Grupo nuevo** | Vacía aprendices, colegio y firmas para empezar otra ficha desde cero. Los grupos guardados **no se borran**: siguen en la lista |
| Guardar grupo actual | Guarda sin necesidad de generar la bitácora |
| Exportar a archivo | Descarga `Grupo-3192383.json` con todo el grupo, firmas incluidas |
| Importar de archivo | Carga ese archivo en cualquier computador |
| Eliminar grupo | Borra el grupo guardado. No borra lo que esté escrito en el formulario |

La lista «Grupo guardado» empieza siempre con **«➕ Grupo nuevo — escribir los datos desde cero»**,
que hace lo mismo que el botón. Es la forma de pasar de una ficha a otra sin perder ninguna.

**Los grupos viven en el navegador donde se guardaron.** Si el aprendiz cambia de computador, usa
otro navegador o el equipo del aula se limpia, la lista aparece vacía. Para eso está exportar:
el archivo `.json` se guarda en Drive o en una USB y se importa donde sea. *Limpiar formulario* no
borra los grupos guardados.

**Para el instructor:** puede armar el grupo de cada ficha una sola vez, exportarlo y entregar el
`.json` a los aprendices. Ellos lo importan y arrancan con todos los datos puestos. Tenga en cuenta
que ese archivo lleva datos personales (documentos, direcciones, teléfonos y firmas), así que
compártalo solo con el grupo que corresponde.

---

## Grados y calendario de bitácoras

En el paso 1 se elige el **grado escolar**, y de esa elección dependen el calendario y la lista
de competencias:

| | Grado 11 | Grado 10 |
|---|---|---|
| Frecuencia | Dos por mes (quincenal) | Una por mes |
| Períodos | Impar: del 1 al 15 · Par: del 15 al 30 | Del 1 al 30 del mes |
| Bitácora 1 | Mayo de 2026 | Agosto de 2026 |
| Total | 12 (mayo a octubre de 2026) | 12 (agosto de 2026 a septiembre de 2027) |

Dos reglas del calendario escolar colombiano están incorporadas:

- **Diciembre y enero no llevan bitácora.** El año escolar va hasta noviembre y se retoma en
  febrero, así que el conteo pasa por encima de esos dos meses: en grado 10, la bitácora 4 es
  noviembre de 2026 y la 5 es **febrero de 2027**.
- **El cierre nunca pasa del último día real del mes.** Febrero cierra el 28 (o el 29 en año
  bisiesto), no el 30.

Para ajustar cualquiera de las dos cosas, edita `CONFIG.grados` en el `comun.js`:

```js
grados: {
  '10': {
    etiqueta: 'Grado 10 — una bitácora por mes',
    anioBase: 2026, mesBase: 8,     // bitácora 1 = agosto de 2026
    bitacorasPorMes: 1,             // 1 = mensual · 2 = quincenal
    totalBitacoras: 12,
    mesesExcluidos: [12, 1],        // vacaciones
    competencias: [ ... ]
  }
}
```

> ⚠️ **Pendiente:** las competencias de grado 10 están sin cargar; solo aparece
> "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA". Reemplaza esa lista por las de la ficha de grado 10,
> con el formato `código - NOMBRE EN MAYÚSCULAS`.

Al cambiar de grado, la app rehace la lista de bitácoras, recalcula el período y **reconcilia las
competencias**: si una actividad tenía seleccionada una competencia que no existe en el otro
grado, la cambia por la primera de la lista nueva, para que nunca quede apuntando a algo que no
corresponde.

El grado **no se escribe en el Excel**: el formato GFPI-F-147 no tiene ese campo. Solo determina
el calendario y las competencias.

---

## Los nombres salen como nombre propio

En los **cuatro formatos**, los nombres de persona se escriben con mayúscula inicial, venga el dato
como venga: `SANTIAGO PINEDA MEJIA` y `santiago pineda mejia` salen los dos como
**`Santiago Pineda Mejia`**. Así el documento se ve igual aunque cada aprendiz escriba su nombre a
su manera. Alcanza a los aprendices y a los dos instructores, en el texto de los formatos y también
en el **nombre de los archivos**.

Tres detalles del tratamiento:

- **Las partículas van en minúscula**: `Juan de la Cruz`, no `Juan De La Cruz`.
- **Las siglas se respetan**: `SENA-Washington Leon Nieto Arce` no se convierte en `Sena-Washington`.
  La lista está en `SIGLAS_NOMBRE` (`comun.js`), por si hay que agregar alguna.
- **Los acentos no se inventan**: si el nombre se escribió sin tilde, sigue sin tilde. Adivinar
  acentos daría errores en apellidos.

No se tocan el nombre de la **entidad** (`COLEGIO TECNICO PALERMO I.E.D` va así en el formato) ni el
**cargo** (`INSTRUCTOR`): no son nombres de persona.

---

## La ficha es obligatoria

El **número de grupo (ficha)** no puede faltar en ninguno de los cuatro formularios: sin él los
documentos no se pueden asociar al grupo en SofiaPlus. Si está vacío, ninguno de los cuatro genera
nada y el error señala el campo.

Va impreso en el **GFPI-F-023** (los dos momentos, en «No. Grupo») y en la **bitácora GFPI-F-147**
(celda `B27`, «Número de grupo», **centrado**: la plantilla trae esa celda alineada a la izquierda
mientras todas las de su fila están centradas, y el número quedaba descolgado contra el borde). El **acta GFPI-F-193** no tiene casilla para la ficha, pero
también la exige: queda guardada con el grupo y de ahí sale en los otros tres.

**Al cambiar de colegio, la ficha se actualiza.** La app distingue una ficha que puso ella al elegir
un colegio —la reemplaza por la del colegio nuevo— de una que se escribió a mano, que respeta. Antes
se conservaba siempre, y al pasar de un colegio a otro el documento salía con la ficha del anterior.

---

## Las fechas van con el año completo

En todos los formatos, las fechas que escribe la app salen como **26/11/2026**, con las cuatro
cifras del año. Son documentos que se archivan y se consultan años después: «26/11/26» obliga a
adivinar el siglo, y en una carpeta con bitácoras de 2026 y 2027 la diferencia importa.

Vale para los cuatro formatos: las fechas de la etapa productiva y del registro en SofiaPlus en el
GFPI-F-023, las del momento de seguimiento y el renglón de cierre en el Momento 2, las de inicio y
fin de cada actividad en la bitácora, y el eco que aparece bajo cada campo de fecha.

> Los **rótulos impresos** del formato oficial siguen diciendo «(dd/mm/aa)» o «(DD/MM/AA)». No se
> tocan: son texto del SENA, y la regla de esta aplicación es escribir dentro del formato sin
> alterar lo que el SENA publicó. Si el centro autoriza cambiarlos, se hace en un momento.

---

## Colegios

En el paso 3 el aprendiz elige su colegio de una lista desplegable y la app llena solos el
**nombre, la ficha, la dirección, el NIT y el correo electrónico**. Los cuatro colegios vienen
cargados:

| Colegio | Ficha | Dirección | NIT | Correo |
|---|---|---|---|
| COLEGIO FERNANDO SOTO APARICIO I.E.D | 3218687 | CRA 73 A Bis B No. 36-47 Sur | 860.532.538-3 | corpoandares@gmail.com |
| COLEGIO MAGDALENA ORTEGA DE NARIÑO I.E.D | 3191165 | CRA 69B # 78A-36 | 830.000.292 | lnalmagdalenaorteg10@educacionbogota.edu.co |
| COLEGIO REPUBLICA ESTADOS UNIDOS DE AMERICA I.E.D | 3191243 | CRA 23 # 24c-22 Sur | 899-99735-4 | coldirepeeuudeamer18@educacionbogota.edu.co |
| COLEGIO TECNICO PALERMO I.E.D | 3450523 | CRA 23 # 49-27 | 830036734-4 | cedpalermocedip13@educacionbogota.edu.co |

El **correo de la entidad** solo se imprime en el GFPI-F-023, que es el único de los tres formatos
que lo pide; en la bitácora queda guardado con el grupo para que el 023 lo encuentre ya escrito.

Los tres campos quedan editables. Si alguien los modifica a mano, la lista pasa sola a
**"Otra entidad"**, para que nunca quede afirmando un colegio que no corresponde a lo escrito.
Elegir un colegio de la lista es obligatorio.

Para agregar, quitar o corregir colegios, edita el arreglo `CONFIG.colegios` de `comun.js`:

```js
colegios: [
  { nombre: 'COLEGIO ...', direccion: 'CRA ...', nit: '...', ficha: '3218687',
    correo: 'algo@educacionbogota.edu.co' },
  ...
]
```

**La ficha ya escrita no se pisa.** Si el aprendiz cambia de colegio después de haber escrito un
número de ficha, la app cambia la dirección y el NIT pero respeta la ficha: puede haber más de un
grupo por colegio, y sobrescribirla en silencio sería peor que dejarla.

---

## Firmas

En el paso 6 se carga **una imagen por firmante**: cada aprendiz, el instructor de seguimiento y
el ente co-formador. La app las inserta como imágenes centradas sobre la línea que les
corresponde en el formato (`C80:E80` para el aprendiz 1, `H80:I80` para el 2, `C83:E83` para el 3,
`B91:E91` para el instructor y `H91:I91` para el co-formador).

- Formatos aceptados: **PNG o JPG**. Consejo para los aprendices: firmar en una hoja blanca y
  tomar la foto de cerca, con buena luz.
- La imagen se reduce a 600 px de ancho antes de guardarla, así una foto de celular de 4 MB no
  desborda el borrador local.
- Las firmas de los **aprendices son obligatorias**; las del instructor y el ente co-formador son
  **opcionales**, porque muchas veces se firman a mano sobre el impreso. Si el instructor prefiere,
  puede guardar una imagen de su firma y reutilizarla cada quincena.

---

## Una copia por aprendiz

Al pulsar **Generar bitácoras** se descarga un archivo por cada aprendiz registrado, todos con el
mismo contenido y el nombre de cada uno:

```
GFPI-F-147-Bitacora-1_Danna Lizeth Puerto Diaz.xlsx
GFPI-F-147-Bitacora-1_Lizeth Mariana Caita Beltran.xlsx
GFPI-F-147-Bitacora-1_Victor Santiago Pérez Mosqueda.xlsx
```

El libro se serializa **una sola vez** y ese mismo contenido se descarga con cada nombre, así las
copias son idénticas byte a byte y no hay forma de que una quede desactualizada.

La primera vez, Chrome o Edge preguntan «¿Descargar varios archivos?»: hay que elegir **Permitir**.
La app lo advierte en pantalla justo debajo del botón.

---

## Cómo usarla

### Opción A — publicarla en GitHub Pages (recomendada para los aprendices)

1. Crea un repositorio nuevo en GitHub, por ejemplo `bitacora-gfpi`.
2. Sube **todos** los archivos a la raíz: las cuatro páginas `.html`, `comun.js`, `docx.js`,
   `exceljs.min.js`, `fflate.min.js` y las tres plantillas oficiales.
3. En el repositorio: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / carpeta `/ (root)`** → *Save*.
4. A los dos minutos la app queda en `https://TU-USUARIO.github.io/bitacora-gfpi/`.
5. Comparte ese enlace. Los aprendices no instalan nada.

### Opción B — en el computador del aula, sin publicar

Con Python instalado, dentro de la carpeta:

```bash
python -m http.server 8000
```

y abre `http://localhost:8000` en el navegador.

### Opción C — doble clic en `index.html`

**También funciona.** Al abrir la página con doble clic (dirección `file://`) el navegador prohíbe
leer archivos de la carpeta, así que las plantillas van además **empaquetadas dentro de la app**:
`plantillas-023.js`, `plantillas-193.js` y `plantillas-147.js` las llevan en base64, y un `.js` sí
se puede cargar sin servidor.

La app intenta primero leer el `.docx` o `.xlsx` de la carpeta —así, si el SENA publica una versión
nueva, el sitio publicado la usa de inmediato— y solo si no puede recurre a la copia empaquetada.
El botón para elegir la plantilla a mano queda como último recurso.

> Si reemplazas una plantilla oficial, vuelve a ejecutar `python3 empaquetar_plantillas.py` para
> regenerar esas copias. Si lo olvidas, las opciones A y B siguen usando el archivo nuevo; solo el
> modo doble clic se quedaría con el anterior.

---

## Cómo cambiar los datos precargados

Todo está en el bloque `CONFIG`, al inicio de `comun.js`. Un solo archivo para los tres
formatos:

```js
const CONFIG = {
  anioBase: 2026,        // la bitácora 1 arranca en mayo de 2026
  mesBase: 5,
  maxAprendices: 3,
  maxActividades: 5,
  dominioInstitucional: 'educacionbogota.edu.co',
  predeterminados: { entidad: '...', instructorNombre: '...', ... },
  competencias: [ '220501096 - DESARROLLAR LA SOLUCIÓN...', ... ]
};
```

- **Colegios** → edita `colegios` (ver la sección anterior).
- **Otra ficha, instructor o jefe inmediato** → cambia `predeterminados`. Ojo: eso solo afecta a
  los grupos **nuevos**. Los grupos ya guardados en el navegador conservan lo que tenían; para
  actualizarlos hay que abrirlos, corregir la tarjeta «Instructores» y volver a guardar.
- **Competencias de tu ficha** → agrégalas al arreglo `competencias` **del grado que corresponda**,
  con el formato `código - NOMBRE`.
- **Otro calendario** → `anioBase`, `mesBase`, `bitacorasPorMes`, `totalBitacoras` y
  `mesesExcluidos` de cada grado (ver la sección «Grados y calendario»).
- **Grupos de 4 o 5 aprendices** → sube `maxAprendices` hasta 5; el mapa de celdas, las anclas de
  firma y la tabla del Momento 2 ya contemplan las cinco filas del formato.
- **Factores del Momento 2** → si el SENA cambia las columnas de la tabla de valoración, se editan
  `FACTORES_TECNICOS` y `FACTORES_ACTITUDINALES` en `comun.js`, y los `paraId` correspondientes en
  `FILAS_TECNICOS` / `FILAS_ACTITUDINALES` de `seguimiento.html`.

---

## Cómo correr las pruebas

```bash
npm install exceljs@4.4.0 fflate
node probar.js                  # núcleo del Excel y los dos calendarios
node probar_seguimiento.js      # núcleo del Momento 2: valoraciones, fechas y firmas

pip install playwright openpyxl lxml && playwright install chromium
python3 prueba_ui.py            # la bitácora en un navegador: diligenciar, firmar y descargar
python3 prueba_formatos.py      # el menú y los tres formatos de Word
python3 verificar.py            # los 11 criterios de aceptación del Excel
python3 verificar_seguimiento.py  # el Momento 2 celda por celda (necesita probar_seguimiento.js)
python3 prueba_sin_servidor.py  # los cuatro formatos abiertos con doble clic (file://)
```

`prueba_ui.py` debe correr antes que `prueba_formatos.py`: la primera deja el archivo del grupo que
la segunda importa, que es justamente lo que demuestra que los formatos comparten los datos.

`probar.js` **extrae el código del propio `bitacora.html`** (el bloque entre las marcas
`==== INICIO NUCLEO ====` y `==== FIN NUCLEO ====`) en lugar de tener una copia. Así la prueba
nunca se desincroniza de lo que corre en el navegador.

---

## Decisiones de diseño que conviene conocer

**La plantilla no está vacía.** Trae datos de un diligenciamiento anterior: un período en `E11`,
una "X" en proyecto productivo, el programa de formación, la tabla ARL completa y una fecha de
entrega. La función `limpiarResiduos()` los borra antes de escribir. Si algún día cambian la
plantilla, revisa esa función.

**ExcelJS conserva las imágenes, pero descarta las formas.** La plantilla marca la clasificación
"Pública" con una *forma* dibujada (un cuadro rojo), no con una imagen: al reescribir el archivo
desaparecía. La app la vuelve a insertar como imagen en la misma posición
(`restaurarCuadroPublica()`). Los tres logos sí sobreviven sin ayuda.

**Fechas en el formulario.** Los campos `<input type="date">` los dibuja el navegador con el
formato de **su** idioma: en inglés se ven `07/15/2026` y en español `15/07/2026`. La página no
puede cambiarlo. Por eso, bajo cada campo, la app muestra la fecha en `dd/mm/aaaa` — exactamente como
va a quedar escrita en el formato oficial. Está en `comun.js` (`ecoDeFechas`) y funciona en los tres
formularios.

**Fechas en el archivo.** Se escriben como objetos `Date` en UTC con formato `dd/mm/yyyy`. Usar fechas locales
hace que Excel muestre el día anterior; usar texto hace que dejen de ser fechas (el archivo de
ejemplo tenía una escrita como `15/7/0206`).

**Documentos y teléfonos van como texto**, para no perder ceros iniciales ni terminar en
notación científica.

**Colores heredados.** Algunas celdas de la plantilla traen fuente gris casi invisible
(`A5A5A5`, `BFBFBF`) o azul de hipervínculo, residuo de usos anteriores. Cuando la app escribe
un dato real en una de ellas lo pasa a negro; no toca ningún otro atributo.

**Tabla ARL.** La plantilla trae fórmulas `=B15`…`=B17` que copiaban el nombre *con su numeral*
("1. Danna…"). La app las reemplaza por el nombre en texto plano. Los cuatro campos van en `SI`
y el nivel de riesgo en 1, según lo definido para el programa.

**Firmas como imagen, no en la celda.** Las celdas `B80`, `F80`, `B83`, `C92`, `H92`… contienen
los *rótulos* ("Firma del aprendiz 1"), no el espacio de firma: la línea es el borde inferior de
las celdas vecinas. Por eso las firmas se insertan como imágenes ancladas dos filas arriba de esa
línea y centradas sobre ella, sin tocar ningún texto del formato.

**Actividades.** El formato trae tres bloques completamente formateados (filas 53, 55 y 57) y dos
a medio formatear (59 y 61: solo tienen fusionada la columna de competencias). Al usar los slots
4 y 5 la app replica fusiones, bordes y alto de fila desde la fila 53.

---

## Fase 2 (preparada, no implementada)

La capa de datos ya está aislada tras la interfaz `BitacoraStore` (`guardar`, `listar`, `obtener`),
implementada hoy sobre `localStorage`. Sustituirla por Supabase o Firebase no exige tocar el
formulario ni el núcleo de escritura del Excel.

Con backend, los grupos guardados dejarían de vivir en cada navegador: el aprendiz entraría con su
correo institucional y encontraría su grupo desde cualquier equipo, sin exportar ni importar nada.
Mientras tanto, el archivo `.json` cumple esa función. El panel del instructor sería una segunda página
que consume ese mismo store: filtros por grupo, período y aprendiz; estado de la entrega;
regeneración del `.xlsx` de cualquier registro; y exportación masiva a `.zip` por ficha y quincena.
