/* =========================================================================
   CÓDIGO COMÚN A LOS TRES FORMATOS
   =========================================================================
   Aquí vive todo lo que comparten la Bitácora (GFPI-F-147), el Acta de inicio
   (GFPI-F-193) y la Planeación y Seguimiento (GFPI-F-023):

     · CONFIG            → colegios, grados, competencias y datos precargados.
                           Es el ÚNICO sitio donde hay que editarlos.
     · Perfiles de grupo → los aprendices, el colegio, el instructor y las
                           firmas, guardados una vez y reutilizados por los
                           tres formatos: un grupo guardado desde la bitácora
                           aparece también en el acta y en la planeación.
     · Utilidades de texto y fecha, y la descarga de archivos.

   Este archivo se carga antes que el código propio de cada formulario.
   ========================================================================= */

/* =========================================================================
   CONFIGURACIÓN EDITABLE POR EL INSTRUCTOR
   Cambia estos valores y la aplicación queda adaptada a tu centro/ficha.
   ========================================================================= */
const CONFIG = {
  // Nombre del archivo de plantilla que está junto a este index.html
  archivoPlantilla: 'GFPI-F-147-Bitacora-#_Modelo.xlsx',
  hoja: 'Formato Bitácora Art. Media',

  // Cada grado tiene su propio calendario de bitácoras y sus competencias.
  //   bitacorasPorMes: 2 → quincenal (impar: del 1 al 15; par: del 15 al 30)
  //   bitacorasPorMes: 1 → mensual   (del 1 al 30 del mes)
  //   mesBase: mes de la bitácora 1 (1 = enero). El ciclo puede cruzar de año.
  grados: {
    '11': {
      etiqueta: 'Grado 11 — dos bitácoras por mes (quincenal)',
      anioBase: 2026, mesBase: 5, bitacorasPorMes: 2, totalBitacoras: 12,
      mesesExcluidos: [12, 1],        // vacaciones: el año escolar se retoma en febrero
      // GFPI-F-023: plantilla propia del grado y los valores que ya trae escritos
      plantillaPlaneacion: 'GFPI-F-023_V06.__Formato_de_Planeación_Seguimiento_y_Evaluación_de_Etapa_Productiva_Grado11_Modelo.docx',
      etiquetaArchivo: 'Grado11',
      planeacion: {
        fechaFinLectiva: '26/11/2026',
        alternativaRegistrada: 'PROYECTO PRODUCTIVO',
        fechaRegistroSofia: '2026-05-08',
        fechaInicioEP: '2026-05-01',
        fechaFinEP: '2026-11-26'
      },
      competencias: [
        '220501096 - DESARROLLAR LA SOLUCIÓN DE SOFTWARE DE ACUERDO CON EL DISEÑO Y METODOLOGÍAS DE DESARROLLO',
        '220501113 - ADMINISTRAR BASE DE DATOS DE ACUERDO CON LOS ESTÁNDARES Y REQUISITOS TÉCNICOS',
        '220501046 - ANALIZAR LOS REQUISITOS DEL CLIENTE PARA CONSTRUIR EL SISTEMA DE INFORMACIÓN',
        '220501097 - APLICAR PRÁCTICAS DE PRUEBA DE SOFTWARE PARA VERIFICAR EL FUNCIONAMIENTO',
        'RESULTADOS DE APRENDIZAJE ETAPA PRACTICA'
      ]
    },
    '10': {
      etiqueta: 'Grado 10 — una bitácora por mes',
      anioBase: 2026, mesBase: 8, bitacorasPorMes: 1, totalBitacoras: 12,
      mesesExcluidos: [12, 1],        // vacaciones: el año escolar se retoma en febrero
      plantillaPlaneacion: 'GFPI-F-023_V06.__Formato_de_Planeación_Seguimiento_y_Evaluación_de_Etapa_Productiva_Grado10_Modelo.docx',
      etiquetaArchivo: 'Grado10',
      planeacion: {
        fechaFinLectiva: '26/11/2027',
        alternativaRegistrada: 'PROYECTO PRODUCTIVO',
        fechaRegistroSofia: '2026-03-09',
        fechaInicioEP: '2026-07-01',
        fechaFinEP: '2027-11-26'
      },
      // ⚠️ PENDIENTE: reemplaza esta lista por las competencias de la ficha de
      //    grado 10, con el mismo formato «código - NOMBRE EN MAYÚSCULAS».
      competencias: [
        'RESULTADOS DE APRENDIZAJE ETAPA PRACTICA'
      ]
    }
  },
  gradoPredeterminado: '11',

  maxAprendices: 3,    // el formato admite 5 filas, pero el grupo trabaja con 3
  maxActividades: 5,

  dominioInstitucional: 'educacionbogota.edu.co',

  // Colegios con los que se trabaja. Al elegir uno en el formulario se llenan
  // solos la ficha, la dirección, el NIT y el correo de la entidad.
  // Para agregar o quitar colegios, edita esta lista: no hay que tocar nada más.
  colegios: [
    { nombre: 'COLEGIO FERNANDO SOTO APARICIO I.E.D',
      direccion: 'CRA 73 A Bis B No. 36-47 Sur', nit: '860.532.538-3', ficha: '3218687',
      correo: 'corpoandares@gmail.com' },
    { nombre: 'COLEGIO MAGDALENA ORTEGA DE NARIÑO I.E.D',
      direccion: 'CRA 69B # 78A-36', nit: '830.000.292', ficha: '3191165',
      correo: 'lnalmagdalenaorteg10@educacionbogota.edu.co' },
    { nombre: 'COLEGIO REPUBLICA ESTADOS UNIDOS DE AMERICA I.E.D',
      direccion: 'CRA 23 # 24c-22 Sur', nit: '899-99735-4', ficha: '3191243',
      correo: 'coldirepeeuudeamer18@educacionbogota.edu.co' },
    { nombre: 'COLEGIO TECNICO PALERMO I.E.D',
      direccion: 'CRA 23 # 49-27', nit: '830036734-4', ficha: '3450523',
      correo: 'cedpalermocedip13@educacionbogota.edu.co' }
  ],

  // Valores precargados (editables por el aprendiz en el formulario)
  predeterminados: {
    grupo: '',
    modalidadFormacion: 'Presencial',
    programa: 'Programación de Software',
    modalidadEjecucion: 'Presencial',
    colegio: '',              // vacío = el aprendiz debe elegirlo en la lista
    entidad: '',
    nit: '',
    direccionEntidad: '',
    jefeNombre: 'Washington Nieto Arce',
    jefeDocumento: '',
    jefeCargo: 'Instructor',
    jefeTelefono: '',
    jefeCorreo: 'wnieto@sena.edu.co',
    instructorNombre: 'Daniel Alberto Ardila Carrasquilla',
    instructorDocumento: '',
    instructorCorreo: 'dardilac@sena.edu.co',
    instructorTelefono: '',
    alternativa: 'proyectoProductivo'
  },

  tiposDocumento: ['T.I', 'C.C', 'C.E', 'PPT'],

  // GFPI-F-023 Momento 2 (seguimiento). A diferencia del Momento 1, aquí hay
  // UNA sola plantilla para los dos grados: lo único que cambia entre grados es
  // la fecha fin de la etapa lectiva, y esa la escribe la app.
  plantillaSeguimiento:
    'GFPI-F-023_V06.__Formato_de_Planeación_Seguimiento_y_Evaluación_de_Etapa_Productiva_Momento2_Modelo.docx',
  ciudadPredeterminada: 'Bogotá'
};

/* Los factores que valora el instructor en el Momento 2, en el mismo orden en
   que aparecen las columnas del formato oficial. Si el SENA cambia el formato,
   se corrigen aquí y en MAPA_SEGUIMIENTO (seguimiento.html). */
const FACTORES_TECNICOS = [
  ['aplicacionConocimiento',    'Aplicación de conocimiento'],
  ['mejoraContinua',            'Mejora continua'],
  ['fortalecimientoOcupacional','Fortalecimiento ocupacional'],
  ['oportunidadCalidad',        'Oportunidad y calidad'],
  ['responsabilidadAmbiental',  'Responsabilidad ambiental'],
  ['administracionRecursos',    'Administración de recursos'],
  ['seguridadSalud',            'Seguridad y salud en el trabajo'],
  ['documentacionEP',           'Documentación etapa productiva']
];

const FACTORES_ACTITUDINALES = [
  ['relacionesInterpersonales', 'Relaciones interpersonales'],
  ['trabajoEquipo',             'Trabajo en equipo'],
  ['solucionProblemas',         'Solución de problemas'],
  ['cumplimiento',              'Cumplimiento'],
  ['organizacion',              'Organización']
];

/* Valoración vacía de un aprendiz: todo «S» (satisfactorio), que es el caso
   normal; el instructor baja a «M» solo lo que haya que mejorar. */
function valoracionVacia() {
  const v = { observaciones: '' };
  FACTORES_TECNICOS.concat(FACTORES_ACTITUDINALES).forEach(([clave]) => { v[clave] = 'S'; });
  return v;
}


/* Limpia espacios raros: la plantilla y los archivos de ejemplo traen
   espacios de no separación (NBSP, U+00A0) que rompen comparaciones y búsquedas. */
function normalizar(texto) {
  if (texto === null || texto === undefined) return '';
  return String(texto).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
}

/* Convierte 'aaaa-mm-dd' (lo que entrega <input type="date">) en un Date UTC.
   Se usa UTC a propósito: ExcelJS convierte a número de serie con base UTC,
   y con fechas locales el archivo termina mostrando el día anterior. */
function aFechaUTC(iso) {
  if (iso instanceof Date) return iso;
  if (!iso) return null;
  const [a, m, d] = String(iso).split('-').map(Number);
  return new Date(Date.UTC(a, m - 1, d));
}

/* Formato "d/mm/aaaa", igual al que usa el formato oficial: "Desde 1/07/2026 hasta 15/07/2026" */
function fechaTexto(fecha) {
  const f = aFechaUTC(fecha);
  if (!f) return '';
  return `${f.getUTCDate()}/${String(f.getUTCMonth() + 1).padStart(2, '0')}/${f.getUTCFullYear()}`;
}

function ultimoDiaDelMes(anio, mes) {
  return new Date(Date.UTC(anio, mes + 1, 0)).getUTCDate();   // día 0 del mes siguiente
}

/* Devuelve el mes que ocupa la posición `indice` (base 0) de la secuencia
   escolar que arranca en anioBase/mesBase, saltando los meses de vacaciones.
   El año escolar en Colombia va hasta noviembre y se retoma en febrero, así que
   diciembre y enero no llevan bitácora y el conteo pasa por encima de ellos. */
function mesDeSecuencia(anioBase, mesBase, indice, mesesExcluidos) {
  const excluidos = mesesExcluidos || [];
  if (excluidos.length >= 12) return { anio: anioBase, mes: mesBase - 1 };  // no dejar bucle infinito
  let anio = anioBase, mes = mesBase - 1;                    // mes base 0
  const excluido = m => excluidos.includes(m + 1);
  const avanzar = () => { mes++; if (mes > 11) { mes = 0; anio++; } };
  while (excluido(mes)) avanzar();                           // por si el mes base cae en vacaciones
  for (let k = 0; k < indice; k++) { do { avanzar(); } while (excluido(mes)); }
  return { anio, mes };
}

/* Calendario de bitácoras, según el grado:
   - Grado 11 (dos por mes): la impar va del 1 al 15 y la par del 15 al 30 del
     mismo mes. Bitácora 1 = mayo de 2026.
   - Grado 10 (una por mes): del 1 al 30 del mes. Bitácora 1 = agosto de 2026.
   Diciembre y enero se saltan (vacaciones escolares) y el cierre nunca pasa del
   último día real del mes: en febrero es 28, o 29 en año bisiesto, no 30. */
function periodoDeBitacora(n, calendario) {
  const cal = calendario || {};
  const anioBase = cal.anioBase || 2026;
  const mesBase = cal.mesBase || 5;
  const porMes = cal.bitacorasPorMes === 1 ? 1 : 2;

  const desplazamiento = porMes === 1 ? (n - 1) : Math.floor((n - 1) / 2);
  const { anio, mes } = mesDeSecuencia(anioBase, mesBase, desplazamiento, cal.mesesExcluidos);
  const cierre = Math.min(30, ultimoDiaDelMes(anio, mes));

  if (porMes === 1) {
    return { inicio: new Date(Date.UTC(anio, mes, 1)), fin: new Date(Date.UTC(anio, mes, cierre)) };
  }
  const esPrimeraQuincena = n % 2 === 1;
  return {
    inicio: new Date(Date.UTC(anio, mes, esPrimeraQuincena ? 1 : 15)),
    fin:    new Date(Date.UTC(anio, mes, esPrimeraQuincena ? 15 : cierre))
  };
}

/* Escapa texto que se va a insertar en HTML */
function esc(t) {
  return String(t ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/* ---------- Generación de los archivos ---------- */
function descargar(datos, nombre, tipo) {
  const blob = new Blob([datos], {
    type: tipo || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nombre;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/* =========================================================================
   PERFILES DE GRUPO
   Lo que se repite bitácora tras bitácora: los aprendices con sus documentos,
   correos y direcciones, el colegio, el instructor, el jefe inmediato y las
   firmas. Se guarda aparte del borrador, con su propio nombre, y puede
   exportarse a un archivo .json para llevarlo a otro computador o para que el
   instructor lo reparta ya armado.
   (Ojo con los nombres: `estado.grupo` es el número de ficha; "grupo guardado"
   es este perfil completo.)
   ========================================================================= */
const CLAVE_GRUPOS = 'bitacora-gfpi-f147-grupos';
const TIPO_ARCHIVO_GRUPO = 'grupo-bitacora-gfpi-f147';
const VERSION_ARCHIVO_GRUPO = 1;

// Campos del formulario que forman parte del perfil (no incluye número de
// bitácora, fechas ni actividades: eso cambia en cada entrega).
const CAMPOS_GRUPO = ['grado', 'grupo', 'programa', 'modalidadFormacion', 'modalidadEjecucion',
  'colegio', 'entidad', 'nit', 'direccionEntidad', 'correoEntidad',
  // Instructor de seguimiento
  'instructorNombre', 'instructorDocumento', 'instructorCorreo', 'instructorTelefono',
  // Instructor técnico — en la bitácora y en el 023 se rotula «ente co-formador»
  'jefeNombre', 'jefeDocumento', 'jefeCargo', 'jefeCorreo', 'jefeTelefono',
  'alternativa'];

function leerGrupos() {
  try { return JSON.parse(localStorage.getItem(CLAVE_GRUPOS) || '[]'); } catch (e) { return []; }
}
function escribirGrupos(grupos) {
  localStorage.setItem(CLAVE_GRUPOS, JSON.stringify(grupos));   // sin try: el llamador avisa si falla
}

function nombreDeGrupo(datos) {
  const ficha = normalizar(datos.grupo) || 'sin ficha';
  const nombres = (datos.aprendices || [])
    .map(a => normalizar(a.nombre).split(' ')[0]).filter(Boolean).join(', ');
  return `Ficha ${ficha}${nombres ? ' — ' + nombres : ''}`;
}

function datosDelGrupoActual() {
  const datos = { aprendices: JSON.parse(JSON.stringify(estado.aprendices)),
                  firmas: JSON.parse(JSON.stringify(estado.firmas)) };
  CAMPOS_GRUPO.forEach(c => { datos[c] = estado[c]; });
  return datos;
}

function pintarSelectorGrupos() {
  const grupos = leerGrupos();
  $('#grupoGuardado').innerHTML =
    '<option value="">➕ Grupo nuevo — escribir los datos desde cero</option>' +
    grupos.map(g => `<option value="${esc(g.id)}">${esc(g.nombre)}</option>`).join('');
  $('#grupoGuardado').value = estado.grupoGuardado || '';
  $('#btnEliminarGrupo').classList.toggle('oculto', !estado.grupoGuardado);
}

function avisoGrupo(texto, esError) {
  $('#estadoGrupo').textContent = esError ? '' : texto;
  marcarError('grupoGuardado', esError ? texto : '');
  if (!esError && texto) setTimeout(() => { $('#estadoGrupo').textContent = ''; }, 6000);
}

/* Vuelca un perfil en el formulario.
   Cada formulario define su propia `refrescarInterfaz()`: comun.js no sabe qué
   campos pinta cada página, solo le avisa que los datos cambiaron. */
function aplicarGrupo(datos) {
  estado.sinGrupo = false;
  CAMPOS_GRUPO.forEach(c => { if (datos[c] !== undefined) estado[c] = datos[c]; });
  estado.aprendices = (datos.aprendices || [aprendizVacio()]).slice(0, CONFIG.maxAprendices);
  estado.firmas = datos.firmas || {};
  if (typeof refrescarInterfaz === 'function') refrescarInterfaz();
  if (typeof guardarBorrador === 'function') guardarBorrador();
}

/* Empieza un grupo desde cero: limpia los aprendices, las firmas y los datos
   del colegio, pero NO borra los grupos ya guardados: siguen en la lista.
   Sirve cuando el instructor pasa de una ficha a otra, o cuando un aprendiz
   quiere escribir sus propios datos en un equipo donde ya hay un grupo. */
function nuevoGrupo() {
  const hayDatos = (estado.aprendices || []).some(a => normalizar(a.nombre));
  const selector = document.getElementById('grupoGuardado');
  if (hayDatos && !confirm(
      '¿Empezar un grupo nuevo?\n\n' +
      'Se vacían los aprendices, el colegio y las firmas de este formulario.\n' +
      'El grupo guardado NO se borra: puedes volver a elegirlo en la lista.')) {
    if (selector) selector.value = estado.grupoGuardado || '';
    return false;
  }
  estado.grupoGuardado = '';
  estado.sinGrupo = true;              // no volver a cargar uno solo al recargar
  CAMPOS_GRUPO.forEach(c => {
    estado[c] = c === 'grado' ? CONFIG.gradoPredeterminado : (CONFIG.predeterminados[c] ?? '');
  });
  estado.aprendices = [aprendizVacio()];
  estado.firmas = {};
  if (typeof refrescarInterfaz === 'function') refrescarInterfaz();
  if (typeof guardarBorrador === 'function') guardarBorrador();
  pintarSelectorGrupos();
  avisoGrupo('Grupo nuevo: escribe los datos y al generar quedará guardado.');
  return true;
}

function guardarGrupoActual() {
  const datos = datosDelGrupoActual();
  if (!normalizar(datos.aprendices?.[0]?.nombre)) {
    avisoGrupo('Escribe al menos el nombre del primer aprendiz antes de guardar el grupo.', true);
    if (typeof mostrarPaso === 'function') mostrarPaso(1);
    return;
  }
  const grupos = leerGrupos();
  const nombre = nombreDeGrupo(datos);
  // Si ya existe un grupo con la misma ficha y los mismos aprendices, se actualiza
  const existente = grupos.find(g => g.id === estado.grupoGuardado) ||
                    grupos.find(g => g.nombre === nombre);
  const registro = existente || { id: 'g' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7) };
  registro.nombre = nombre;
  registro.datos = datos;
  registro.actualizado = new Date().toISOString().slice(0, 10);
  if (!existente) grupos.push(registro);
  try {
    escribirGrupos(grupos);
  } catch (e) {
    avisoGrupo('No se pudo guardar: el navegador se quedó sin espacio. Exporta el grupo a un archivo o elimina alguno guardado.', true);
    return;
  }
  estado.grupoGuardado = registro.id;
  pintarSelectorGrupos();
  if (typeof guardarBorrador === 'function') guardarBorrador();
  avisoGrupo(`✓ Guardado como «${nombre}». La próxima bitácora ya no habrá que reescribirlo.`);
}

function eliminarGrupoActual() {
  const grupos = leerGrupos();
  const g = grupos.find(x => x.id === estado.grupoGuardado);
  if (!g) return;
  if (!confirm(`¿Eliminar el grupo «${g.nombre}»? Los datos escritos en el formulario no se borran.`)) return;
  try { escribirGrupos(grupos.filter(x => x.id !== g.id)); } catch (e) { /* nada que hacer */ }
  estado.grupoGuardado = '';
  pintarSelectorGrupos();
  if (typeof guardarBorrador === 'function') guardarBorrador();
  avisoGrupo('Grupo eliminado.');
}

function exportarGrupo() {
  const datos = datosDelGrupoActual();
  const nombre = nombreDeGrupo(datos);
  const archivo = { tipo: TIPO_ARCHIVO_GRUPO, version: VERSION_ARCHIVO_GRUPO, nombre, datos };
  const texto = JSON.stringify(archivo, null, 2);
  descargar(new TextEncoder().encode(texto),
            `Grupo-${normalizar(datos.grupo) || 'sin-ficha'}.json`,
            'application/json');
  avisoGrupo('✓ Archivo del grupo descargado. Guárdalo o compártelo con el grupo.');
}

async function importarGrupo(archivo) {
  try {
    const contenido = JSON.parse(await archivo.text());
    if (contenido.tipo !== TIPO_ARCHIVO_GRUPO || !contenido.datos || !Array.isArray(contenido.datos.aprendices)) {
      throw new Error('El archivo no es un grupo de bitácora válido.');
    }
    const grupos = leerGrupos();
    const nombre = contenido.nombre || nombreDeGrupo(contenido.datos);
    const existente = grupos.find(g => g.nombre === nombre);
    const registro = existente || { id: 'g' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7) };
    registro.nombre = nombre;
    registro.datos = contenido.datos;
    registro.actualizado = new Date().toISOString().slice(0, 10);
    if (!existente) grupos.push(registro);
    escribirGrupos(grupos);
    estado.grupoGuardado = registro.id;
    aplicarGrupo(contenido.datos);
    pintarSelectorGrupos();
    avisoGrupo(`✓ Grupo «${nombre}» importado y cargado.`);
  } catch (e) {
    avisoGrupo('No se pudo importar: ' + e.message, true);
  }
}


/* -------------------------------------------------------------------------
   CARGA DE LAS PLANTILLAS OFICIALES
   Al abrir la página con doble clic (dirección file://) el navegador prohíbe
   leer archivos de la carpeta, y por eso aparecía «Selecciona la plantilla
   oficial». Un .js sí se puede cargar sin servidor, así que las plantillas van
   también empaquetadas en base64 (plantillas-023.js, -193.js y -147.js,
   generados por empaquetar_plantillas.py).

   Orden: primero se intenta leer el archivo de la carpeta —así, si el SENA
   publica una versión nueva y se reemplaza el .docx, el sitio publicado la usa
   de inmediato— y solo si no se puede, se recurre a la copia empaquetada.
   ------------------------------------------------------------------------- */
function base64AUint8(base64) {
  const limpio = String(base64).replace(/^data:[^;]+;base64,/i, '');
  const binario = atob(limpio);
  const bytes = new Uint8Array(binario.length);
  for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
  return bytes;
}

function plantillaEmbebida(nombre) {
  const b64 = (window.PLANTILLAS_EMBEBIDAS || {})[nombre];
  return b64 ? base64AUint8(b64) : null;
}

/* Devuelve la plantilla como Uint8Array, o null si no se encontró por ningún
   camino (entonces la página ofrece elegirla a mano). */
async function obtenerPlantilla(nombre) {
  try {
    // encodeURIComponent es indispensable: el "#" del nombre de la bitácora se
    // interpretaría como el inicio de un fragmento de URL.
    const resp = await fetch(encodeURIComponent(nombre));
    if (resp.ok) return new Uint8Array(await resp.arrayBuffer());
  } catch (e) { /* file:// o sin servidor: se usa la copia empaquetada */ }
  return plantillaEmbebida(nombre);
}

/* Formato de correo y de documento/teléfono: los tres formatos validan igual */
const RE_CORREO = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const RE_DIGITOS = /^\d{7,12}$/;

/* -------------------------------------------------------------------------
   ECO DE FECHAS
   Los campos <input type="date"> los dibuja el navegador con el formato de SU
   idioma: en inglés se ven como 07/15/2026 y en español como 15/07/2026. La
   página no puede cambiarlo. Para que no quede duda de qué día se eligió, bajo
   cada campo se muestra la fecha en dd/mm/aaaa, que es exactamente como va a
   quedar escrita en el formato oficial.
   ------------------------------------------------------------------------- */
function fechaCortaVisible(iso) {
  const f = aFechaUTC(iso);
  if (!f) return '';
  const dd = String(f.getUTCDate()).padStart(2, '0');
  const mm = String(f.getUTCMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${f.getUTCFullYear()}`;
}

let ecoOcupado = false;
function ecoDeFechas() {
  if (ecoOcupado) return;
  ecoOcupado = true;
  document.querySelectorAll('input[type="date"]').forEach(campo => {
    let eco = campo.nextElementSibling;
    if (!eco || !eco.classList.contains('fecha-eco')) {
      eco = document.createElement('span');
      eco.className = 'fecha-eco';
      eco.setAttribute('aria-hidden', 'true');   // el valor ya lo anuncia el propio campo
      campo.insertAdjacentElement('afterend', eco);
    }
    const texto = fechaCortaVisible(campo.value);
    if (eco.textContent !== texto) eco.textContent = texto;
  });
  ecoOcupado = false;
}

/* Se activa sola: al cargar, al escribir y cada vez que un formulario repinta
   sus listas (actividades, aprendices, integrantes). */
function activarEcoDeFechas() {
  const estilo = document.createElement('style');
  estilo.textContent =
    '.fecha-eco{display:block;min-height:1em;margin-top:3px;font-size:.74rem;' +
    'color:var(--suave,#5c665e);font-variant-numeric:tabular-nums;letter-spacing:.02em}' +
    '.fecha-eco:not(:empty)::before{content:"→ ";opacity:.6}';
  document.head.appendChild(estilo);

  document.addEventListener('input', ecoDeFechas);
  document.addEventListener('change', ecoDeFechas);
  new MutationObserver(() => ecoDeFechas())
    .observe(document.body, { childList: true, subtree: true });
  ecoDeFechas();
}
// La comprobación de `document` va primero: este archivo también se ejecuta en
// Node cuando corren las pruebas, y ahí no existe.
if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', activarEcoDeFechas);
  } else {
    activarEcoDeFechas();
  }
}

/* -------------------------------------------------------------------------
   INSTRUCTORES
   Los dos instructores son los mismos en los tres formatos y cambian poco: se
   capturan una vez y quedan en el grupo, junto con los aprendices. El técnico
   es la misma persona que la bitácora y el 023 rotulan «ente co-formador».
   ------------------------------------------------------------------------- */
const INSTRUCTORES = [
  { clave: 'instructor', titulo: 'Instructor de seguimiento',
    campos: [
      ['nombre', 'instructorNombre', 'Nombre completo', 'text'],
      ['documento', 'instructorDocumento', 'Cédula', 'text'],
      ['correo', 'instructorCorreo', 'Correo electrónico', 'email'],
      ['telefono', 'instructorTelefono', 'Teléfono', 'text']
    ] },
  { clave: 'coformador', titulo: 'Instructor técnico',
    nota: 'En la bitácora y en el GFPI-F-023 este mismo dato se rotula «ente co-formador».',
    campos: [
      ['nombre', 'jefeNombre', 'Nombre completo', 'text'],
      ['documento', 'jefeDocumento', 'Cédula', 'text'],
      ['cargo', 'jefeCargo', 'Cargo', 'text'],
      ['correo', 'jefeCorreo', 'Correo electrónico', 'email'],
      ['telefono', 'jefeTelefono', 'Teléfono', 'text']
    ] }
];

/* Reduce la imagen antes de guardarla: una foto de celular pesa varios MB y no
   cabría en el borrador local. 600 px de ancho sobran para una firma. */
function procesarImagenFirma(archivo, anchoMaximo) {
  const tope = anchoMaximo || 600;
  return new Promise((resolve, reject) => {
    const lector = new FileReader();
    lector.onerror = () => reject(new Error('No se pudo leer la imagen.'));
    lector.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error('El archivo no es una imagen válida.'));
      img.onload = () => {
        const escala = Math.min(1, tope / img.naturalWidth);
        const ancho = Math.max(1, Math.round(img.naturalWidth * escala));
        const alto = Math.max(1, Math.round(img.naturalHeight * escala));
        const lienzo = document.createElement('canvas');
        lienzo.width = ancho; lienzo.height = alto;
        const ctx = lienzo.getContext('2d');
        ctx.fillStyle = '#fff';                 // los JPG no tienen transparencia
        ctx.fillRect(0, 0, ancho, alto);
        ctx.drawImage(img, 0, 0, ancho, alto);
        resolve({ datos: lienzo.toDataURL('image/png'), ancho, alto });
      };
      img.src = lector.result;
    };
    lector.readAsDataURL(archivo);
  });
}

function pintarInstructores() {
  const caja = document.getElementById('tarjetaInstructores');
  if (!caja) return;
  caja.innerHTML = INSTRUCTORES.map(inst => {
    const firma = (estado.firmas || {})[inst.clave];
    const campos = inst.campos.map(([, clave, rotulo, tipo]) => `
      <div class="campo${clave === 'instructorNombre' || clave === 'jefeNombre' ? ' ancho' : ''}">
        <label for="inst-${clave}">${rotulo}</label>
        <input type="${tipo}" id="inst-${clave}" data-inst="${clave}"
               ${tipo === 'text' && /Documento|Telefono/.test(clave) ? 'inputmode="numeric"' : ''}
               value="${esc(estado[clave] || '')}">
        <span class="msj-error" data-error="${clave}"></span>
      </div>`).join('');
    return `<div class="bloque">
      <div class="bloque-cab"><strong>${inst.titulo}</strong></div>
      ${inst.nota ? `<p class="pista" style="margin:-4px 0 12px">${inst.nota}</p>` : ''}
      <div class="rejilla">${campos}</div>
      <div style="margin-top:14px">
        <label style="font-size:.82rem;font-weight:600">Firma</label>
        ${firma ? `<img class="firma-previa" src="${firma.datos}" alt="Firma de ${esc(estado[inst.campos[0][1]] || inst.titulo)}">`
                : '<div class="firma-linea"></div>'}
        <input type="file" accept="image/png,image/jpeg" data-firma-inst="${inst.clave}">
        ${firma ? `<button type="button" class="btn-min" data-quitar-firma-inst="${inst.clave}">Quitar firma</button>` : ''}
        <span class="msj-error" data-error="firma-${inst.clave}"></span>
      </div>
    </div>`;
  }).join('');
}

/* Conecta la tarjeta. Cada página la llama una vez, después de definir su
   `estado` y su `refrescarInterfaz()`. */
function conectarInstructores() {
  if (!document.getElementById('tarjetaInstructores')) return;
  document.addEventListener('input', ev => {
    const clave = ev.target.dataset && ev.target.dataset.inst;
    if (!clave) return;
    estado[clave] = ev.target.value;
    if (typeof alCambiarInstructor === 'function') alCambiarInstructor(clave);
    if (typeof guardarBorrador === 'function') guardarBorrador();
  });
  document.addEventListener('change', async ev => {
    const clave = ev.target.dataset && ev.target.dataset.firmaInst;
    if (!clave) return;
    const archivo = ev.target.files && ev.target.files[0];
    if (!archivo) return;
    marcarError('firma-' + clave, '');
    try {
      estado.firmas[clave] = await procesarImagenFirma(archivo);
      pintarInstructores();
      // La página puede depender de la firma (el Momento 2 avisa si falta):
      // se le da el mismo aviso que cuando cambia un dato escrito.
      if (typeof alCambiarInstructor === 'function') alCambiarInstructor('firma-' + clave);
      if (typeof guardarBorrador === 'function') guardarBorrador();
    } catch (e) { marcarError('firma-' + clave, e.message); }
  });
  document.addEventListener('click', ev => {
    const clave = ev.target.dataset && ev.target.dataset.quitarFirmaInst;
    if (!clave) return;
    delete estado.firmas[clave];
    pintarInstructores();
    if (typeof alCambiarInstructor === 'function') alCambiarInstructor('firma-' + clave);
    if (typeof guardarBorrador === 'function') guardarBorrador();
  });
  pintarInstructores();
}

/* Muestra u oculta el mensaje de error de un campo */
function marcarError(clave, mensaje) {
  const span = document.querySelector(`[data-error="${clave}"]`);
  if (span) span.textContent = mensaje || '';
}

/* Al quitar un integrante del medio, las firmas deben correrse con él para que
   no queden asignadas a otra persona. */
function reindexarFirmasAprendices(indiceEliminado) {
  const totalAnterior = estado.aprendices.length + 1;   // el splice ya ocurrió
  const nuevas = {};
  Object.keys(estado.firmas).forEach(k => {             // instructor y co-formador no se mueven
    if (!k.startsWith('aprendiz')) nuevas[k] = estado.firmas[k];
  });
  let destino = 0;
  for (let i = 0; i < totalAnterior; i++) {
    if (i === indiceEliminado) continue;
    const f = estado.firmas['aprendiz' + i];
    if (f) nuevas['aprendiz' + destino] = f;
    destino++;
  }
  estado.firmas = nuevas;
}

/* Agrega un integrante al equipo. Devuelve su índice, o null si ya son el
   máximo. El equipo es el mismo para los tres formatos: se guarda en el grupo. */
function agregarIntegrante() {
  if (estado.aprendices.length >= CONFIG.maxAprendices) return null;
  estado.aprendices.push(aprendizVacio());
  return estado.aprendices.length - 1;
}

function quitarIntegrante(indice) {
  if (estado.aprendices.length <= 1) return false;
  estado.aprendices.splice(indice, 1);
  reindexarFirmasAprendices(indice);
  return true;
}

function aprendizVacio() {
  return { nombre:'', tipoDocumento:'T.I', numeroDocumento:'', telefono:'',
           correoInstitucional:'', correoPersonal:'', direccion:'', nivelRiesgo:1 };
}

/* Conecta los botones de la tarjeta «Grupo de trabajo». Cada página la llama
   una vez, después de definir su `estado` y su `refrescarInterfaz()`. */
function conectarTarjetaGrupo() {
  const btn = (id, fn) => { const e = document.getElementById(id); if (e) e.addEventListener('click', fn); };
  btn('btnGuardarGrupo', guardarGrupoActual);
  btn('btnEliminarGrupo', eliminarGrupoActual);
  btn('btnExportarGrupo', exportarGrupo);
  btn('btnImportarGrupo', () => document.getElementById('archivoGrupo').click());
  const sel = document.getElementById('grupoGuardado');
  btn('btnNuevoGrupo', nuevoGrupo);
  if (sel) sel.addEventListener('change', () => {
    marcarError('grupoGuardado', '');
    if (!sel.value) { nuevoGrupo(); return; }     // «Grupo nuevo»
    estado.grupoGuardado = sel.value;
    const g = leerGrupos().find(x => x.id === sel.value);
    if (g) { aplicarGrupo(g.datos); avisoGrupo(`✓ Datos de «${g.nombre}» cargados.`); }
    pintarSelectorGrupos();
  });
  const arch = document.getElementById('archivoGrupo');
  if (arch) arch.addEventListener('change', async () => {
    const f = arch.files && arch.files[0];
    arch.value = '';
    if (f) await importarGrupo(f);
  });
  cargarUltimoGrupo();
  pintarSelectorGrupos();
}

/* Al entrar a un formato con el formulario en blanco, se carga solo el último
   grupo guardado: es lo que hace que los tres formatos compartan los datos sin
   que el aprendiz tenga que ir a buscarlos. No se hace si ya escribió algo, ni
   si pulsó «Grupo nuevo» a propósito. */
function cargarUltimoGrupo() {
  if (estado.grupoGuardado || estado.sinGrupo) return;
  if ((estado.aprendices || []).some(a => normalizar(a.nombre))) return;
  const grupos = leerGrupos();
  if (!grupos.length) return;
  const ultimo = grupos.slice().sort(
    (a, b) => String(b.actualizado || '').localeCompare(String(a.actualizado || '')))[0];
  estado.grupoGuardado = ultimo.id;
  aplicarGrupo(ultimo.datos);
  avisoGrupo(`✓ Datos de «${ultimo.nombre}» cargados. Si no es tu grupo, elige otro o pulsa «Grupo nuevo».`);
}

/* Exportación para las pruebas en Node (en el navegador no hace nada) */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG, normalizar, aFechaUTC, fechaTexto, ultimoDiaDelMes,
                     mesDeSecuencia, periodoDeBitacora, esc, CAMPOS_GRUPO,
                     CLAVE_GRUPOS, TIPO_ARCHIVO_GRUPO, VERSION_ARCHIVO_GRUPO, nuevoGrupo,
                     agregarIntegrante, quitarIntegrante, reindexarFirmasAprendices,
                     RE_CORREO, RE_DIGITOS, fechaCortaVisible, INSTRUCTORES,
                     FACTORES_TECNICOS, FACTORES_ACTITUDINALES, valoracionVacia };
}
