/* =========================================================================
   MOTOR DE FORMATOS WORD (.docx)
   =========================================================================
   Un .docx es un archivo .zip que por dentro tiene, entre otras cosas,
   word/document.xml con todo el texto. Este motor abre la plantilla oficial,
   reemplaza el texto de párrafos concretos y vuelve a empaquetarla. Igual que
   con el Excel: se EDITA la plantilla, nunca se construye un documento nuevo,
   así los logos, las tablas, los encabezados y los textos legales quedan
   exactamente como los publicó el SENA.

   ¿Cómo se identifica un párrafo? Por su atributo `w14:paraId`, un
   identificador único que Word le asigna a cada párrafo y que no cambia al
   escribir dentro de él. Es el equivalente a la referencia de celda del Excel:
   estable, verificable y legible. Los mapas de cada formato se obtuvieron
   comparando la plantilla vacía con un documento ya diligenciado por los
   aprendices, así que no hay adivinanza.

   Depende de fflate (fflate.min.js) para descomprimir y comprimir el zip.
   ========================================================================= */

/* ==== INICIO NUCLEO DOCX ==== */

/* Escapa lo que va dentro de una etiqueta XML */
function escaparXml(texto) {
  return String(texto === null || texto === undefined ? '' : texto)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/* Encuentra los límites del párrafo con el paraId indicado.
   Los párrafos pueden anidarse (cuadros de texto), así que no sirve una
   expresión regular perezosa: hay que contar apertura y cierre. */
function limitesParrafo(xml, paraId) {
  const marca = `w14:paraId="${paraId}"`;
  const posMarca = xml.indexOf(marca);
  if (posMarca < 0) return null;
  const inicio = xml.lastIndexOf('<w:p ', posMarca);
  if (inicio < 0) return null;

  // ¿Es un párrafo vacío autocerrado?  <w:p ... />
  const finEtiqueta = xml.indexOf('>', posMarca);
  if (xml[finEtiqueta - 1] === '/') return { inicio, fin: finEtiqueta + 1, autocerrado: true };

  let profundidad = 1, i = finEtiqueta + 1;
  while (profundidad > 0 && i < xml.length) {
    const abre = xml.indexOf('<w:p ', i);
    const cierra = xml.indexOf('</w:p>', i);
    if (cierra < 0) return null;
    if (abre >= 0 && abre < cierra) {
      const f = xml.indexOf('>', abre);
      if (xml[f - 1] !== '/') profundidad++;      // los autocerrados no abren nivel
      i = f + 1;
    } else {
      profundidad--;
      i = cierra + 6;
    }
  }
  return { inicio, fin: i, autocerrado: false };
}

/* Los textos de ejemplo de la plantilla («<Nombre del proyecto>», «<Ciudad y
   fecha>»…) están en ROJO a propósito: son marcas para quien diligencia. Si se
   hereda ese color, el dato real sale rojo en el documento final. Lo que la app
   escribe es información, no una marca, así que va en negro. Solo se cambia el
   color: negrita, tamaño, fuente y alineación se conservan. */
function aNegro(rPr) {
  if (!rPr) return rPr;
  if (/<w:color\b/.test(rPr)) return rPr.replace(/<w:color\b[^>]*\/>/g, '<w:color w:val="000000"/>');
  return rPr;   // sin color explícito: hereda el del estilo, que ya es negro
}

/* Propiedades de fuente del párrafo. Word guarda en <w:pPr><w:rPr> el formato
   de la marca de párrafo: es justo el que aplicaría si alguien escribiera ahí
   a mano, así que el texto insertado sale con la fuente correcta del formato. */
function propiedadesDeFuente(parrafoXml) {
  const pPr = parrafoXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  if (pPr) {
    const rPr = pPr[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    if (rPr) return aNegro(rPr[0]);
  }
  const primerRun = parrafoXml.match(/<w:r\b[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/);
  return primerRun ? aNegro(primerRun[1]) : '';
}

/* -------------------------------------------------------------------------
   FORMATO FORZADO DE UNA CELDA
   Un valor del mapa puede ser texto suelto o { texto, alinear, negrita,
   tamano }. Hace falta porque las plantillas no siempre son coherentes consigo
   mismas: en la tabla de valoración del Momento 2, por ejemplo, unas casillas
   vienen centradas y en negrita y otras alineadas a la izquierda y más
   pequeñas. Respetando cada celda tal cual, la columna de «S»/«M» sale
   desordenada. Estas opciones son la excepción, no la regla: sin ellas el
   motor conserva el formato del formato oficial, que es lo correcto.
   ------------------------------------------------------------------------- */
function opcionesDeValor(valor) {
  return (valor && typeof valor === 'object' && !Array.isArray(valor))
    ? valor : { texto: valor };
}

/* El orden dentro de <w:pPr> importa: <w:jc> va después del espaciado y antes
   de <w:rPr>. Word tolera el desorden; LibreOffice no siempre. */
function conAlineacion(pPr, alinear) {
  if (!alinear) return pPr;
  const jc = `<w:jc w:val="${alinear}"/>`;
  if (!pPr) return '<w:pPr>' + jc + '</w:pPr>';
  if (/<w:jc\b/.test(pPr)) return pPr.replace(/<w:jc\b[^>]*\/>/, jc);
  if (pPr.includes('<w:rPr>')) return pPr.replace('<w:rPr>', jc + '<w:rPr>');
  return pPr.replace('</w:pPr>', jc + '</w:pPr>');
}

/* Igual con <w:rPr>: <w:b/> va justo después de <w:rFonts>, y <w:sz> al final. */
function conNegrita(rPr) {
  if (!rPr) return '<w:rPr><w:b/><w:bCs/></w:rPr>';
  if (/<w:b\/>/.test(rPr)) return rPr;
  const fuentes = rPr.match(/<w:rFonts\b[^>]*\/>/);
  return fuentes ? rPr.replace(fuentes[0], fuentes[0] + '<w:b/><w:bCs/>')
                 : rPr.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>');
}
function conTamano(rPr, medios) {
  if (!medios) return rPr;
  const sz = `<w:sz w:val="${medios}"/><w:szCs w:val="${medios}"/>`;
  if (!rPr) return `<w:rPr>${sz}</w:rPr>`;
  if (/<w:sz\b/.test(rPr)) {
    return rPr.replace(/<w:sz\b[^>]*\/>/, `<w:sz w:val="${medios}"/>`)
              .replace(/<w:szCs\b[^>]*\/>/, `<w:szCs w:val="${medios}"/>`);
  }
  return rPr.replace('</w:rPr>', sz + '</w:rPr>');
}
function conFormato(rPr, op) {
  let r = rPr;
  if (op.negrita) r = conNegrita(r);
  if (op.tamano) r = conTamano(r, op.tamano);
  return r;
}

/* Convierte un texto (que puede traer saltos de línea) en runs de Word */
function runsDeTexto(texto, rPr) {
  const lineas = String(texto).split(/\r?\n/);
  return lineas.map((linea, i) => {
    const salto = i > 0 ? '<w:br/>' : '';
    return `<w:r>${rPr}${salto}<w:t xml:space="preserve">${escaparXml(linea)}</w:t></w:r>`;
  }).join('');
}

/* Reemplaza el contenido de un párrafo conservando sus propiedades.
   `valor` es el texto, o { texto, alinear, negrita, tamano } cuando hay que
   forzar el formato de la celda (ver opcionesDeValor). */
function escribirParrafo(xml, paraId, valor) {
  const lim = limitesParrafo(xml, paraId);
  if (!lim) return { xml, ok: false };
  const op = opcionesDeValor(valor);
  const original = xml.slice(lim.inicio, lim.fin);
  const rPr = conFormato(propiedadesDeFuente(original), op);

  let cabecera, pPr = '';
  if (lim.autocerrado) {
    cabecera = original.slice(0, -2) + '>';                 // <w:p .../>  →  <w:p ...>
  } else {
    cabecera = original.slice(0, original.indexOf('>') + 1);
    const m = original.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    // También la marca de párrafo va a negro: si no, lo que alguien escriba
    // después en Word saldría en el rojo del texto de ejemplo.
    if (m) pPr = aNegro(m[0]);
  }
  pPr = conAlineacion(pPr, op.alinear);
  const nuevo = cabecera + pPr + runsDeTexto(op.texto, rPr) + '</w:p>';
  return { xml: xml.slice(0, lim.inicio) + nuevo + xml.slice(lim.fin), ok: true };
}

/* -------------------------------------------------------------------------
   INSERCIÓN DE FIRMAS (imágenes)
   Meter una imagen en un .docx son tres cosas a la vez:
     1. el archivo en word/media/
     2. una relación en word/_rels/document.xml.rels que le dé un rId
     3. un <w:drawing> dentro de un párrafo, apuntando a ese rId
   Si falta cualquiera de las tres, Word abre el documento pero muestra el
   recuadro rojo de «imagen no encontrada».
   ------------------------------------------------------------------------- */
const EMU_POR_PIXEL_DOCX = 9525;
const ANCHO_MAX_FIRMA_EMU = 1900000;   // ~5,2 cm: el ancho de la celda de firma
const ALTO_MAX_FIRMA_EMU  = 550000;    // ~1,5 cm, como las firmas que ya trae el formato

function base64ABytes(base64) {
  const limpio = String(base64).replace(/^data:[^;]+;base64,/i, '');
  if (typeof atob === 'function') {
    const binario = atob(limpio);
    const bytes = new Uint8Array(binario.length);
    for (let i = 0; i < binario.length; i++) bytes[i] = binario.charCodeAt(i);
    return bytes;
  }
  return new Uint8Array(Buffer.from(limpio, 'base64'));   // Node (pruebas)
}

function extensionDeImagen(datos) {
  const m = String(datos).match(/^data:image\/([a-z+]+);base64,/i);
  return m && /jpe?g/i.test(m[1]) ? 'jpeg' : 'png';
}

/* Declara el tipo de archivo en [Content_Types].xml si aún no está */
function asegurarTipoContenido(entradas, extension) {
  const clave = '[Content_Types].xml';
  const dec = new TextDecoder('utf-8'), cod = new TextEncoder();
  let xml = dec.decode(entradas[clave]);
  if (new RegExp(`Extension="${extension}"`, 'i').test(xml)) return;
  const tipo = extension === 'jpeg' ? 'image/jpeg' : 'image/png';
  xml = xml.replace('<Types ', '<Types ')
           .replace(/(<Types[^>]*>)/, `$1<Default Extension="${extension}" ContentType="${tipo}"/>`);
  entradas[clave] = cod.encode(xml);
}

/* Agrega la relación y devuelve el rId nuevo */
function agregarRelacionImagen(entradas, nombreArchivo) {
  const clave = 'word/_rels/document.xml.rels';
  const dec = new TextDecoder('utf-8'), cod = new TextEncoder();
  let xml = dec.decode(entradas[clave]);
  let n = 1000;
  while (xml.includes(`Id="rId${n}"`)) n++;              // un rId que no exista
  const rId = 'rId' + n;
  xml = xml.replace('</Relationships>',
    `<Relationship Id="${rId}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
    `Target="media/${nombreArchivo}"/></Relationships>`);
  entradas[clave] = cod.encode(xml);
  return rId;
}

function xmlDeImagen(rId, cx, cy, id, nombre) {
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${nombre}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>' +
    '</wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${nombre}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

/* Pone la imagen dentro del párrafo indicado, conservando su alineación */
function escribirImagen(entradas, xml, paraId, firma, indice) {
  const lim = limitesParrafo(xml, paraId);
  if (!lim) return { xml, ok: false };

  const extension = extensionDeImagen(firma.datos);
  const nombreArchivo = `firma${indice}.${extension}`;
  entradas['word/media/' + nombreArchivo] = base64ABytes(firma.datos);
  asegurarTipoContenido(entradas, extension);
  const rId = agregarRelacionImagen(entradas, nombreArchivo);

  const anchoPx = Number(firma.ancho) || 200, altoPx = Number(firma.alto) || 60;
  const maxAncho = Number(firma.maxAnchoEmu) || ANCHO_MAX_FIRMA_EMU;
  const maxAlto = Number(firma.maxAltoEmu) || ALTO_MAX_FIRMA_EMU;
  const escala = Math.min(maxAncho / (anchoPx * EMU_POR_PIXEL_DOCX),
                          maxAlto / (altoPx * EMU_POR_PIXEL_DOCX));
  const cx = Math.round(anchoPx * EMU_POR_PIXEL_DOCX * escala);
  const cy = Math.round(altoPx * EMU_POR_PIXEL_DOCX * escala);

  const original = xml.slice(lim.inicio, lim.fin);
  let cabecera, pPr = '';
  if (lim.autocerrado) {
    cabecera = original.slice(0, -2) + '>';
  } else {
    cabecera = original.slice(0, original.indexOf('>') + 1);
    const m = original.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    if (m) pPr = m[0];
  }
  const nuevo = cabecera + pPr + xmlDeImagen(rId, cx, cy, 900 + indice, 'Firma') + '</w:p>';
  return { xml: xml.slice(0, lim.inicio) + nuevo + xml.slice(lim.fin), ok: true };
}

/* -------------------------------------------------------------------------
   LISTAS CON VIÑETA (a., b., c. …)
   El formato trae un número fijo de renglones —el acta, por ejemplo, trae dos
   para los aprendices— pero el equipo puede ser más grande. Meter a todos en
   el mismo renglón deja «a.» con tres nombres y «b.» vacío, que es justo lo
   que se veía. Aquí se llena un renglón por persona y, si faltan, se clona el
   último: como el párrafo lleva su `<w:numPr>`, Word le pone la letra que
   sigue sin que haya que calcular nada.
   ------------------------------------------------------------------------- */

/* Copia un párrafo (con su viñeta y su formato) y le pone otro texto.
   Se le quitan paraId y textId: son identificadores únicos y Word los vuelve
   a generar; dejarlos repetidos ensucia el documento. */
function clonarParrafo(modeloXml, texto) {
  const cabeceraFin = modeloXml.indexOf('>') + 1;
  let cabecera = modeloXml.slice(0, cabeceraFin)
    .replace(/\s+w14:paraId="[^"]*"/g, '')
    .replace(/\s+w14:textId="[^"]*"/g, '');
  if (cabecera.endsWith('/>')) cabecera = cabecera.slice(0, -2) + '>';
  const pPr = aNegro((modeloXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0]);
  const rPr = propiedadesDeFuente(modeloXml);
  return cabecera + pPr + runsDeTexto(texto, rPr) + '</w:p>';
}

/* Límites del párrafo anterior al que empieza en `inicio`, o null.
   Sirve para detectar el renglón en blanco que el formato deja entre viñeta y
   viñeta: si se clona solo la viñeta, la lista queda con espacios desiguales. */
function limitesAnterior(xml, inicio) {
  const fin = xml.lastIndexOf('</w:p>', inicio);
  if (fin < 0) return null;
  const ini = xml.lastIndexOf('<w:p ', fin);
  if (ini < 0) return null;
  const bloque = xml.slice(ini, fin + 6);
  if (bloque.indexOf('<w:p ', 1) > 0) return null;      // hay anidamiento: mejor no tocar
  return { ini, fin: fin + 6, bloque };
}
function parrafoAnterior(xml, inicio) {
  const lim = limitesAnterior(xml, inicio);
  return lim ? lim.bloque : null;
}

/* Quita del documento un renglón de la lista, junto con el párrafo en blanco
   que lo separa del anterior. Se usa con los renglones que el formato trae de
   más: dejarlos vacíos deja una viñeta «b.» sin nada al lado, que en un
   documento oficial se lee como un error. */
function borrarParrafo(xml, paraId) {
  const lim = limitesParrafo(xml, paraId);
  if (!lim) return { xml, ok: false };
  const ant = limitesAnterior(xml, lim.inicio);
  const desde = (ant && esSeparador(ant.bloque)) ? ant.ini : lim.inicio;
  return { xml: xml.slice(0, desde) + xml.slice(lim.fin), ok: true };
}

function esSeparador(parrafoXml) {
  return !!parrafoXml && !/<w:t[ >]/.test(parrafoXml) && !parrafoXml.includes('<w:numPr>');
}

/* Escribe `textos` en los renglones `paraIds`, clonando el último si hacen
   falta más, y vaciando los que sobren. */
function escribirLista(xml, paraIds, textos) {
  const faltantes = [];
  const usados = Math.min(paraIds.length, textos.length);
  for (let i = 0; i < usados; i++) {
    const r = escribirParrafo(xml, paraIds[i], textos[i]);
    if (!r.ok) faltantes.push(paraIds[i]);
    xml = r.xml;
  }
  // Los renglones que sobran se quitan, de atrás hacia adelante para que los
  // que faltan por borrar no se muevan de sitio.
  for (let i = paraIds.length - 1; i >= usados; i--) {
    const r = borrarParrafo(xml, paraIds[i]);
    if (!r.ok) faltantes.push(paraIds[i]);
    xml = r.xml;
  }
  if (textos.length > paraIds.length) {
    const ultimo = paraIds[paraIds.length - 1];
    const lim = limitesParrafo(xml, ultimo);
    if (!lim) return { xml, faltantes: faltantes.concat(ultimo) };
    const modelo = xml.slice(lim.inicio, lim.fin);
    const anterior = parrafoAnterior(xml, lim.inicio);
    const separador = esSeparador(anterior) ? clonarParrafo(anterior, '') : '';
    let extras = '';
    for (let i = paraIds.length; i < textos.length; i++) {
      extras += separador + clonarParrafo(modelo, textos[i]);
    }
    xml = xml.slice(0, lim.fin) + extras + xml.slice(lim.fin);
  }
  return { xml, faltantes };
}

/* -------------------------------------------------------------------------
   BLOQUEO DE SOLO LECTURA
   Word guarda la restricción de edición en word/settings.xml, en la etiqueta
   <w:documentProtection>. Con w:edit="readOnly" y w:enforcement="1" el
   documento se abre bloqueado: se ve y se imprime, pero la cinta de edición
   queda deshabilitada. Si además lleva contraseña, Word pide esa contraseña
   para quitar la restricción.

   IMPORTANTE, y hay que decirlo claro: esto es un SELLO, no cifrado. El .docx
   sigue siendo un .zip y quien sepa puede borrar esta línea. Sirve para que
   nadie cambie una calificación por descuido o por las buenas; no resiste a
   quien quiera romperlo a propósito. El documento realmente cerrado es un PDF.

   La contraseña no se guarda en ninguna parte: solo viaja su huella, y de la
   huella no se puede volver a la contraseña.
   ------------------------------------------------------------------------- */

/* Vueltas de la función de resumen. Word escribe 100000; 50000 tarda la mitad
   en el navegador y para un sello es de sobra. Word usa el número que venga
   escrito en el archivo, así que cualquiera de los dos le sirve. */
const VUELTAS_PROTECCION = 50000;

function aBase64(bytes) {
  if (typeof btoa === 'function') {
    let s = '';
    bytes.forEach(b => { s += String.fromCharCode(b); });
    return btoa(s);
  }
  return Buffer.from(bytes).toString('base64');   // Node (pruebas)
}

/* La contraseña, en UTF-16 little endian: es como la codifica Word. */
function utf16le(texto) {
  const bytes = new Uint8Array(texto.length * 2);
  for (let i = 0; i < texto.length; i++) {
    const c = texto.charCodeAt(i);
    bytes[2 * i] = c & 0xFF;
    bytes[2 * i + 1] = c >>> 8;
  }
  return bytes;
}

function concatenar(a, b) {
  const r = new Uint8Array(a.length + b.length);
  r.set(a); r.set(b, a.length);
  return r;
}

function enteroLE(n) {
  return new Uint8Array([n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]);
}

/* Huella de la contraseña, según el algoritmo de Word (MS-OFFCRYPTO 2.3.7.1):
     H0 = SHA512(sal + contraseña)
     Hn = SHA512(Hn-1 + número de vuelta, 4 bytes little endian)
   Es asíncrona porque usa la criptografía del navegador; se calcula UNA vez
   por tanda, no una por documento. */
async function huellaDeContrasena(contrasena, cripto, vueltas) {
  const sub = (cripto || (typeof crypto !== 'undefined' ? crypto : null));
  if (!sub || !sub.subtle) throw new Error('Este navegador no permite calcular la contraseña.');
  const giros = vueltas || VUELTAS_PROTECCION;
  const sal = new Uint8Array(16);
  sub.getRandomValues(sal);
  let h = new Uint8Array(await sub.subtle.digest('SHA-512', concatenar(sal, utf16le(contrasena))));
  for (let i = 0; i < giros; i++) {
    h = new Uint8Array(await sub.subtle.digest('SHA-512', concatenar(h, enteroLE(i))));
  }
  return { hash: aBase64(h), sal: aBase64(sal), vueltas: giros };
}

/* La etiqueta que entiende Word. Sin huella, el bloqueo no lleva contraseña:
   se quita desde «Revisar → Restringir edición», que ya es un aviso claro. */
function xmlDeProteccion(huella) {
  let etiqueta = '<w:documentProtection w:edit="readOnly" w:enforcement="1"';
  if (huella && huella.hash) {
    etiqueta += ' w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash"' +
                ' w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14"' +   // 14 = SHA-512
                ` w:cryptSpinCount="${huella.vueltas}"` +
                ` w:hash="${escaparXml(huella.hash)}" w:salt="${escaparXml(huella.sal)}"`;
  }
  return etiqueta + '/>';
}

/* Inserta la etiqueta en word/settings.xml. El orden dentro de <w:settings>
   está fijado por el esquema: <w:documentProtection> va justo antes de
   <w:defaultTabStop>. Si el archivo ya trae una protección, se reemplaza. */
function protegerSoloLectura(entradas, huella) {
  const clave = 'word/settings.xml';
  if (!entradas[clave]) return false;
  const dec = new TextDecoder('utf-8'), cod = new TextEncoder();
  let xml = dec.decode(entradas[clave]);
  const etiqueta = xmlDeProteccion(huella);
  if (/<w:documentProtection\b[^>]*\/>/.test(xml)) {
    xml = xml.replace(/<w:documentProtection\b[^>]*\/>/, etiqueta);
  } else if (xml.includes('<w:defaultTabStop')) {
    xml = xml.replace('<w:defaultTabStop', etiqueta + '<w:defaultTabStop');
  } else {
    xml = xml.replace(/(<w:settings[^>]*>)/, '$1' + etiqueta);
  }
  entradas[clave] = cod.encode(xml);
  return true;
}

/* Aplica un mapa { paraId: texto } sobre el document.xml.
   Devuelve el XML nuevo y la lista de paraId que no se encontraron, para que
   la aplicación pueda avisar en vez de entregar un documento incompleto. */
function aplicarMapaDocx(xml, valores) {
  let resultado = xml;
  const faltantes = [];
  Object.keys(valores).forEach(paraId => {
    const valor = valores[paraId];
    if (valor === undefined || valor === null) return;
    const r = escribirParrafo(resultado, paraId, valor);
    if (!r.ok) faltantes.push(paraId);
    resultado = r.xml;
  });
  return { xml: resultado, faltantes };
}

/* Abre la plantilla, aplica el mapa y devuelve el .docx listo para descargar.
   `fflateRef` es el objeto fflate (unzipSync / zipSync).
   `bufferPlantilla` es un Uint8Array con el .docx original.
   `opciones.proteccion` bloquea el documento en solo lectura: `true` sin
   contraseña, o la huella que devuelve huellaDeContrasena() para pedirla. */
function generarDocx(fflateRef, bufferPlantilla, valores, imagenes, listas, opciones) {
  const entradas = fflateRef.unzipSync(new Uint8Array(bufferPlantilla));
  if (!entradas['word/document.xml']) throw new Error('El archivo no parece un documento de Word válido.');

  const decodificador = new TextDecoder('utf-8');
  const codificador = new TextEncoder();
  const xml = decodificador.decode(entradas['word/document.xml']);
  let { xml: nuevoXml, faltantes } = aplicarMapaDocx(xml, valores);

  // Listas con viñeta: [{ paraIds: [...], textos: [...] }]
  (listas || []).forEach(lista => {
    const r = escribirLista(nuevoXml, lista.paraIds, lista.textos);
    faltantes = faltantes.concat(r.faltantes);
    nuevoXml = r.xml;
  });

  // Firmas: cada una es { paraId: {datos, ancho, alto} }
  let i = 0;
  Object.keys(imagenes || {}).forEach(paraId => {
    const firma = imagenes[paraId];
    if (!firma || !firma.datos) return;
    const r = escribirImagen(entradas, nuevoXml, paraId, firma, ++i);
    if (!r.ok) faltantes.push(paraId);
    nuevoXml = r.xml;
  });

  entradas['word/document.xml'] = codificador.encode(nuevoXml);

  const proteccion = (opciones || {}).proteccion;
  if (proteccion && !protegerSoloLectura(entradas, proteccion === true ? null : proteccion)) {
    throw new Error('La plantilla no trae word/settings.xml: no se puede bloquear.');
  }

  // level 6: mismo nivel de compresión que usa Word, tamaño y velocidad razonables
  const salida = fflateRef.zipSync(entradas, { level: 6 });
  return { datos: salida, faltantes };
}

/* ==== FIN NUCLEO DOCX ==== */

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { escaparXml, limitesParrafo, propiedadesDeFuente, runsDeTexto,
                     escribirParrafo, aplicarMapaDocx, generarDocx,
                     base64ABytes, extensionDeImagen, escribirImagen,
                     clonarParrafo, escribirLista, parrafoAnterior, limitesAnterior,
                     borrarParrafo, esSeparador, aNegro,
                     opcionesDeValor, conAlineacion, conNegrita, conTamano,
                     huellaDeContrasena, xmlDeProteccion, protegerSoloLectura,
                     VUELTAS_PROTECCION, utf16le };
}
