/**
 * ENCUESTA DE SEGUIMIENTO INSTITUCIONAL Y MEJORA CONTINUA · ESTUDIANTES  (v2.0)
 * Backend en Google Apps Script, ligado a la hoja que contiene la base de estudiantes.
 *
 * ── PRIMERA VEZ ─────────────────────────────────────────────────────────────
 *   1. En la hoja de la base: Extensiones → Apps Script → pega este código → Guardar.
 *   2. Ejecuta configuracionInicial() (autoriza permisos). Crea las pestañas de trabajo
 *      sin tocar la base de RR. HH.
 *   3. Recarga la hoja: aparece el menú «Encuesta». Revisa CONFIG y CARRERAS y usa
 *      «Encuesta → Actualizar padrón de habilitados».
 *   4. Implementar → Nueva implementación → Aplicación web
 *        Ejecutar como: Yo · Quién tiene acceso: Cualquier persona
 *      Copia la URL que termina en /exec y pégala en API_URL_DEFECTO del index.html.
 *
 * ── PARA ACTUALIZAR ESTE CÓDIGO DESPUÉS ─────────────────────────────────────
 *   Implementar → Administrar implementaciones → ✏️ → Versión: «Nueva versión» → Implementar.
 *   (Nunca «Nueva implementación»: cambia la URL.)
 *
 * Todo lo que cambia de un semestre a otro (planes, semestres, sedes, pestañas de
 * origen, fechas, nombres de carreras, excluidos) se maneja desde la hoja, sin tocar código.
 */

const VERSION_APP = '2.0';
const ZONA = 'America/La_Paz';
const H = {             // nombres de pestañas
  CONFIG: 'CONFIG', CARRERAS: 'CARRERAS', PADRON: 'PADRON', EXCLUIDOS: 'EXCLUIDOS',
  AGREGADOS: 'AGREGADOS', RESP: 'RESPUESTAS', PAPELERA: 'PAPELERA', RESUMEN: 'RESUMEN', BITACORA: 'BITACORA'
};
const SEDES = { LPZ: 'La Paz', EAT: 'El Alto', CBB: 'Cochabamba', SCZ: 'Santa Cruz' };
const CAT_TEC = [
  'Herramientas de creación y diseño', 'Plataformas de gestión y evaluación', 'Recursos audiovisuales y multimedia',
  'Bases de datos y recursos de investigación', 'Simuladores y entornos interactivos', 'Recursos de investigación y análisis',
  'Comunicación y colaboración', 'Herramientas específicas por disciplina o uso', 'Plataformas IA', 'Otros recursos y plataformas varias'
];
// Nombres sugeridos. La pestaña CARRERAS manda: ahí se corrigen y se agregan códigos nuevos.
const CARRERAS_SUGERIDAS = [
  ['NGE', 'Negocios y Gestión Empresarial'], ['AFC', 'Auditoría Financiera y Control de Gestión'], ['DDP', 'Diseño Digital y Producción Transmedia'],
  ['ADM', 'Administración de Empresas'], ['AHT', 'Administración Hotelera y Turismo'],
  ['ARQ', 'Arquitectura'], ['BYF', 'Bioquímica y Farmacia'], ['CPU', 'Contaduría Pública'], ['DER', 'Derecho'],
  ['DGP', 'Diseño Gráfico y Producción Crossmedia'], ['ENF', 'Enfermería'], ['ICO', 'Ingeniería Comercial'],
  ['IEC', 'Ingeniería Económica'], ['IEF', 'Ingeniería Económica y Financiera'], ['MED', 'Medicina'], ['ODO', 'Odontología'],
  ['PSI', 'Psicología'], ['PYM', 'Publicidad y Marketing'], ['SIS', 'Ingeniería de Sistemas'],
  ['CPD', 'Comunicación y Periodismo Digital'], ['PER', 'Periodismo'], ['THO', 'Turismo y Hotelería'], ['GAC', 'Gastronomía y Artes Culinarias']
];
// Columnas que se buscan en las bases de RR. HH. (por nombre, en cualquier orden).
const COLS_BASE = {
  correo: ['correo_institucional', 'correo', 'email', 'correo_electronico'],
  sede: ['sede'], modalidad: ['modalidad'],
  plan: ['anho_plan_estudio', 'anio_plan_estudio', 'plan', 'plan_estudio', 'ano_plan_estudio'],
  carrera: ['carrera', 'cod_carrera', 'codigo_carrera'],
  semestre: ['semestre_pertenencia', 'semestre'],
  n1: ['primer_nombre', 'nombres', 'nombre'], n2: ['segundo_nombre'],
  a1: ['primer_apellido', 'apellidos', 'apellido_paterno'], a2: ['segundo_apellido', 'apellido_materno']
};
const MAIL_RE = /^[a-z0-9._%+-]+@unifranz\.edu\.bo$/;

/* ═════════════════════════════ MENÚ ═════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Encuesta')
    .addItem('▶ Abrir encuesta', 'menuAbrir')
    .addItem('■ Cerrar encuesta', 'menuCerrar')
    .addSeparator()
    .addItem('↻ Actualizar padrón de habilitados', 'menuActualizarPadron')
    .addItem('ⓘ Ver estado de la encuesta', 'menuEstado')
    .addSeparator()
    .addItem('Anular la respuesta de un estudiante…', 'menuAnular')
    .addItem('Restaurar una respuesta anulada…', 'menuRestaurar')
    .addSeparator()
    .addItem('Excluir a un estudiante…', 'menuExcluir')
    .addItem('Reincorporar a un estudiante excluido…', 'menuReincorporar')
    .addItem('Agregar estudiantes a mano (pestaña AGREGADOS)', 'menuIrAgregados')
    .addSeparator()
    .addItem('Iniciar un nuevo periodo…', 'menuNuevoPeriodo')
    .addItem('Reconstruir RESUMEN', 'menuResumen')
    .addItem('Borrar filas de prueba', 'menuBorrarPruebas')
    .addSeparator()
    .addItem('Configuración inicial / reparar pestañas', 'configuracionInicial')
    .addToUi();
}

/* ═════════════════════ CONFIGURACIÓN INICIAL ═════════════════════ */
function configuracionInicial() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(ZONA);
  crearConfig_(); crearCarreras_(); hojaSimple_(H.EXCLUIDOS, ['Correo', 'Motivo', 'Fecha', 'Registrado por']);
  hojaSimple_(H.AGREGADOS, ['correo_institucional', 'sede', 'modalidad', 'anho_plan_estudio', 'carrera', 'semestre_pertenencia', 'primer_nombre', 'segundo_nombre', 'primer_apellido', 'segundo_apellido', 'nota']);
  hojaResp_(); hojaPapelera_(); hojaSimple_(H.BITACORA, ['Fecha', 'Usuario', 'Acción', 'Detalle']);
  const r = generarPadron_();
  CacheService.getScriptCache().remove('carreras_padron');
  crearResumen_();
  onOpen();
  bitacora_('Configuración inicial', r.texto);
  aviso_('Configuración lista', r.texto + '\n\nRevisa la pestaña CONFIG y los nombres en CARRERAS. Las celdas amarillas en CARRERAS son códigos nuevos sin nombre.');
}

function crearConfig_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(H.CONFIG);
  const filas = [
    ['Encuesta activa', 'SÍ', 'SÍ = los estudiantes pueden responder. NO = cerrada.'],
    ['Fecha de apertura', '', 'Opcional. Antes de esta fecha la encuesta aparece cerrada.'],
    ['Fecha de cierre', '', 'Opcional. Después de esta fecha (incluido el día completo) aparece cerrada.'],
    ['Periodo', 'II-2026', 'Nombre del periodo. Se usa al archivar respuestas en «Nuevo periodo».'],
    ['Pestañas de origen', primeraBase_(), 'Pestañas con bases de RR. HH., separadas por punto y coma. AGREGADOS se incluye siempre.'],
    ['Planes habilitados', '2026', 'Ej.: 2026  ·  2017; 2026  ·  vacío = todos'],
    ['Semestres habilitados', '', 'Ej.: 1  ·  1; 2  ·  vacío = todos'],
    ['Sedes habilitadas', '', 'Códigos: LPZ; EAT; CBB; SCZ  ·  vacío = todas'],
    ['Modalidades habilitadas', '', 'Ej.: PRESENCIAL  ·  vacío = todas'],
    ['Carreras habilitadas', '', 'Códigos, ej.: NGE; ICO  ·  vacío = todas'],
    ['Mensaje de encuesta cerrada', 'La encuesta no está disponible en este momento.', 'Texto que ve el estudiante cuando está cerrada.']
  ];
  if (!sh) {
    sh = ss.insertSheet(H.CONFIG, 0);
    sh.getRange(1, 1, 1, 3).setValues([['Parámetro', 'Valor', 'Ayuda']]);
    sh.getRange(2, 1, filas.length, 3).setValues(filas);
  } else {   // agrega solo los parámetros que falten; nunca pisa valores existentes
    const exist = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().map(r => String(r[0]).trim());
    filas.filter(f => exist.indexOf(f[0]) < 0).forEach(f => sh.appendRow(f));
  }
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#3a3a3a').setFontColor('#ffffff');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 520);
  sh.setFrozenRows(1);
  const fila = buscarFila_(sh, 'Encuesta activa');
  if (fila) sh.getRange(fila, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SÍ', 'NO'], true).build());
  ['Fecha de apertura', 'Fecha de cierre'].forEach(k => { const f = buscarFila_(sh, k); if (f) sh.getRange(f, 2).setNumberFormat('dd/MM/yyyy'); });
  sh.getRange(2, 2, sh.getLastRow() - 1, 1).setBackground('#fff8e1');
}
function primeraBase_() {
  const ss = SpreadsheetApp.getActive();
  const propias = Object.keys(H).map(k => H[k]);
  const sh = ss.getSheets().find(s => propias.indexOf(s.getName()) < 0 && s.getName().indexOf('RESPUESTAS') !== 0);
  return sh ? sh.getName() : 'Base_de_estudiantes';
}

function crearCarreras_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(H.CARRERAS);
  if (!sh) {
    sh = ss.insertSheet(H.CARRERAS);
    sh.getRange(1, 1, 1, 2).setValues([['Código', 'Nombre que ve el estudiante']]);
    sh.getRange(2, 1, CARRERAS_SUGERIDAS.length, 2).setValues(CARRERAS_SUGERIDAS);
  }
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#3a3a3a').setFontColor('#ffffff');
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 360); sh.setFrozenRows(1);
}

function hojaSimple_(nombre, cab) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, cab.length).setValues([cab]);
  sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#3a3a3a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

/* ═════════════════════════ LECTURA DE CONFIG ═════════════════════════ */
function config_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.CONFIG);
  const c = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { c[String(r[0]).trim()] = r[1]; });
  const lista = v => String(v || '').split(/[,;\n]/).map(x => norm_(x)).filter(String);
  return {
    activa: /^s/i.test(String(c['Encuesta activa'] || 'SÍ').trim()),
    apertura: c['Fecha de apertura'] instanceof Date ? c['Fecha de apertura'] : null,
    cierre: c['Fecha de cierre'] instanceof Date ? c['Fecha de cierre'] : null,
    periodo: String(c['Periodo'] || '').trim(),
    origenes: String(c['Pestañas de origen'] || 'Base_de_estudiantes').split(/[,;\n]/).map(x => x.trim()).filter(String),
    planes: lista(c['Planes habilitados']), semestres: lista(c['Semestres habilitados']),
    sedes: lista(c['Sedes habilitadas']), modalidades: lista(c['Modalidades habilitadas']),
    carreras: lista(c['Carreras habilitadas']),
    mensaje: String(c['Mensaje de encuesta cerrada'] || 'La encuesta no está disponible en este momento.')
  };
}
function setConfig_(clave, valor) {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.CONFIG);
  const f = buscarFila_(sh, clave);
  if (f) sh.getRange(f, 2).setValue(valor);
}
function abierta_(cfg) {
  if (!cfg.activa) return false;
  const hoy = new Date();
  if (cfg.apertura && hoy < cfg.apertura) return false;
  if (cfg.cierre) { const fin = new Date(cfg.cierre); fin.setHours(23, 59, 59, 999); if (hoy > fin) return false; }
  return true;
}
function nombresCarreras_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.CARRERAS), m = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
    const k = String(r[0]).trim().toUpperCase(); if (k) m[k] = String(r[1] || '').trim() || k;
  });
  return m;
}

/* ═════════════════════════ PADRÓN ═════════════════════════ */
const CAB_PADRON = ['Correo', 'Nombre completo', 'Primer nombre', 'Cód. sede', 'Sede', 'Cód. carrera', 'Carrera', 'Semestre', 'Plan', 'Modalidad', 'Pestaña de origen', 'Estado'];

function leerBase_(sh) {
  const datos = sh.getDataRange().getValues();
  if (datos.length < 2) return { filas: [], error: null };
  const cab = datos[0].map(h => norm_(h).replace(/\s+/g, '_'));
  const idx = {};
  Object.keys(COLS_BASE).forEach(k => { idx[k] = cab.findIndex(h => COLS_BASE[k].indexOf(h) >= 0); });
  const faltan = ['correo', 'carrera'].filter(k => idx[k] < 0);
  if (faltan.length) return { filas: [], error: 'En «' + sh.getName() + '» no encuentro la columna: ' + faltan.join(', ') };
  const val = (r, k) => idx[k] >= 0 ? String(r[idx[k]] === null ? '' : r[idx[k]]).trim() : '';
  return {
    error: null,
    filas: datos.slice(1).map(r => ({
      correo: val(r, 'correo').toLowerCase().replace(/\s+/g, ''),
      sede: val(r, 'sede').toUpperCase(), modalidad: val(r, 'modalidad').toUpperCase(),
      plan: val(r, 'plan'), carrera: val(r, 'carrera').toUpperCase(), semestre: val(r, 'semestre'),
      n1: val(r, 'n1'), n2: val(r, 'n2'), a1: val(r, 'a1'), a2: val(r, 'a2'), origen: sh.getName()
    })).filter(e => e.correo)
  };
}
function cumple_(e, cfg) {
  const ok = (lista, v) => !lista.length || lista.indexOf(norm_(v)) >= 0;
  return ok(cfg.planes, e.plan) && ok(cfg.semestres, e.semestre) && ok(cfg.sedes, e.sede) && ok(cfg.modalidades, e.modalidad) && ok(cfg.carreras, e.carrera);
}
function excluidos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.EXCLUIDOS), s = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { const c = String(r[0]).trim().toLowerCase(); if (c) s[c] = true; });
  return s;
}
function filaPadron_(e, nombres) {
  const titulo = s => String(s || '').toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());
  const nombre = [e.n1, e.n2, e.a1, e.a2].filter(String).map(titulo).join(' ').replace(/\s+/g, ' ');
  return [e.correo, nombre, titulo(String(e.n1).split(/\s+/)[0]), e.sede, SEDES[e.sede] || e.sede, e.carrera, nombres[e.carrera] || e.carrera,
    Number(e.semestre) || e.semestre, Number(e.plan) || e.plan, e.modalidad, e.origen];
}

function generarPadron_() {
  const ss = SpreadsheetApp.getActive(), cfg = config_(), nombres = nombresCarreras_(), exc = excluidos_();
  const errores = [], vistos = {}, salida = [], codigos = {};
  let leidos = 0, fueraFiltro = 0, nExcl = 0, dup = 0;
  cfg.origenes.concat([H.AGREGADOS]).forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) { if (nombre !== H.AGREGADOS) errores.push('No existe la pestaña «' + nombre + '»'); return; }
    const b = leerBase_(sh);
    if (b.error) { errores.push(b.error); return; }
    b.filas.forEach(e => {
      leidos++;
      const manual = nombre === H.AGREGADOS;          // los agregados a mano no pasan por los filtros
      if (!manual && !cumple_(e, cfg)) { fueraFiltro++; return; }
      if (exc[e.correo]) { nExcl++; return; }
      if (vistos[e.correo]) { dup++; return; }
      vistos[e.correo] = true; codigos[e.carrera] = true;
      salida.push(filaPadron_(e, nombres));
    });
  });
  // códigos de carrera nuevos → se agregan a CARRERAS (en amarillo) para ponerles nombre
  const shC = ss.getSheetByName(H.CARRERAS) || (crearCarreras_(), ss.getSheetByName(H.CARRERAS));
  const nuevos = Object.keys(codigos).filter(k => k && !nombres[k]).sort();
  nuevos.forEach(k => { shC.appendRow([k, '']); shC.getRange(shC.getLastRow(), 1, 1, 2).setBackground('#fff59d'); });

  let sh = ss.getSheetByName(H.PADRON);
  if (!sh) sh = ss.insertSheet(H.PADRON);
  sh.clear();
  sepFormulas_(sh);
  sh.getRange(1, 1, 1, CAB_PADRON.length).setValues([CAB_PADRON]).setFontWeight('bold').setBackground('#3a3a3a').setFontColor('#ffffff');
  salida.sort((x, y) => (x[3] + x[6] + x[1]).localeCompare(y[3] + y[6] + y[1], 'es'));
  if (salida.length) { sh.getRange(2, 1, salida.length, salida[0].length).setValues(salida); ponerFormulaEstado_(sh, salida.length); }
  sh.setFrozenRows(1); sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 260);

  const texto = 'Padrón: ' + salida.length + ' estudiantes habilitados.\n' +
    'Filas leídas: ' + leidos + ' · fuera de los filtros: ' + fueraFiltro + ' · excluidos: ' + nExcl + ' · correos repetidos: ' + dup + '.' +
    (nuevos.length ? '\nCódigos de carrera nuevos sin nombre: ' + nuevos.join(', ') + ' (complétalos en CARRERAS).' : '') +
    (errores.length ? '\n\n⚠ ' + errores.join('\n⚠ ') : '');
  return { n: salida.length, texto: texto };
}
// Estado vivo (Respondió / Pendiente). La fórmula cubre exactamente las filas escritas, para no
// inflar getLastRow() con celdas vacías de un ARRAYFORMULA abierto.
function ponerFormulaEstado_(sh, n) {
  const c = colLetra_(cabResp_().indexOf('Correo') + 1), ref = H.RESP + '!$' + c + '$2:$' + c;
  sh.getRange(2, CAB_PADRON.length).setFormula(fx_('=ARRAYFORMULA(IF(COUNTIF(' + ref + ',A2:A' + (n + 1) + ')>0,"Respondió","Pendiente"))'));
}
function formulaEstadoFila_(sh, r) {
  const c = colLetra_(cabResp_().indexOf('Correo') + 1);
  sh.getRange(r, CAB_PADRON.length).setFormula(fx_('=IF(COUNTIF(' + H.RESP + '!$' + c + '$2:$' + c + ',A' + r + ')>0,"Respondió","Pendiente")'));
}

// Busca un correo en el padrón; si no está, lo busca en vivo en las bases (por si la base se actualizó).
function buscarEstudiante_(correo, cfg) {
  const ss = SpreadsheetApp.getActive();
  if (excluidos_()[correo]) return { motivo: 'no_encontrado' };
  const sh = ss.getSheetByName(H.PADRON);
  if (sh && sh.getLastRow() > 1) {
    const f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(correo).matchEntireCell(true).findNext();
    if (f) return { fila: sh.getRange(f.getRow(), 1, 1, CAB_PADRON.length - 1).getValues()[0] };
  }
  let fuera = false;
  const nombres = nombresCarreras_();
  const hojas = cfg.origenes.concat([H.AGREGADOS]);
  for (let i = 0; i < hojas.length; i++) {
    const b = ss.getSheetByName(hojas[i]);
    if (!b) continue;
    const f = b.createTextFinder(correo).matchEntireCell(true).findNext();
    if (!f) continue;
    const base = leerFila_(b, f.getRow());
    if (!base || base.correo !== correo) continue;
    if (hojas[i] !== H.AGREGADOS && !cumple_(base, cfg)) { fuera = true; continue; }
    const fila = filaPadron_(base, nombres);
    if (sh) { sh.appendRow(fila); formulaEstadoFila_(sh, sh.getLastRow()); CacheService.getScriptCache().remove('carreras_padron'); }
    return { fila: fila };
  }
  return { motivo: fuera ? 'no_habilitado' : 'no_encontrado' };
}
function leerFila_(sh, r) {
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => norm_(h).replace(/\s+/g, '_'));
  const fila = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
  const get = k => { const i = cab.findIndex(h => COLS_BASE[k].indexOf(h) >= 0); return i >= 0 ? String(fila[i] === null ? '' : fila[i]).trim() : ''; };
  return { correo: get('correo').toLowerCase().replace(/\s+/g, ''), sede: get('sede').toUpperCase(), modalidad: get('modalidad').toUpperCase(),
    plan: get('plan'), carrera: get('carrera').toUpperCase(), semestre: get('semestre'), n1: get('n1'), n2: get('n2'), a1: get('a1'), a2: get('a2'), origen: sh.getName() };
}
function listaCarrerasPadron_() {
  const cache = CacheService.getScriptCache(), k = 'carreras_padron';
  const c = cache.get(k); if (c) return JSON.parse(c);
  const sh = SpreadsheetApp.getActive().getSheetByName(H.PADRON), nombres = nombresCarreras_(), vistos = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 6, sh.getLastRow() - 1, 1).getValues().forEach(r => { const k2 = String(r[0]).trim(); if (k2) vistos[k2] = true; });
  const out = Object.keys(vistos).map(cod => ({ cod: cod, nombre: nombres[cod] || cod })).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  cache.put(k, JSON.stringify(out), 600);
  return out;
}

/* ═════════════════════════ RESPUESTAS ═════════════════════════ */
function cabResp_() {
  return [
    'Marca temporal', 'ID respuesta', 'Correo', 'Nombre completo', '1. Sede', 'Cód. sede', '2. Carrera', 'Cód. carrera', '3. Semestre', 'Plan', 'Modalidad',
    'Datos corregidos por el estudiante', 'Sede según base', 'Carrera según base',
    '4. ¿Comprendes cómo los Proyectos integradores se relacionan con tu carrera y futuro profesional?',
    '5. Califica del 1 al 5: ¿Qué tan satisfecho te sientes con las experiencias de aprendizaje del semestre?',
    '6. ¿Te sientes satisfecho con el aprendizaje que tienes con tus docentes?',
    '7. ¿Percibes que las actividades experienciales del Proyecto Integrador son planificadas y coordinadas entre los docentes del semestre?',
    '8. ¿Percibes que las codocencias contribuyeron a las experiencias de aprendizaje del Proyecto Integrador?',
    '9. ¿Sientes que tus evaluaciones se están enfocando en la calificación (notas) o en el aprendizaje que desarrollas?',
    '10. ¿Estás satisfecho con la retroalimentación que te brinda tus docentes en las experiencias y en las evaluaciones?',
    '11. ¿Cómo te sientes con el desarrollo del semestre?',
    '12. Observaciones o sugerencias',
    'RRDDTT 1. ¿Qué recursos didácticos y digitales se utilizan efectivamente en tus clases? (máx. 3)'
  ].concat(CAT_TEC.map((c, i) => 'RRDDTT 1.' + (i + 1) + ' ' + c + ' (1 = sí)')).concat([
    'RRDDTT 2. ¿Con qué frecuencia se utilizan los recursos didácticos y digitales en tu clase? (1 = Nunca · 5 = Siempre)',
    'RRDDTT 3. ¿En qué medida los recursos que utilizas te ayudan a desarrollar tus experiencias de aprendizaje en clase? (1 = No me ayudan nada · 5 = Me ayudan siempre)',
    'RRDDTT 4. Tienes sugerencias de algún recursos tecnológicos/didácticos y/o Comentarios',
    'Duración (segundos)', 'Dispositivo', 'Fecha y hora local (dispositivo)', 'Periodo', 'Versión app', 'Datos completos (JSON)'
  ]);
}
function hojaResp_(nombre) {
  const ss = SpreadsheetApp.getActive(), C = cabResp_();
  nombre = nombre || H.RESP;
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, C.length).setValues([C]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#3a3a3a').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1); sh.setFrozenColumns(3); sh.setRowHeight(1, 90);
    sh.setColumnWidths(1, C.length, 140); sh.setColumnWidth(3, 280);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  } else {
    const act = sh.getRange(1, 1, 1, C.length).getValues()[0];
    for (let i = 0; i < C.length; i++) if (String(act[i]).trim() !== C[i])
      throw new Error('La fila 1 de «' + nombre + '» no coincide en la columna ' + (i + 1) + '. Esperado: «' + C[i] + '». No se guardó nada.');
  }
  return sh;
}
function hojaPapelera_() {
  const ss = SpreadsheetApp.getActive(), C = ['Fecha de anulación', 'Anulado por', 'Motivo'].concat(cabResp_());
  let sh = ss.getSheetByName(H.PAPELERA);
  if (!sh) { sh = ss.insertSheet(H.PAPELERA); sh.getRange(1, 1, 1, C.length).setValues([C]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#7a3b3b').setWrap(true); sh.setFrozenRows(1); }
  return sh;
}
function respuestaDe_(correo, sh) {
  sh = sh || SpreadsheetApp.getActive().getSheetByName(H.RESP);
  if (!sh || sh.getLastRow() < 2 || !correo) return null;
  const f = sh.getRange(2, 3, sh.getLastRow() - 1, 1).createTextFinder(correo).matchEntireCell(true).findNext();
  return f ? { fila: f.getRow(), valores: sh.getRange(f.getRow(), 1, 1, sh.getLastColumn()).getValues()[0] } : null;
}
const fmtFecha_ = d => d instanceof Date ? Utilities.formatDate(d, ZONA, 'dd/MM/yyyy') : String(d || '');

/* ═════════════════════════ WEB APP ═════════════════════════ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const accion = p.action || 'ping';
  let out;
  try {
    if (accion === 'ping') {
      const sh = SpreadsheetApp.getActive().getSheetByName(H.RESP), cfg = config_();
      out = { ok: true, servicio: 'encuesta-seguimiento-estudiantes', version_app: VERSION_APP, abierta: abierta_(cfg), periodo: cfg.periodo, respuestas: sh ? Math.max(0, sh.getLastRow() - 1) : 0 };
    } else if (accion === 'validar') out = validar_(p.correo);
    else if (accion === 'verificar') out = { ok: true, id: p.id || '', existe: existeId_(p.id) };
    else if (accion === 'guardar') out = guardar_(JSON.parse(p.data || '{}'));
    else out = { ok: false, error: 'Acción desconocida' };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return responder_(out, p.callback);
}
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = guardar_(body.data || body);
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return responder_(out);
}
function responder_(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][\w$]{0,80}$/.test(cb)) return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function validar_(correoIn) {
  const cfg = config_();
  if (!abierta_(cfg)) return { ok: true, habilitado: false, motivo: 'cerrada', mensaje: cfg.mensaje };
  const correo = String(correoIn || '').trim().toLowerCase();
  if (!MAIL_RE.test(correo)) return { ok: true, habilitado: false, motivo: 'no_encontrado' };
  const r = buscarEstudiante_(correo, cfg);
  if (!r.fila) return { ok: true, habilitado: false, motivo: r.motivo };
  const ya = respuestaDe_(correo);
  if (ya) return { ok: true, habilitado: false, motivo: 'ya_respondio', fecha: fmtFecha_(ya.valores[0]) };
  const f = r.fila, nombres = nombresCarreras_();
  return {
    ok: true, habilitado: true, carreras: listaCarrerasPadron_(),
    estudiante: { nombre: f[2], sede: f[3], sede_nombre: f[4], carrera_cod: f[5], carrera: nombres[f[5]] || f[6], semestre: f[7], plan: f[8], modalidad: f[9] }
  };
}

function guardar_(d) {
  if (!d || typeof d !== 'object') throw new Error('Datos vacíos');
  const id = txt_(d.id, 60);
  if (!/^(ES|PRUEBA)-[A-Z0-9-]{4,40}$/.test(id)) throw new Error('ID inválido');
  const prueba = id.indexOf('PRUEBA-') === 0;
  const correo = String(d.correo || '').trim().toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = hojaResp_();
    if (existeId_(id, sh)) return { ok: true, id: id, duplicado: true };
    let fila = null;
    if (!prueba) {
      const cfg = config_();
      if (!abierta_(cfg)) return { ok: false, codigo: 'cerrada' };
      const r = buscarEstudiante_(correo, cfg);
      if (!r.fila) return { ok: false, codigo: r.motivo };
      const ya = respuestaDe_(correo, sh);
      if (ya) return { ok: false, codigo: 'ya_respondio', fecha: fmtFecha_(ya.valores[0]) };
      fila = r.fila;
    }
    sh.appendRow(filaResp_(d, id, correo, fila));
    CacheService.getScriptCache().remove('resumen');
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function filaResp_(d, id, correo, P) {
  const idx = Array.isArray(d.r01_idx) ? d.r01_idx.map(Number).filter(n => n >= 1 && n <= CAT_TEC.length).slice(0, 3) : [];
  P = P || [correo, 'PRUEBA', '', d.sede_cod, d.sede, d.carrera_cod, d.carrera, d.semestre, d.plan, d.modalidad];
  const corr = !!d.datos_corregidos;
  return [
    new Date(), id, correo, P[1],
    txt_(corr ? d.sede : P[4], 40), txt_(corr ? d.sede_cod : P[3], 10), txt_(corr ? d.carrera : P[6], 80), txt_(corr ? d.carrera_cod : P[5], 10),
    P[7], P[8], P[9], corr ? 'Sí' : 'No', corr ? P[4] : '', corr ? P[6] : '',
    siNo_(d.p04), esc_(d.p05), esc_(d.p06), esc_(d.p07), esc_(d.p08), txt_(d.p09, 20), esc_(d.p10), esc_(d.p11), txt_(d.p12, 1000),
    idx.map(n => CAT_TEC[n - 1]).join(' | ')
  ].concat(CAT_TEC.map((_, i) => idx.indexOf(i + 1) >= 0 ? 1 : 0)).concat([
    esc_(d.r02), esc_(d.r03), txt_(d.r04, 500),
    Number(d.dur_s) || '', d.movil ? 'Móvil' : 'Computadora', txt_(d.fecha_local, 40), config_().periodo, txt_(d.version, 10),
    JSON.stringify(d).slice(0, 45000)
  ]);
}
function txt_(v, max) { let s = (v === null || v === undefined) ? '' : String(v).trim().slice(0, max || 500); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function esc_(v) { const n = Number(v); return (n >= 1 && n <= 5 && Math.round(n) === n) ? n : ''; }
function siNo_(v) { const s = String(v || '').trim(); return /^s/i.test(s) ? 'Sí' : /^n/i.test(s) ? 'No' : ''; }
function existeId_(id, sh) {
  if (!id) return false;
  sh = sh || SpreadsheetApp.getActive().getSheetByName(H.RESP);
  if (!sh || sh.getLastRow() < 2) return false;
  return !!sh.getRange(2, 2, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
}

/* ═════════════════════════ ACCIONES DEL MENÚ ═════════════════════════ */
function menuAbrir() { setConfig_('Encuesta activa', 'SÍ'); bitacora_('Abrir encuesta', ''); aviso_('Encuesta abierta', 'Los estudiantes ya pueden responder.' + avisoFechas_()); }
function menuCerrar() { setConfig_('Encuesta activa', 'NO'); bitacora_('Cerrar encuesta', ''); aviso_('Encuesta cerrada', 'Nadie puede responder hasta que la abras de nuevo.'); }
function avisoFechas_() { const c = config_(); return (c.apertura || c.cierre) ? '\n(Recuerda que también rigen las fechas de CONFIG.)' : ''; }

function menuActualizarPadron() {
  const r = generarPadron_();
  CacheService.getScriptCache().remove('carreras_padron');
  bitacora_('Actualizar padrón', r.texto);
  aviso_('Padrón actualizado', r.texto);
}
function menuEstado() {
  const ss = SpreadsheetApp.getActive(), cfg = config_();
  const pad = ss.getSheetByName(H.PADRON), resp = ss.getSheetByName(H.RESP);
  const nPad = pad ? Math.max(0, pad.getLastRow() - 1) : 0;
  let nResp = 0;
  if (resp && resp.getLastRow() > 1) nResp = resp.getRange(2, 2, resp.getLastRow() - 1, 1).getValues().filter(r => String(r[0]).indexOf('PRUEBA-') !== 0).length;
  aviso_('Estado de la encuesta',
    'Estado: ' + (abierta_(cfg) ? 'ABIERTA' : 'CERRADA') + '\nPeriodo: ' + cfg.periodo +
    '\nFiltros → planes: ' + (cfg.planes.join(', ') || 'todos') + ' · semestres: ' + (cfg.semestres.join(', ') || 'todos') +
    ' · sedes: ' + (cfg.sedes.join(', ').toUpperCase() || 'todas') + ' · modalidades: ' + (cfg.modalidades.join(', ').toUpperCase() || 'todas') +
    '\nPestañas de origen: ' + cfg.origenes.join(', ') +
    '\n\nHabilitados en el padrón: ' + nPad + '\nRespuestas recibidas: ' + nResp + (nPad ? ' (' + Math.round(nResp / nPad * 100) + '%)' : ''));
}

// Anular = mover la respuesta a PAPELERA (nada se pierde) para que el estudiante pueda volver a responder.
function menuAnular() {
  const ui = SpreadsheetApp.getUi();
  const correo = pedir_('Anular respuesta', 'Correo institucional del estudiante cuya respuesta quieres anular:');
  if (!correo) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(H.RESP);
  const r = respuestaDe_(correo, sh);
  if (!r) return aviso_('Sin respuesta', 'No hay una respuesta registrada con el correo:\n' + correo);
  const v = r.valores;
  const motivo = pedir_('Motivo', 'Respuesta encontrada:\n• ' + v[3] + '\n• ' + v[6] + ' · ' + v[4] + '\n• Enviada el ' + fmtFecha_(v[0]) + '\n• ID ' + v[1] + '\n\nEscribe el motivo de la anulación (queda en la bitácora):', true);
  if (motivo === null) return;
  const ok = ui.alert('Confirmar anulación', '¿Anular la respuesta de ' + correo + '?\n\nLa fila se moverá a PAPELERA (se puede restaurar) y el estudiante podrá responder de nuevo.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const actual = respuestaDe_(correo, sh);          // se vuelve a buscar dentro del candado
    if (!actual) return aviso_('Sin cambios', 'La respuesta ya no estaba en RESPUESTAS.');
    hojaPapelera_().appendRow([new Date(), usuario_(), motivo || '(sin motivo)'].concat(actual.valores));
    sh.deleteRow(actual.fila);
  } finally { lock.releaseLock(); }
  bitacora_('Anular respuesta', correo + ' · ID ' + v[1] + ' · motivo: ' + (motivo || '—'));
  aviso_('Respuesta anulada', 'Listo. ' + correo + ' ya puede volver a responder la encuesta.\nLa respuesta anterior quedó en PAPELERA.');
}
function menuRestaurar() {
  const ui = SpreadsheetApp.getUi();
  let id = pedir_('Restaurar respuesta', 'ID de la respuesta a restaurar (columna «ID respuesta» de PAPELERA, ej.: ES-ABC12-XYZ987):');
  if (!id) return;
  id = id.toUpperCase();
  const pap = SpreadsheetApp.getActive().getSheetByName(H.PAPELERA);
  if (!pap || pap.getLastRow() < 2) return aviso_('Papelera vacía', 'No hay respuestas anuladas.');
  const f = pap.getRange(2, 5, pap.getLastRow() - 1, 1).createTextFinder(id.toUpperCase()).matchEntireCell(true).findNext();
  if (!f) return aviso_('No encontrada', 'No hay ninguna respuesta con el ID ' + id + ' en PAPELERA.');
  const fila = pap.getRange(f.getRow(), 1, 1, pap.getLastColumn()).getValues()[0], valores = fila.slice(3, 3 + cabResp_().length), correo = valores[2];
  const resp = hojaResp_();
  if (respuestaDe_(correo, resp)) return aviso_('No se puede restaurar', correo + ' ya tiene una respuesta nueva en RESPUESTAS. Anúlala primero si quieres recuperar la anterior.');
  if (ui.alert('Confirmar', '¿Restaurar la respuesta ' + id + ' de ' + correo + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  resp.appendRow(valores); pap.deleteRow(f.getRow());
  bitacora_('Restaurar respuesta', correo + ' · ID ' + id);
  aviso_('Restaurada', 'La respuesta volvió a RESPUESTAS.');
}

function menuExcluir() {
  const correo = pedir_('Excluir estudiante', 'Correo institucional del estudiante que NO debe responder:');
  if (!correo) return;
  const motivo = pedir_('Motivo', 'Motivo de la exclusión (ej.: retiro, cambio de carrera):', true);
  if (motivo === null) return;
  const ss = SpreadsheetApp.getActive(), ex = ss.getSheetByName(H.EXCLUIDOS);
  if (excluidos_()[correo]) return aviso_('Ya estaba excluido', correo);
  ex.appendRow([correo, motivo, new Date(), usuario_()]);
  const pad = ss.getSheetByName(H.PADRON);
  if (pad && pad.getLastRow() > 1) { const f = pad.getRange(2, 1, pad.getLastRow() - 1, 1).createTextFinder(correo).matchEntireCell(true).findNext(); if (f) pad.deleteRow(f.getRow()); }
  bitacora_('Excluir estudiante', correo + ' · ' + motivo);
  const ya = respuestaDe_(correo);
  aviso_('Estudiante excluido', correo + ' ya no podrá ingresar.' + (ya ? '\n\nOjo: ya tenía una respuesta registrada. Si no debe contar, anúlala con «Anular la respuesta de un estudiante».' : ''));
}
function menuReincorporar() {
  const correo = pedir_('Reincorporar estudiante', 'Correo institucional a quitar de EXCLUIDOS:');
  if (!correo) return;
  const ex = SpreadsheetApp.getActive().getSheetByName(H.EXCLUIDOS);
  if (!ex || ex.getLastRow() < 2) return aviso_('Sin excluidos', 'La lista está vacía.');
  const vals = ex.getRange(2, 1, ex.getLastRow() - 1, 1).getValues();
  let n = 0;
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim().toLowerCase() === correo) { ex.deleteRow(i + 2); n++; }
  if (!n) return aviso_('No estaba excluido', correo);
  bitacora_('Reincorporar estudiante', correo);
  aviso_('Reincorporado', correo + ' vuelve a estar habilitado si cumple los filtros de CONFIG (o si está en AGREGADOS). Se sumará al padrón cuando ingrese o al actualizar el padrón.');
}
function menuIrAgregados() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.AGREGADOS) || hojaSimple_(H.AGREGADOS, ['correo_institucional', 'sede', 'modalidad', 'anho_plan_estudio', 'carrera', 'semestre_pertenencia', 'primer_nombre', 'segundo_nombre', 'primer_apellido', 'segundo_apellido', 'nota']);
  SpreadsheetApp.getActive().setActiveSheet(sh);
  aviso_('Agregar estudiantes a mano', 'Escribe una fila por estudiante (correo, sede LPZ/EAT/CBB/SCZ, código de carrera, etc.).\nQuienes estén aquí quedan habilitados aunque no cumplan los filtros de CONFIG.\nAl terminar usa «Actualizar padrón de habilitados».');
}

function menuNuevoPeriodo() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive(), cfg = config_();
  const nuevo = pedir_('Nuevo periodo', 'Periodo actual: ' + cfg.periodo + '\n\nLas respuestas actuales se guardarán en la pestaña «RESPUESTAS ' + cfg.periodo + '» y RESPUESTAS quedará vacía.\n\nEscribe el nombre del NUEVO periodo (ej.: I-2027):', true);
  if (!nuevo) return;
  const archivo = 'RESPUESTAS ' + (cfg.periodo || Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'));
  if (ss.getSheetByName(archivo)) return aviso_('Ya existe', 'Ya hay una pestaña «' + archivo + '». Cambia el nombre del periodo actual en CONFIG y vuelve a intentar.');
  if (ui.alert('Confirmar', '¿Archivar ' + cfg.periodo + ' e iniciar ' + nuevo + '?\n\nDespués revisa en CONFIG los planes, semestres y pestañas de origen, y actualiza el padrón.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const resp = ss.getSheetByName(H.RESP);
    resp.copyTo(ss).setName(archivo);
    if (resp.getLastRow() > 1) resp.deleteRows(2, resp.getLastRow() - 1);
    setConfig_('Periodo', nuevo);
  } finally { lock.releaseLock(); }
  crearResumen_();
  bitacora_('Nuevo periodo', cfg.periodo + ' → ' + nuevo + ' (archivo: ' + archivo + ')');
  aviso_('Nuevo periodo iniciado', 'Respuestas de ' + cfg.periodo + ' archivadas en «' + archivo + '».\nAhora: revisa CONFIG → Actualizar padrón → Abrir encuesta.');
}
function menuResumen() { crearResumen_(); aviso_('RESUMEN', 'Pestaña RESUMEN reconstruida.'); }
function menuBorrarPruebas() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.RESP);
  if (!sh || sh.getLastRow() < 2) return aviso_('Sin filas', 'No hay respuestas.');
  const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  let n = 0;
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]).indexOf('PRUEBA-') === 0) { sh.deleteRow(i + 2); n++; }
  bitacora_('Borrar filas de prueba', n + ' filas');
  aviso_('Filas de prueba', 'Borradas: ' + n);
}

/* ═════════════════════════ RESUMEN (fórmulas vivas) ═════════════════════════ */
function crearResumen_() {
  const ss = SpreadsheetApp.getActive();
  let rs = ss.getSheetByName(H.RESUMEN);
  if (!rs) rs = ss.insertSheet(H.RESUMEN);
  rs.clear();
  sepFormulas_(rs);
  const C = cabResp_();
  const col = t => { const i = C.findIndex(h => h.indexOf(t) === 0); if (i < 0) throw new Error('Columna no encontrada: ' + t); return colLetra_(i + 1); };
  const R = c => H.RESP + '!$' + c + '$2:$' + c;
  const ID = R(col('ID respuesta')), SEDE = R(col('Cód. sede')), CARR = R(col('Cód. carrera'));
  const P = c => H.PADRON + '!$' + c + '$2:$' + c;
  const real = '"<>PRUEBA-*"', W = 7, rows = [], fmt = [];
  const push = (r, t) => { while (r.length < W) r.push(''); rows.push(r); if (t) fmt.push([rows.length, t]); };
  const sedes = Object.keys(SEDES);

  push(['RESUMEN · Encuesta de Seguimiento Institucional y Mejora Continua (Estudiantes)'], 'titulo');
  push(['Se actualiza solo. Excluye filas de prueba (ID que empieza con PRUEBA-).'], 'nota');
  push(['']);
  push(['Total de respuestas', '=COUNTIFS(' + ID + ',"<>",' + ID + ',' + real + ')'], 'total');
  push(['Habilitados en el padrón', '=COUNTA(' + P('A') + ')'], 'total');
  push(['Cobertura', '=IFERROR(B4/B5,0)'], 'totalpct');
  push(['']);
  push(['AVANCE POR SEDE', 'Habilitados', 'Respondieron', 'Cobertura'], 'cab');
  sedes.forEach(s => { const r = rows.length + 1; push([SEDES[s], '=COUNTIF(' + P('D') + ',"' + s + '")', '=COUNTIFS(' + P('D') + ',"' + s + '",' + P('L') + ',"Respondió")', '=IFERROR(C' + r + '/B' + r + ',0)'], 'pct4'); });
  push(['']);
  const ESC = [['5.', 'Satisfacción con las experiencias de aprendizaje'], ['6.', 'Satisfacción con el aprendizaje con docentes'], ['7.', 'Actividades planificadas y coordinadas (acuerdo)'], ['8.', 'Aporte de las codocencias (acuerdo)'], ['10.', 'Satisfacción con la retroalimentación'], ['11.', 'Cómo se siente con el desarrollo del semestre'], ['RRDDTT 2.', 'Frecuencia de uso de recursos (1 Nunca – 5 Siempre)'], ['RRDDTT 3.', 'Los recursos ayudan al aprendizaje (1 – 5)']];
  push(['PROMEDIOS (escala 1 a 5)', 'General'].concat(sedes.map(s => SEDES[s])), 'cab');
  ESC.forEach(q => {
    const c = R(col(q[0] + ' '));
    push([q[0] + ' ' + q[1], '=IFERROR(AVERAGEIFS(' + c + ',' + ID + ',' + real + '),"–")'].concat(sedes.map(s => '=IFERROR(AVERAGEIFS(' + c + ',' + SEDE + ',"' + s + '",' + ID + ',' + real + '),"–")')), 'prom');
  });
  push(['']);
  const dist = (tit, ct, ops) => {
    push([tit, 'N', '%'], 'cab');
    const c = R(col(ct));
    ops.forEach(o => { push([o, '=COUNTIFS(' + c + ',"' + o + '",' + ID + ',' + real + ')', '=IFERROR(B' + (rows.length + 1) + '/$B$4,0)'], 'pct'); });
    push(['']);
  };
  dist('4. PROYECTOS INTEGRADORES Y CARRERA', '4. ', ['Sí', 'No']);
  dist('9. ENFOQUE DE LAS EVALUACIONES', '9. ', ['Notas', 'Aprendizaje']);
  push(['RRDDTT 1. RECURSOS MÁS USADOS (máx. 3 por estudiante)', 'N', '% de estudiantes'], 'cab');
  CAT_TEC.forEach((c, i) => { const cc = R(col('RRDDTT 1.' + (i + 1) + ' ')); push([(i + 1) + '. ' + c, '=SUMIFS(' + cc + ',' + ID + ',' + real + ')', '=IFERROR(B' + (rows.length + 1) + '/$B$4,0)'], 'pct'); });
  push(['']);
  push(['POR CARRERA', 'Habilitados', 'Respondieron', 'Cobertura', 'Prom. 5', 'Prom. 11', 'Prom. RRDDTT 3'], 'cab');
  const shC = ss.getSheetByName(H.CARRERAS);
  const cods = shC && shC.getLastRow() > 1 ? shC.getRange(2, 1, shC.getLastRow() - 1, 2).getValues().filter(r => String(r[0]).trim()) : [];
  cods.forEach(rw => {
    const k = String(rw[0]).trim().toUpperCase(), r = rows.length + 1;
    push([k + ' · ' + (rw[1] || k), '=COUNTIF(' + P('F') + ',"' + k + '")', '=COUNTIFS(' + CARR + ',"' + k + '",' + ID + ',' + real + ')', '=IFERROR(C' + r + '/B' + r + ',"–")',
      '=IFERROR(AVERAGEIFS(' + R(col('5. ')) + ',' + CARR + ',"' + k + '",' + ID + ',' + real + '),"–")',
      '=IFERROR(AVERAGEIFS(' + R(col('11. ')) + ',' + CARR + ',"' + k + '",' + ID + ',' + real + '),"–")',
      '=IFERROR(AVERAGEIFS(' + R(col('RRDDTT 3.')) + ',' + CARR + ',"' + k + '",' + ID + ',' + real + '),"–")'], 'carr');
  });

  rs.getRange(1, 1, rows.length, W).setValues(rows.map(r => r.map(v => typeof v === 'string' && v.charAt(0) === '=' ? fx_(v) : v)));
  rs.setColumnWidth(1, 430); rs.setColumnWidths(2, W - 1, 115);
  rs.getRange(1, 1, rows.length, W).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  fmt.forEach(f => {
    const n = f[0], t = f[1];
    if (t === 'titulo') rs.getRange(n, 1).setFontSize(14).setFontWeight('bold');
    if (t === 'nota') rs.getRange(n, 1).setFontColor('#777777').setFontStyle('italic');
    if (t === 'total') rs.getRange(n, 1, 1, 2).setFontWeight('bold').setFontSize(12);
    if (t === 'totalpct') rs.getRange(n, 1, 1, 2).setFontWeight('bold').setFontSize(12).setNumberFormat('0.0%');
    if (t === 'cab') rs.getRange(n, 1, 1, W).setFontWeight('bold').setBackground('#3a3a3a').setFontColor('#ffffff');
    if (t === 'pct') rs.getRange(n, 3).setNumberFormat('0.0%');
    if (t === 'pct4') rs.getRange(n, 4).setNumberFormat('0.0%');
    if (t === 'prom') rs.getRange(n, 2, 1, W - 1).setNumberFormat('0.00');
    if (t === 'carr') { rs.getRange(n, 4).setNumberFormat('0.0%'); rs.getRange(n, 5, 1, 3).setNumberFormat('0.00'); }
  });
}

/* ═════════════════════════ UTILIDADES ═════════════════════════ */
function norm_(s) { return String(s === null || s === undefined ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
// Separador de argumentos de las fórmulas: depende de la configuración regional de la hoja
// («,» en inglés; «;» en español y otras con coma decimal). Se prueba escribiendo =SUM(2,5)
// en la celda A1 de una pestaña recién vaciada y se guarda en las propiedades del script.
function sepFormulas_(sh) {
  try {
    const c = sh.getRange(1, 1);
    for (const s of [',', ';']) {
      c.setFormula('=SUM(2' + s + '5)'); SpreadsheetApp.flush();
      if (c.getValue() === 7) { c.clearContent(); PropertiesService.getScriptProperties().setProperty('SEP', s); return s; }
    }
    c.clearContent();
  } catch (e) {}
  return sep_();
}
function sep_() { try { return PropertiesService.getScriptProperties().getProperty('SEP') || ','; } catch (e) { return ','; } }
// Cambia las comas separadoras por «;» si hace falta (las comas dentro de "texto" no se tocan).
function fx_(f) { return sep_() === ',' ? f : f.replace(/("[^"]*")|,/g, (m, q) => q || ';'); }
function colLetra_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function buscarFila_(sh, clave) {
  if (!sh || sh.getLastRow() < 1) return 0;
  const v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === clave) return i + 1;
  return 0;
}
function usuario_() { try { return Session.getActiveUser().getEmail() || 'administrador'; } catch (e) { return 'administrador'; } }
function bitacora_(accion, detalle) {
  try { const sh = SpreadsheetApp.getActive().getSheetByName(H.BITACORA) || hojaSimple_(H.BITACORA, ['Fecha', 'Usuario', 'Acción', 'Detalle']); sh.appendRow([new Date(), usuario_(), accion, String(detalle || '').slice(0, 2000)]); } catch (e) {}
}
function aviso_(titulo, texto) { try { SpreadsheetApp.getUi().alert(titulo, texto, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(titulo + '\n' + texto); } }
function pedir_(titulo, texto, crudo) {
  const ui = SpreadsheetApp.getUi(), r = ui.prompt(titulo, texto, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const t = r.getResponseText().trim();
  return crudo ? t : t.toLowerCase();
}
