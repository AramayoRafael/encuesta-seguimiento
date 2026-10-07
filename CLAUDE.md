# Encuesta de Seguimiento Institucional y Mejora Continua — UNIFRANZ

Proyecto de Rafael Aramayo (UNIFRANZ, Cochabamba). App web para encuestar a estudiantes y backend
en Google Apps Script sobre Google Sheets. Viene de una conversación en Cowork (oct-2026); este
archivo resume todo lo decidido allí. Léelo completo antes de tocar código.

## Cómo trabaja Rafael
- Prefiere ejecución directa, poco ida y vuelta y justificaciones breves.
- Todo en **castellano neutro** (tú), nunca voseo. Textos para estudiantes: simples y amables.
- Aspecto profesional en Sheets: encabezados gris oscuro `#3a3a3a`, sin paletas llamativas.

## Estructura
| Ruta | Qué es |
|---|---|
| `src.html` | **Fuente de la app.** Se edita aquí. Tiene `__LOGO_FULL__` y `__LOGO_WORD__` como marcadores. |
| `build.py` | Genera `index.html` (logos en base64) y valida la sintaxis JS con `node --check`. Correr siempre tras editar `src.html`. |
| `index.html` | Archivo que se publica. **No editar a mano.** |
| `apps-script/Codigo.gs` | Backend completo (web app + menú «Encuesta» en Sheets). |
| `apps-script/appsscript.json` | Manifiesto: zona America/La_Paz, web app «Ejecutar como yo» + acceso anónimo. |
| `apps-script/.clasp.json.example` | Plantilla para `clasp` (copiar a `.clasp.json` con el scriptId real). |
| `tests/e2e_test.py` | Recorrido completo de la app con backend simulado (Playwright). |
| `tests/backend.test.js` | Prueba de humo del script con una hoja simulada en Node. |
| `docs/…xlsx` | Instrumento original con las preguntas textuales. |
| `assets/` | Logos (fondo transparente) usados por `build.py`. |

## Arquitectura
- **Front:** un solo HTML, tema oscuro futurista (campo estelar en canvas, cuadrícula animada, tarjetas de vidrio con borde cónico giratorio, warp entre pantallas, confeti). Respeta `prefers-reduced-motion`. Probado a 390 px y 1366 px.
- **Hosting:** GitHub Pages en `AramayoRafael/encuesta-seguimiento` (público, rama `main`, raíz) → https://aramayorafael.github.io/encuesta-seguimiento/ Vercel quedó como opción futura (enlaces de prueba, proxy para ocultar la URL del script); su plan Hobby es para uso no comercial.
- **Backend:** Apps Script **ligado a la hoja de la base de estudiantes** (Google Sheet «Base de datos Estudiantes 1er semestre P2026», propiedad de la cuenta @unifranz de Rafael, pestaña de RR. HH. `Base_de_estudiantes`, ~23 000 filas, 13 columnas: sede, modalidad, anho_plan_estudio, carrera, semestre_pertenencia, primer_nombre, segundo_nombre, primer_apellido, segundo_apellido, ci, nacionalidad, correo_institucional, celular).
- **URL del script:** se incrusta en `API_URL_DEFECTO` dentro de `src.html`. Ya incrustada (implementación `AKfycbxHT-22…MBT0`, 07-oct-2026). Sin URL la app corre en **modo de prueba** (acepta cualquier @unifranz con datos ficticios). `?api=<url>` en la dirección permite probar otra URL sin guardarla. Triple toque en «v2.0» (abajo a la derecha) o `?admin=1` abre el panel técnico (ping, fila de prueba, reenviar pendientes).

### Lecciones técnicas (vienen de la Ficha 360, no repetir errores)
- GET por **JSONP** (`&callback=`). POST con `fetch` `text/plain` → si falla, `no-cors` + `action=verificar` (3 intentos) → plan C `action=guardar` por JSONP. Si todo falla, la respuesta queda en cola local (`encsi_cola1`) y se reenvía sola al volver a abrir.
- Errores definitivos del servidor (`ya_respondio`, `no_encontrado`, `no_habilitado`, `cerrada`) **no** se encolan.
- Implementación: acceso **«Cualquier persona»** a secas. «Cualquier persona con cuenta de Google» rompe las llamadas anónimas. Como la hoja es de una cuenta Workspace, verificar que el dominio permita esa opción; si no, copiar la base a la cuenta personal de Rafael.
- Para actualizar el script: **Administrar implementaciones → ✏️ → Nueva versión**. Nunca «Nueva implementación» (cambia la URL). Con clasp: `clasp push` + `clasp deploy -i <deploymentId>`.
- No hay comparación de versiones app↔script (en la Ficha 360 causaba bucles de recarga).
- GitHub Pages cachea: repartir el enlace con `?v=N` tras cada publicación.
- Las respuestas guardan el **correo** del estudiante (decisión de Rafael), por eso la bienvenida dice «Confidencial», no «Anónima».
- La base tiene CI y celulares: **nunca** llevar la base al front. El script solo devuelve primer nombre, sede, carrera, semestre, plan y modalidad.

## Flujo de la app (v2.0)
Bienvenida («13 preguntas») → **Correo** (solo se pide `@unifranz.edu.bo`; se normaliza a minúsculas sin espacios) → validación en el servidor → **Módulo 01 · Experiencia del semestre** (9 preguntas) → **Módulo 02 · Recursos didácticos y digitales** (4 preguntas) → Revisión → Envío → código de respuesta.
- Sede, carrera y semestre (preguntas 1–3 del instrumento) **no se muestran**: se toman del padrón y se guardan con la respuesta. (Existió una pantalla de confirmación/corrección; Rafael pidió quitarla. El payload aún trae `datos_corregidos`, siempre `false`.)
- Preguntas y opciones textuales (no «corregir» la redacción del instrumento):
  - 4 Proyectos integradores: **Sí / No**.
  - 5, 6, 10, 11: escala 1–5 «Nada satisfecho … Totalmente satisfecho» (carita animada).
  - 7, 8: escala 1–5 «Nada de acuerdo … Totalmente de acuerdo» (medidor de aguja).
  - 9: **Notas / Aprendizaje**.
  - 12: abierta opcional (1000 caracteres).
  - RRDDTT 1: 10 categorías de recursos tecnológicos, **máximo 3**.
  - RRDDTT 2: escala 1–5, **1 = Nunca, 5 = Siempre** (barras).
  - RRDDTT 3: escala 1–5, **1 = No me ayudan nada, 5 = Me ayudan siempre** (medidor).
  - RRDDTT 4: abierta opcional (500).
- Borrador local `encsi_draft2`; al reabrir ofrece continuar.

## Contrato app ↔ script
- `GET ?action=ping` → `{ok, version_app, abierta, periodo, respuestas}`
- `GET ?action=validar&correo=` → `{ok, habilitado:true, estudiante:{nombre, sede, sede_nombre, carrera_cod, carrera, semestre, plan, modalidad}, carreras:[{cod,nombre}]}` o `{ok, habilitado:false, motivo:'no_encontrado'|'no_habilitado'|'ya_respondio'|'cerrada', fecha?, mensaje?}`
- `GET ?action=verificar&id=` → `{ok, existe}`
- `POST {action:'guardar', data:{id, correo, sede, sede_cod, carrera, carrera_cod, semestre, plan, modalidad, datos_corregidos, p04…p12, r01, r01_idx, r02, r03, r04, dur_s, movil, fecha_local, version}}` → `{ok, id}` o `{ok:false, codigo, fecha?}`. El script **revalida** el correo y toma nombre/semestre/plan del padrón, no del cliente. IDs `ES-…`; los `PRUEBA-…` se saltan la validación y RESUMEN los excluye.

## Script: pestañas y menú «Encuesta»
- **CONFIG** (Parámetro | Valor | Ayuda): Encuesta activa SÍ/NO, fechas de apertura/cierre, Periodo (II-2026), Pestañas de origen, Planes habilitados (**2026**), Semestres, Sedes (LPZ, EAT, CBB, SCZ), Modalidades, Carreras, mensaje de cerrada. Vacío = todos. El filtro visual de Sheets no cuenta; manda CONFIG.
- **CARRERAS**: código → nombre visible. Prellenada con sugerencias; los códigos nuevos que aparezcan en la base se agregan solos en amarillo. Ej.: en el plan 2026, ADM pasó a **NGE (Negocios y Gestión Empresarial)**. Revisar con Rafael IEC/IEF y cualquier código nuevo. La base muestra ARQ en Santa Cruz (plan 2017); Rafael cree que no se oferta en 2026: lo decide el padrón, no la app.
- **PADRON**: generado (correo, nombre, sedes, carrera, semestre, plan, modalidad, origen, **Estado** Respondió/Pendiente con fórmula viva). Si un correo no está en PADRON, el script lo busca en vivo en las bases y lo agrega si cumple los filtros.
- **EXCLUIDOS** (bajas reversibles), **AGREGADOS** (alta manual, ignora filtros), **RESPUESTAS**, **PAPELERA** (respuestas anuladas, restaurables), **RESUMEN** (fórmulas: avance por sede y carrera, promedios, distribuciones), **BITACORA**.
- Menú: Abrir/Cerrar · Actualizar padrón · Ver estado · **Anular la respuesta de un estudiante** (mueve a PAPELERA y le permite responder de nuevo) · Restaurar anulada · Excluir/Reincorporar · Agregados · **Nuevo periodo** (archiva en «RESPUESTAS <periodo>») · Reconstruir RESUMEN · Borrar filas de prueba · Configuración inicial.
- Reglas: la fila 1 de RESPUESTAS es el contrato (no reordenar ni borrar columnas; `hojaResp_` aborta si no coincide). Las columnas de las bases se buscan por nombre (alias en `COLS_BASE`). Escrituras con `LockService`. Textos que empiezan con `= + - @` se guardan con apóstrofo.

## Estado y pendientes (al 07-oct-2026)
- [x] Interfaz v2.0 aprobada por Rafael, salvo detalles que vaya pidiendo.
- [x] Script v2.0 escrito y probado solo con hoja simulada.
- [x] Script implementado; ping y validar responden de forma anónima.
- [x] Repo GitHub creado y Pages activo (07-oct-2026).
- [x] URL incrustada y publicada (enlace `?v=2`).
- [ ] Prueba real: ping, fila de prueba, ingreso con un correo real del padrón, anular y volver a responder, borrar pruebas.
- [ ] Opcional: `clasp` para subir el script sin copiar y pegar; Vercel si se quiere ocultar la URL del script.

## Comandos
```bash
python3 build.py                 # genera index.html
node tests/backend.test.js       # prueba del script con hoja simulada
python3 tests/e2e_test.py        # recorrido completo de la app (Playwright)
# clasp (una vez): npm i -g @google/clasp && clasp login
#   activar https://script.google.com/home/usersettings → API de Google Apps Script
#   cp apps-script/.clasp.json.example apps-script/.clasp.json  (poner scriptId)
# luego: cd apps-script && clasp push && clasp deploy -i <deploymentId> -d "descripción"
```
