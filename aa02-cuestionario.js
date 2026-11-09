/*
 * AuditX — AA-02 Cuestionario de continuidad del cliente (vista del papel)
 *
 * Qué hace:
 *  - Dibuja el procedimiento previo (4 criterios Sí/No que indican si hay que reevaluar al cliente)
 *    y los 24 requerimientos de la hoja «AA-02 Cuestionario Contin.»
 *    (Respuesta Sí/No/N/A, Riesgo, Ref. de P/T, Observaciones).
 *  - Calcula la clasificación de riesgo del cliente con la misma regla de la hoja.
 *  - Verifica pendientes, bloquea y firma el papel (la firma se invalida si el contenido cambia).
 *  - Guarda dentro del libro (parte XML personalizada) y, fuera de Excel, en localStorage.
 *
 * Punto de integración con el menú de Pre-Planificación (taskpane.js):
 *  - Escucha  «auditx:navigate»  con detail.id === "AA-02" para mostrarse.
 *  - Emite    «auditx:estado-papel»  con el estado, el avance y el riesgo (alimenta el índice vivo, M2-02).
 *  - Emite    «auditx:volver»  al pulsar la flecha de regreso.
 *
 * Códigos del documento maestro (los mismos de AA-01): M3-01 (cuestionario digital con conclusión
 * y firma), M2-04 (firma) y M1-02 (escala de riesgo configurable).
 */

/* global Office, Excel, document, window, console */

(function () {
  "use strict";

  /* =====================================================================
   * 1. CONFIGURACIÓN (lo que Auditool puede ajustar sin tocar la lógica)
   * ===================================================================== */
  var CONFIG = {
    ID_PAPEL: "AA-02",
    VERSION_ESQUEMA: 1,
    NAMESPACE_XML: "http://auditool.org/auditx/aa02/v1",
    CLAVE_LOCAL: "auditx.aa02.v1",
    AUTOGUARDADO_MS: 600,

    // Escala de riesgo de la firma, de menor a mayor (hoy: hoja de control, celdas L351:L355).
    // El puntaje de cada nivel es su posición: 1 = primero … 5 = último. Se personaliza en M1-02.
    NIVELES_RIESGO: ["MÍNIMO", "BAJO", "MEDIO", "ALTO", "CRÍTICO"],

    TIPOS_RIESGO: ["Negocio", "Fraude", "Procesos"],
    FRECUENCIAS: ["Anual", "Mensual", "Trimestral", "Diario", "Por evento"],
    MAX_RIESGOS_EEFF: 5,
    MAX_REGISTRO: 12,

    // Reglas de completitud. Los «bloqueos» impiden firmar; los «avisos» solo advierten.
    EXIGIR_RIESGO_EN_TODAS: false,   // false: solo se exige nivel de riesgo cuando la respuesta es «No»
    EXIGIR_JUSTIFICACION: false,     // true: «No» y «N/A» deben llevar observación para poder firmar

    // Ganchos de integración. Reemplácelos cuando existan (SSO de Microsoft 365, creación de encargo).
    obtenerUsuario: function () { return Promise.resolve(null); },   // → { nombre, correo }
    obtenerEncargo: function () { return Promise.resolve({}); }      // → { entidad, nit, desde, hasta, ... }
  };

  /* =====================================================================
   * 2. DATOS DEL CUESTIONARIO (extraídos de la hoja original)
   *    fila  = fila de la hoja «AA-02 Cuestionario Contin.» (sirve para migrar datos antiguos)
   *    resp  = se contesta Sí / No / N/A
   *    riesgo= lleva nivel de riesgo y entra en la clasificación del cliente
   *    nota  = texto de apoyo que la hoja trae en la misma celda o en la fila de abajo
   * ===================================================================== */
  var SECCIONES = [
    // La hoja no tiene secciones: son 24 requerimientos seguidos. Se dibujan sin acordeón.
    { n: 1, titulo: "Requerimientos de continuidad", reqs: [
      {id: "1", fila: 26, num: 1, texto: "¿La firma cuenta con los recursos, tiempo y competencia suficiente para aceptar este encargo?", resp: true, riesgo: true},
      {id: "2", fila: 27, num: 2, texto: "¿El equipo de auditoría cuenta con la competencia y experiencia adecuadas para ejecutar el encargo conforme a los requerimientos profesionales?", resp: true, riesgo: true},
      {id: "3", fila: 28, num: 3, texto: "¿Existe algún asunto de independencia o conflicto de intereses si acepta este encargo de auditoría?", resp: true, riesgo: true},
      {id: "4", fila: 29, num: 4, texto: "¿El auditor y/o los colaboradores son independientes y están libres de conflictos?", nota: "Considere: Relaciones personales/familiares con personas influyentes en la actividad, clientes claves, o inclusive, servicios distintos a la auditoría, como contabilidad, tesorería, intereses financieros y otras relaciones comerciales con el cliente.", resp: true, riesgo: true},
      {id: "5", fila: 30, num: 5, texto: "¿El auditor dispone de expertos y profesionales en temas específicos que son necesarios para el desarrollo del encargo de auditoría?", resp: true, riesgo: true},
      {id: "6", fila: 31, num: 6, texto: "¿Al aceptar continuar con el encargo de auditoría, el auditor cumple con las políticas de independencia y los valores éticos?", resp: true, riesgo: true},
      {id: "7", fila: 32, num: 7, texto: "Si se han identificado conflictos de interés, falta de independencia u otras amenazas, ¿Se han tomado las medidas apropiadas para eliminar, reducir o controlar a un nivel aceptable dichas amenazas?", resp: true, riesgo: true},
      {id: "8", fila: 33, num: 8, texto: "¿Hay alguna razón (o evento reciente) que pone en duda la integridad de los principales propietarios, la alta dirección y los encargados del gobierno de la entidad?", nota: "Considere las operaciones de la entidad, incluyendo las prácticas comerciales, la reputación de la empresa, y la historia de las infracciones éticas o reglamentarias.", resp: true, riesgo: true},
      {id: "9", fila: 35, num: 9, texto: "¿Existen riesgos relevantes relacionados con cumplimiento legal, incluyendo prevención de lavado de activos u otros actos ilícitos?", resp: true, riesgo: true},
      {id: "10", fila: 36, num: 10, texto: "¿Se han seguido los requisitos de continuación según el manual de gestión de calidad de la firma y políticas correspondientes?", resp: true, riesgo: true},
      {id: "11", fila: 37, num: 11, texto: "¿Hay algún problema identificado en auditorías anteriores y otras contrataciones con el cliente que deban abordarse y tenerse en cuenta para aceptar este encargo?", resp: true, riesgo: true},
      {id: "12", fila: 38, num: 12, texto: "¿Existen nuevas circunstancias que aumenten el riesgo de compromiso y continuidad?", resp: true, riesgo: true},
      {id: "13", fila: 39, num: 13, texto: "¿Hay nuevas circunstancias que pueden afectar la independencia como auditores de la entidad, con base a las políticas de la Firma o los requerimientos de independencia profesionales o reguladores?", resp: true, riesgo: true},
      {id: "14", fila: 40, num: 14, texto: "¿Ha realizado el cliente recientemente adquisiciones significativas y/o numerosas?", resp: true, riesgo: true},
      {id: "15", fila: 41, num: 15, texto: "¿Hay alguna indicación de que en los últimos tres años la entidad del cliente haya pedido cambios del Líder o del Gerente del encargo?", resp: true, riesgo: true},
      {id: "16", fila: 42, num: 16, texto: "¿Ha existido algún desacuerdo con la entidad o se ha comunicado algún asunto significativo a su regulador?", resp: true, riesgo: true},
      {id: "17", fila: 43, num: 17, texto: "¿Ha sido el director ejecutivo o el controller (o equivalente) sustituido en el último año?", resp: true, riesgo: true},
      {id: "18", fila: 44, num: 18, texto: "¿Existió algún problema relacionado con el pago de honorarios de auditoría?", resp: true, riesgo: true},
      {id: "19", fila: 45, num: 19, texto: "¿Hay algún asunto relacionado con alguna imposición de limitaciones al alcance?", resp: true, riesgo: true},
      {id: "20", fila: 46, num: 20, texto: "¿Hay algún asunto relacionado con asignar de forma adecuada personal de auditoría?", resp: true, riesgo: true},
      {id: "21", fila: 47, num: 21, texto: "El auditor y/o su equipo del encargo posee:", nota: "• Conocimiento de la industria.\n• Experiencia con los requerimientos reglamentarios.\n• Capacidad de adquirir habilidades y conocimientos necesarios.", resp: true, riesgo: true},
      {id: "22", fila: 48, num: 22, texto: "¿Hay personal calificado para llevar a cabo la revisión de gestión de calidad?", resp: true, riesgo: true},
      {id: "23", fila: 49, num: 23, texto: "¿Se han revisado los factores de riesgo que influyen en el negocio del cliente?", resp: true, riesgo: true},
      {id: "24", fila: 50, num: 24, texto: "¿Se ha evaluado los hallazgos identificados en el último/más reciente trabajo realizado (cierre anterior)?", resp: true, riesgo: true},
    ] },
  ];

  // Procedimiento previo (hoja, filas 17 a 20). Si alguno es «Sí» se recomienda reevaluar al cliente.
  var CRITERIOS_PREVIOS = [
    "Solicitud por parte del cliente existente para efectuar un trabajo de auditoría",
    "Existencia de nuevos requisitos legales, regulatorios o profesionales que afecten los procedimientos de auditoría",
    "Cambios significativos en la gerencia u otros miembros clave del personal",
    "Cambios relevantes en la naturaleza, tamaño o estructura operativa del cliente"
  ];

  /* =====================================================================
   * 3. UTILIDADES
   * ===================================================================== */
  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

  function getPath(obj, ruta) {
    var partes = ruta.split("."), cur = obj;
    for (var i = 0; i < partes.length; i++) { if (cur == null) return undefined; cur = cur[partes[i]]; }
    return cur;
  }
  function setPath(obj, ruta, valor) {
    var partes = ruta.split("."), cur = obj;
    for (var i = 0; i < partes.length - 1; i++) {
      if (cur[partes[i]] == null) cur[partes[i]] = /^\d+$/.test(partes[i + 1]) ? [] : {};
      cur = cur[partes[i]];
    }
    cur[partes[partes.length - 1]] = valor;
  }
  function hoy() {
    var d = new Date();
    return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
  }
  function ahoraISO() { return new Date().toISOString(); }
  function horaCorta() { var d = new Date(); return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2); }
  function copiar(o) { return JSON.parse(JSON.stringify(o)); }
  function escXml(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function debounce(fn, ms) { var t; return function () { clearTimeout(t); t = setTimeout(fn, ms); }; }

  /* =====================================================================
   * 4. ESTADO
   * ===================================================================== */
  function nuevoEstado() {
    return {
      esquema: CONFIG.VERSION_ESQUEMA,
      meta: {
        entidad: "", nit: "", desde: "", hasta: "",
        preparadoPor: "", revisadoPor: "", supervisadoPor: "", aprobadoPor: ""
      },
      previo: { c1: "", c2: "", c3: "", c4: "" },   // criterios del procedimiento previo (SI / NO)
      r: {},                                   // respuestas por id de requerimiento
      eeff: { riesgos: [""], estrategia: "", enfoque: "" },
      registro: [],                            // riesgos identificados y controles
      conclusion: { comentarios: "", decision: "" },
      firma: null,                             // { por, fecha, hash, invalida? }
      historial: [],
      actualizado: null
    };
  }

  var estado = nuevoEstado();
  var itemsPorId = {};          // id → definición del requerimiento
  var itemsRespondibles = [];   // lista plana de los que se contestan
  var seccionDeItem = {};       // id → número de sección
  var raiz = null;
  var usuarioActual = null;

  function indexarDatos() {
    itemsPorId = {}; itemsRespondibles = []; seccionDeItem = {};
    SECCIONES.forEach(function (sec) {
      sec.reqs.forEach(function (rq) {
        var todos = [rq].concat(rq.hijos || []);
        todos.forEach(function (it) {
          itemsPorId[it.id] = it; seccionDeItem[it.id] = sec.n;
          if (it.resp) itemsRespondibles.push(it);
        });
      });
    });
  }

  function respuesta(id) {
    if (!estado.r[id]) estado.r[id] = { r: "", riesgo: "", ref: "", com: "" };
    return estado.r[id];
  }

  /* =====================================================================
   * 5. CÁLCULOS (misma regla de la hoja original)
   * ===================================================================== */

  // Clasificación del cliente. En la hoja: puntaje = suma de los niveles elegidos (1…5) y el cliente
  // queda en el nivel cuyo tope (N × nivel) alcanza esa suma, es decir, el promedio redondeado hacia
  // arriba. Las preguntas sin nivel no cuentan (en la hoja quedaban como celdas vacías).
  function calcularRiesgo() {
    var niveles = CONFIG.NIVELES_RIESGO, conteo = niveles.map(function () { return 0; }), suma = 0, n = 0;
    itemsRespondibles.forEach(function (it) {
      if (!it.riesgo) return;
      var idx = niveles.indexOf(respuesta(it.id).riesgo);
      if (idx < 0) return;
      conteo[idx]++; suma += idx + 1; n++;
    });
    if (!n) return { n: 0, suma: 0, promedio: 0, nivel: 0, etiqueta: "", conteo: conteo };
    var promedio = suma / n;
    var nivel = Math.min(niveles.length, Math.max(1, Math.ceil(promedio - 1e-9)));
    return { n: n, suma: suma, promedio: promedio, nivel: nivel, etiqueta: niveles[nivel - 1], conteo: conteo };
  }

  // Procedimiento previo: la hoja cuenta los criterios marcados «SI» y, si hay al menos uno,
  // recomienda reevaluar al cliente existente y completar el checklist como anexo del memorando.
  function calcularPrevio() {
    var p = estado.previo, claves = ["c1", "c2", "c3", "c4"];
    var si = claves.filter(function (k) { return p[k] === "SI"; }).length;
    var sin = claves.filter(function (k) { return !p[k]; }).length;
    return { si: si, sinResponder: sin, reevaluar: si >= 1 };
  }

  function calcularAvance() {
    var total = itemsRespondibles.length, hechas = 0, porSeccion = {};
    SECCIONES.forEach(function (s) { porSeccion[s.n] = { total: 0, hechas: 0 }; });
    itemsRespondibles.forEach(function (it) {
      var s = porSeccion[seccionDeItem[it.id]], ok = !!respuesta(it.id).r;
      s.total++; if (ok) { s.hechas++; hechas++; }
    });
    return { total: total, hechas: hechas, porSeccion: porSeccion };
  }

  // Devuelve { bloqueos: [...], avisos: [...] }. Cada pendiente indica dónde está para poder ir a él.
  function evaluar() {
    var bloqueos = [], avisos = [], m = estado.meta, c = estado.conclusion;
    function b(grupo, texto, destino) { bloqueos.push({ grupo: grupo, texto: texto, destino: destino }); }

    if (!m.entidad.trim()) b("Datos del encargo", "Falta el nombre de la entidad.", { campo: "meta.entidad" });
    CRITERIOS_PREVIOS.forEach(function (t, i) {
      if (!estado.previo["c" + (i + 1)]) b("Procedimiento previo", "Falta responder el criterio " + (i + 1) + ": " + resumen(t, 60), { campo: "previo.c" + (i + 1) });
    });

    var sinResponder = 0, noSinRiesgo = 0, sinRiesgoSi = 0, sinJustificar = 0;
    itemsRespondibles.forEach(function (it) {
      var r = respuesta(it.id), etiqueta = (it.etq ? it.etq + " " : "") + resumen(it.texto, 70), sec = seccionDeItem[it.id];
      var grupo = "Sección " + sec + ". " + tituloSeccion(sec);
      if (!r.r) { b(grupo, "Sin responder: " + (it.num ? it.num + " · " : "") + etiqueta, { id: it.id }); sinResponder++; return; }
      if (it.riesgo && !r.riesgo) {
        if (r.r === "NO" || CONFIG.EXIGIR_RIESGO_EN_TODAS) {
          b(grupo, "Falta el nivel de riesgo: " + (it.num ? it.num + " · " : "") + etiqueta, { id: it.id, campo: "riesgo" }); noSinRiesgo++;
        } else if (r.r === "SI") { sinRiesgoSi++; }
      }
      if ((r.r === "NO" || r.r === "NA") && !r.com.trim()) {
        if (CONFIG.EXIGIR_JUSTIFICACION) b(grupo, "Falta justificar la respuesta " + (r.r === "NO" ? "«No»" : "«N/A»") + " en observaciones: " + etiqueta, { id: it.id, campo: "com" });
        sinJustificar++;
      }
    });
    if (sinRiesgoSi) avisos.push(sinRiesgoSi + " respuestas «Sí» no tienen nivel de riesgo; no entran en la clasificación del cliente.");
    if (sinJustificar && !CONFIG.EXIGIR_JUSTIFICACION) avisos.push(sinJustificar + " respuestas «No» o «N/A» no tienen observaciones que las justifiquen.");

    if (!c.decision) b("Conclusión", "Falta elegir si se acepta o se declina el encargo.", { campo: "conclusion.decision" });
    if (estado.registro.some(function (x) { return x.descripcion.trim() && !x.nivel; })) {
      avisos.push("Hay riesgos identificados sin nivel de riesgo.");
    }
    return { bloqueos: bloqueos, avisos: avisos };
  }

  function resumen(t, max) { t = String(t).replace(/\s+/g, " ").trim(); return t.length > max ? t.slice(0, max - 1) + "…" : t; }
  function tituloSeccion(n) { for (var i = 0; i < SECCIONES.length; i++) if (SECCIONES[i].n === n) return SECCIONES[i].titulo; return ""; }

  /* =====================================================================
   * 6. FIRMA
   *    El hash cubre todo el contenido del papel. Si cambia algo después de firmar,
   *    el hash deja de coincidir y la firma se marca como inválida.
   * ===================================================================== */
  function contenidoFirmable() {
    return JSON.stringify({ meta: estado.meta, previo: estado.previo, r: estado.r, eeff: estado.eeff, registro: estado.registro, conclusion: estado.conclusion });
  }
  function hashTexto(texto) {
    var cripto = window.crypto && window.crypto.subtle;
    if (cripto && window.TextEncoder) {
      return cripto.digest("SHA-256", new TextEncoder().encode(texto)).then(function (buf) {
        return "sha256:" + Array.prototype.map.call(new Uint8Array(buf), function (x) { return ("0" + x.toString(16)).slice(-2); }).join("");
      });
    }
    // Respaldo (contextos sin crypto.subtle): no es criptográfico, solo detecta cambios.
    var h = 5381; for (var i = 0; i < texto.length; i++) h = ((h << 5) + h + texto.charCodeAt(i)) | 0;
    return Promise.resolve("djb2:" + (h >>> 0).toString(16));
  }
  function verificarFirma() {
    if (!estado.firma) return Promise.resolve();
    return hashTexto(contenidoFirmable()).then(function (h) { estado.firma.invalida = (h !== estado.firma.hash); });
  }

  function iniciarFirma() {
    var ev = evaluar();
    if (ev.bloqueos.length) { abrirPendientes(); return; }
    $("#a2-firmaNombre", raiz).textContent = nombreFirmante();
    $("#a2-firmaConfirma", raiz).hidden = false;
    $("#a2-btnFirmar", raiz).disabled = true;
  }
  function nombreFirmante() { return (usuarioActual && usuarioActual.nombre) || estado.meta.aprobadoPor.trim() || "el socio responsable"; }

  function confirmarFirma() {
    hashTexto(contenidoFirmable()).then(function (h) {
      estado.firma = { por: nombreFirmante(), correo: (usuarioActual && usuarioActual.correo) || "", fecha: ahoraISO(), hash: h, decision: estado.conclusion.decision, riesgo: calcularRiesgo().etiqueta };
      registrarHistorial("firma");
      $("#a2-firmaConfirma", raiz).hidden = true;
      refrescarTodo(); guardarYa();
    });
  }
  function reabrir() {
    registrarHistorial(estado.firma && estado.firma.invalida ? "reapertura-firma-invalida" : "reapertura");
    estado.firma = null;
    $("#a2-firmaReabrirConfirma", raiz).hidden = true;
    refrescarTodo(); guardarYa();
  }
  function registrarHistorial(evento) {
    estado.historial.push({ evento: evento, por: nombreFirmante(), fecha: ahoraISO() });
    if (estado.historial.length > 50) estado.historial.shift();
  }

  /* =====================================================================
   * 7. ALMACENAMIENTO
   *    Dentro de Excel: parte XML personalizada del libro (viaja con el archivo).
   *    Fuera de Excel (vista previa en navegador): localStorage.
   *    Para llevarlo a SharePoint/Graph basta con reemplazar estas dos funciones.
   * ===================================================================== */
  function dentroDeExcel() { return typeof Office !== "undefined" && typeof Excel !== "undefined" && Office.context && Office.context.host === Office.HostType.Excel; }

  var Almacen = {
    cargar: function () {
      if (dentroDeExcel()) {
        return Excel.run(function (ctx) {
          var partes = ctx.workbook.customXmlParts.getByNamespace(CONFIG.NAMESPACE_XML);
          partes.load("items"); return ctx.sync().then(function () {
            if (!partes.items.length) return null;
            var xml = partes.items[0].getXml(); return ctx.sync().then(function () {
              var doc = new DOMParser().parseFromString(xml.value, "application/xml");
              var nodo = doc.documentElement; return nodo && nodo.textContent ? JSON.parse(nodo.textContent) : null;
            });
          });
        }).catch(function (e) { console.warn("AA-02: no se pudo leer del libro", e); return Almacen._local(); });
      }
      return Promise.resolve(Almacen._local());
    },
    guardar: function (datos) {
      var json = JSON.stringify(datos);
      if (dentroDeExcel()) {
        var xml = '<aa02 xmlns="' + CONFIG.NAMESPACE_XML + '">' + escXml(json) + "</aa02>";
        return Excel.run(function (ctx) {
          var partes = ctx.workbook.customXmlParts.getByNamespace(CONFIG.NAMESPACE_XML);
          partes.load("items"); return ctx.sync().then(function () {
            if (partes.items.length) partes.items[0].setXml(xml); else ctx.workbook.customXmlParts.add(xml);
            return ctx.sync();
          });
        });
      }
      try { window.localStorage.setItem(CONFIG.CLAVE_LOCAL, json); } catch (e) { return Promise.reject(e); }
      return Promise.resolve();
    },
    _local: function () {
      try { var t = window.localStorage.getItem(CONFIG.CLAVE_LOCAL); return t ? JSON.parse(t) : null; } catch (e) { return null; }
    }
  };

  var guardarConRetraso = debounce(function () { guardarYa(); }, CONFIG.AUTOGUARDADO_MS);

  function guardarYa() {
    setGuardado("Guardando…", "");
    estado.actualizado = ahoraISO();
    return Almacen.guardar(estado).then(function () {
      setGuardado("Guardado a las " + horaCorta(), "");
      notificarEstado();
    }).catch(function (e) {
      console.error("AA-02: error al guardar", e);
      setGuardado("No se pudo guardar. Se reintentará al próximo cambio.", "error");
    });
  }
  function setGuardado(texto, tipo) { var el = $("#a2-estadoGuardado", raiz); el.textContent = texto; el.setAttribute("data-estado", tipo); }

  function estadoPapel() {
    if (estado.firma) return estado.firma.invalida ? "invalida" : "firmado";
    var av = calcularAvance(), ev = evaluar();
    if (!ev.bloqueos.length) return "completo";
    var algo = av.hechas > 0 || estado.meta.entidad.trim();
    return algo ? "borrador" : "sin-iniciar";
  }
  function notificarEstado() {
    var av = calcularAvance(), rg = calcularRiesgo();
    document.dispatchEvent(new CustomEvent("auditx:estado-papel", {
      detail: { id: CONFIG.ID_PAPEL, estado: estadoPapel(), avance: av.total ? av.hechas / av.total : 0, riesgo: rg.etiqueta, decision: estado.conclusion.decision || null }
    }));
  }

  /* =====================================================================
   * 8. DIBUJO
   * ===================================================================== */
  function plantilla(id) { return document.getElementById("a2-" + id).content.firstElementChild.cloneNode(true); }

  function llenarSelect(sel, opciones, vacio) {
    sel.innerHTML = "";
    var o = document.createElement("option"); o.value = ""; o.textContent = vacio || "Seleccione"; sel.appendChild(o);
    opciones.forEach(function (t) { var op = document.createElement("option"); op.value = t; op.textContent = t; sel.appendChild(op); });
  }

  function crearItem(it) {
    var nodo = plantilla("tpl-item"), r = respuesta(it.id);
    nodo.setAttribute("data-id", it.id);
    $(".aa-item-num", nodo).textContent = it.num ? String(it.num) : (it.etq || "");
    $(".aa-item-text", nodo).textContent = it.texto;
    if (it.nota) { var n = $(".aa-item-nota", nodo); n.textContent = it.nota; n.hidden = false; }
    if (it.ayuda) { var a = $(".aa-item-ayuda", nodo); a.textContent = it.ayuda; a.hidden = false; }
    $$('input[type="radio"]', nodo).forEach(function (rd) { rd.name = "a2-r-" + it.id; });
    llenarSelect($('[data-tipo="riesgo"] select', nodo), CONFIG.NIVELES_RIESGO, "Nivel");
    $('[data-tipo="riesgo"]', nodo).hidden = !it.riesgo;
    pintarItem(nodo, r);
    return nodo;
  }

  function pintarItem(nodo, r) {
    nodo.setAttribute("data-resp", r.r || "");
    $$('input[type="radio"]', nodo).forEach(function (rd) { rd.checked = (rd.value === r.r); });
    $('[data-campo="riesgo"]', nodo).value = r.riesgo || "";
    var cont = $('[data-tipo="riesgo"]', nodo), idx = CONFIG.NIVELES_RIESGO.indexOf(r.riesgo) + 1;
    cont.setAttribute("data-nivel", idx > 0 ? String(Math.min(idx, 5)) : "");
    $('[data-campo="ref"]', nodo).value = r.ref || "";
    $('[data-campo="com"]', nodo).value = r.com || "";
    $(".aa-dot", nodo).hidden = !(r.ref || r.com);
  }

  function crearInfo(it, esNota) {
    var nodo = plantilla("tpl-info");
    $(".aa-item-num", nodo).textContent = it.num ? String(it.num) : (it.etq || "");
    $(".aa-item-text", nodo).textContent = it.texto;
    if (esNota) nodo.classList.add("is-nota");
    return nodo;
  }

  function dibujarSecciones() {
    var cont = $("#a2-secciones", raiz); cont.innerHTML = "";
    var plano = SECCIONES.length === 1;   // un solo bloque: sin acordeón
    SECCIONES.forEach(function (sec) {
      var s = plantilla("tpl-seccion"), cuerpo = $(".aa-sec-body", s), btn = $(".aa-sec-toggle", s);
      s.setAttribute("data-sec", sec.n);
      $(".aa-sec-num", s).textContent = sec.n;
      $(".aa-sec-title", s).textContent = sec.titulo;
      var idc = "a2-sec-" + sec.n; cuerpo.id = idc; btn.setAttribute("aria-controls", idc);
      if (plano) { s.classList.add("is-plana"); cuerpo.hidden = false; btn.setAttribute("aria-expanded", "true"); }
      sec.reqs.forEach(function (rq) {
        var hijos = rq.hijos || [], grupo = document.createElement("div"); grupo.className = "aa-grupo";
        grupo.appendChild(rq.resp ? crearItem(rq) : crearInfo(rq, !rq.num));
        if (hijos.length) {
          var h = document.createElement("div"); h.className = "aa-grupo-hijos";
          hijos.forEach(function (k) { h.appendChild(k.resp ? crearItem(k) : crearInfo(k, true)); });
          grupo.appendChild(h);
        }
        cuerpo.appendChild(grupo);
      });
      cont.appendChild(s);
    });
  }

  function dibujarListaEEFF() {
    var cont = $("#a2-listaRiesgosEEFF", raiz); cont.innerHTML = "";
    estado.eeff.riesgos.forEach(function (txt, i) {
      var f = plantilla("tpl-riesgo-eeff"), ta = $("textarea", f);
      ta.value = txt; ta.setAttribute("data-bind", "eeff.riesgos." + i);
      var q = $("button", f); q.setAttribute("data-quitar-eeff", i);
      q.hidden = estado.eeff.riesgos.length <= 1;
      cont.appendChild(f);
    });
    $("#a2-btnAgregarRiesgoEEFF", raiz).hidden = estado.eeff.riesgos.length >= CONFIG.MAX_RIESGOS_EEFF;
  }

  function dibujarRegistro() {
    var cont = $("#a2-listaRegistro", raiz); cont.innerHTML = "";
    estado.registro.forEach(function (x, i) {
      var f = plantilla("tpl-registro");
      $(".aa-reg-num", f).textContent = "Riesgo " + (i + 1);
      f.setAttribute("data-reg", i);
      llenarSelect($('[data-campo="tipo"]', f), CONFIG.TIPOS_RIESGO, "Tipo");
      llenarSelect($('[data-campo="nivel"]', f), CONFIG.NIVELES_RIESGO, "Nivel");
      llenarSelect($('[data-campo="frecuencia"]', f), CONFIG.FRECUENCIAS, "Frecuencia");
      $$("[data-campo]", f).forEach(function (el) { el.value = x[el.getAttribute("data-campo")] || ""; });
      $("button.aa-icon-btn", f).setAttribute("data-quitar-reg", i);
      cont.appendChild(f);
    });
    $("#a2-btnAgregarRegistro", raiz).hidden = estado.registro.length >= CONFIG.MAX_REGISTRO;
    $("#a2-contadorRegistro", raiz).textContent = estado.registro.length ? String(estado.registro.length) : "";
  }

  // Sincroniza todos los campos con data-bind (datos del encargo, conclusión, textos largos).
  function pintarBinds() {
    $$("[data-bind]", raiz).forEach(function (el) {
      var v = getPath(estado, el.getAttribute("data-bind"));
      if (el.type === "radio") el.checked = (el.value === v);
      else if (el.type === "checkbox") el.checked = !!v;
      else if (document.activeElement !== el) el.value = v == null ? "" : v;
    });
  }

  function pintarAvance() {
    var av = calcularAvance(), pct = av.total ? Math.round((av.hechas / av.total) * 100) : 0;
    $("#a2-rellenoAvance", raiz).style.width = pct + "%";
    $("#a2-barraAvance", raiz).setAttribute("aria-valuenow", String(pct));
    $("#a2-textoAvance", raiz).textContent = av.hechas + " de " + av.total + " respondidas (" + pct + " %)";
    $("#a2-resumenRequerimientos", raiz).textContent = av.total - av.hechas ? (av.total - av.hechas) + " por responder" : "Todas respondidas";
    $$(".aa-sec", raiz).forEach(function (s) {
      var d = av.porSeccion[s.getAttribute("data-sec")], c = $(".aa-count", s);
      c.textContent = d.hechas + "/" + d.total; c.setAttribute("data-completo", String(d.total > 0 && d.hechas === d.total));
    });
    var faltan = estado.meta.entidad.trim() ? 0 : 1;
    var cd = $("#a2-contadorDatos", raiz); cd.textContent = faltan ? faltan + " por completar" : "Completo"; cd.setAttribute("data-completo", String(!faltan));
    var pv = calcularPrevio(), cp = $("#a2-contadorPrevio", raiz);
    cp.textContent = pv.sinResponder ? pv.sinResponder + " por responder" : "Completo"; cp.setAttribute("data-completo", String(!pv.sinResponder));
  }

  function pintarPrevio() {
    var pv = calcularPrevio(), caja = $("#a2-resultadoPrevio", raiz);
    if (pv.reevaluar) {
      caja.textContent = "Según los criterios anteriores se recomienda realizar la reevaluación del cliente existente, y evaluar la aprobación para futuros clientes. Posteriormente complete el siguiente checklist como anexo del memorando de reevaluación junto con la justificación, métodos y conclusiones.";
      caja.hidden = false;
    } else if (!pv.sinResponder) {
      caja.textContent = "Ningún criterio aplica; no se recomienda reevaluar al cliente por este procedimiento.";
      caja.hidden = false;
    } else { caja.hidden = true; }
  }

  function pintarRiesgo() {
    var rg = calcularRiesgo(), tope = $("#a2-chipRiesgoTop", raiz), chip = $("#a2-chipRiesgo", raiz), det = $("#a2-detalleRiesgo", raiz);
    var nivel = rg.nivel ? String(rg.nivel) : "";
    tope.setAttribute("data-nivel", nivel); chip.setAttribute("data-nivel", nivel);
    tope.textContent = rg.nivel ? "Riesgo: " + rg.etiqueta : "Riesgo: sin calificar";
    chip.textContent = rg.nivel ? rg.etiqueta : "Sin calificar";
    det.textContent = rg.nivel
      ? "Promedio de " + rg.promedio.toFixed(2).replace(".", ",") + " sobre 5, calculado con " + rg.n + " requerimientos que tienen nivel de riesgo."
      : "Califique el nivel de riesgo de los requerimientos para obtener la clasificación.";
    var dist = $("#a2-distribucionRiesgo", raiz), ley = $("#a2-leyendaRiesgo", raiz);
    dist.innerHTML = ""; ley.innerHTML = "";
    CONFIG.NIVELES_RIESGO.forEach(function (nom, i) {
      var cant = rg.conteo[i], sp = document.createElement("span");
      sp.className = "d" + Math.min(i + 1, 5); sp.style.width = (rg.n ? (cant / rg.n) * 100 : 0) + "%"; dist.appendChild(sp);
      var li = document.createElement("li"), sw = document.createElement("i");
      sw.className = "d" + Math.min(i + 1, 5); sw.style.background = "var(--r" + Math.min(i + 1, 5) + ")";
      li.appendChild(sw); li.appendChild(document.createTextNode(nom + ": " + cant)); ley.appendChild(li);
    });
    dist.setAttribute("aria-label", "Distribución por nivel de riesgo: " + CONFIG.NIVELES_RIESGO.map(function (n, i) { return n + " " + rg.conteo[i]; }).join(", "));
    $("#a2-avisoRiesgoAlto", raiz).hidden = !(estado.conclusion.decision === "Aceptar" && rg.nivel >= 4);
  }

  function pintarEstadoYFirma() {
    var est = estadoPapel(), chip = $("#a2-chipEstado", raiz);
    var textos = { "sin-iniciar": "Sin iniciar", borrador: "En borrador", completo: "Completo, sin firmar", firmado: "Firmado", invalida: "Firma inválida" };
    chip.textContent = textos[est]; chip.setAttribute("data-estado", est); raiz.setAttribute("data-estado", est);

    var ev = evaluar(), bloque = $("#a2-bloqueFirma", raiz), txt = $("#a2-firmaEstado", raiz);
    bloque.setAttribute("data-firma", est === "firmado" ? "firmado" : est === "invalida" ? "invalida" : "");
    var badge = $("#a2-badgePendientes", raiz); badge.textContent = ev.bloqueos.length; badge.hidden = !ev.bloqueos.length;

    var btnF = $("#a2-btnFirmar", raiz);
    btnF.hidden = false;
    if (estado.firma) {
      var f = estado.firma, cuando = new Date(f.fecha).toLocaleString("es-CO");
      txt.textContent = est === "invalida"
        ? "La firma de " + f.por + " (" + cuando + ") ya no es válida: el contenido cambió después de firmar."
        : "Firmado por " + f.por + " el " + cuando + " · Decisión: " + (f.decision === "Aceptar" ? "aceptar" : "declinar") + " el encargo.";
      btnF.textContent = est === "invalida" ? "Reabrir" : "Reabrir para editar";
      btnF.disabled = false; btnF.setAttribute("data-modo", "reabrir");
    } else {
      txt.textContent = ev.bloqueos.length ? "Sin firmar. Faltan " + ev.bloqueos.length + " pendientes para poder firmar." : "Listo para firmar.";
      btnF.textContent = "Firmar"; btnF.setAttribute("data-modo", "firmar");
      btnF.disabled = !!ev.bloqueos.length || !$("#a2-firmaConfirma", raiz).hidden;
    }
    aplicarBloqueo(!!estado.firma && est === "firmado");
  }

  function aplicarBloqueo(bloqueado) {
    $$("input, select, textarea, .aa-icon-btn, #a2-btnAgregarRiesgoEEFF, #a2-btnAgregarRegistro", raiz).forEach(function (el) {
      if (el.closest("[data-siempre-activo]") || el.hasAttribute("data-siempre-activo") || el.readOnly) return;
      el.disabled = bloqueado;
    });
  }

  function refrescarTodo() {
    pintarAvance(); pintarPrevio(); pintarRiesgo(); pintarEstadoYFirma();
    $("#a2-panelPendientes", raiz).hidden ? null : dibujarPendientes();
  }

  /* =====================================================================
   * 9. PENDIENTES
   * ===================================================================== */
  function abrirPendientes() { $("#a2-panelPendientes", raiz).hidden = false; dibujarPendientes(); $("#a2-panelPendientes", raiz).scrollIntoView({ behavior: "smooth", block: "start" }); }
  function dibujarPendientes() {
    var ev = evaluar(), cuerpo = $("#a2-pendCuerpo", raiz); cuerpo.innerHTML = "";
    if (!ev.bloqueos.length) {
      var ok = document.createElement("p"); ok.className = "aa-pend-ok"; ok.textContent = "No hay pendientes: el papel está completo y se puede firmar."; cuerpo.appendChild(ok);
    } else {
      var grupos = {}, orden = [];
      ev.bloqueos.forEach(function (p, i) { if (!grupos[p.grupo]) { grupos[p.grupo] = []; orden.push(p.grupo); } grupos[p.grupo].push({ p: p, i: i }); });
      orden.forEach(function (g) {
        var h = document.createElement("p"); h.className = "aa-pend-grupo"; h.textContent = g + " (" + grupos[g].length + ")"; cuerpo.appendChild(h);
        var ul = document.createElement("ul"); ul.className = "aa-pend-lista";
        grupos[g].slice(0, 12).forEach(function (x) {
          var li = document.createElement("li"), b = document.createElement("button"); b.type = "button"; b.textContent = x.p.texto;
          b.addEventListener("click", function () { irA(x.p.destino); }); li.appendChild(b); ul.appendChild(li);
        });
        if (grupos[g].length > 12) { var m = document.createElement("li"); m.className = "aa-hint"; m.textContent = "… y " + (grupos[g].length - 12) + " más en esta sección."; ul.appendChild(m); }
        cuerpo.appendChild(ul);
      });
    }
    if (ev.avisos.length) {
      var ha = document.createElement("p"); ha.className = "aa-pend-grupo"; ha.textContent = "Avisos (no impiden firmar)"; cuerpo.appendChild(ha);
      var ua = document.createElement("ul"); ua.className = "aa-pend-lista";
      ev.avisos.forEach(function (t) { var li = document.createElement("li"); li.className = "aa-hint"; li.textContent = t; ua.appendChild(li); });
      cuerpo.appendChild(ua);
    }
  }

  function alternar(btn, panel, abrir) {
    var v = typeof abrir === "boolean" ? abrir : btn.getAttribute("aria-expanded") !== "true";
    btn.setAttribute("aria-expanded", String(v)); panel.hidden = !v;
  }

  function irA(dest) {
    var el = null;
    if (dest.id) {
      var nodo = $('.aa-item[data-id="' + dest.id + '"]', raiz);
      var sec = nodo && nodo.closest(".aa-sec");
      if (sec) alternar($(".aa-sec-toggle", sec), $(".aa-sec-body", sec), true);
      if (nodo && dest.campo === "com") { alternar($(".aa-detalle-toggle", nodo), $(".aa-detalle", nodo), true); }
      el = nodo && (dest.campo ? $('[data-campo="' + dest.campo + '"]', nodo) : $('input[type="radio"]', nodo));
      if (nodo) { nodo.classList.add("is-pendiente-foco"); setTimeout(function () { nodo.classList.remove("is-pendiente-foco"); }, 1800); }
    } else if (dest.campo) {
      el = $('[data-bind="' + dest.campo + '"]', raiz);
      var card = el && el.closest(".aa-card-body");
      if (card && card.hidden) alternar($('[aria-controls="' + card.id + '"]', raiz), card, true);
    }
    if (el) { el.scrollIntoView({ behavior: "smooth", block: "center" }); try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
  }

  /* =====================================================================
   * 10. EVENTOS
   * ===================================================================== */
  function alCambiar() {
    // Todo cambio posterior a la firma la invalida (por si el bloqueo visual se saltó).
    if (estado.firma && !estado.firma.invalida) { verificarFirma().then(function () { refrescarTodo(); }); }
    refrescarTodo(); guardarConRetraso(); setGuardado("Cambios sin guardar…", "");
  }

  function alEditarItem(e) {
    var el = e.target, nodo = el.closest(".aa-item");
    if (!nodo) return false;
    var campo = el.getAttribute("data-campo"); if (!campo) return true;
    var id = nodo.getAttribute("data-id"), r = respuesta(id);
    r[campo] = el.value;
    if (campo === "r") {
      if (r.r === "NO") { alternar($(".aa-detalle-toggle", nodo), $(".aa-detalle", nodo), true); }
    }
    pintarItem(nodo, r);
    return true;
  }

  function alEditarBind(el) {
    var ruta = el.getAttribute("data-bind"), v;
    if (el.type === "checkbox") v = el.checked;
    else if (el.type === "radio") { if (!el.checked) return; v = el.value; }
    else v = el.value;
    setPath(estado, ruta, v);
  }

  function conectarEventos() {
    var cambio = function (e) {
      var el = e.target;
      if (alEditarItem(e)) { alCambiar(); return; }
      var reg = el.closest(".aa-reg");
      if (reg && el.getAttribute("data-campo")) {
        var i = +reg.getAttribute("data-reg"); estado.registro[i][el.getAttribute("data-campo")] = el.value; alCambiar(); return;
      }
      if (el.hasAttribute("data-bind")) { alEditarBind(el); alCambiar(); }
    };
    raiz.addEventListener("input", function (e) { if (e.target.matches('input[type="text"], input[type="date"], textarea')) cambio(e); });
    raiz.addEventListener("change", function (e) { if (!e.target.matches('input[type="text"], input[type="date"], textarea')) cambio(e); });

    raiz.addEventListener("click", function (e) {
      var t = e.target.closest("button"); if (!t) return;
      if (t.classList.contains("aa-sec-toggle")) { alternar(t, $(".aa-sec-body", t.closest(".aa-sec"))); return; }
      if (t.classList.contains("aa-detalle-toggle")) { alternar(t, $(".aa-detalle", t.closest(".aa-item"))); return; }
      if (t.classList.contains("aa-card-toggle")) { alternar(t, document.getElementById(t.getAttribute("aria-controls"))); return; }
      if (t.hasAttribute("data-quitar-eeff")) { estado.eeff.riesgos.splice(+t.getAttribute("data-quitar-eeff"), 1); if (!estado.eeff.riesgos.length) estado.eeff.riesgos.push(""); dibujarListaEEFF(); alCambiar(); return; }
      if (t.hasAttribute("data-quitar-reg")) { estado.registro.splice(+t.getAttribute("data-quitar-reg"), 1); dibujarRegistro(); alCambiar(); return; }
    });

    $("#a2-btnAgregarRiesgoEEFF", raiz).addEventListener("click", function () {
      if (estado.eeff.riesgos.length < CONFIG.MAX_RIESGOS_EEFF) { estado.eeff.riesgos.push(""); dibujarListaEEFF(); alCambiar(); }
    });
    $("#a2-btnAgregarRegistro", raiz).addEventListener("click", function () {
      if (estado.registro.length < CONFIG.MAX_REGISTRO) {
        estado.registro.push({ descripcion: "", tipo: "", nivel: "", control: "", frecuencia: "", refPT: "" });
        dibujarRegistro(); alCambiar();
        var ult = $$(".aa-reg", raiz).pop(); if (ult) { alternar($("#a2-regToggle", raiz), $("#a2-regCuerpo", raiz), true); $("textarea", ult).focus(); }
      }
    });

    $("#a2-btnPendientes", raiz).addEventListener("click", function () { var p = $("#a2-panelPendientes", raiz); if (p.hidden) abrirPendientes(); else p.hidden = true; });
    $("#a2-btnCerrarPendientes", raiz).addEventListener("click", function () { $("#a2-panelPendientes", raiz).hidden = true; });

    $("#a2-btnFirmar", raiz).addEventListener("click", function () {
      if (this.getAttribute("data-modo") === "reabrir") { $("#a2-firmaReabrirConfirma", raiz).hidden = false; $("#a2-bloqueFirma", raiz).scrollIntoView({ behavior: "smooth", block: "center" }); }
      else iniciarFirma();
    });
    $("#a2-btnConfirmarFirma", raiz).addEventListener("click", confirmarFirma);
    $("#a2-btnCancelarFirma", raiz).addEventListener("click", function () { $("#a2-firmaConfirma", raiz).hidden = true; pintarEstadoYFirma(); });
    $("#a2-btnConfirmarReabrir", raiz).addEventListener("click", reabrir);
    $("#a2-btnCancelarReabrir", raiz).addEventListener("click", function () { $("#a2-firmaReabrirConfirma", raiz).hidden = true; });

    $("#a2-btnVolver", raiz).addEventListener("click", function () { guardarYa(); document.dispatchEvent(new CustomEvent("auditx:volver", { detail: { desde: CONFIG.ID_PAPEL } })); });

    // Integración con el menú de Pre-Planificación.
    document.addEventListener("auditx:navigate", function (e) {
      if (e.detail && e.detail.id === CONFIG.ID_PAPEL) { document.body.hidden = false; window.scrollTo(0, 0); }
    });
  }

  /* =====================================================================
   * 11. ARRANQUE
   * ===================================================================== */
  function fusionar(base, guardado) {
    // Tolera datos de versiones anteriores: completa lo que falte sin perder lo guardado.
    var out = copiar(base);
    if (!guardado || typeof guardado !== "object") return out;
    ["meta", "previo", "eeff", "conclusion"].forEach(function (k) { if (guardado[k]) Object.keys(guardado[k]).forEach(function (c) { out[k][c] = guardado[k][c]; }); });
    out.r = guardado.r || {}; out.registro = guardado.registro || []; out.firma = guardado.firma || null;
    out.historial = guardado.historial || []; out.actualizado = guardado.actualizado || null;
    if (!out.eeff.riesgos || !out.eeff.riesgos.length) out.eeff.riesgos = [""];
    return out;
  }

  function iniciar() {
    raiz = document.getElementById("aa02");
    if (!raiz || raiz.getAttribute("data-listo")) return;
    raiz.setAttribute("data-listo", "1");
    indexarDatos();
    conectarEventos();

    Promise.all([Almacen.cargar(), CONFIG.obtenerUsuario(), CONFIG.obtenerEncargo()]).then(function (res) {
      usuarioActual = res[1];
      estado = fusionar(nuevoEstado(), res[0]);
      var enc = res[2] || {};
      Object.keys(enc).forEach(function (k) { if (k in estado.meta && !String(estado.meta[k]).trim()) estado.meta[k] = enc[k]; });
      if (usuarioActual && usuarioActual.nombre && !estado.meta.preparadoPor.trim()) estado.meta.preparadoPor = usuarioActual.nombre;
      return verificarFirma();
    }).catch(function (e) { console.error("AA-02: error al iniciar", e); }).then(function () {
      dibujarSecciones(); dibujarListaEEFF(); dibujarRegistro(); pintarBinds();
      // Abre la primera sección con pendientes para retomar donde se quedó.
      var av = calcularAvance(), primera = SECCIONES.filter(function (s) { return av.porSeccion[s.n].hechas < av.porSeccion[s.n].total; })[0] || SECCIONES[0];
      var secEl = $('.aa-sec[data-sec="' + primera.n + '"]', raiz);
      if (secEl && !secEl.classList.contains("is-plana")) alternar($(".aa-sec-toggle", secEl), $(".aa-sec-body", secEl), true);
      refrescarTodo(); notificarEstado();
      setGuardado(estado.actualizado ? "Guardado" : "Sin cambios", "");
    });
  }

  window.AuditX = window.AuditX || {};
  window.AuditX.AA02 = {
    iniciar: iniciar,
    config: CONFIG,
    obtenerEstado: function () { return copiar(estado); },
    evaluar: evaluar,
    calcularRiesgo: calcularRiesgo,
    exportarJSON: function () { return JSON.stringify(estado, null, 2); }
  };

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(function () { iniciar(); });
  else if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
