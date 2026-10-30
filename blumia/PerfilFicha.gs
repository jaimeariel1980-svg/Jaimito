/* ============================================================
   FICHA DEL PUESTO a partir del perfil del puesto (30/10, Ariel: «el perfil es dios de todo el proceso»).
   perfilFicha_(cfg, uso) devuelve el perfil acordado con el cliente como texto «Sección › Campo: valor» para pasárselo a la IA de cada etapa.
   Usos:
     'cv'        análisis de CV, informe y asistente interno (incluye señales de alerta; sin sueldos)
     'candidato' lo que ve el candidato: flyer, avisos, mensajes de hunting (sin señales de alerta, sin sueldos; con la búsqueda ciega, sin datos de la empresa)
   NUNCA entran edad, género, estado civil ni situación familiar: son información, no criterio (discriminatorio y prohibido por la Ley 23.592).
   ============================================================ */

const PERFIL_NIVELES_ = ['', 'Básico', 'Intermedio', 'Avanzado', 'Experto'];

function perfilFicha_(cfg, uso) {
  try {
    const u = ultimoPerfilDatos_(cfg);
    if (!u) return '';
    const d = JSON.parse(u.texto), K = PERFIL_CAMPOS_, v = d.v || [];
    const ciega = busquedaCiega_(cfg), cand = uso === 'candidato';
    const SEC_FUERA = /solicitante|pasos del proceso|notas internas|firmas/i;
    const IDX_DEMO = { 19: 1, 20: 1, 21: 1 }; // edad mínima / máxima / excluyente
    const IDX_PLATA = { 58: 1, 59: 1, 60: 1, 61: 1, 64: 1, 57: 1 };
    const IDX_EMPRESA = { 0: 1, 1: 1, 3: 1, 4: 1, 5: 1, 6: 1 };
    const L = []; let previoOk = false;
    K.campos.forEach((c, idx) => {
      let ok = false;
      const sec = String(c.s || ''), lab = String(c.l || '');
      if (!SEC_FUERA.test(sec) && !(NO_RELLENAR_PERFIL_.test(lab + ' ' + sec)) && !IDX_DEMO[c.i] && !IDX_PLATA[c.i]
        && !(cand && /red flags/i.test(sec)) && !(cand && ciega && IDX_EMPRESA[c.i])) {
        let x = v[idx];
        if (c.t === 'select') { const k = Number(x); x = (k > 0 && c.o && c.o[k]) ? c.o[k] : ''; }
        x = String(x == null ? '' : x).trim();
        if (x) {
          ok = true;
          if (/^(¿es excluyente\?|nivel|nivel requerido)$/i.test(lab.trim())) { if (previoOk && L.length) L[L.length - 1] += ' [' + lab.replace(/[¿?]/g, '') + ': ' + x + ']'; }
          else L.push((sec ? sec + ' › ' : '') + (lab || c.h || '') + ': ' + x);
        }
      }
      previoOk = ok && !/^(¿es excluyente\?|nivel|nivel requerido)$/i.test(lab.trim()) ? true : (ok && previoOk);
    });
    // competencias 6 a 10 y resumen general (van al final del arreglo para no correr los datos viejos)
    for (let k = 0; k < 5; k++) {
      const nom = String(v[81 + 2 * k] || '').trim(), niv = PERFIL_NIVELES_[Number(v[82 + 2 * k]) || 0];
      if (nom) L.push('Competencias requeridas › Competencia ' + (6 + k) + ': ' + nom + (niv ? ' [Nivel requerido: ' + niv + ']' : ''));
    }
    const res = String(v[91] || '').trim();
    if (res && !cand) L.push('Resumen general del perfil: ' + res);
    return L.join('\n').slice(0, 14000);
  } catch (e) { return ''; }
}

// Bloque listo para pegar en el prompt del análisis de CV (vacío si la búsqueda todavía no tiene perfil cargado)
function perfilFichaBloque_(cfg) {
  const f = perfilFicha_(cfg, 'cv');
  if (!f) return '';
  return '\n\nFICHA DEL PUESTO acordada con el cliente (perfil del puesto). Usala como contexto principal del puesto; los campos marcados «Excluyente» pesan más que los demás. Si algo contradice los requisitos de arriba, priorizá esta ficha. ' +
    'No uses edad, género, estado civil, situación familiar ni nacionalidad para puntuar ni para decidir.\n' + f.slice(0, 7000) + '\n';
}

/* Perfil del puesto para CONSULTAR (referencia en el tablero, en los formularios de screening y competencias, en el hunting…):
   secciones con «campo: valor», solo lectura. Para el equipo de Blumia: incluye señales de alerta y condiciones, nunca edad, género, estado civil ni datos del solicitante. */
function perfilReferencia_(perfilId) {
  const cfg = getConfig(perfilId);
  const hero = getHeroPostulacion(perfilId);
  const out = { pos: String(hero.pos || cfg.nombre || perfilId), fecha: '', secciones: [], vacio: true };
  const u = ultimoPerfilDatos_(cfg);
  if (!u) return out;
  out.fecha = u.fecha;
  const d = JSON.parse(u.texto), K = PERFIL_CAMPOS_, v = d.v || [];
  const SEC_FUERA = /solicitante|pasos del proceso|notas internas|firmas/i, DEMO = { 19: 1, 20: 1, 21: 1 };
  const AUX = /^(¿es excluyente\?|nivel|nivel requerido)$/i;
  const map = {}, orden = [];
  const poner = (sec, lab, val) => { if (!map[sec]) { map[sec] = []; orden.push(sec); } map[sec].push([lab, val]); };
  K.campos.forEach((c, idx) => {
    const sec = String(c.s || ''), lab = String(c.l || '').trim();
    if (SEC_FUERA.test(sec) || DEMO[c.i] || NO_RELLENAR_PERFIL_.test(lab + ' ' + sec)) return;
    let x = v[idx];
    if (c.t === 'select') { const k = Number(x); x = (k > 0 && c.o && c.o[k]) ? c.o[k] : ''; }
    x = String(x == null ? '' : x).trim();
    if (!x) return;
    if (AUX.test(lab)) { const it = map[sec] && map[sec][map[sec].length - 1]; if (it && !it.aux) it[1] += ' · ' + (/nivel/i.test(lab) ? 'nivel ' + x.toLowerCase() : x.toLowerCase()); return; }
    poner(sec, lab || c.h || '', x);
  });
  for (let k = 0; k < 5; k++) {
    const nom = String(v[81 + 2 * k] || '').trim(), niv = PERFIL_NIVELES_[Number(v[82 + 2 * k]) || 0];
    if (nom) poner('Competencias requeridas', 'Competencia ' + (6 + k), nom + (niv ? ' · nivel ' + niv.toLowerCase() : ''));
  }
  const res = String(v[91] || '').trim();
  if (res) poner('Resumen general', 'Resumen', res);
  out.secciones = orden.map(s => ({ t: s, items: map[s].slice(0, 40).map(i => [String(i[0]).slice(0, 120), String(i[1]).slice(0, 3000)]) }));
  out.vacio = !out.secciones.length;
  return out;
}
