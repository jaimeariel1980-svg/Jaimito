/* Competencias del formulario a partir del perfil del puesto (4/11, Ariel: «las competencias del formulario no son las del perfil»).
   Lee las competencias que se acordaron con el cliente (hasta 10, con su nivel requerido) y propone la lista del formulario con una descripción y 3 preguntas de ejemplos concretos
   por competencia (Claude). NO guarda nada: la pantalla las carga en el editor para que la persona las revise y toque «Guardar cambios». Mismo criterio de siempre (pesos 1 y 2): Experto = 2, el resto = 1; se puede cambiar a mano en el editor. */
function competenciasDelPerfil_(perfilId) {
  const cfg = getConfig(perfilId);
  const u = ultimoPerfilDatos_(cfg);
  if (!u) throw new Error('Esta búsqueda todavía no tiene el perfil del puesto cargado (⚙ Configurar → Perfil del puesto).');
  const v = (JSON.parse(u.texto).v) || [];
  const base = [];
  const poner = (nom, niv) => {
    nom = String(nom == null ? '' : nom).trim();
    if (!nom) return;
    const i = nom.indexOf(':');
    const n = (i > 0 ? nom.slice(0, i) : nom).trim(), d = (i > 0 ? nom.slice(i + 1) : '').trim();
    base.push({ nombre: n.slice(0, 120), descripcion: d.slice(0, 300), peso: (Number(niv) || 0) >= 4 ? 2 : 1 });
  };
  for (let k = 0; k < 5; k++) poner(v[46 + 2 * k], v[47 + 2 * k]);
  for (let k = 0; k < 5; k++) poner(v[81 + 2 * k], v[82 + 2 * k]);
  if (!base.length) throw new Error('El perfil del puesto no tiene competencias cargadas.');
  const sistema = 'Sos un consultor senior de selección. Para cada competencia de un perfil de puesto armás: una descripción y 3 preguntas de entrevista por competencias que pidan ejemplos concretos de situaciones reales vividas (formato «Contame de una vez que…», «Describí una situación en la que…»). ' +
    'CRITERIO DE REDACCIÓN DE LAS DESCRIPCIONES (igual para todas, para que la lista se lea pareja): una sola oración de 10 a 20 palabras, en tercera persona del presente, que arranque con un verbo conjugado («Prospecta…», «Construye…», «Lee…», «Se compromete…») y describa la conducta observable que se espera; sin repetir el nombre de la competencia, sin enumerar sustantivos sueltos, sin entre paréntesis y terminada en punto. Ejemplos del estilo: «Prospecta y abre cuentas nuevas en ciclos de venta largos y complejos.» · «Construye y sostiene vínculos de largo plazo con decisores y hace venta cruzada.» · «Se compromete con objetivos de ventas y margen y los cumple.» Si el cliente dio una aclaración, usala como insumo pero reescribila con este criterio. Las preguntas tienen que ser específicas del puesto y del sector. Prohibido preguntar por edad, género, estado civil, hijos, embarazo, religión, nacionalidad o cualquier dato personal protegido. Español rioplatense, sin tecnicismos de RRHH. ' +
    'Devolvé SOLO JSON válido: {"competencias":[{"nombre":"igual al recibido","descripcion":"una oración","preguntas":["…","…","…"]}]} en el mismo orden y con la misma cantidad que recibiste.';
  const contexto = 'PUESTO:\n' + perfilFicha_(cfg, 'cv').slice(0, 5000) + '\n\nCOMPETENCIAS A DESARROLLAR (nombre, y descripción si el cliente la dio):\n' + base.map((c, i) => (i + 1) + '. ' + c.nombre + (c.descripcion ? ' — ' + c.descripcion : '')).join('\n');
  let ia = [];
  try { ia = (extraerJSON_(procesarConClaude(sistema, contexto, 'analisis')).competencias) || []; } catch (e) { Logger.log('competenciasDelPerfil_: ' + e); }
  const out = base.map((c, i) => {
    const r = ia[i] || {};
    return {
      nombre: c.nombre,
      descripcion: descripcionPareja_(r.descripcion || c.descripcion || ''),
      peso: c.peso,
      preguntas: (Array.isArray(r.preguntas) ? r.preguntas : []).map(q => limpiarTxt_(q)).filter(Boolean).slice(0, 4)
    };
  });
  return { comps: out, fecha: u.fecha, conIA: out.some(c => c.preguntas.length) };
}

// deja la descripción con el mismo formato en todas: primera letra en mayúscula y un solo punto final
function descripcionPareja_(t) {
  t = limpiarTxt_(t).replace(/\s+/g, ' ').replace(/[\s.;:,]+$/, '');
  if (!t) return '';
  return t.charAt(0).toUpperCase() + t.slice(1) + '.';
}
