const CONTACTO_PREF = 'Postu Enfe 2026 ', CONTACTO_GRUPO = 'Postulantes Enfermería 2026';

function grupoContactos_() {
  const g = (People.ContactGroups.list({ pageSize: 1000 }).contactGroups || []).find(x => x.name === CONTACTO_GRUPO);
  return g ? g.resourceName : People.ContactGroups.create({ contactGroup: { name: CONTACTO_GRUPO } }).resourceName;
}

function contactosPorTel_() {
  const out = {};
  let tok;
  do {
    const r = People.People.Connections.list('people/me', { pageSize: 1000, personFields: 'names,phoneNumbers', pageToken: tok });
    (r.connections || []).forEach(p => (p.phoneNumbers || []).forEach(t => { const k = telWa_(t.value); if (k.length > 10) out[k] = p; }));
    tok = r.nextPageToken;
  } while (tok);
  return out;
}

function agendarCandidatos() {
  const sh = hoja_('Candidatos');
  if (!sh || sh.getLastRow() < 2) return 'Sin candidatos';
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_CAND.length).getValues();
  const grp = grupoContactos_(), ya = contactosPorTel_(), crear = [], vistos = {};
  let renombrados = 0;
  datos.forEach(r => {
    const nom = String(r[COL.nombre - 1] || '').replace(/\s+/g, ' ').trim(), tel = String(r[COL.tel - 1] || '');
    if (!nom || tel.replace(/\D/g, '').length < 8) return;
    const k = telWa_(tel), nombre = CONTACTO_PREF + nom, p = ya[k];
    if (vistos[k]) return;
    vistos[k] = 1;
    if (!p) {
      crear.push({ contactPerson: { names: [{ givenName: nombre }], phoneNumbers: [{ value: '+' + k, type: 'mobile' }], memberships: [{ contactGroupMembership: { contactGroupResourceName: grp } }] } });
      return;
    }
    const actual = p.names && p.names[0] ? String(p.names[0].displayName || '') : '';
    if (actual.indexOf(CONTACTO_PREF) === 0 && actual !== nombre) {
      try {
        People.People.updateContact({ etag: p.etag, names: [{ givenName: nombre }] }, p.resourceName, { updatePersonFields: 'names' });
        renombrados++;
      } catch (e) { registrarError_('renombrar contacto', e); }
    }
  });
  for (let i = 0; i < crear.length; i += 200) {
    People.People.batchCreateContacts({ contacts: crear.slice(i, i + 200), readMask: 'names' });
    if (i + 200 < crear.length) Utilities.sleep(2000);
  }
  PropertiesService.getScriptProperties().setProperty('CONTACTOS_OK', '1');
  const msg = 'Contactos: ' + crear.length + ' agendados nuevos, ' + renombrados + ' actualizados (' + Object.keys(vistos).length + ' candidatos con teléfono).';
  Logger.log(msg);
  return msg;
}

function agendarCandidatosAuto_() {
  if (PropertiesService.getScriptProperties().getProperty('CONTACTOS_OK') === '1') agendarCandidatos();
}
