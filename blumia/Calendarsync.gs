function syncCalendarConfirmados() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    const datos = sh.getDataRange().getValues();

    eventos.forEach(ev => {
      const titulo = norm(ev.getTitle());
      const desc = norm(ev.getDescription());

      let estado = null;
      if (titulo.indexOf('screening') > -1) estado = 'Screening agendado';
      else if (titulo.indexOf('competencia') > -1) estado = 'Competencias agendada';
      else if (titulo.indexOf('empresa') > -1) estado = 'Entrevista empresa agendada';
      if (!estado) return;

      for (let i = 1; i < datos.length; i++) {
        const nombre = norm(datos[i][1]).replace(/\s+/g,'');
        const email = norm(datos[i][21]);
        if (!nombre) continue;
        const mParen = ev.getTitle().match(/\(([^)]+)\)/);
        const nombreEv = mParen ? norm(mParen[1]).replace(/\s+/g,'') : '';
        const matchMail = email && email.includes('@') && (titulo.indexOf(email) > -1 || desc.indexOf(email) > -1);
        const matchNombre = nombreEv && (nombre === nombreEv || (nombreEv.length >= 10 && (nombre.includes(nombreEv) || nombreEv.includes(nombre))));
       if (matchMail || matchNombre) {
          const celda = sh.getRange(i + 1, 9);
          const actual = String(celda.getValue()||'').toLowerCase();
          const noTocar = actual.includes('eliminado') || actual.includes('descartado') || actual.includes('avanza a empresa') || actual.includes('en cliente') || actual.includes('contratado') || actual.includes('reserva') || actual.includes('entrevista empresa');
          if (celda.getValue() !== estado && !noTocar) {
            celda.setValue(estado).setBackground('#D1E9FF').setHorizontalAlignment('center').setFontWeight('bold');
          }
        }
      }
    });
  });
}