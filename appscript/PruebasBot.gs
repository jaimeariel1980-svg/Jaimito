// Corredor de pruebas del bot (PRUEBAS_BOT_AZCU.md). Se ejecuta a mano desde el editor; no confirma tarjetas.
// Requiere la propiedad del script AZCU_PRUEBAS = SI (se apaga sola al terminar) para que no pueda dispararlo cualquiera desde el /exec.
function azcuPruebasBot(){
  var P = PropertiesService.getScriptProperties();
  if(P.getProperty('AZCU_PRUEBAS')!=='SI') throw new Error('Para correr las pruebas cargá la propiedad del script AZCU_PRUEBAS = SI y ejecutá de nuevo. Se apaga sola al terminar.');
  var t0 = Date.now(), T = azcuPruebasTokens_(), lista = azcuPruebasLista_(T), i = +P.getProperty('AZCU_PRUEBAS_I') || 0, run = P.getProperty('AZCU_PRUEBAS_ID') || '';
  var ss = SpreadsheetApp.openById(MAESTRO_ID), sh = ss.getSheetByName('Pruebas_Bot') || ss.insertSheet('Pruebas_Bot');
  if(!run){ run = Utilities.formatDate(new Date(), AZCU_TZ, 'dd/MM HH:mm')+' · '+azcuModelo_(); P.setProperty('AZCU_PRUEBAS_ID', run); }
  if(sh.getLastRow()===0) sh.appendRow(['Corrida','#','Frase','Esperado','Respuesta','Tarjetas','Llamadas','Seg','Resultado','Motivo']);
  var pin = P.getProperty('AZCU_PIN');
  for(; i<lista.length; i++){
    if(Date.now()-t0 > 290000){ P.setProperty('AZCU_PRUEBAS_I', String(i)); Logger.log('Pausa en la prueba '+(i+1)+' de '+lista.length+'. Volvé a ejecutar azcuPruebasBot para seguir.'); return; }
    var q = lista[i], r = {texto:'', acciones:[], adjuntos:[]}, ini = Date.now(), n0, res;
    if(q.skip){ sh.appendRow([run, q.n, q.frase, q.esperado, '', '', '', '', '⏭ no corrida', q.skip]); continue; }
    n0 = AZCU_NLLAM_;
    try{ r = azcuChat(pin, q.hist.concat([{role:'user', content:q.frase}]), q.ctx) || r; }catch(e){ r = {texto:'ERROR: '+e.message, acciones:[], adjuntos:[]}; }
    var seg = Math.round((Date.now()-ini)/100)/10, llam = AZCU_NLLAM_-n0;
    res = azcuPruebasJuzgar_(q, r, seg, llam, T);
    sh.appendRow([run, q.n, q.frase, q.esperado, String(r.texto||'').slice(0,700), (r.acciones||[]).map(function(a){ return a.tipo+' → '+a.resumen; }).join(' | ').slice(0,700), llam, seg, res.v, res.m]);
  }
  var f = sh.getDataRange().getValues().filter(function(x){ return x[0]===run && /^[✅🟡❌]/.test(String(x[8])); }), c = {'✅':0,'🟡':0,'❌':0};
  f.forEach(function(x){ c[String(x[8]).charAt(0)==='✅' ? '✅' : (String(x[8]).indexOf('🟡')===0 ? '🟡' : '❌')]++; });
  var seg2 = f.reduce(function(a,x){ return a+(+x[7]||0); }, 0), ll2 = f.reduce(function(a,x){ return a+(+x[6]||0); }, 0);
  sh.appendRow([run, '', 'RESUMEN', '✅ '+c['✅']+'  🟡 '+c['🟡']+'  ❌ '+c['❌'], '', '', ll2, Math.round(seg2), '', 'Promedio '+(f.length ? Math.round(seg2/f.length*10)/10 : 0)+' s y '+(f.length ? Math.round(ll2/f.length*10)/10 : 0)+' llamadas por frase']);
  P.deleteProperty('AZCU_PRUEBAS_I'); P.deleteProperty('AZCU_PRUEBAS_ID'); P.setProperty('AZCU_PRUEBAS', 'NO');
  f.filter(function(x){ return String(x[8]).charAt(0)!=='✅'; }).forEach(function(x){ Logger.log(x[8]+' #'+x[1]+' «'+x[2]+'» → '+x[9]+'\n   Respuesta: '+String(x[4]).replace(/\n+/g,' / ').slice(0,260)+(x[5] ? '\n   Tarjetas: '+String(x[5]).slice(0,200) : '')); });
  Logger.log('Listo ('+run+'): ✅ '+c['✅']+'  🟡 '+c['🟡']+'  ❌ '+c['❌']+'. Detalle en la pestaña Pruebas_Bot de la planilla Maestra.');
}

function azcuPruebasTokens_(){
  var ds = azcuDatos_().filter(function(p){ return p.det && p.det.calle && p.det.numero && !/^EXT-/i.test(p.id) && p.etapa!=='Otras inmobiliarias'; }).sort(function(a,b){ return a.id<b.id ? -1 : 1; }), por = {}, uni = [], dos = null, bv = null, T = {};
  ds.forEach(function(p){ var k = azcuSinAc_(p.det.calle); (por[k] = por[k] || []).push(p); });
  Object.keys(por).forEach(function(k){ if(por[k].length===1) uni.push(por[k][0]); });
  var porDir = {}; ds.forEach(function(p){ var k = azcuSinAc_(p.det.calle+' '+p.det.numero); (porDir[k] = porDir[k] || []).push(p); });
  Object.keys(porDir).forEach(function(k){ if(!dos && porDir[k].length>1) dos = porDir[k]; });
  ds.forEach(function(p){ if(!bv && /^(av|avda|bv|bvd|boulevard|pje|pasaje)\.?\s/i.test(String(p.det.calle).trim())) bv = p; });
  var paso = Math.max(1, Math.floor(uni.length/6));
  for(var i=1;i<=6;i++){ var p = uni[(i-1)*paso] || uni[uni.length-1]; if(p){ T['A'+i] = p.det.calle+' '+p.det.numero; T['C'+i] = String(p.det.calle).trim(); T['I'+i] = p.id; T['D'+i] = p.dir; var nm = String(p.prop || (p.duenos && p.duenos[0] && p.duenos[0].nombre) || '').trim().split(/\s+/); T['O'+i] = nm.length>1 ? nm[nm.length-1].replace(/[^A-Za-zÁÉÍÓÚáéíóúÑñ]/g,'') : ''; } }
  if(T.A2) T.X2 = T.C2+' '+(+T.A2.split(' ').pop()*10+7);
  if(dos){ T.D = dos[0].det.calle+' '+dos[0].det.numero; T.U1 = dos[0].dir; T.U2 = dos[1].dir; }
  if(bv){ T.BV = String(bv.det.calle).replace(/^(av|avda|bv|bvd|boulevard|pje|pasaje)\.?\s+/i,'').trim()+' '+bv.det.numero; }
  if(T.A1){ var w = T.C1.split(' ').pop(); T.TIPEO = T.A1.replace(w, w.length>=5 ? w.charAt(0)+w.charAt(1)+w.charAt(3)+w.charAt(2)+w.slice(4) : w); }
  if(T.A2) T.N2 = azcuPruebasPalabras_(+T.A2.split(' ').pop()), T.C2s = T.C2;
  return T;
}
function azcuPruebasPalabras_(n){
  var U = ['cero','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez','once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte'], D = ['','','veinti','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'], C = ['','ciento','doscientos','trescientos','cuatrocientos','quinientos','seiscientos','setecientos','ochocientos','novecientos'];
  function m(x){ var o = ''; if(x>=100){ o += (x===100 ? 'cien' : C[Math.floor(x/100)]); x = x%100; if(x) o += ' '; } if(x<=20 && x>0) o += U[x]; else if(x>20){ var d = Math.floor(x/10), u = x%10; o += d===2 ? (u ? 'veinti'+U[u] : 'veinte') : D[d]+(u ? ' y '+U[u] : ''); } return o; }
  var mil = Math.floor(n/1000), r = n%1000;
  return (mil ? (mil===1 ? 'mil' : m(mil)+' mil') : '') + (mil && r ? ' ' : '') + (r ? m(r) : '');
}
function azcuPruebasFicha_(T, k){ return 'Datos y operación\n• Dirección: '+T['D'+k]+', Rosario\n• Código: '+T['I'+k]+'\n• Tipo: Departamento\n• Etapa: Publicada\n• Origen: Propia\n\n¿Querés ver los datos técnicos o el historial de interacciones?'; }
function azcuPruebasLista_(T){
  var N = 0, L = [];
  function sub(x){ return String(x==null ? '' : x).replace(/\{([A-Za-z]\w*)\}/g, function(m,k){ return T[k]!==undefined ? T[k] : m; }); }
  function falta(x){ var m = String(x||'').match(/\{[A-Za-z]\w*\}/g) || [], k; for(var i=0;i<m.length;i++){ k = m[i].slice(1,-1); if(T[k]===undefined) return k; } return ''; }
  function add(frase, esperado, o){
    o = o || {}; N++;
    var h = (o.hist||[]).map(function(x){ return {role:x[0]==='u'?'user':'assistant', content:sub(x[1])}; });
    var q = {n:N, frase:sub(frase), esperado:esperado, hist:h, ctx:{nombre:'Prueba', fichaId:o.ficha ? (T['I'+o.ficha]||'') : ''}, chk:o, skip:o.skip||''};
    var f = falta(frase) || (o.hist||[]).map(function(x){ return falta(x[1]); }).filter(Boolean)[0] || ((o.tx && falta(o.tx)) || (o.no && falta(o.no))) || (o.ficha && !T['I'+o.ficha] ? 'I'+o.ficha : '');
    if(f && !q.skip) q.skip = 'Falta en la cartera: '+f;
    q.chk = {t:o.t, min:o.min, n:o.n, tx:sub(o.tx), no:sub(o.no), cr:sub(o.cr), max:o.max, adj:o.adj};
    L.push(q);
  }
  var FA = function(k){ return azcuPruebasFicha_(T,k); };
  add('hola!','Saludo corto y cálido, sin menú largo.',{t:[],max:260});
  add('gracias genia','Una línea, sin herramientas.',{t:[],max:180});
  add('cuánto sale {A1}?','Da el precio, sin ficha completa.',{t:[],no:'Datos y operaci',max:420});
  add('precio de {X2}','«No encuentro», ofrece parecidas, no inventa precio.',{t:[],tx:'no encuentro|no (la )?encontr|no hay|parecid',no:'\\d{2,3}\\.\\d{3}'});
  add('contame de {D}','Pregunta cuál unidad (por código, con botones).',{t:[],tx:'unidades|cu[aá]l',no:'Datos y operaci',adj:true});
  add('qué precio tienen {A3} y {A4}?','Da los dos precios, sin preguntar piso.',{t:[],no:'cu[aá]l (es|de)|qu[eé] piso|qu[eé] dpto'});
  add('{BV} qué etapa tiene?','La encuentra y dice la etapa.',{t:[],tx:'etapa|publicad|captaci|reserva|vendid|baja'});
  add('tenés algo en {C1}?','Lista las propiedades de la calle.',{t:[],tx:'{C1}'});
  add('y en {C1}?','Habla de la calle pedida, no de la ficha abierta.',{ficha:5,t:[],tx:'{C1}',no:'{C5} \\d'});
  add('y el dueño?','Dueño de la propiedad de la charla (el nombre real, nunca inventado).',{hist:[['u','contame de {A1}'],['a',FA(1)]],t:[],tx:'{O1}',no:'Datos y operaci|{C2}'});
  add('en la ficha tampoco está el tel?','Responde el teléfono (o que no está), sin otra ficha.',{hist:[['u','contame de {A1}'],['a',FA(1)]],t:[],no:'Datos y operaci|{C2}|{C3}',max:500});
  add('y la otra?','Muestra la unidad que no se mostró.',{hist:[['u','contame de {D}'],['a','Hay dos unidades en {D}: {U1} y {U2}. ¿Cuál querés ver?'],['u','la primera'],['a','Datos de {U1}: Departamento, Publicada.']],t:[],no:'a cu[aá]l te refer'});
  add('la otra','Pregunta a cuál (no hay contexto).',{t:[],tx:'cu[aá]l'});
  add('mandame la ubicación','Ubicación de la ficha abierta.',{ficha:2,t:[],adj:true});
  add('mandame la ubicación de {A1}','Ubicación de la pedida, no de la ficha abierta.',{ficha:2,t:[],no:'{C2} \\d'});
  add('visita en {A2}, Marcos Tejo, le interesó, quiere volver con la esposa','Tarjeta directa de visita con observaciones.',{t:['cargar'],cr:'esposa'});
  add('vino una pareja a ver {A3}','Arma la tarjeta (con lo que hay) o pregunta lo que falta, corto.',{tx:'nombre|qui[eé]n|resultado|c[oó]mo|detalle|tarjeta',max:300});
  add('no, 54000','Tarjeta nueva con 54.000, misma propiedad.',{hist:[['u','visité {A2} con Marcos Tejo, ofreció 50000'],['a','¡Dale! Dejé la tarjeta lista, confirmala acá abajo 👇']],min:1,no:'Datos y operaci'});
  add('dejalo, mejor no','Sin tarjeta nueva; dice que no confirme.',{hist:[['u','visité {A2} con Marcos Tejo, ofreció 50000'],['a','¡Dale! Dejé la tarjeta lista, confirmala acá abajo 👇']],t:[],tx:'confirm|sin efecto|no hago|no hace falta|no armo'});
  add('agendá visita mañana a las 15 en {A4} con Ana López','Tarjeta de agendar con la fecha de mañana.',{t:['agendar'],cr:'15:00'});
  add('el 2/11 a las 15 visita en {A4} con Ana','2 de noviembre a las 15, no número de calle.',{t:['agendar'],cr:'02/11|2/11'});
  add('fui con un cliente a ver una casa de la inmobiliaria Bertollo en Pérez Bulnes 6637, le gustó','Visita externa sin buscar en la cartera.',{t:['visitaext']});
  add('Ana Gómez hizo una oferta de 90 mil por {A2}','Tarjeta de propuesta con el monto.',{t:['propuesta'],cr:'90'});
  add('rechazaron la propuesta de {A2}','Pide el motivo antes de armar la tarjeta.',{t:[],tx:'motivo|por qu[eé]'});
  add('aceptaron la propuesta de {A2}','Pregunta de quién es la propuesta, o arma la tarjeta y aclara que pasa a Reserva.',{tx:'reserva|qui[eé]n|nombre|cliente'});
  add('pasá {A3} a publicada','Tarjeta de cambio de etapa, o avisa que ya está en Publicada.',{tx:'ya est[aá]|publicad'});
  add('bajale el precio a {A2} a 62 mil dólares','Tarjeta de editar precio con 62.000.',{t:['campos'],cr:'62'});
  add('se cayó la visita de mañana en {A4}','La ubica y arma la tarjeta de cancelar (o dice que no hay).',{tx:'visita|no (encuentro|hay|figura)'});
  add('recordame mañana a las 9 llamar a Juan','Recordatorio personal.',{skip:'Crea el recordatorio sin tarjeta: probar a mano.'});
  add('avisá que mañana no hay reunión','Pregunta a quién va.',{t:[],tx:'equipo|vos|persona|qui[eé]n'});
  add('decile a Marcos que mañana vaya a {A3} a las 10','Tarjeta de aviso a Marcos (no una visita).',{min:1,cr:'Marcos',t:['avisopersona']});
  add('qué tengo hoy?','Agenda del día.',{t:[]});
  add('cómo viene la semana?','Resumen de la semana.',{t:[]});
  add('qué papeles faltan en toda la cartera?','Lista completa, sin «y N más».',{t:[],no:'y \\d+ m[aá]s'});
  add('pasame las consultas de WhatsApp del finde','Todas, con datos y cuántas sin atender.',{t:[]});
  add('qué está atrasado?','Atrasos por vendedor.',{t:[]});
  add('cuánto es la comisión del 4% sobre 95 mil dólares?','USD 3.800, sin mandar a la calculadora.',{t:[],tx:'3\\.?800',no:'calculadora'});
  add('120 mil menos 5%','114.000',{t:[],tx:'114\\.?000'});
  add('llamé al 341-5551234 y no atendió','No hace una resta.',{no:'-5\\.5|5\\.550'});
  add('341-5551234','No responde -5.550.893.',{no:'-5\\.5|5\\.550'});
  add('mandale la encuesta al cliente de {A6}','Botones de encuesta de esa propiedad.',{t:[],tx:'encuesta'});
  add('pasame el formulario de captación para mandarle a un dueño','Link del formulario, sin pedir propiedad.',{t:[],tx:'formulario|link|enlace',no:'qu[eé] propiedad|cu[aá]l propiedad'});
  add('haceme el pdf de la ficha de {A1}','PDF de esa ficha.',{skip:'Genera un archivo en Drive: probar a mano.'});
  add('qué es mejor, crédito hipotecario UVA o tradicional?','Máximo 3 líneas, sin listas.',{t:[],max:450,no:'^\\s*1\\.'});
  add('redactame un mail para el dueño de {A2} contándole que bajamos el precio','Mail completo con el dato real.',{t:[],tx:'asunto|estimad|hola|saludos'});
  add('recuperá la visita que borré ayer','Mira la Papelera y pregunta cuál (o arma la tarjeta de restaurar).',{tx:'papelera|restaur|recuper|borrad|no (hay|encuentro|figura)'});
  add('dame info de {TIPEO}','Tolera el tipeo: da la ficha, no dice «no encuentro».',{t:[],no:'no encuentro|no (la )?encontr',tx:'Datos y operaci|precio|etapa'});
  add('che q onda la de {C2}?','Entiende la calle pedida.',{t:[],tx:'{C2}'});
  add('cargame que hoy vinieron 3 personas a ver {A3}: Ana (le gustó), Pedro (no le interesó, caro) y Luis (va a hacer una oferta)','Una tarjeta por persona.',{n:3});
  add('ehh visita en {C2} al {N2} con marcos tejo le encantó','Entiende la dirección dictada y arma la tarjeta.',{t:['cargar']});
  return L;
}
function azcuPruebasJuzgar_(q, r, seg, llam, T){
  var c = q.chk, txt = azcuSinAc_(String(r.texto||'')), tipos = (r.acciones||[]).map(function(a){ return a.tipo; }), res = (r.acciones||[]).map(function(a){ return azcuSinAc_(a.resumen||''); }).join(' | '), m = [], amar = [], ad = (r.adjuntos||[]).length;
  function re(x){ try{ return new RegExp(azcuSinAc_(x), 'im'); }catch(e){ return null; } }
  if(/^ERROR/.test(String(r.texto||''))) return {v:'❌', m:String(r.texto).slice(0,200)};
  if(!String(r.texto||'').trim() && !tipos.length && !ad) m.push('respuesta vacía');
  if(c.t){ var a = tipos.slice().sort().join(','), b = c.t.slice().sort().join(','); if(a!==b) m.push('tarjetas: esperaba ['+b+'] y salió ['+a+']'); }
  if(c.min && tipos.length<c.min) m.push('esperaba al menos '+c.min+' tarjeta(s)');
  if(c.n && tipos.length<c.n) m.push('esperaba '+c.n+' tarjetas y salieron '+tipos.length);
  if(c.tx && !(re(c.tx).test(txt) || (c.cr && false) || re(c.tx).test(res))) m.push('no aparece /'+c.tx+'/');
  if(c.cr && !re(c.cr).test(res)) m.push('la tarjeta no trae /'+c.cr+'/');
  if(c.no && re(c.no).test(txt)) m.push('aparece lo prohibido /'+c.no+'/');
  if(c.adj && !ad) m.push('faltan botones');
  if(c.max && txt.length>c.max) amar.push('larga ('+txt.length+' caracteres)');
  if(seg>8) amar.push('lenta ('+seg+' s)');
  if(llam>4) amar.push(llam+' llamadas al modelo');
  if(m.length) return {v:'❌', m:m.join('; ')};
  if(amar.length) return {v:'🟡', m:amar.join('; ')};
  return {v:'✅', m:''};
}
