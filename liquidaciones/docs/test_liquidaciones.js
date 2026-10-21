const fs=require('fs'),vm=require('vm'),assert=require('assert');
const dir=__dirname+'/..';
const ctx={Logger:{log(){}},SpreadsheetApp:{}};vm.createContext(ctx);
['Nucleo.gs','Config.gs','Extras.gs','Nocturnidad.gs','Feriados.gs'].forEach(f=>vm.runInContext(fs.readFileSync(dir+'/'+f,'utf8'),ctx,{filename:f}));
const run=c=>{const v=vm.runInContext('JSON.stringify('+c+')',ctx);return v===undefined?undefined:JSON.parse(v)};
const H=(s)=>{const m=s.split(':');return +m[0]*60+(+m[1]||0)};
// parse
assert.deepStrictEqual(run('parseTurno("23/06")'),{ini:1380,fin:1800});
assert.deepStrictEqual(run('parseTurno("13/17:30")'),{ini:780,fin:1050});
assert.deepStrictEqual(run('parseTurno("17/24")'),{ini:1020,fin:1440});
assert.strictEqual(run('parseTurno("ACA")'),null);
assert.strictEqual(run('parseFranjas("13/17 y 19/23").length'),2);
// calcularExtra: octubre 2026: lun 5, sab 10, dom 11, lun 12 feriado
const mk=(dia,fer)=>`({anio:2026,mes:10,dia:${dia},feriados:${fer||'{}'},noLab:{}})`;
let r=run(`calcularExtra(300,13*60,${mk(5)})`);assert.deepStrictEqual([r.h50,r.h100],[5,0]);          // lun 13-18 => 50
r=run(`calcularExtra(240,19*60,${mk(5)})`);assert.deepStrictEqual([r.h50,r.h100],[3,1]);              // lun 19-23: 3 al 50, 1 al 100 (22-23)
r=run(`calcularExtra(360,0,${mk(5)})`);assert.deepStrictEqual([r.h50,r.h100],[0,6]);                  // 0-6 => 100
r=run(`calcularExtra(300,10*60,${mk(10)})`);assert.deepStrictEqual([r.h50,r.h100],[3,2]);              // sab 10-15: 10-13 50, 13-15 100
r=run(`calcularExtra(240,8*60,${mk(11)})`);assert.deepStrictEqual([r.h50,r.h100],[0,4]);              // dom 100
r=run(`calcularExtra(240,8*60,${mk(12,'{"10-12":1}')})`);assert.deepStrictEqual([r.h50,r.h100],[0,4]); // feriado 100
r=run(`calcularExtra(50,13*60,${mk(5)})`);assert.deepStrictEqual([r.h50,r.h100],[0,0]);                // <1h
r=run(`calcularExtra(100,13*60,${mk(5)})`);assert.deepStrictEqual([r.h50,r.h100],[1.5,0]);             // hachazo 90
r=run(`calcularExtra(360,18*60,${mk(10)})`);assert.deepStrictEqual([r.h50,r.h100],[0,6]);              // sab 18-24 => 100
r=run(`calcularExtra(300,20*60,${mk(9)})`);assert.deepStrictEqual([r.h50,r.h100],[2,3]);               // vie 20-01: 2 al 50, 3 al 100
// identidad
assert.ok(run('apellidoCoincide("YACUZZI","YACCUZZI Maria")'));
assert.ok(run('apellidoCoincide("DELLABARCA","BARCA DELLA Juan")'));
assert.ok(!run('apellidoCoincide("LOMBARDO","ALONSO Cecilia")'));
// consolidado GRH
let c=run('consolidarGRH_(12,2,10,6,true,30)');assert.strictEqual(c.alerta,'reconversión 2h');
c=run('consolidarGRH_(20,16,10,6,true,30)');assert.strictEqual(c.pagar,30);assert.ok(/EXCEDE/.test(c.alerta));
c=run('consolidarGRH_(5,0,0,0,false,30)');assert.strictEqual(c.alerta,'NO FIGURA EN EL EXPEDIENTE (5 h sin autorización)');
// procesarAgente_ con fichadas simuladas
const base={idFichadas:'x'};
const sim=(fich,a)=>{for(let i=1;i<=31;i++)fich[i]=fich[i]||{};ctx._f=fich;ctx._a=a;return run(`(function(){_fich=_f;_diasMes=31;var out=[];var c={anio:2026,mes:10,feriados:{"10-12":1},noLab:{}};procesarAgente_(_a,"GUARDIA",{},c,out);return out;})()`)};
const ag=(o)=>Object.assign({leg:'1',ape:'PEREZ',hab:{},cam:{},ext:{},lic:{}},o);
const F=(o)=>{const x={};for(const d in o)x[d]={'1':{marcas:o[d].map(H),txt:o[d].slice(),nombre:'PEREZ Juan'}};return x};
const R=(r)=>({ent:r[5],sal:r[6],fr:r[8],h50:r[9],h100:r[10],est:r[11],cond:r[12],fich:r[13]});
r=null;
// 1) extra 14/19 sobre 7/14, fichó 6:58:12-19:10:40 => 5h al 50; entrada/salida tal cual la tarjeta
r=R(sim(F({5:['6:58:12','19:10:40']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0]);
assert.deepStrictEqual([r.ent,r.sal,r.h50,r.h100,r.est,r.fich],['6:58:12','19:10:40',5,0,'AUTORIZADA','[6:58:12 · 19:10:40]']);
// 2) extra declarada pero se fue a las 14:05 => no se paga
r=R(sim(F({5:['6:58','14:05']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0]);assert.deepStrictEqual([r.est,r.h50],['A REVISAR',0]);
// 3) realizó menos: declarada 14/19, salió 17:20 => 3h
r=R(sim(F({5:['6:58','17:20']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0]);assert.deepStrictEqual([r.h50,r.est],[3,'AUTORIZADA']);assert.ok(/menos/.test(r.cond));
// 4) extra 0/6 tras turno 18/24 del día anterior, sin marcar a medianoche
let rows=sim(F({5:['17:55'],6:['6:02']}),ag({hab:{5:'18/24'},ext:{6:'0/6'}}));
assert.strictEqual(rows.length,1);r=R(rows[0]);assert.deepStrictEqual([r.h50,r.h100,r.ent,r.sal],[0,6,'17:55(-1)','6:02']);assert.ok(/17:55\(-1\)/.test(r.fich));
// 5) sin declarar, fichó de más: 7/14 sale 17:05 => A REVISAR no autorizada 3h
rows=sim(F({6:['6:58','17:05']}),ag({hab:{6:'7/14'}}));
assert.strictEqual(rows.length,1);r=R(rows[0]);assert.ok(/NO autorizada/.test(r.cond));assert.deepStrictEqual([r.h50,r.ent,r.sal],[3,'6:58','17:05']);
// 6) turno noche normal => nada; feriado no genera
assert.strictEqual(sim(F({5:['22:55'],6:['6:03']}),ag({hab:{5:'23/6'}})).length,0);
assert.strictEqual(sim(F({12:['6:58','17:05']}),ag({hab:{12:'7/14'}})).length,0);
// 7) 18/24 + cola 0/6 no da falso positivo
assert.strictEqual(sim(F({5:['17:58'],6:['6:02','17:55'],7:['6:01']}),ag({hab:{5:'18/24',6:'0/6'}})).length,0);
// 8) OMISIÓN: una sola fichada => nunca hay extra
r=R(sim(F({5:['6:58']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0]);assert.deepStrictEqual([r.ent,r.sal,r.est,r.h50,r.fich],['6:58','','A REVISAR',0,'[6:58]']);assert.ok(/omisión/.test(r.cond));
assert.strictEqual(sim(F({6:['6:58']}),ag({hab:{6:'7/14'}})).length,0);
// 9) 4 fichadas: las intermedias no cuentan, jornada completa 6:00-17:30 => extra 13/17 = 4h
r=R(sim(F({5:['6:00','13:00','13:30','17:30']}),ag({hab:{5:'6/13'},ext:{5:'13/17'}}))[0]);
assert.deepStrictEqual([r.est,r.h50,r.fr],['AUTORIZADA',4,'13:00 a 17:00']);assert.ok(/intermedia/.test(r.cond));
// 10) RABINOVICH: 12/19 + extra 7/12, fichó 6:52 7:42 12:11 19:00 => intermedias no cuentan => 5h
r=R(sim(F({8:['19:15'],9:['6:52:52','7:42:41','12:11:37','19:00:10']}),ag({hab:{9:'12/19'},ext:{9:'7/12'}}))[0]);
assert.deepStrictEqual([r.est,r.h50,r.fr,r.ent,r.sal],['AUTORIZADA',5,'07:00 a 12:00','6:52:52','19:00:10']);assert.strictEqual(r.fich,'[6:52:52 · 7:42:41 · 12:11:37 · 19:00:10]');
// 11) 3 fichadas (una en el medio) => se paga igual
r=R(sim(F({5:['6:00','13:00','17:10']}),ag({hab:{5:'6/13'},ext:{5:'13/17'}}))[0]);
assert.deepStrictEqual([r.est,r.h50],['AUTORIZADA',4]);
// 12) SARBA: extra dentro del propio turno => no es extra
r=R(sim(F({1:['11:57','22:12']}),ag({hab:{1:'17/24'},ext:{1:'19/22'}}))[0]);assert.ok(/dentro del turno/.test(r.cond));
// 13) PAY/NIETO: 0/6 que es la cola de un 23/6 del día anterior => no es extra; si el día anterior era franco, se paga una vez
r=R(sim(F({9:['23:53'],10:['6:05']}),ag({hab:{9:'23/6'},ext:{10:'0/6'}}))[0]);assert.ok(/dentro del turno/.test(r.cond));
r=R(sim(F({9:['23:53'],10:['6:05']}),ag({hab:{9:'F'},ext:{10:'0/6'}}))[0]);assert.deepStrictEqual([r.est,r.h100,r.ent,r.sal],['AUTORIZADA',6,'23:53(-1)','6:05']);
// 14) NARVAEZ: 13/24 + extra 0/6 + 6/13, fichó 13:05(-1) 23:59(-1) 13:10 => jornada completa, 23:59 es intermedia => 6h al 100
r=R(sim(F({13:['13:05','23:59'],14:['13:10']}),ag({hab:{13:'13/24',14:'6/13'},ext:{14:'0/6'}}))[0]);assert.deepStrictEqual([r.est,r.h100,r.ent,r.sal],['AUTORIZADA',6,'13:05(-1)','13:10']);
// 15) Nieva: 6/13 sin extra declarada, 5:51:03-17:10:47 => entrada y salida como en la tarjeta
r=R(sim(F({13:['5:51:03','17:10:47']}),ag({hab:{13:'6/13'}}))[0]);assert.deepStrictEqual([r.ent,r.sal,r.fich,r.h50],['5:51:03','17:10:47','[5:51:03 · 17:10:47]',4]);
// 17) CANALES: 23/6 con extra 19/23 y licencia vieja "13 LAO al 11/09/2026" en la planilla; la salida está en la pestaña del día siguiente
rows=sim(F({8:['6:05:05','18:48:58'],9:['6:02:30','18:37:44'],10:['6:00:15']}),ag({hab:{8:'23/6',9:'23/6',10:'F'},ext:{8:'19/23',9:'19/23'},lic:{8:'13 LAO al 11/09/2026',9:'13 LAO al 11/09/2026'}}));
r=R(rows[0]);assert.deepStrictEqual([r.est,r.h50,r.h100,r.ent,r.sal],['AUTORIZADA',3,1,'18:48:58','6:02:30(+1)']);
r=R(rows[1]);assert.deepStrictEqual([r.est,r.h50,r.h100,r.ent,r.sal],['AUTORIZADA',3,1,'18:37:44','6:00:15(+1)']);
// misma extra sin turno conocido (licencia vigente): igual encuentra la salida del día siguiente
r=R(sim(F({8:['18:48:58'],9:['6:02:30']}),ag({hab:{8:'23/6'},ext:{8:'19/23'},lic:{8:'LAO al 30/10/2026'}}))[0]);assert.deepStrictEqual([r.est,r.h50,r.h100],['AUTORIZADA',3,1]);
// extra en franco 19/23, al otro día entra 5:50 a su 6/13: no se mezcla esa entrada
r=R(sim(F({5:['18:55','23:05'],6:['5:50','13:02']}),ag({hab:{5:'F',6:'6/13'},ext:{5:'19/23'}}))[0]);assert.deepStrictEqual([r.est,r.h50,r.h100],['AUTORIZADA',3,1]);
// 16) identidad
const f2=F({5:['6:58','19:10']});f2[5]['1'].nombre='ALONSO Cecilia';
assert.strictEqual(R(sim(f2,ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0]).est,'REVISAR IDENTIDAD');
// 18) orden y colores: autorizada primero (blanco), luego cada tipo
assert.strictEqual(run("tipoFila_(['','','','','','','','','',1,0,'AUTORIZADA',''])"),0);
assert.strictEqual(run("tipoFila_(['','','','','','','','','',0,0,'A REVISAR','omisión: una sola fichada'])"),2);
assert.strictEqual(run("tipoFila_(['','','','','','','','','',4,0,'A REVISAR','extra NO autorizada (no hay franja declarada)'])"),1);
// 19) horario de la tarjeta vs turno de la planilla
const fh=F({1:['11:57:56','22:12:41']});fh[1]['1'].horario='12:00 a 19:00';
r=sim(fh,ag({hab:{1:'17/24'},ext:{1:'19/22'}}))[0];assert.strictEqual(r[14],'12:00 a 19:00 ≠ planilla 17/24');
const fh2=F({5:['6:58','19:10']});fh2[5]['1'].horario='07:00 a 14:00';
assert.strictEqual(sim(fh2,ag({hab:{5:'7/14'},ext:{5:'14/19'}}))[0][14],'07:00 a 14:00');
// 20) comparación de fotos de extras
const dif=run(`compararFotos_({"2 piso|1|5":"13/17","2 piso|1|6":"19/23"},{"2 piso|1|5":"13/19","2 piso|1|7":"7/12"})`);
assert.deepStrictEqual(dif.map(x=>x.tipo).sort(),['Agregada','Modificada','Quitada']);
// 21) Excel GRH: por encabezado y con columnas fijas
let g=run(`parsearGRH_([['LEGAJO','Apellido','AREA','x','x','x','x','HS 50%','HS 100%'],['62319','NIEVA','Enfermería','','','','',10,6],['1','X','Médicos','','','','',5,5]])`);
assert.deepStrictEqual(g,{'62319':{a50:10,a100:6,ape:'NIEVA'}});
g=run(`parsearGRH_([['Planilla'],['NRO','Nombre','Área','HS 100','HS 50'],[]])`);assert.deepStrictEqual(g,{});
g=run(`parsearGRH_([['División Enfermería'],['Legajo','Agente','Área','','','','','Hs 50 %','Hs 100 %'],[57696,'CANALES','ENFERMERIA','','','','',12,0]])`);
assert.deepStrictEqual(g,{'57696':{a50:12,a100:0,ape:'CANALES'}});
// 22) consolidado mensual: Burgos 15,5 al 50 con GRH 10/6 => reconversión 5,5, sin endoso
c=run('consolidarGRH_(15.5,0,10,6,true,30)');assert.deepStrictEqual([c.total,c.pagar,c.reconv,c.endoso,c.alerta],[15.5,15.5,5.5,0,'reconversión 5.5h']);
assert.strictEqual(run("colorAlerta_('reconversión 5.5h')"),'#fff3b8');
// 23) NOCTURNIDAD
assert.strictEqual(run('minNoche_(1380,1800)'),420);assert.strictEqual(run('minNoche_(1080,1800)'),540);assert.strictEqual(run('minNoche_(360,780)'),0);
const noc=(fich,a)=>{for(let i=1;i<=31;i++)fich[i]=fich[i]||{};ctx._f=fich;ctx._a=a;return run(`(function(){_fich=_f;_diasMes=31;_mes=10;_anio=2026;var f=[],g={};procesarNocheAgente_(_a,"GUARDIA",{},f,g);return {f:f,g:g};})()`)};
// 23/6 entró 22:55 salió 6:03(+1) => 7 h (desde las 23)
let n=noc(F({5:['22:55:10'],6:['6:03:20']}),ag({hab:{5:'23/6'}}));assert.deepStrictEqual([n.f[0][9],n.f[0][5],n.f[0][6]],[7,'22:55:10','6:03:20(+1)']);
// 18/06 => 9 h nocturnas (21 a 6)
n=noc(F({5:['17:55'],6:['6:05']}),ag({hab:{5:'18/06'}}));assert.strictEqual(n.f[0][9],9);
// 17/24 => 3 h (21 a 24)
n=noc(F({5:['16:58','0:02']}),ag({hab:{5:'17/24'}}));assert.strictEqual(n.f.length,1);
n=noc(F({5:['16:58'],6:['0:02']}),ag({hab:{5:'17/24'}}));assert.strictEqual(n.f[0][9],3);
// turno de día => nada
assert.strictEqual(noc(F({5:['5:55','13:05']}),ag({hab:{5:'6/13'}})).f.length,0);
// extra 19/23 antes del 23/6: la extra no cuenta, solo 23 a 6 => 7 h
n=noc(F({5:['18:59'],6:['6:04']}),ag({hab:{5:'23/6'},ext:{5:'19/23'}}));assert.deepStrictEqual([n.f[0][9],n.f[0][5]],[7,'18:59']);
// llegó tarde 0:30 => 5,5 h
n=noc(F({6:['0:30','6:02']}),ag({hab:{5:'23/6'}}));assert.strictEqual(n.f[0][9],5.5);
// omisión: solo entrada => no se paga
n=noc(F({5:['22:55']}),ag({hab:{5:'23/6'}}));assert.ok(/OMISIÓN/.test(n.f[0][10]));assert.deepStrictEqual(n.g['1'].omis,[5]);assert.strictEqual(n.g['1'].min,0);
// omisión: solo salida
n=noc(F({6:['6:03']}),ag({hab:{5:'23/6'}}));assert.ok(/falta la entrada/.test(n.f[0][10]));
// no vino (sin fichadas) => se ignora
assert.strictEqual(noc(F({}),ag({hab:{5:'23/6'}})).f.length,0);
// 18/24 + cola 0/6 => una sola jornada de 9 h
n=noc(F({5:['17:58'],6:['6:02']}),ag({hab:{5:'18/24',6:'0/6'}}));assert.strictEqual(n.f.length,1);assert.strictEqual(n.f[0][9],9);
// dos noches seguidas 23/6
n=noc(F({5:['22:55'],6:['6:03','22:50'],7:['6:01']}),ag({hab:{5:'23/6',6:'23/6'}}));assert.deepStrictEqual(n.f.map(r=>r[9]),[7,7]);assert.strictEqual(n.g['1'].min,840);
// Butori: programado 23/6, fichó 17:00-0:00 sin cambio cargado => 3 h nocturnas (21-24) y aviso
n=noc(F({5:['17:00:12'],6:['0:00:40']}),ag({hab:{5:'23/6'}}));assert.strictEqual(n.f[0][9],3);assert.ok(/HORARIO DISTINTO/.test(n.f[0][10]));assert.deepStrictEqual(n.g['1'].dist,[5]);
// con el cambio cargado (17/24) no hay aviso
n=noc(F({5:['17:00'],6:['0:00']}),ag({hab:{5:'23/6'},cam:{5:'17/24'}}));assert.strictEqual(n.f[0][9],3);assert.ok(!/DISTINTO/.test(n.f[0][10]));
// extra declarada antes del turno: no es horario distinto
n=noc(F({5:['18:59'],6:['6:04']}),ag({hab:{5:'23/6'},ext:{5:'19/23'}}));assert.ok(!/DISTINTO/.test(n.f[0][10]));
// 24) FERIADOS (12/10 feriado)
const fer=(fich,a)=>{for(let i=1;i<=31;i++)fich[i]=fich[i]||{};ctx._f=fich;ctx._a=a;return run(`(function(){_fich=_f;_diasMes=31;_mes=10;_anio=2026;var c={anio:2026,mes:10,feriados:{"10-12":1},noLab:{}},f=[],g={};feriadoAgente_(_a,"GUARDIA",{},c,f,g);return {f:f,g:g};})()`)};
// 6/13 el feriado, fichó 5:55-13:05 => 7 h (dentro del turno)
let q=fer(F({12:['5:55','13:05']}),ag({hab:{12:'6/13'}}));assert.deepStrictEqual([q.f.length,q.f[0][9]],[1,7]);
// 6/13 + extra 13/17 el feriado: feriado = 7 h del turno (la extra va por extras)
q=fer(F({12:['5:55','17:05']}),ag({hab:{12:'6/13'},ext:{12:'13/17'}}));assert.strictEqual(q.f[0][9],7);
// 23/6 del 11 (víspera): cuenta 0 a 6 del 12 => 6 h; 23/6 del 12: 23 a 24 => 1 h
q=fer(F({11:['22:55'],12:['6:03','22:50'],13:['6:01']}),ag({hab:{11:'23/6',12:'23/6'}}));assert.deepStrictEqual(q.f.map(r=>r[9]),[6,1]);assert.strictEqual(q.g['1'].min,420);
// FDO pero vino a trabajar 7-14 => cobra feriado 7 h
q=fer(F({12:['6:58','14:02']}),ag({hab:{12:'FDO'}}));assert.strictEqual(q.f[0][9],7.1);assert.ok(/franco por feriado/.test(q.f[0][10]));
// omisión el feriado => no se paga
q=fer(F({12:['5:55']}),ag({hab:{12:'6/13'}}));assert.ok(/OMISIÓN/.test(q.f[0][10]));assert.strictEqual(q.g['1'].min,0);
// no vino => nada
assert.strictEqual(fer(F({}),ag({hab:{12:'6/13'}})).f.length,0);
// 25) NO LABORABLE (sábado 10/10 no; usamos lunes 5 como no laborable): jornada completa como extra 50/100
ctx._f=F({5:['5:58','13:03']});for(let i=1;i<=31;i++)ctx._f[i]=ctx._f[i]||{};ctx._a=ag({hab:{5:'6/13'}});
let nl=run(`(function(){_fich=_f;_diasMes=31;_mes=10;_anio=2026;var out=[];procesarAgente_(_a,"GUARDIA",{},{anio:2026,mes:10,feriados:{},noLab:{"10-5":1}},out);return out;})()`);
assert.strictEqual(nl.length,1);assert.deepStrictEqual([nl[0][11],nl[0][9],nl[0][10]],['AUTORIZADA',7,0]);assert.ok(/no laborable/.test(nl[0][12]));
// 26) SADOFE: no cobra feriado; en día de semana lo fichado es extra (a revisar si no está declarada)
q=fer(F({12:['5:55','18:05']}),ag({hab:{3:'6/18',4:'6/18',10:'6/18',11:'6/18',12:'6/18'}}));assert.strictEqual(q.f.length,0);
ctx._f=F({7:['6:58','14:05']});for(let i=1;i<=31;i++)ctx._f[i]=ctx._f[i]||{};ctx._a=ag({hab:{3:'6/18',4:'6/18',10:'6/18',11:'6/18'}});
let sd=run(`(function(){_fich=_f;_diasMes=31;var out=[];procesarAgente_(_a,"GUARDIA",{},{anio:2026,mes:10,feriados:{"10-12":1},noLab:{}},out);return out;})()`);
assert.strictEqual(sd.length,1);assert.ok(/SADOFE/.test(sd[0][12]));assert.strictEqual(sd[0][9],7);
// no laborables fijos 2026
assert.deepStrictEqual(run('noLabFijos_(2026,11)'),['11-13']);assert.deepStrictEqual(run('noLabFijos_(2026,12)'),['12-7']);assert.deepStrictEqual(run('noLabFijos_(2026,10)'),[]);
console.log('OK todos los tests');
