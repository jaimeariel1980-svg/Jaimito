const fs=require('fs'),vm=require('vm'),assert=require('assert');
const dir=__dirname+'/..';
const ctx={Logger:{log(){}},SpreadsheetApp:{}};vm.createContext(ctx);
['Nucleo.gs','Config.gs','Extras.gs'].forEach(f=>vm.runInContext(fs.readFileSync(dir+'/'+f,'utf8'),ctx,{filename:f}));
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
c=run('consolidarGRH_(5,0,0,0,false,30)');assert.strictEqual(c.alerta,'SIN aut GRH');
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
// 9) 4 fichadas con extra hecha: 6:00-13:00 y 13:30-17:30 => extra 13/17 cuenta 13:30-17:00 = 3.5h
r=R(sim(F({5:['6:00','13:00','13:30','17:30']}),ag({hab:{5:'6/13'},ext:{5:'13/17'}}))[0]);
assert.deepStrictEqual([r.est,r.h50,r.fr],['AUTORIZADA',3.5,'13:30 a 17:00']);
// 10) RABINOVICH: 12/19 + extra 7/12, fichó 6:52 7:42 12:11 19:00 => solo estuvo 7:00-7:42 => <1h, no se paga
r=R(sim(F({8:['19:15'],9:['6:52','7:42','12:11','19:00']}),ag({hab:{9:'12/19'},ext:{9:'7/12'}}))[0]);
assert.deepStrictEqual([r.est,r.h50,r.fr],['A REVISAR',0,'07:00 a 07:42']);assert.strictEqual(r.fich,'[6:52 · 7:42 · 12:11 · 19:00]');
// 11) 3 fichadas (impar) => no se paga solo, se muestra cuánto sería
r=R(sim(F({5:['6:00','13:00','17:10']}),ag({hab:{5:'6/13'},ext:{5:'13/17'}}))[0]);
assert.deepStrictEqual([r.est,r.h50],['A REVISAR',0]);assert.ok(/3 fichadas/.test(r.cond));
// 12) SARBA: extra dentro del propio turno => no es extra
r=R(sim(F({1:['11:57','22:12']}),ag({hab:{1:'17/24'},ext:{1:'19/22'}}))[0]);assert.ok(/dentro del turno/.test(r.cond));
// 13) PAY/NIETO: 0/6 que es la cola de un 23/6 del día anterior => no es extra; si el día anterior era franco, se paga una vez
r=R(sim(F({9:['23:53'],10:['6:05']}),ag({hab:{9:'23/6'},ext:{10:'0/6'}}))[0]);assert.ok(/dentro del turno/.test(r.cond));
r=R(sim(F({9:['23:53'],10:['6:05']}),ag({hab:{9:'F'},ext:{10:'0/6'}}))[0]);assert.deepStrictEqual([r.est,r.h100,r.ent,r.sal],['AUTORIZADA',6,'23:53(-1)','6:05']);
// 14) NARVAEZ: 0/6 sin presencia (13:05 23:59 del día anterior y 13:10) => no se paga
r=R(sim(F({13:['13:05','23:59'],14:['13:10']}),ag({hab:{13:'13/24',14:'6/13'},ext:{14:'0/6'}}))[0]);assert.deepStrictEqual([r.est,r.h50,r.h100],['A REVISAR',0,0]);
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
console.log('OK todos los tests');
