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
const F=(o)=>{const x={};for(const d in o)x[d]={'1':{marcas:o[d].map(H),nombre:'PEREZ Juan'}};return x};
// 1) extra 14/19 sobre 7/14, fichó 06:58-19:10 => 5h al 50
let rows=sim(F({5:['6:58','19:10']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}));
assert.strictEqual(rows.length,1);assert.deepStrictEqual([rows[0][9],rows[0][10],rows[0][11]],[5,0,'AUTORIZADA']);
// 2) extra declarada pero se fue a las 14:05 => no se paga
rows=sim(F({5:['6:58','14:05']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}));
assert.strictEqual(rows[0][11],'A REVISAR');assert.strictEqual(rows[0][9],0);
// 3) realizó menos: declarada 14/19, salió 17:20 => 3h (hachazo 3.33->3)
rows=sim(F({5:['6:58','17:20']}),ag({hab:{5:'7/14'},ext:{5:'14/19'}}));
assert.deepStrictEqual([rows[0][9],rows[0][11]],[3,'AUTORIZADA']);assert.ok(/menos/.test(rows[0][12]));
// 4) extra 0/6 tras turno 18/24 del día anterior (salida 06:02 en pestaña del día 6)
rows=sim(F({5:['17:55','23:58'],6:['6:02']}),ag({hab:{5:'18/24'},ext:{6:'0/6'}}));
assert.strictEqual(rows.length,1);assert.deepStrictEqual([rows[0][9],rows[0][10]],[0,6]);
// 5) sin declarar, fichó de más: 7/14 sale 17:05 => A REVISAR no autorizada 3h
rows=sim(F({6:['6:58','17:05']}),ag({hab:{6:'7/14'}}));
assert.strictEqual(rows.length,1);assert.ok(/NO autorizada/.test(rows[0][12]));assert.strictEqual(rows[0][9],3);assert.deepStrictEqual([rows[0][5],rows[0][6]],['06:58','17:05']);
// 6) turno noche 23/6 sin extras, salida normal => nada; y feriado no genera
rows=sim(F({5:['22:55'],6:['6:03']}),ag({hab:{5:'23/6'}}));assert.strictEqual(rows.length,0);
rows=sim(F({12:['6:58','17:05']}),ag({hab:{12:'7/14'}}));assert.strictEqual(rows.length,0);
// 7) 18/24 + cola 0/6 no da falso positivo
rows=sim(F({5:['17:58'],6:['6:02','17:55'],7:['6:01']}),ag({hab:{5:'18/24',6:'0/6'}}));assert.strictEqual(rows.length,0);
// 8) identidad
const f2=F({5:['6:58','19:10']});f2[5]['1'].nombre='ALONSO Cecilia';
rows=sim(f2,ag({hab:{5:'7/14'},ext:{5:'14/19'}}));assert.strictEqual(rows[0][11],'REVISAR IDENTIDAD');
console.log('OK todos los tests');
