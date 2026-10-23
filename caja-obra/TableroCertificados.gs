// ===== TABLERO CAJA DE OBRA: certificados de honorarios =====
// Usan la misma lógica que Anita (certEmitir_ / certCorregir_ en Código.gs).

function tbCertCarpeta_(){
  var it=DriveApp.getFoldersByName('Certificados NJ Arquitectura');
  return it.hasNext()?it.next():DriveApp.createFolder('Certificados NJ Arquitectura');
}
function tbGuardarPdf_(blob){
  var f=tbCertCarpeta_().createFile(blob.setName(blob.getName().replace(/\.pdf$/,'')+'_'+Utilities.formatDate(new Date(),'GMT-3','yyyyMMdd-HHmm')+'.pdf'));
  try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); }catch(x){}
  return f.getUrl();
}

function tbCertEmitir(clave,obra,pct,mes,nEsperado){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||'');
    if(obrasLista().indexOf(obra)<0) throw new Error('Esa obra no existe.');
    var r=certEmitir_(obra,Number(pct)||0,tbTxt_(mes,40),Number(nEsperado)||0);
    var url=''; try{ url=tbGuardarPdf_(r.pdf); }catch(e){}
    SpreadsheetApp.flush();
    return {datos:tableroDatos_(),cert:{n:r.n,mes:r.etiqueta,total:r.total,acum:r.acum,reaj:r.reaj,url:url}};
  });
}

function tbCertCorregir(clave,obra,mes,pct){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||''); pct=Number(pct)||0;
    if(obrasLista().indexOf(obra)<0) throw new Error('Esa obra no existe.');
    if(!(pct>0&&pct<1)) throw new Error('Ese % no parece un IPC mensual válido (0 a 100%).');
    certCorregir_(obra,String(mes||''),pct);
    SpreadsheetApp.flush();
    return tableroDatos_();
  });
}

function tbCertPdf(clave,obra,n){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||''); n=Number(n)||0;
    var c=ss().getSheetByName('Certificados'), last=c.getLastRow();
    for(var r=18;r<=last;r++){
      var v=c.getRange(r,1,1,8).getValues()[0];
      if(v[0]===obra&&Number(v[1])===n){
        var fecha=(v[7] instanceof Date)?v[7]:new Date();
        var pdf=certPDF(certParams(obra),n,mesLindo(v[2]),Number(v[3])||0,Number(v[4])||0,Number(v[5])||0,Number(v[6])||0,fecha);
        return {url:tbGuardarPdf_(pdf)};
      }
    }
    throw new Error('No encontré ese certificado.');
  });
}

function tbCertParams(clave,obra,mensual,plazo,baseISO){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||''); mensual=Math.round(Number(mensual)||0); plazo=Math.round(Number(plazo)||0);
    if(obrasLista().indexOf(obra)<0) throw new Error('Esa obra no existe.');
    if(!(mensual>0)) throw new Error('El honorario mensual tiene que ser mayor a 0.');
    if(!(plazo>=1&&plazo<=120)) throw new Error('El plazo va de 1 a 120 meses.');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(String(baseISO||''))) throw new Error('Poné la fecha base.');
    var c=ss().getSheetByName('Certificados'), v=c.getRange(6,1,10,4).getValues(), f=0;
    for(var i=0;i<v.length;i++){ if(v[i][0]===obra){ f=6+i; break; } }
    if(!f){ for(var j=0;j<v.length;j++){ if(!v[j][0]){ f=6+j; break; } } }
    if(!f) throw new Error('No hay lugar en la tabla de parámetros de Certificados.');
    c.getRange(f,1,1,4).setValues([[obra,mensual,plazo,Utilities.parseDate(baseISO+' 12:00','GMT-3','yyyy-MM-dd HH:mm')]]);
    c.getRange(f,2).setNumberFormat('"$"#,##0');
    c.getRange(f,4).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();
    return tableroDatos_();
  });
}

function tbCertBorrar(clave,obra,n){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||''); n=Number(n)||0;
    var s=ss(), c=s.getSheetByName('Certificados'), last=c.getLastRow(), fila=0, maxN=0;
    for(var r=18;r<=last;r++){
      if(c.getRange(r,1).getValue()===obra){ var k=Number(c.getRange(r,2).getValue())||0; if(k>maxN) maxN=k; if(k===n) fila=r; }
    }
    if(!fila) throw new Error('No encontré ese certificado.');
    if(n!==maxN) throw new Error('Solo se puede borrar el último certificado de la obra.');
    var v=c.getRange(fila,1,1,8).getValues()[0];
    var pp=s.getSheetByName('Papelera certificados');
    if(!pp){
      pp=s.insertSheet('Papelera certificados');
      pp.getRange(1,1,1,10).setValues([['Borrado','Por','Obra','N°','Mes','% IPC','Inflación acumulada','Reajuste','Honorarios del mes','Fecha']]).setFontWeight('bold');
      pp.setFrozenRows(1);
    }
    pp.appendRow([new Date(),'Tablero'].concat(v));
    c.deleteRow(fila);
    var mv=s.getSheetByName('Movimientos'), vv=mv.getDataRange().getValues(), etiqueta=String(mesLindo(v[2])).toLowerCase();
    for(var j=vv.length-1;j>=1;j--){
      if(vv[j][1]==='Ingreso'&&vv[j][2]==='Honorarios'&&vv[j][10]===obra&&String(vv[j][3]).toLowerCase().indexOf(etiqueta)>=0&&Math.round(Number(vv[j][5])||0)===Math.round(Number(v[6])||0)){
        var pm=s.getSheetByName('Papelera');
        if(pm) pm.appendRow([new Date(),'Tablero'].concat(vv[j].slice(0,13)));
        mv.deleteRow(j+1); break;
      }
    }
    SpreadsheetApp.flush();
    return tableroDatos_();
  });
}
