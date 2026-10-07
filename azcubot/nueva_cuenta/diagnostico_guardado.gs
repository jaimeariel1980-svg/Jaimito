function diagnosticoGuardado(){
  var mae=hojaMaestra_(), head=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0], falt=[];
  Object.keys(MAP_CAMPOS).forEach(function(k){ if(head.indexOf(MAP_CAMPOS[k])<0) falt.push(MAP_CAMPOS[k]); });
  Logger.log('Hoja: '+mae.getName()+' · Planilla: '+mae.getParent().getName()+' · '+mae.getLastColumn()+' columnas');
  Logger.log(falt.length ? '❌ Faltan columnas: '+falt.join(' | ') : '✅ Están todas las columnas');
}
