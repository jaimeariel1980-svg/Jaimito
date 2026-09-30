function testAgenda(){
  guardarAgenda({codigo:'TEST', fecha:'01/01/2030', hora:'10:00', interesado:'Prueba', telefono:'3410000000', vendedor:'Test', nota:'borrar'});
  Logger.log('Fila agregada en Agenda OK');
}
