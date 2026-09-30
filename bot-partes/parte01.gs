var MAESTRO_ID = '1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';

var CHAT_ID_ALERTAS = '5692711315';

var TZ = 'America/Argentina/Buenos_Aires';

var MIN_ANTES = 30;

var DIAS_PAPELES = 10;

var HOJA = 'Ventas - Base de Datos Maestra';

var C = { COD:0, ACT:2, DUENO:3, TEL:9, MAIL:10, CALLE:43, NUM:44, DOCS:69, PRECIO:108, MONEDA:109, ETAPA:111, TOKKO:119, COMPL:67 };

var DOCS = ['Plano edificación','Plano mensura','Escritura','DNI propietario','CUIL propietario','Título propiedad','TGI Municipal','API Provincial','EPE Luz','Aguas Santafesinas','Litoral Gas'];


function guardarClaves(){
  var p = PropertiesService.getScriptProperties();
  p.setProperty('TG_TOKEN', 'PEGAR_TOKEN_NUEVO');
  p.setProperty('OPENAI_KEY', 'PEGAR_KEY_NUEVA');
  Logger.log('Claves guardadas OK');
}

function TG(){ return PropertiesService.getScriptProperties().getProperty('TG_TOKEN'); }

function OAI(){ return PropertiesService.getScriptProperties().getProperty('OPENAI_KEY'); }

function ss(){ return SpreadsheetApp.openById(MAESTRO_ID); }

function prop(k,v){ var p=PropertiesService.getScriptProperties(); if(v===undefined) return p.getProperty(k); p.setProperty(k,v); }

function fmt(d,f){ return Utilities.formatDate(d,TZ,f); }

function hoy(){ return fmt(new Date(),'dd/MM/yyyy'); }

function tg(m,q){ return UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/'+m+(q||''),{muteHttpExceptions:true}); }


function setWebhookManual(){
  var URL_EXEC = 'https://script.google.com/macros/s/AKfycbx4bNGJvySzo-hhQb5Kv092QdIb28CffgDGV3RMvun0lZ1U7QRKV0RITSMuPPGhHXaF/exec';
  Logger.log(tg('setWebhook','?url='+encodeURIComponent(URL_EXEC)).getContentText());
}

function pararTodo(){ Logger.log(tg('deleteWebhook','?drop_pending_updates=true').getContentText()); }
