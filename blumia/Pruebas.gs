function probarClaves() {
  const chk = (nombre, clave, url, headers) => {
    if (!clave) return nombre + ': FALTA la propiedad del script';
    const r = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
    const c = r.getResponseCode();
    return nombre + ' (termina en ' + clave.slice(-4) + '): ' + (c === 200 ? 'OK' : 'ERROR ' + c + ' ' + r.getContentText().slice(0, 150));
  };
  const out = [
    chk('OpenAI', OPENAI_API_KEY, 'https://api.openai.com/v1/models', { Authorization: 'Bearer ' + OPENAI_API_KEY }),
    chk('Anthropic', ANTHROPIC_API_KEY, 'https://api.anthropic.com/v1/models', { 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' })
  ].join('\n');
  Logger.log(out);
  return out;
}
