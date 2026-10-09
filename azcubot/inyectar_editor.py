import re,sys
ED=open('editor_pdf.js',encoding='utf-8').read()
for f in ['../appscript/Precarga.html','paginas/Precarga.html','../appscript/Tablero.html','Tablero.html']:
    h=open(f,encoding='utf-8').read()
    h2=re.sub(r'/\*ED_PDF_INI\*/.*?/\*ED_PDF_FIN\*/\n?',lambda m:ED,h,flags=re.S)
    if h2==h: print('sin cambios',f)
    open(f,'w',encoding='utf-8').write(h2)
