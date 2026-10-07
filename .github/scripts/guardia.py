import sys, os, unicodedata

prev, real = sys.argv[1], sys.argv[2]
EXT = ('.gs', '.js', '.html')


def norm_name(n):
    n = unicodedata.normalize('NFD', n)
    n = ''.join(c for c in n if not unicodedata.combining(c))
    return n[:-3] + '.gs' if n.endswith('.js') else n


def leer(d):
    out = {}
    for f in os.listdir(d):
        p = os.path.join(d, f)
        if not os.path.isfile(p):
            continue
        if not (f.endswith(EXT) or f == 'appsscript.json'):
            continue
        txt = open(p, encoding='utf8', errors='replace').read().replace('\r', '')
        out[norm_name(f)] = '\n'.join(l.rstrip() for l in txt.split('\n')).strip()
    return out


a, b = leer(prev), leer(real)
dif = sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))
for k in dif:
    estado = 'solo en el Apps Script real (alguien lo creó a mano)' if k not in a else 'falta en el Apps Script real' if k not in b else 'modificado a mano'
    print('- %s: %s' % (k, estado))
sys.exit(1 if dif else 0)
