# Cómo se publica

- Los cambios se piden por Claude y llegan como pedido de cambios (pull request) en GitHub.
- Al aprobarlo con **Merge**:
  - `appscript/` → se publica solo en el Apps Script de azcuenaga (GitHub Actions + clasp).
  - `pami/` → se publica solo en el Apps Script de dotacionpp1.
  - `azcubot/netlify/` → Cloudflare Pages publica la web `azcubot.pages.dev`.
- Antes de publicar Apps Script, el sistema compara con lo último publicado; si alguien lo tocó a mano, frena y guarda la copia en la rama `sync-azcu` o `sync-pami`.
- No editar el código directo en Apps Script.

Cualquier duda, avisá a Ariel.
