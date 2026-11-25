"""Gera site/index.html (publicado na Cloudflare) a partir de pagina/fonte.html.

pagina/fonte.html é o mesmo arquivo publicado como artifact no Claude, que não leva
<!doctype>, <html>, <head> nem charset (o Claude acrescenta). Sem charset, a Cloudflare
serve o texto e o navegador quebra os acentos. Este script acrescenta o cabeçalho.

Uso (da raiz do repositório): python3 -I pagina/gerar_site.py
"""
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
CABECALHO = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<link rel="icon" href="/favicon.svg" type="image/svg+xml">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:site_name" content="Autoria em Forma">
<meta property="og:title" content="Formatação ABNT e revisão de TCC em Brasília | Autoria em Forma">
<meta property="og:description" content="TCC, monografia e artigo no padrão do seu curso, da sua faculdade e da ABNT, com cada correção marcada e explicada. Calcule seu orçamento.">
<meta property="og:url" content="https://site.autoriaemforma.workers.dev/">
<meta property="og:image" content="https://site.autoriaemforma.workers.dev/og.png?v=2">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>
"""

fonte = (RAIZ / "pagina" / "fonte.html").read_text(encoding="utf-8")
fim = fonte.index("</style>") + len("</style>")
saida = CABECALHO + fonte[:fim] + "\n</head>\n<body>\n" + fonte[fim:] + "\n</body>\n</html>\n"
(RAIZ / "site" / "index.html").write_text(saida, encoding="utf-8")
print("site/index.html gerado")

# Cabeçalhos de segurança da página (Cloudflare lê site/_headers). A CSP libera só os scripts
# desta página, pelo hash; se um script mudar, rodar este arquivo de novo atualiza os hashes.
# Checkup de segurança e HTTP Observatory, 22/10/2026.
import base64, hashlib, re

scripts = re.findall(r"<script>(.*?)</script>", saida, flags=re.S)
assert not re.search(r"<script[^>]*\bsrc=", saida), "script externo: incluir a origem na CSP"
hashes = " ".join("'sha256-" + base64.b64encode(hashlib.sha256(x.encode("utf-8")).digest()).decode() + "'" for x in scripts)
csp = "; ".join([
    "default-src 'none'",
    "script-src " + hashes,
    "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
    "font-src https://fonts.gstatic.com",
    "img-src 'self' data:",
    "connect-src 'self'",
    "base-uri 'none'",
    "form-action 'none'",
    "frame-ancestors 'none'",
])
cabecalhos = f"""# Gerado por pagina/gerar_site.py. Não editar à mão.
/*
  Content-Security-Policy: {csp}
  X-Content-Type-Options: nosniff
  X-Frame-Options: DENY
  Referrer-Policy: strict-origin-when-cross-origin
  Permissions-Policy: camera=(), microphone=(), geolocation=(), payment=(), usb=()
  Cross-Origin-Opener-Policy: same-origin
"""
assert all(len(l) <= 2000 for l in cabecalhos.splitlines())
(RAIZ / "site" / "_headers").write_text(cabecalhos, encoding="utf-8")
print(f"site/_headers gerado ({len(scripts)} scripts na CSP)")
