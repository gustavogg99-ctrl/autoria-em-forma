// Worker "site": serve a página estática (pasta site/) e cuida de duas rotas próprias.
//   POST /api/e   recebe eventos da página (rastreador no fim de pagina/fonte.html)
//   GET  /painel  painel de acessos, protegido por senha (usuário qualquer)
// Todo o resto vai direto para os arquivos estáticos (wrangler.jsonc: run_worker_first).
//
// Privacidade: não guarda IP, user agent nem identificador de pessoa. Só data, tipo de
// evento, origem do acesso, domínio de quem mandou, celular/computador e país.
//
// Segurança (checkup de 22/10/2026):
//   - /api/e só aceita envio vindo da própria página (cabeçalho Origin obrigatório) e no máximo
//     LIMITE_EVENTOS por visitante por minuto. O IP serve só de chave do limitador da Cloudflare,
//     em memória; não é gravado.
//   - /painel aceita no máximo LIMITE_PAINEL tentativas por visitante por minuto (força bruta) e
//     manda cabeçalhos de segurança, com CSP de nonce para o script do painel.
//   - Os cabeçalhos da página estática estão em site/_headers (gerado por pagina/gerar_site.py).

const TIPOS = new Set(["visita", "viu_precos", "calc_inicio", "calc_enviar", "whats_direto", "copiar_numero", "exemplo"]);
const FACULDADES = new Set(["UnB", "CEUB", "UCB", "IESB", "UDF", "UNICEPLAC", "UNIP", "IDP", "Outra faculdade do DF"]);
const BOT = /bot|crawl|spider|slurp|preview|fetch|monitor|headless|lighthouse|facebookexternalhit|whatsapp\//i;

export default {
  async fetch(req, env, ctx) {
    const url = new URL(req.url);
    if (url.pathname === "/api/e") return coletar(req, env, ctx, url);
    if (url.pathname === "/painel" || url.pathname.startsWith("/painel/")) return painel(req, env, url);
    return env.ASSETS.fetch(req);
  },
};

// ---------- coleta ----------

function diaBrasilia(d = new Date()) {
  return new Date(d.getTime() - 3 * 3600 * 1000).toISOString().slice(0, 10); // Brasília: UTC-3, sem horário de verão
}

function limpar(s, max) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9._-]/g, "").slice(0, max);
}

// Origem do acesso: o parâmetro ?origem= do link vence; sem ele, o domínio de quem mandou.
function classificar(origem, refHost, proprio) {
  const o = limpar(origem, 40);
  if (o) return o === "whats" || o === "wpp" || o === "zap" ? "whatsapp" : o;
  const h = refHost.replace(/^www\./, "");
  if (!h || h === proprio) return "direto";
  if (/(^|\.)google\./.test(h)) return "google";
  if (/(^|\.)bing\.com$/.test(h)) return "bing";
  if (/duckduckgo\.com$/.test(h)) return "duckduckgo";
  if (/yahoo\./.test(h)) return "yahoo";
  if (/whatsapp\.com$|^wa\.me$|^l\.wl\.co$/.test(h)) return "whatsapp";
  if (/instagram\.com$/.test(h)) return "instagram";
  if (/facebook\.com$|^fb\.me$/.test(h)) return "facebook";
  if (/linkedin\.com$|^lnkd\.in$/.test(h)) return "linkedin";
  if (/^t\.co$|twitter\.com$|(^|\.)x\.com$/.test(h)) return "x";
  if (/tiktok\.com$/.test(h)) return "tiktok";
  if (/claude\.ai$/.test(h)) return "claude";
  return "site:" + h.slice(0, 60);
}

function faixa(p) {
  const n = Math.floor(Number(p));
  if (!n || n < 1) return "?";
  if (n <= 20) return "até 20 p.";
  if (n <= 40) return "21–40 p.";
  if (n <= 80) return "41–80 p.";
  if (n <= 120) return "81–120 p.";
  return "121+ p.";
}

// Limitador da Cloudflare (wrangler.jsonc, "ratelimits"). Sem o binding (teste antigo), libera.
async function dentroDoLimite(limitador, req, rota) {
  if (!limitador) return true;
  const ip = req.headers.get("cf-connecting-ip") || "sem-ip";
  try { return (await limitador.limit({ key: rota + ":" + ip })).success; } catch { return true; }
}

const BASICOS = { "cache-control": "no-store", "x-content-type-options": "nosniff" };

async function coletar(req, env, ctx, url) {
  const vazio = new Response(null, { status: 204, headers: BASICOS });
  if (req.method !== "POST") return new Response("método", { status: 405, headers: BASICOS });
  const ua = req.headers.get("user-agent") || "";
  if (BOT.test(ua)) return vazio;
  // Navegador sempre manda Origin num POST; sem ele (script) ou de outro site, descarta.
  const origemHdr = req.headers.get("origin");
  let origemOk = false;
  try { origemOk = !!origemHdr && new URL(origemHdr).host === url.host; } catch { origemOk = false; }
  if (!origemOk) return vazio;
  if (!(await dentroDoLimite(env.LIMITE_EVENTOS, req, "e"))) return new Response(null, { status: 429, headers: BASICOS });
  const corpo = await req.text();
  if (corpo.length > 1024) return vazio;
  let d;
  try { d = JSON.parse(corpo); } catch { return vazio; }
  if (!d || !TIPOS.has(d.t)) return vazio;

  let refHost = "";
  try { refHost = d.r ? new URL(d.r).host.toLowerCase() : ""; } catch { /* referenciador inválido */ }
  const origem = classificar(d.o, refHost, url.host);
  let detalhe = null;
  if (d.t === "calc_enviar") {
    const serv = d.s === "rev" ? "formatação + revisão" : "formatação";
    const fac = FACULDADES.has(d.f) ? d.f : "sem faculdade";
    detalhe = serv + " · " + fac + " · " + faixa(d.p);
  }
  const celular = req.headers.get("sec-ch-ua-mobile") === "?1" || /Mobi|Android|iPhone|iPad/i.test(ua);
  const pais = String((req.cf && req.cf.country) || "").slice(0, 2) || null;
  const agora = new Date();

  const gravar = async () => {
    await env.DB.prepare(
      "INSERT INTO eventos (ts, dia, tipo, origem, ref_host, detalhe, dispositivo, pais) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    ).bind(agora.toISOString(), diaBrasilia(agora), d.t, origem, refHost.slice(0, 80) || null, detalhe,
      celular ? "celular" : "computador", pais).run();
    // limpeza: de vez em quando apaga eventos com mais de 400 dias
    if (Math.random() < 0.01) await env.DB.prepare("DELETE FROM eventos WHERE dia < ?").bind(diaBrasilia(new Date(agora.getTime() - 400 * 86400000))).run();
  };
  ctx.waitUntil(gravar().catch((e) => console.error("gravar", e)));
  return vazio;
}

// ---------- painel ----------

async function sha256(s) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join("");
}

async function autorizado(req, env) {
  const h = req.headers.get("authorization") || "";
  if (!h.startsWith("Basic ") || !env.PAINEL_HASH) return false;
  let senha = "";
  try { senha = atob(h.slice(6)).split(":").slice(1).join(":"); } catch { return false; }
  const a = await sha256(senha), b = env.PAINEL_HASH.toLowerCase();
  let dif = a.length ^ b.length;
  for (let i = 0; i < Math.min(a.length, b.length); i++) dif |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return dif === 0;
}

const CABECALHOS = {
  "cache-control": "no-store",
  "x-robots-tag": "noindex, nofollow",
  "x-frame-options": "DENY",
  "x-content-type-options": "nosniff",
  "referrer-policy": "no-referrer",
  "cross-origin-opener-policy": "same-origin",
  "permissions-policy": "camera=(), microphone=(), geolocation=(), payment=(), usb=()",
};

function cspPainel(nonce) {
  return "default-src 'none'; script-src 'nonce-" + nonce + "'; style-src 'unsafe-inline'; img-src 'self'; " +
    "base-uri 'none'; form-action 'none'; frame-ancestors 'none'";
}

async function painel(req, env, url) {
  if (!(await dentroDoLimite(env.LIMITE_PAINEL, req, "painel"))) {
    return new Response("Muitas tentativas. Espere um minuto e tente de novo.", {
      status: 429,
      headers: { ...CABECALHOS, "content-type": "text/plain; charset=utf-8", "retry-after": "60" },
    });
  }
  if (!(await autorizado(req, env))) {
    return new Response("Senha necessária.", {
      status: 401,
      headers: { ...CABECALHOS, "content-type": "text/plain; charset=utf-8", "www-authenticate": 'Basic realm="Painel Autoria em Forma", charset="UTF-8"' },
    });
  }
  const dias = [7, 30, 90].includes(Number(url.searchParams.get("dias"))) ? Number(url.searchParams.get("dias")) : 30;
  const hoje = diaBrasilia();
  const inicio = diaBrasilia(new Date(Date.now() - (dias - 1) * 86400000));
  const q = (sql) => env.DB.prepare(sql).bind(inicio).all().then((r) => r.results);

  const [porDia, porOrigem, dispositivos, paises, detalhes, ultimos] = await Promise.all([
    q("SELECT dia, tipo, COUNT(*) n FROM eventos WHERE dia >= ? GROUP BY dia, tipo"),
    q("SELECT origem, tipo, COUNT(*) n FROM eventos WHERE dia >= ? GROUP BY origem, tipo"),
    q("SELECT dispositivo k, COUNT(*) n FROM eventos WHERE dia >= ? AND tipo = 'visita' GROUP BY dispositivo ORDER BY n DESC"),
    q("SELECT COALESCE(pais, '?') k, COUNT(*) n FROM eventos WHERE dia >= ? AND tipo = 'visita' GROUP BY pais ORDER BY n DESC LIMIT 8"),
    q("SELECT detalhe k, COUNT(*) n FROM eventos WHERE dia >= ? AND tipo = 'calc_enviar' GROUP BY detalhe ORDER BY n DESC LIMIT 15"),
    q("SELECT ts, tipo, origem, dispositivo, detalhe FROM eventos WHERE dia >= ? ORDER BY id DESC LIMIT 25"),
  ]);
  const dados = { dias, inicio, hoje, porDia, porOrigem, dispositivos, paises, detalhes, ultimos };
  const json = JSON.stringify(dados).replace(/</g, "\\u003c");
  const nonce = btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(16))));
  const html = PAINEL_HTML.replace("<script>", '<script nonce="' + nonce + '">').replace("__DADOS__", () => json);
  return new Response(html, {
    headers: { ...CABECALHOS, "content-type": "text/html; charset=utf-8", "content-security-policy": cspPainel(nonce) },
  });
}

const PAINEL_HTML = `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Painel de acessos</title>
<style>
:root{color-scheme:light;--plano:#f9f9f7;--sup:#fcfcfb;--t1:#0b0b0b;--t2:#52514e;--t3:#898781;--grade:#e1e0d9;--eixo:#c3c2b7;--borda:rgba(11,11,11,.10);--s1:#2a78d6;--s1h:#1c5cab;--bom:#006300}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--plano:#0d0d0d;--sup:#1a1a19;--t1:#f0efec;--t2:#c3c2b7;--t3:#898781;--grade:#2c2c2a;--eixo:#383835;--borda:rgba(255,255,255,.10);--s1:#3987e5;--s1h:#86b6ef;--bom:#0ca30c}}
*{box-sizing:border-box}
body{margin:0;background:var(--plano);color:var(--t1);font:15px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:1040px;margin:0 auto;padding:20px 16px 48px}
header{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:12px;margin-bottom:16px}
h1{font-size:20px;margin:0}
h2{font-size:15px;margin:0 0 12px}
.sub{color:var(--t2);font-size:13px}
nav.periodo{display:flex;gap:4px;background:var(--sup);border:1px solid var(--borda);border-radius:8px;padding:3px}
nav.periodo a{padding:4px 12px;border-radius:6px;color:var(--t2);text-decoration:none;font-size:14px}
nav.periodo a[aria-current]{background:var(--t1);color:var(--sup);font-weight:600}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:12px}
.rola{overflow-x:auto}
.card{background:var(--sup);border:1px solid var(--borda);border-radius:10px;padding:16px}
.tile .rot{color:var(--t2);font-size:13px}
.tile .num{font-size:30px;font-weight:650;margin-top:2px}
.tile .nota{color:var(--t3);font-size:12px}
.grade2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px}
@media (max-width:760px){.grade2{grid-template-columns:1fr}}
.grafico{position:relative}
svg{display:block;width:100%;height:auto;overflow:visible}
.tt{position:absolute;pointer-events:none;background:var(--t1);color:var(--sup);font-size:12px;padding:6px 8px;border-radius:6px;white-space:nowrap;transform:translate(-50%,-110%);opacity:0;transition:opacity .1s}
.tt.on{opacity:1}
table{width:100%;border-collapse:collapse;font-size:14px}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--grade)}
th{color:var(--t2);font-weight:600;font-size:12px}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
.barra{display:grid;grid-template-columns:minmax(90px,38%) 1fr auto;align-items:center;gap:8px;padding:5px 0;font-size:14px}
.barra .trilho{height:10px}
.barra .ench{height:10px;background:var(--s1);border-radius:0 4px 4px 0}
.barra .v{font-variant-numeric:tabular-nums;color:var(--t2);font-size:13px;min-width:5ch;text-align:right}
.vazio{color:var(--t3);font-size:14px}
details{margin-top:10px}
summary{cursor:pointer;color:var(--t2);font-size:13px}
.rodape{color:var(--t3);font-size:12px;margin-top:20px}
.dica{color:var(--t3);font-size:12px;margin-top:8px}
code{font-size:12px;background:var(--plano);padding:1px 4px;border-radius:4px}
</style>
</head>
<body>
<main>
<header>
  <div><h1>Painel de acessos</h1><div class="sub" id="periodo-txt"></div></div>
  <nav class="periodo" id="periodo"></nav>
</header>

<section class="tiles" id="tiles"></section>

<section class="card">
  <h2>Visitas por dia</h2>
  <div class="grafico" id="g-dia"><div class="tt" id="tt"></div></div>
  <details><summary>Ver em tabela</summary><table id="t-dia"></table></details>
</section>

<div class="grade2">
  <section class="card"><h2>Funil</h2><div id="funil"></div>
    <p class="dica">Cada etapa conta uma vez por visita. "Enviaram orçamento" é o clique no botão da calculadora; "chamaram direto" soma o link do WhatsApp e o botão "Copiar número".</p></section>
  <section class="card"><h2>De onde vieram</h2><div id="origens"></div>
    <p class="dica">Marque cada link que você manda: <code>?origem=whats</code>, <code>?origem=indicacao-nome</code>, <code>?origem=status</code>. Sem marca, quem vem pelo WhatsApp costuma aparecer como "direto".</p></section>
</div>

<section class="card" style="margin-top:12px">
  <h2>Resultado por origem</h2>
  <div class="rola"><table id="t-origem"></table></div>
</section>

<div class="grade2">
  <section class="card"><h2>Orçamentos enviados pela calculadora</h2><div id="detalhes"></div></section>
  <section class="card"><h2>Aparelho e país</h2><div id="disp"></div><div id="paises" style="margin-top:12px"></div></section>
</div>

<section class="card" style="margin-top:12px">
  <h2>Últimos eventos</h2>
  <table id="t-ultimos"></table>
</section>

<p class="rodape">Seus próprios acessos não contam: abrir este painel marca o navegador como seu. Em outro aparelho, abra o site com <code>?eu=1</code> uma vez (<code>?eu=0</code> desfaz). Visitas = sessões do navegador, não pessoas únicas. Horários de Brasília.</p>
</main>

<script>
try { localStorage.setItem("af_eu", "1"); } catch (e) {}
var D = __DADOS__;
var NOMES = { visita: "Visita", viu_precos: "Viu os preços", calc_inicio: "Usou a calculadora", calc_enviar: "Enviou orçamento", whats_direto: "Clicou no WhatsApp", copiar_numero: "Copiou o número", exemplo: "Viu o exemplo" };
var $ = function (id) { return document.getElementById(id); };
var esc = function (s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); };
var fmt = function (n) { return Number(n).toLocaleString("pt-BR"); };
var pct = function (a, b) { return b ? Math.round((a / b) * 100) + "%" : "–"; };
var dataBR = function (d) { var p = d.split("-"); return p[2] + "/" + p[1]; };

// período
$("periodo").innerHTML = [7, 30, 90].map(function (d) { return '<a href="?dias=' + d + '"' + (d === D.dias ? ' aria-current="page"' : "") + ">" + d + " dias</a>"; }).join("");
$("periodo-txt").textContent = dataBR(D.inicio) + " a " + dataBR(D.hoje) + " · atualizado ao abrir a página";

// totais por tipo
var tot = {}; D.porDia.forEach(function (r) { tot[r.tipo] = (tot[r.tipo] || 0) + r.n; });
var g = function (t) { return tot[t] || 0; };
var direto = g("whats_direto") + g("copiar_numero");
$("tiles").innerHTML = [
  ["Visitas", g("visita"), "no período"],
  ["Usaram a calculadora", g("calc_inicio"), pct(g("calc_inicio"), g("visita")) + " das visitas"],
  ["Enviaram orçamento", g("calc_enviar"), pct(g("calc_enviar"), g("visita")) + " das visitas"],
  ["Chamaram direto", direto, "link do WhatsApp ou número copiado"]
].map(function (t) { return '<div class="card tile"><div class="rot">' + t[0] + '</div><div class="num">' + fmt(t[1]) + '</div><div class="nota">' + t[2] + "</div></div>"; }).join("");

// visitas por dia (todos os dias do período, inclusive os zerados)
var dias = [], vis = {};
D.porDia.forEach(function (r) { if (r.tipo === "visita") vis[r.dia] = r.n; });
for (var i = 0; i < D.dias; i++) { var d = new Date(D.inicio + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + i); var k = d.toISOString().slice(0, 10); dias.push({ dia: k, n: vis[k] || 0 }); }
(function () {
  var W = Math.max(320, $("g-dia").clientWidth || 960), H = W < 600 ? 200 : 240, ml = 36, mr = 8, mt = 10, mb = 26, iw = W - ml - mr, ih = H - mt - mb;
  var max = Math.max(4, Math.max.apply(null, dias.map(function (x) { return x.n; })));
  var passo = Math.ceil(max / 4); max = passo * 4;
  var bw = iw / dias.length, gap = Math.min(2, bw * 0.25);
  var s = '<svg viewBox="0 0 ' + W + " " + H + '" role="img" aria-label="Visitas por dia">';
  for (var y = 0; y <= max; y += passo) { var yy = mt + ih - (y / max) * ih; s += '<line x1="' + ml + '" x2="' + (W - mr) + '" y1="' + yy + '" y2="' + yy + '" stroke="var(' + (y ? "--grade" : "--eixo") + ')" stroke-width="1"/><text x="' + (ml - 6) + '" y="' + (yy + 4) + '" text-anchor="end" font-size="11" fill="var(--t3)">' + y + "</text>"; }
  var cada = Math.ceil(dias.length / Math.max(3, Math.floor(iw / 80)));
  dias.forEach(function (x, i) {
    var h = (x.n / max) * ih, xx = ml + i * bw + gap / 2, w = Math.max(1, bw - gap), yy = mt + ih - h, r = Math.min(4, w / 2, h);
    if (x.n) s += '<path class="b" d="M' + xx + " " + (mt + ih) + "V" + (yy + r) + "Q" + xx + " " + yy + " " + (xx + r) + " " + yy + "H" + (xx + w - r) + "Q" + (xx + w) + " " + yy + " " + (xx + w) + " " + (yy + r) + "V" + (mt + ih) + 'Z" fill="var(--s1)"/>';
    s += '<rect class="alvo" data-i="' + i + '" x="' + (ml + i * bw) + '" y="' + mt + '" width="' + bw + '" height="' + ih + '" fill="transparent"/>';
    if (i % cada === 0) s += '<text x="' + (ml + i * bw + bw / 2) + '" y="' + (H - 8) + '" text-anchor="middle" font-size="11" fill="var(--t3)">' + dataBR(x.dia) + "</text>";
  });
  s += "</svg>";
  var box = $("g-dia"), tt = $("tt");
  box.insertAdjacentHTML("afterbegin", s);
  var svg = box.querySelector("svg");
  box.addEventListener("pointermove", function (e) {
    var a = e.target.closest && e.target.closest(".alvo");
    if (!a) { tt.classList.remove("on"); return; }
    var x = dias[+a.dataset.i], rb = box.getBoundingClientRect(), ra = a.getBoundingClientRect();
    tt.textContent = dataBR(x.dia) + " · " + x.n + (x.n === 1 ? " visita" : " visitas");
    tt.style.left = (ra.left - rb.left + ra.width / 2) + "px";
    tt.style.top = (ra.top - rb.top + ra.height * (1 - x.n / max)) + "px";
    tt.classList.add("on");
  });
  box.addEventListener("pointerleave", function () { tt.classList.remove("on"); });
  $("t-dia").innerHTML = "<tr><th>Dia</th><th class=n>Visitas</th></tr>" + dias.slice().reverse().map(function (x) { return "<tr><td>" + dataBR(x.dia) + '</td><td class="n">' + x.n + "</td></tr>"; }).join("");
})();

function barras(lista, alvo, total) {
  if (!lista.length) { $(alvo).innerHTML = '<p class="vazio">Nada no período.</p>'; return; }
  var max = Math.max.apply(null, lista.map(function (x) { return x[1]; })) || 1;
  $(alvo).innerHTML = lista.map(function (x) {
    return '<div class="barra"><span>' + esc(x[0]) + '</span><div class="trilho"><div class="ench" style="width:' + (x[1] / max) * 100 + '%"></div></div><span class="v">' + fmt(x[1]) + (total ? " · " + pct(x[1], total) : "") + "</span></div>";
  }).join("");
}

barras([["Visitas", g("visita")], ["Viram os preços", g("viu_precos")], ["Usaram a calculadora", g("calc_inicio")], ["Enviaram orçamento", g("calc_enviar")], ["Chamaram direto", direto]], "funil", g("visita"));

// por origem
var orig = {};
D.porOrigem.forEach(function (r) { var o = orig[r.origem] = orig[r.origem] || {}; o[r.tipo] = r.n; });
var ords = Object.keys(orig).map(function (k) { return [k, orig[k].visita || 0, orig[k]]; }).sort(function (a, b) { return b[1] - a[1]; });
barras(ords.filter(function (x) { return x[1]; }).map(function (x) { return [x[0], x[1]]; }), "origens", g("visita"));
$("t-origem").innerHTML = ords.length ? "<tr><th>Origem</th><th class=n>Visitas</th><th class=n>Calculadora</th><th class=n>Enviaram</th><th class=n>Chamaram direto</th><th class=n>Conversão</th></tr>" + ords.map(function (x) {
  var o = x[2], dir = (o.whats_direto || 0) + (o.copiar_numero || 0), conv = (o.calc_enviar || 0) + dir;
  return "<tr><td>" + esc(x[0]) + '</td><td class="n">' + (o.visita || 0) + '</td><td class="n">' + (o.calc_inicio || 0) + '</td><td class="n">' + (o.calc_enviar || 0) + '</td><td class="n">' + dir + '</td><td class="n">' + pct(conv, o.visita || 0) + "</td></tr>";
}).join("") : '<tr><td class="vazio">Nada no período.</td></tr>';

barras(D.detalhes.map(function (x) { return [x.k, x.n]; }), "detalhes");
barras(D.dispositivos.map(function (x) { return [x.k, x.n]; }), "disp", g("visita"));
barras(D.paises.map(function (x) { return [x.k, x.n]; }), "paises");

$("t-ultimos").innerHTML = D.ultimos.length ? "<tr><th>Quando</th><th>Evento</th><th>Origem</th><th>Aparelho</th></tr>" + D.ultimos.map(function (r) {
  var t = new Date(r.ts).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo", day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
  return "<tr><td>" + t + "</td><td>" + esc(NOMES[r.tipo] || r.tipo) + (r.detalhe ? ' <span class="sub">(' + esc(r.detalhe) + ")</span>" : "") + "</td><td>" + esc(r.origem) + "</td><td>" + esc(r.dispositivo) + "</td></tr>";
}).join("") : '<tr><td class="vazio">Nenhum evento ainda.</td></tr>';
</script>
</body>
</html>`;
