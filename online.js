// Versão online da tela (GitHub Pages): troca a API local pela pasta do sistema no SharePoint.
// A tela continua chamando api("/api/..."); aqui cada rota vira leitura de estado/*.json ou um
// pedido em pedidos/, que o trabalhador (o computador com a ponte ligada) executa.
(() => {
const CFG = window.VEIGA_CONFIG;
if (!CFG || /^0{8}-/.test(CFG.clientId) || /^0{8}-/.test(CFG.tenantId)) {  // config.js ainda com o exemplo
  window.veigaOnline = {
    api: async () => { throw new Error("Sistema ainda não configurado."); },
    async iniciar() {
      document.getElementById("app").innerHTML = `<div class="card vazio"><h2>Quase pronto</h2>
        <p>O site está no ar, mas o login com a conta @veiga.law ainda depende da liberação da TI.</p>
        <p class="sub">Assim que a TI registrar o aplicativo, a esteira aparece aqui.</p></div>`;
      return false;
    },
  };
  return;
}
const GRAPH = "https://graph.microsoft.com/v1.0";
const ESCOPOS = ["Files.ReadWrite.All", "Sites.ReadWrite.All"];
const TIPOS = { esocial: [".zip"], extratos: [".xlsx"], darf: [".pdf", ".xlsx"], cadastro: [".pdf", ".xlsx", ".docx"] };
const msal = new window.msal.PublicClientApplication({
  auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenantId}`, redirectUri: location.origin + location.pathname },
  cache: { cacheLocation: "localStorage" },
});
let conta = null, raiz = null;

async function token() {
  try { return (await msal.acquireTokenSilent({ scopes: ESCOPOS, account: conta })).accessToken; }
  catch (e) { await msal.acquireTokenRedirect({ scopes: ESCOPOS, account: conta }); }
}
async function graph(url, opts = {}) {
  const r = await fetch(url.startsWith("http") ? url : GRAPH + url,
    { ...opts, headers: { Authorization: "Bearer " + await token(), ...(opts.headers || {}) } });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error(`SharePoint respondeu ${r.status}. ${(await r.text()).slice(0, 200)}`);
  return r.status === 204 ? true : r.json();
}
const item = caminho => `/drives/${raiz.drive}/items/${raiz.id}` +
  (caminho ? ":/" + caminho.split("/").map(encodeURIComponent).join("/") + ":" : "");

async function ler(caminho, texto = false) {
  const i = await graph(item(caminho) + "?$select=id,@microsoft.graph.downloadUrl");
  if (!i) return null;
  const r = await fetch(i["@microsoft.graph.downloadUrl"], { cache: "no-store" });
  return texto ? r.text() : r.json();
}
async function listar(pasta) {
  const r = await graph(item(pasta) + "/children?$top=999");
  return r ? r.value : [];
}
async function gravar(caminho, corpo) {
  const bytes = corpo instanceof Blob ? corpo : new Blob([corpo]);
  if (bytes.size <= 4 * 1024 * 1024)
    return graph(item(caminho) + "/content", { method: "PUT", body: bytes });
  const s = await graph(item(caminho) + "/createUploadSession",
    { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) });
  const passo = 5 * 1024 * 1024;  // múltiplo de 320 KiB, como o Graph exige
  for (let ini = 0; ini < bytes.size; ini += passo) {
    const fim = Math.min(ini + passo, bytes.size);
    const r = await fetch(s.uploadUrl, { method: "PUT", body: bytes.slice(ini, fim),
      headers: { "Content-Range": `bytes ${ini}-${fim - 1}/${bytes.size}` } });
    if (!r.ok) throw new Error(`Falha no envio (${r.status}).`);
  }
}

// ---------------------------------------------------------------- pedidos ao trabalhador
async function pedir(pedido) {
  const nome = `${new Date().toISOString().replace(/\D/g, "").slice(0, 17)}_${pedido.acao}_${Math.random().toString(36).slice(2, 6)}.json`;
  await gravar("pedidos/" + nome, JSON.stringify({ ...pedido, quem: conta.name, criado_em: new Date().toISOString() }));
  for (let i = 0; i < 20; i++) {  // espera até ~40 s pela resposta
    await new Promise(r => setTimeout(r, 2000));
    const feito = await ler("pedidos/feitos/" + nome);
    if (feito) {
      if (!feito.resultado.ok) throw new Error(feito.resultado.erro);
      return feito.resultado;
    }
  }
  throw new Error("Pedido registrado, mas o trabalhador ainda não respondeu. Ele será feito assim que o computador que processa estiver ligado.");
}

// ---------------------------------------------------------------- as rotas da tela
const cache = {};
async function cliente(cid) {
  const c = await ler(`estado/cliente-${cid}.json`);
  if (!c) throw new Error("Cliente ainda não publicado. Tente de novo em alguns segundos.");
  cache[cid] = c;
  const entradas = [];
  for (const tipo of Object.keys(TIPOS))
    for (const i of await listar(`${c.pasta_sp}/entrada/${tipo}`))
      if (!i.folder) entradas.push({ id: i.id, tipo, nome: i.name, tamanho: i.size, criado_em: i.createdDateTime });
  c.arquivos = [...entradas, ...c.arquivos.filter(a => a.tipo === "saida")];  // a entrada vem direto da pasta
  return c;
}
async function enviarArquivos(cid, fd) {
  const c = cache[cid] || await cliente(cid);
  const tipo = fd.get("tipo"), aceitos = [], recusados = [];
  for (const f of fd.getAll("arquivos")) {
    if (!TIPOS[tipo].some(x => f.name.toLowerCase().endsWith(x))) { recusados.push(f.name); continue; }
    await gravar(`${c.pasta_sp}/entrada/${tipo}/${f.name}`, f);
    aceitos.push(f.name);
  }
  return { aceitos, recusados };
}

async function api(url, opts = {}) {
  const metodo = (opts.method || "GET").toUpperCase();
  const corpo = typeof opts.body === "string" ? JSON.parse(opts.body) : {};
  let m;
  if (metodo === "GET") {
    if (url === "/api/etapas") {
      const e = await ler("estado/etapas.json");
      if (!e) throw new Error("O sistema ainda não foi publicado nesta pasta do SharePoint.");
      const parado = Date.now() - new Date(e.trabalhador?.vivo_em || 0) > 5 * 60 * 1000;
      document.getElementById("parado").hidden = !parado;
      return e;
    }
    if (url === "/api/clientes") return (await ler("estado/clientes.json")) || [];
    if ((m = url.match(/^\/api\/clientes\/(\d+)$/))) return cliente(m[1]);
    if (url.startsWith("/api/movimentos")) return (await ler("estado/movimentos.json")) || [];
    if (url === "/api/metricas") return ler("estado/metricas.json");
    if ((m = url.match(/^\/api\/jobs\/(\d+)\/log$/))) return { log: await ler(`estado/logs/job-${m[1]}.txt`, true) };
  }
  if (metodo === "POST") {
    if (url === "/api/clientes") return pedir({ acao: "cadastrar", dados: corpo });
    if ((m = url.match(/^\/api\/clientes\/(\d+)\/arquivos$/))) return enviarArquivos(m[1], opts.body);
    if ((m = url.match(/^\/api\/clientes\/(\d+)\/upload\/(concluir|reabrir)$/)))
      return pedir({ acao: m[2] + "_upload", cliente_id: +m[1] });
    if ((m = url.match(/^\/api\/clientes\/(\d+)\/etapas\/(\w+)\/(rodar|aprovar|devolver)$/)))
      return pedir({ acao: m[3], cliente_id: +m[1], etapa: m[2], nota: corpo.nota });
  }
  if (metodo === "DELETE" && (m = url.match(/^\/api\/arquivos\/([^?]+)/)))
    return graph(`/drives/${raiz.drive}/items/${m[1]}`, { method: "DELETE" });
  throw new Error("Rota não disponível na versão online: " + url);
}

// ---------------------------------------------------------------- entrada
window.veigaOnline = {
  api,
  async iniciar() {
    const volta = await msal.handleRedirectPromise();
    conta = volta?.account || msal.getAllAccounts()[0];
    if (!conta) { await msal.loginRedirect({ scopes: ESCOPOS }); return false; }
    const link = "u!" + btoa(unescape(encodeURIComponent(CFG.pastaSharePoint))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    const r = await graph(`/shares/${link}/driveItem`);
    if (!r) throw new Error("Não achei a pasta do sistema no SharePoint. Confira o link em config.js.");
    raiz = { drive: r.parentReference.driveId, id: r.id };
    const eu = document.getElementById("eu");
    eu.value = conta.name; eu.readOnly = true; eu.title = conta.username;
    return true;
  },
};
})();
