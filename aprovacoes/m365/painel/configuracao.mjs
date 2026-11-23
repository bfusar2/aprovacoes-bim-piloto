import { identifier } from "../../trimble.mjs";
import { graphSettings } from "./graph.mjs";
import { projectSettings } from "./model.mjs";

const fields = new Set([
  "region", "sharePointSite", "intakeList", "boardList", "modelo", "codePrefix",
  "revisionSeparator", "revisionPrefix", "revisionDigits", "formats", "folders", "entra", "identidades",
]);
const entraFields = new Set([
  "tenantId", "clientId", "siteCollectionId", "webId", "solicitacoesId", "entregasId",
]);
const key = (projectId) => "app-trimble.aprovacoes.config.v1." + identifier(projectId);

function storage(name) {
  try { return globalThis[name]; }
  catch { return undefined; }
}

function onlyKeys(object, allowed) {
  return object && typeof object === "object" && !Array.isArray(object) &&
    Object.keys(object).every((name) => allowed.has(name));
}

function validate(projectId, raw) {
  if (!onlyKeys(raw, fields) || !onlyKeys(raw.folders, new Set(["wip"])) ||
      !onlyKeys(raw.folders.wip, new Set(["id", "name"])) ||
      (raw.entra !== undefined && !onlyKeys(raw.entra, entraFields))) {
    throw new Error("Use somente os campos de configuração do projeto. Cole o modelo sem senhas, tokens ou e-mails.");
  }
  const settings = projectSettings({ [projectId]: raw }, projectId);
  if (raw.entra) graphSettings(settings);
  return settings;
}

export function importSettings(projectId, json, local = storage("localStorage"), session = storage("sessionStorage")) {
  if (typeof json !== "string" || json.length > 12000) throw new Error("O JSON da configuração excede o tamanho permitido.");
  let parsed;
  try { parsed = JSON.parse(json); }
  catch { throw new Error("O JSON da configuração está inválido."); }
  const raw = parsed && Object.hasOwn(parsed, projectId) ? parsed[projectId] : parsed;
  const settings = validate(projectId, raw);
  const value = JSON.stringify(raw);
  const name = key(projectId);
  try {
    if (local) {
      local.setItem(name, value);
      try { session?.removeItem(name); } catch { /* A preferência persistente tem prioridade. */ }
      return { settings, persistent: true };
    }
  } catch { /* Tenta o armazenamento da aba quando o navegador restringe a persistência. */ }
  try {
    if (session) {
      session.setItem(name, value);
      return { settings, persistent: false };
    }
  } catch { /* A mensagem abaixo indica a restrição do navegador. */ }
  throw new Error("O navegador bloqueou o armazenamento desta extensão. Permita o armazenamento do painel para salvar a configuração.");
}

export function localSettings(projectId, local = storage("localStorage"), session = storage("sessionStorage")) {
  const name = key(projectId);
  for (const store of [session, local]) {
    let saved;
    try { saved = store?.getItem(name); }
    catch { continue; }
    if (saved !== null && saved !== undefined) return JSON.parse(saved);
  }
  return null;
}

export function removeSettings(projectId, local = storage("localStorage"), session = storage("sessionStorage")) {
  const name = key(projectId);
  for (const store of [local, session]) {
    try { store?.removeItem(name); }
    catch { /* Um armazenamento indisponível também não pode ser usado para leitura. */ }
  }
}

export async function loadSettings(projectId, fetcher = fetch) {
  const local = localSettings(projectId);
  if (local) return validate(projectId, local);
  const response = await fetcher("./projetos.json", { cache: "no-store", credentials: "omit", redirect: "error" });
  if (!response.ok) throw new Error("Falha ao consultar a configuração pública dos projetos.");
  const projects = await response.json();
  return Object.hasOwn(projects, projectId) ? projectSettings(projects, projectId) : null;
}

export function settingsTemplate(region) {
  const value = typeof region === "string" ? region.toLowerCase().replace(/[^a-z]/g, "") : "";
  return {
    region: ["northamerica", "europe", "unitedkingdom", "asia", "australia"].includes(value) ? value : "northamerica",
    sharePointSite: "https://empresa.sharepoint.com/sites/BIM",
    intakeList: "BIMSolicitacoes",
    boardList: "BIMEntregas",
    modelo: "ENTREGA|1",
    codePrefix: "PRJ-",
    revisionSeparator: "_",
    revisionPrefix: "R",
    revisionDigits: 2,
    formats: ["ifc", "dwg", "pdf"],
    folders: { wip: { id: "ID_PASTA_EM_ELABORACAO", name: "Em elaboração" } },
  };
}
