import { core, trimble } from "../../trimble.mjs";
import { capture } from "./model.mjs";
import { loadSettings, importSettings, localSettings, removeSettings, settingsTemplate } from "./configuracao.mjs";

const byId = (id) => document.getElementById(id);
let workspace;
let project;
let settings;
let token;
let api;
let files = [];
let sequence = 0;
let proposalKey;
let configRequest = 0;

function cacheProposal(command) {
  if (!proposalKey) return;
  try {
    if (command) sessionStorage.setItem(proposalKey, JSON.stringify(command));
    else sessionStorage.removeItem(proposalKey);
  } catch { /* O formulário nativo continua disponível quando o iframe restringe o armazenamento. */ }
}

function message(text) {
  byId("connection").textContent = text;
}

function clearSubmission() {
  cacheProposal();
  byId("submission").hidden = true;
  byId("new-form").removeAttribute("href");
  byId("fields").replaceChildren();
  byId("copy-status").textContent = "";
}

function clearSession(text) {
  ++sequence;
  token = undefined;
  api = undefined;
  files = [];
  clearSubmission();
  byId("selection").hidden = true;
  byId("board").hidden = true;
  byId("board").removeAttribute("href");
  byId("team").hidden = true;
  message(text);
}

async function refreshFiles() {
  if (!api) return;
  const request = ++sequence;
  clearSubmission();
  files = [];
  byId("file").replaceChildren(new Option("Carregando…", ""));
  byId("inspect").disabled = true;
  byId("selection").hidden = false;
  message("Consultando a pasta de elaboração no Trimble…");
  try {
    const loaded = await api.items(settings.folders.wip.id);
    if (request !== sequence) return;
    files = loaded;
    byId("file").replaceChildren(new Option(files.length ? "Selecione um arquivo" : "Pasta sem arquivos disponíveis", ""));
    for (const item of files) byId("file").add(new Option(item.name, item.id));
    message("Projeto conectado. Os dados da revisão serão obtidos do Trimble.");
  } catch (error) {
    if (request !== sequence) return;
    byId("file").replaceChildren(new Option("Listagem indisponível", ""));
    message(error.message);
  }
}

function onToken(value) {
  if (value === "denied") return clearSession("O acesso à sessão foi negado. Revise a permissão da extensão nas configurações do Trimble.");
  if (!project || !settings || typeof value !== "string" || !value ||
      ["pending", "granted"].includes(value) || token === value) return;
  token = value;
  api = core(trimble(settings.region, token), project.id, settings);
  byId("board").href = settings.boardUrl;
  byId("board").hidden = false;
  byId("team").hidden = false;
  void refreshFiles();
}

async function onProjectChange() {
  if (!workspace || !project) return;
  const previous = project.id;
  const request = ++configRequest;
  settings = undefined;
  clearSession("Conferindo o projeto ativo no Trimble…");
  byId("monitor").hidden = true;
  byId("configuration").hidden = true;
  try {
    const current = await workspace.project.getProject();
    if (!project || request !== configRequest) return;
    if (current.id !== previous) {
      location.reload();
      return;
    }
    project = current;
    byId("configuration").hidden = false;
    renderConfig();
    await refreshSettings();
  } catch {
    if (request === configRequest) clearSession("O projeto do Trimble mudou. Reabra a extensão no projeto desejado.");
  }
}

function onEvent(event, args) {
  if (event === "extension.accessToken") onToken(args?.data);
  if (event === "project.onChanged") void onProjectChange();
  if (event === "extension.sessionLogOut" || event === "extension.sessionInvalid") {
    configRequest++;
    project = undefined;
    settings = undefined;
    clearSession("A sessão expirou. Entre no Trimble e reabra a extensão.");
    byId("monitor").hidden = true;
    byId("configuration").hidden = true;
  }
}

async function activateSettings(next) {
  const request = ++configRequest;
  clearSession("Informe a configuração local para acessar os arquivos deste projeto.");
  settings = next;
  byId("monitor").hidden = !settings;
  byId("config-details").open = !settings;
  if (!settings) return;
  byId("folder").textContent = "Pasta · " + settings.folders.wip.name;
  message("Autorize a leitura da sessão para consultar os arquivos no Trimble.");
  try {
    const permission = await workspace.extension.requestPermission("accesstoken");
    if (request === configRequest) onToken(permission);
  } catch (error) {
    if (request === configRequest && project) clearSession(error.message);
  }
}

async function refreshSettings() {
  const request = ++configRequest;
  const id = project.id;
  try {
    const next = await loadSettings(id);
    if (request === configRequest && project?.id === id) await activateSettings(next);
  } catch (error) {
    if (request === configRequest && project?.id === id) clearSession(error.message);
  }
}

function renderConfig() {
  let saved;
  try { saved = localSettings(project.id); }
  catch { /* Permite corrigir dados locais inválidos pelo painel. */ }
  byId("config-json").value = JSON.stringify(saved || settingsTemplate(project.location || project.region), null, 2);
  byId("remove-config").hidden = !saved;
}

byId("config-form").onsubmit = async (event) => {
  event.preventDefault();
  if (!project) return;
  try {
    const { settings: next, persistent } = importSettings(project.id, byId("config-json").value);
    renderConfig();
    byId("config-status").textContent = persistent ? "Configuração salva neste navegador." : "Configuração salva nesta aba.";
    await activateSettings(next);
  } catch (error) { byId("config-status").textContent = error.message; }
};
byId("remove-config").onclick = async () => {
  if (!project) return;
  removeSettings(project.id);
  renderConfig();
  byId("config-status").textContent = "Configuração local removida.";
  await refreshSettings();
};

async function connect() {
  if (window.parent === window) {
    message("Abra este painel pelo menu do projeto no Trimble Connect Web.");
    return;
  }
  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://components.connect.trimble.com/trimble-connect-workspace-api/index.js";
    script.onload = resolve;
    script.onerror = () => reject(new Error("Falha ao carregar a biblioteca do Trimble."));
    document.head.append(script);
  });
  workspace = await window.TrimbleConnectWorkspace.connect(window.parent, onEvent, 30000);
  await workspace.ui.setMenu({ title: "Aprovações BIM", command: "aprovar-bim", icon: new URL("../../icon.svg", location.href).href });
  project = await workspace.project.getProject();
  proposalKey = "bim.proposta." + project.id;
  byId("project").textContent = "Projeto · " + (project.name || project.id);
  byId("config-project").textContent = "Projeto · " + (project.name || project.id);
  byId("configuration").hidden = false;
  byId("config-details").open = true;
  renderConfig();
  await refreshSettings();
}

function showFields(command) {
  const fields = byId("fields");
  fields.replaceChildren();
  for (const [field, value] of Object.entries(command)) {
    const row = document.createElement("div");
    const label = document.createElement("dt");
    label.textContent = field;
    const content = document.createElement("dd");
    content.textContent = String(value);
    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "secondary";
    copy.textContent = "Copiar";
    copy.setAttribute("aria-label", "Copiar " + field);
    copy.onclick = async () => {
      try {
        await navigator.clipboard.writeText(String(value));
        byId("copy-status").textContent = field + " copiado.";
      } catch {
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(content);
        selection?.removeAllRanges();
        selection?.addRange(range);
        byId("copy-status").textContent = "Valor selecionado. Use Ctrl+C ou o comando Copiar do dispositivo.";
      }
    };
    row.append(label, content, copy);
    fields.append(row);
  }
}

byId("file").onchange = () => {
  ++sequence;
  clearSubmission();
  byId("inspect").disabled = !byId("file").value;
};
byId("refresh").onclick = refreshFiles;
byId("inspect").onclick = async () => {
  const selected = files.find((item) => item.id === byId("file").value);
  if (!selected || !api) return;
  const request = ++sequence;
  clearSubmission();
  byId("inspect").disabled = true;
  try {
    const currentProject = await workspace.project.getProject();
    if (currentProject.id !== project.id) {
      location.reload();
      return;
    }
    const latest = await api.file(selected.id);
    if (request !== sequence) return;
    const command = capture(project.id, settings, selected, latest);
    cacheProposal(command);
    byId("summary").textContent = latest.name + " · Versão Trimble " + latest.versionId;
    showFields(command);
    byId("new-form").href = settings.formUrl;
    byId("submission").hidden = false;
    message("Versão, pasta, nomenclatura e formato conferidos no Trimble.");
  } catch (error) {
    if (request === sequence) message(error.message);
  } finally {
    if (request === sequence) byId("inspect").disabled = false;
  }
};

connect().catch((error) => clearSession(error.message));
