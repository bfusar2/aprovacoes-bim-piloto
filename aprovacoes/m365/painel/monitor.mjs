import { identifier } from "../../trimble.mjs";
import { loadSettings } from "./configuracao.mjs";
import { commandKey, graphLists, graphScope, graphSettings } from "./graph.mjs";
import { commandState, countDeliveries, officialEvent, stageLabel, stages } from "./monitor-model.mjs";

const byId = (name) => document.getElementById(name);
const status = (message) => { byId("monitor-status").textContent = message; };
const time = new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" });
const commandLabels = {
  applied: "Confirmado no histórico",
  notice: "Aviso no SharePoint",
  waiting: "Aguardando resultado",
  overdue: "Sem resultado há mais de 15 minutos",
  undated: "Sem data de envio",
};
const state = { epoch: 0, dossiers: new Map() };

function text(tag, value) {
  const element = document.createElement(tag);
  element.textContent = value;
  return element;
}

function dateText(value) {
  const date = Date.parse(value);
  return Number.isFinite(date) ? time.format(date) : "Data indisponível";
}

function cell(value, description) {
  const element = text("td", value);
  if (description) element.append(text("small", description));
  return element;
}

function row(...columns) {
  const element = document.createElement("tr");
  element.append(...columns);
  return element;
}

function itemAction(label, key, request) {
  const element = document.createElement("td");
  const button = text("button", label);
  button.type = "button";
  button.classList.add("secondary");
  button.addEventListener("click", () => void readDossier(key, request));
  element.append(button);
  return element;
}

function renderDeliveries() {
  const rows = [...deliveries.rows.values()];
  const counts = countDeliveries(rows);
  for (const [name, value] of Object.entries(counts)) byId("monitor-" + name).textContent = String(value);
  const filter = byId("monitor-stage").value;
  const visible = rows.filter((entry) => !filter || entry.Estado === filter);
  byId("monitor-deliveries").replaceChildren(...visible.map((entry) => row(
    cell(entry.Title || entry.Chave, entry.Title ? entry.Chave : ""),
    cell(stageLabel(entry.Estado)),
    itemAction("Ver dossiê", entry.Chave),
  )));
  deliveries.summary(rows.length + " entrega(s) carregada(s)" + (filter ? " · " + visible.length + " exibida(s) neste filtro." : "."));
}

function renderQueue() {
  const rows = [...intake.rows.values()].sort((a, b) => Number(b.id) - Number(a.id));
  const filter = byId("monitor-queue-filter").value;
  const now = Date.now();
  const classified = rows.map((entry) => [entry, commandState(entry, state.dossiers.get(entry.key), now)]);
  const overdue = classified.filter(([, state]) => state === "overdue").length;
  const visible = classified.filter(([, state]) => !filter || filter === state);
  byId("monitor-queue").replaceChildren(...visible.map(([entry, stage]) => {
    let action = cell("Chave indisponível");
    try { action = itemAction("Conferir ID", commandKey(entry.key), entry); }
    catch { /* Chaves inválidas podem entrar pelo formulário nativo; não são usadas em consultas. */ }
    const result = cell(commandLabels[stage], stage === "notice" ? entry.result : "");
    if (stage === "overdue") result.classList.add("monitor-overdue");
    return row(cell("#" + entry.id, [entry.key, entry.action].filter(Boolean).join(" · ")),
      cell(dateText(entry.created)), result, action);
  }));
  intake.summary(rows.length + " pedido(s) carregado(s) · " + overdue + " sem resultado há mais de 15 min" +
    (filter ? " · " + visible.length + " exibido(s) neste filtro." : "."));
}

const deliveries = {
  rows: new Map(), getKey: (row) => row.Chave,
  read: (cursor) => state.lists.browse(state.project, cursor),
  more: byId("monitor-more-deliveries"),
  summary: (message) => { byId("monitor-deliveries-status").textContent = message; },
  render: renderDeliveries,
};
const intake = {
  rows: new Map(), getKey: (row) => row.id,
  read: (cursor) => state.lists.intake(cursor),
  more: byId("monitor-more-requests"),
  summary: (message) => { byId("monitor-queue-status").textContent = message; },
  render: renderQueue,
};

function resetFeed(feed) {
  feed.rows.clear();
  feed.busy = undefined;
  feed.next = null;
  feed.more.hidden = true;
  feed.more.disabled = false;
  feed.render();
}

function clearSession(note) {
  state.epoch++;
  state.lists = undefined;
  state.historyRequest = undefined;
  state.dossiers.clear();
  resetFeed(deliveries);
  resetFeed(intake);
  byId("monitor-data").hidden = true;
  byId("monitor-history-card").hidden = true;
  byId("monitor-history").replaceChildren();
  status(note);
}

async function inProject(epoch) {
  if (epoch !== state.epoch) return false;
  const project = await state.workspace.project.getProject();
  if (epoch !== state.epoch) return false;
  if (project.id !== state.project) {
    leaveProject();
    return false;
  }
  return true;
}

async function readFeed(feed) {
  if (feed.busy || !state.lists) return;
  const current = Symbol();
  const epoch = state.epoch;
  feed.busy = current;
  feed.more.disabled = true;
  feed.summary("Consultando o SharePoint…");
  try {
    if (!await inProject(epoch)) return;
    const page = await feed.read(feed.next);
    if (!await inProject(epoch)) return;
    for (const entry of page.rows) feed.rows.set(feed.getKey(entry), entry);
    feed.next = page.next;
    feed.more.hidden = !page.next;
    feed.render();
  } catch (error) {
    if (epoch === state.epoch) feed.summary(error.message);
  } finally {
    if (feed.busy === current) {
      feed.busy = undefined;
      feed.more.disabled = false;
    }
  }
}

function history(note, entries = [], heading = "Histórico oficial") {
  byId("monitor-history-title").textContent = heading;
  byId("monitor-history-summary").textContent = note;
  byId("monitor-history").replaceChildren(...entries.map((event) => text("li",
    [dateText(event?.quando), event?.acao, event?.email, event?.dados?.parecer].filter(Boolean).join(" · ")
  )));
}

async function readDossier(key, command) {
  if (!state.lists) return;
  const token = Symbol();
  state.historyRequest = token;
  const epoch = state.epoch;
  byId("monitor-history-card").hidden = false;
  history("Consultando o dossiê de " + key + "…");
  byId("monitor-history-title").focus();
  try {
    if (!await inProject(epoch)) return;
    const dossier = await state.lists.locate(key, state.project);
    if (!await inProject(epoch) || state.historyRequest !== token) return;
    state.dossiers.set(key, dossier);
    renderQueue();
    const event = command && officialEvent(command, dossier);
    const note = command ? (event ? "Pedido #" + command.id + " confirmado por evento oficial." :
      "Pedido #" + command.id + " sem evento correspondente no dossiê.") : "";
    history([note, dossier ? stageLabel(dossier.stage) + " · Geração " + dossier.generation :
      "Nenhum dossiê encontrado para esta chave."].filter(Boolean).join(" "), dossier?.history || [], "Histórico · " + key);
  } catch (error) {
    if (epoch === state.epoch && state.historyRequest === token) history(error.message);
  }
}

async function refresh() {
  if (!state.lists || deliveries.busy || intake.busy) return;
  state.historyRequest = undefined;
  state.dossiers.clear();
  byId("monitor-history-card").hidden = true;
  resetFeed(deliveries);
  resetFeed(intake);
  await Promise.allSettled([readFeed(deliveries), readFeed(intake)]);
}

function leaveProject() {
  byId("monitor-links").hidden = true;
  byId("monitor-login").disabled = true;
  clearSession("O projeto ou a sessão do Trimble mudou.");
  window.location.replace("./index.html");
}

async function start() {
  if (window.parent === window) throw new Error("Abra o acompanhamento pelo projeto no Trimble Connect Web.");
  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://components.connect.trimble.com/trimble-connect-workspace-api/index.js";
    script.onload = resolve;
    script.onerror = () => reject(new Error("Falha ao carregar a biblioteca do Trimble."));
    document.head.append(script);
  });
  const invalidationEvents = new Set(["extension.sessionInvalid", "extension.sessionLogOut", "project.onChanged"]);
  state.workspace = await window.TrimbleConnectWorkspace.connect(window.parent, (event) => {
    if (invalidationEvents.has(event)) leaveProject();
  }, 30000);
  state.project = identifier((await state.workspace.project.getProject()).id);
  const settings = await loadSettings(state.project);
  if (!settings) throw new Error("Salve a configuração local do projeto na tela Submeter revisão.");
  byId("monitor-board").href = settings.boardUrl;
  byId("monitor-intake").href = settings.formUrl.replace(/NewForm\.aspx$/, "AllItems.aspx");
  byId("monitor-links").hidden = false;
  state.graph = graphSettings(settings);
  if (!window.msal) throw new Error("A biblioteca Microsoft de autenticação não carregou. Consulte as listas no SharePoint.");
  state.msal = new window.msal.PublicClientApplication({
    auth: {
      clientId: state.graph.clientId,
      authority: "https://login.microsoftonline.com/" + state.graph.tenantId,
      redirectUri: new URL("./retorno.html", location.href).href,
    },
    cache: { cacheLocation: "memoryStorage" },
  });
  await state.msal.initialize();
  byId("monitor-login").disabled = false;
  status("Conecte a conta corporativa para consultar os registros do projeto.");
}

async function connect() {
  const login = byId("monitor-login");
  login.disabled = true;
  clearSession("Conectando à conta Microsoft…");
  const epoch = state.epoch;
  try {
    const result = await state.msal.loginPopup({ scopes: [graphScope], prompt: "select_account" });
    const account = result.account;
    if (!await inProject(epoch)) return;
    if (account?.tenantId?.toLowerCase() !== state.graph.tenantId.toLowerCase()) {
      throw new Error("Use uma conta no diretório Microsoft configurado.");
    }
    state.lists = graphLists(state.graph, async () => {
      try {
        return (await state.msal.acquireTokenSilent({ scopes: [graphScope], account })).accessToken;
      } catch { throw new Error("A sessão Microsoft precisa ser renovada. Clique em Conectar conta Microsoft."); }
    });
    byId("monitor-data").hidden = false;
    status("Conta Microsoft conectada · " + account.username);
    await refresh();
  } catch (error) {
    if (epoch === state.epoch) status(error.message);
  } finally {
    if (epoch === state.epoch) login.disabled = false;
  }
}

for (const [stage, label] of Object.entries(stages)) byId("monitor-stage").add(new Option(label, stage));
byId("monitor-stage").addEventListener("change", renderDeliveries);
byId("monitor-queue-filter").addEventListener("change", renderQueue);
byId("monitor-login").addEventListener("click", () => void connect());
byId("monitor-refresh").addEventListener("click", () => void refresh());
deliveries.more.addEventListener("click", () => void readFeed(deliveries));
intake.more.addEventListener("click", () => void readFeed(intake));
start().catch((error) => status(error.message));
