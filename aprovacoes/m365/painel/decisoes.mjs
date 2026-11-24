import { identifier, trimble, core } from "../../trimble.mjs";
import { capture } from "./model.mjs";
import { loadSettings } from "./configuracao.mjs";
import { graphLists, graphSettings, graphScope, steps, returnable, assertActual, compareCopies, describeCopyEvidence, findDestination, reference, decision, commandKey, verifyOperator, dossierIdentities } from "./graph.mjs";
import { PilotArrivalPending, pilotArrival, pilotBuild, pilotCommand, pilotStages } from "./piloto.mjs";

const byId = (id) => document.getElementById(id);
const status = (text) => { byId("status").textContent = text; };
const state = {};
let busy = false;
byId("panel-build").textContent = pilotBuild;

function storage(key, value) {
  const full = key + "." + state.project;
  try {
    if (value === undefined) return JSON.parse(sessionStorage.getItem(full));
    sessionStorage.setItem(full, JSON.stringify(value));
  } catch { return null; }
}

function lock(locked) {
  for (const node of document.querySelectorAll("main button, main input, main select, main textarea")) node.disabled = locked;
  byId("login").disabled = locked || !state.msal;
}
async function perform(operation) {
  if (busy) return;
  busy = true;
  lock(true);
  try { await operation(); }
  catch (error) { status(error.message || "A operação falhou."); }
  finally { busy = false; lock(false); }
}

function applyTrimbleToken(value) {
  if (value === "denied") {
    state.core = undefined;
    state.trimbleRequest = undefined;
    status("O acesso ao Trimble foi negado. A consulta ao SharePoint continua disponível.");
  } else if (typeof value === "string" && value && !["granted", "pending"].includes(value) && state.settings) {
    state.trimbleRequest = trimble(state.settings.region, value);
    state.core = core(state.trimbleRequest, state.project, state.settings);
  }
}
async function confirmProject() {
  const project = await state.workspace.project.getProject();
  if (project.id !== state.project) {
    window.location.replace("./index.html");
    throw new Error("O projeto ativo mudou. Abra o novo projeto.");
  }
}
// identities: mapa administrado conta Microsoft → e-mail no Trimble. Em decisões vem do modelo congelado no
// dossiê. Na submissão, que ainda não tem dossiê, vem da configuração local do projeto.
async function trimbleApi(identities = {}) {
  await confirmProject();
  if (!state.core) applyTrimbleToken(await state.workspace.extension.requestPermission("accesstoken"));
  if (!state.core) throw new Error("Autorize o acesso ao Trimble e repita a conferência.");
  const operadorTrimble = await verifyOperator(state.trimbleRequest, state.project, state.account?.username, identities);
  return { api: state.core, operadorTrimble };
}

async function start() {
  if (window.parent === window) throw new Error("Abra a extensão pelo projeto no Trimble Connect Web.");
  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://components.connect.trimble.com/trimble-connect-workspace-api/index.js";
    script.onload = resolve;
    script.onerror = () => reject(new Error("Falha ao carregar a biblioteca do Trimble."));
    document.head.append(script);
  });
  state.workspace = await window.TrimbleConnectWorkspace.connect(window.parent, (event, args) => {
    if (event === "extension.accessToken") applyTrimbleToken(args?.data);
    if (["extension.sessionInvalid", "extension.sessionLogOut"].includes(event)) window.location.replace("./index.html");
    if (event === "project.onChanged" && state.project) void confirmProject().catch(() => window.location.replace("./index.html"));
  }, 30000);
  state.project = identifier((await state.workspace.project.getProject()).id);
  state.settings = await loadSettings(state.project);
  if (!state.settings) throw new Error("Salve a configuração local do projeto na tela Submeter revisão.");
  state.graph = graphSettings(state.settings);
  if (!window.msal) throw new Error("A biblioteca Microsoft de autenticação não carregou. Use o formulário nativo.");
  const msal = new window.msal.PublicClientApplication({
    auth: { clientId: state.graph.clientId, authority: "https://login.microsoftonline.com/" + state.graph.tenantId,
      redirectUri: new URL("./retorno.html", location.href).href },
    cache: { cacheLocation: "memoryStorage" },
  });
  await msal.initialize();
  state.msal = msal;
  byId("login").disabled = false;
  status("Conecte a conta corporativa que tem acesso às listas do projeto.");
  byId("requests").href = state.settings.formUrl.replace(/NewForm\.aspx$/, "AllItems.aspx");
  byId("board-link").href = state.settings.boardUrl;
}

function resetCopy() {
  state.destination = undefined;
  byId("found-copy").textContent = "";
  byId("check").checked = false;
}
function actionControls() {
  const action = byId("action").value;
  const copy = ["compartilhar", "publicar"].includes(action);
  byId("destination-label").hidden = !copy;
  byId("destination").required = copy;
  byId("destination").value = "";
  resetCopy();
  const descriptions = {
    verificar: "Conferi a abertura, os metadados e a integridade da versão indicada.",
    revisar: "Conferi a compatibilidade técnica com os requisitos e referências do modelo.",
    compartilhar: "Copiei a versão aprovada na pasta compartilhada e conferi o conteúdo da cópia.",
    publicar: "Copiei a versão autorizada na pasta publicada e conferi o conteúdo da cópia.",
    conferir: "Conferi de forma independente a cópia publicada, a versão e sua correspondência com a entrega aceita.",
    devolver: "Identifiquei no parecer a divergência referente à revisão deste dossiê.",
    encerrar: "Solicito o encerramento deste dossiê pelo motivo registrado no parecer.",
  };
  byId("check-label").textContent = descriptions[action] || "Analisei a versão indicada e confirmo a decisão registrada no parecer.";
}
function showDelivery(doc) {
  state.delivery = doc;
  byId("dossier").hidden = !doc;
  if (!doc) return;
  const source = reference(doc);
  const rule = steps[doc.stage];
  const pilotStage = pilotStages[doc.stage];
  byId("pilot-decision").hidden = doc.stage !== "piloto_decisao";
  byId("pilot-resume").hidden = !["piloto_mover", "piloto_avisar"].includes(doc.stage);
  byId("pilot-reconcile").hidden = true;
  byId("pilot-reconcile").reset();
  byId("pilot-state").hidden = !pilotStage;
  byId("pilot-state").textContent = doc.stage === "piloto_envio_pendente"
    ? "Se o envio permanecer pendente, confira a execução no Power Automate e a caixa de enviados antes de qualquer reenvio." : "";
  byId("pilot-decision").reset();
  byId("title").textContent = doc.key;
  byId("stage").textContent = "Etapa · " + (pilotStage || doc.stage) + " | Geração " + doc.generation;
  byId("reference").textContent = doc.record.origem.nome + " · Arquivo " + source.arquivo + " · Versão " + source.versao + " · Pasta " + source.pasta;
  byId("criteria").textContent = "Critérios do modelo · " + doc.record.modelo.criterios;
  const eligible = rule ? doc.record.modelo.papeis[rule[1]] || [] : [];
  byId("owners").textContent = rule ? "Responsáveis pela etapa (" + rule[1] + ") · " + eligible.join(", ") : pilotStage ? "" : "Etapa encerrada.";
  const items = doc.history.map((event) => {
    const li = document.createElement("li");
    const evidence = event.dados?.comparacao;
    const operator = event.dados?.operadorTrimble;
    li.textContent = [event.quando, event.email, event.acao, event.dados?.parecer,
      operator && "Operador Trimble declarado " + operator,
      evidence && "Conferência declarada · " + describeCopyEvidence(evidence),
      evidence?.hashOrigem && "Hash de origem " + evidence.hashOrigem,
      evidence?.hashDestino && "Hash de destino " + evidence.hashDestino,
      event.dados?.versaoDestino && "Versão de destino " + event.dados.versaoDestino,
      event.dados?.aviso,
    ].filter(Boolean).join(" · ");
    return li;
  });
  byId("history").replaceChildren(...items);
  // The SharePoint Author identity and the frozen rules are authoritative in the flow.
  byId("decision").hidden = !rule;
  if (rule) {
    byId("decision").reset();
    byId("action").replaceChildren(new Option(rule[2], rule[0]));
    if (returnable.includes(doc.stage)) byId("action").add(new Option("Devolver para ajustes", "devolver"));
    byId("action").add(new Option("Encerrar dossiê (autor ou líder)", "encerrar"));
    actionControls();
  }
}
byId("action").onchange = actionControls;
byId("destination").oninput = resetCopy;
byId("find-copy").onclick = () => perform(async () => {
  resetCopy();
  const { api } = await trimbleApi(dossierIdentities(state.delivery));
  const item = await findDestination(api, state.delivery, byId("action").value, byId("destination").value.trim());
  state.destination = item;
  byId("destination").value = item.id;
  byId("found-copy").textContent = item.name + " · Arquivo " + item.id + " · Versão " + item.versionId + " · Pasta " + item.parentId;
  status("Confira a identidade da cópia e seu conteúdo antes de registrar a transferência.");
});

function showReceipt(receipt) {
  state.receipt = receipt;
  byId("receipt").hidden = !receipt;
  if (receipt) byId("receipt-info").textContent = "Pedido " + receipt.id + " · " + receipt.key + ". Consulte o histórico para confirmar a aplicação.";
}
async function send(command) {
  const receipt = await state.lists.submit(command);
  storage("bim.recibo", receipt);
  showReceipt(receipt);
  showDelivery(null);
  state.next = null;
  byId("delivery-list").replaceChildren();
  byId("more").hidden = true;
  byId("proposal").hidden = true;
  byId("key").value = command.Chave;
  status("Pedido registrado na fila do SharePoint. O Power Automate processa a decisão.");
}

byId("login").onclick = () => perform(async () => {
  // The popup is initiated directly by the user gesture.
  const auth = await state.msal.loginPopup({ scopes: [graphScope], prompt: "select_account" });
  if (auth.account?.tenantId?.toLowerCase() !== state.graph.tenantId.toLowerCase()) throw new Error("Use uma conta no diretório Microsoft configurado.");
  state.account = auth.account;
  state.lists = graphLists(state.graph, async () => {
    try {
      return (await state.msal.acquireTokenSilent({ scopes: [graphScope], account: state.account })).accessToken;
    } catch { throw new Error("A sessão Microsoft precisa ser renovada. Clique em Conectar conta Microsoft."); }
  });
  showDelivery(null);
  byId("lookup").hidden = false;
  state.next = null;
  byId("delivery-list").replaceChildren();
  byId("more").hidden = true;
  state.proposal = storage("bim.proposta");
  byId("proposal").hidden = !state.proposal;
  if (state.proposal) {
    byId("proposal-name").textContent = state.proposal.Title + " · " + state.proposal.Chave;
    byId("key").value = state.proposal.Chave;
  }
  showReceipt(storage("bim.recibo"));
  status("Conta Microsoft conectada · " + auth.account.username);
});

async function loadDelivery(key) {
  showDelivery(null);
  await confirmProject();
  const doc = await state.lists.locate(key, state.project);
  byId("key").value = key;
  showDelivery(doc);
  status(doc ? "Dossiê oficial lido no SharePoint." : "Ainda não há um dossiê com essa chave no projeto.");
}
byId("load").onclick = () => perform(() => loadDelivery(commandKey(byId("key").value.trim())));
async function browse(more) {
  await confirmProject();
  if (!more) byId("delivery-list").replaceChildren();
  const page = await state.lists.browse(state.project, more ? state.next : null);
  state.next = page.next;
  for (const row of page.rows) {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.className = "secondary";
    button.textContent = row.Chave + " · " + row.Estado;
    button.onclick = () => perform(() => loadDelivery(row.Chave));
    item.append(button);
    byId("delivery-list").append(item);
  }
  byId("more").hidden = !page.next;
  status(page.rows.length ? "Selecione uma entrega para consultar o dossiê." : "Esta página não contém entregas do projeto.");
}
byId("browse").onclick = () => perform(() => browse(false));
byId("more").onclick = () => perform(() => browse(true));

byId("send-proposal").onclick = () => perform(async () => {
  const item = JSON.parse(state.proposal.DadosJson);
  const { api } = await trimbleApi(state.settings.identidades);
  const actual = await api.file(identifier(item.arquivo));
  assertActual(item, actual, state.project, item.nome);
  const command = capture(state.project, state.settings, actual, actual);
  if (command.Chave !== state.proposal.Chave || JSON.parse(command.DadosJson).modelo !== item.modelo) throw new Error("A regra de submissão mudou. Confira a revisão no painel.");
  await send(command);
});

byId("decision").onsubmit = (event) => {
  event.preventDefault();
  void perform(async () => {
    const doc = state.delivery;
    const action = byId("action").value;
    const check = byId("check").checked;
    const note = byId("note").value;
    await confirmProject();
    let destination;
    let comparacao;
    let operadorTrimble;
    if (!["devolver", "encerrar"].includes(action)) {
      const session = await trimbleApi(dossierIdentities(doc));
      const { api } = session;
      operadorTrimble = session.operadorTrimble;
      const ref = reference(doc);
      assertActual(ref, await api.file(ref.arquivo), state.project, doc.record.origem.nome);
      if (action === "compartilhar" || action === "publicar") {
        if (!state.destination) throw new Error("Use Localizar e conferir cópia para ver a versão de destino.");
        destination = await api.file(identifier(byId("destination").value.trim()));
        if (destination.versionId !== state.destination.versionId || destination.id !== state.destination.id) {
          resetCopy();
          throw new Error("A versão da cópia mudou. Localize-a novamente antes de registrar a transferência.");
        }
        comparacao = await compareCopies(api, ref, { arquivo: destination.id, versao: destination.versionId, pasta: destination.parentId });
      } else if (action === "conferir") {
        const shared = doc.record.compartilhado;
        assertActual(shared, await api.file(shared.arquivo, shared.versao), state.project, doc.record.origem.nome);
        comparacao = await compareCopies(api, doc.record.compartilhado, ref);
      }
    }
    await send(decision(doc, action, note, check, destination, { operadorTrimble, comparacao }));
    if (comparacao) status("Pedido registrado na fila. " + describeCopyEvidence(comparacao));
  });
};

async function waitForPilot(command, terminal) {
  const receipt = await state.lists.submit(command);
  storage("bim.recibo", receipt);
  showReceipt(receipt);
  status("Aguardando o Power Automate registrar a etapa no SharePoint.");
  for (let attempt = 0; attempt < 25; attempt++) {
    await confirmProject();
    const doc = await state.lists.locate(receipt.key, state.project);
    if (doc?.history.some((event) => String(event.comando) === receipt.id)) {
      showDelivery(doc);
      if (!terminal || ["publicado", "devolvido"].includes(doc.stage)) return doc;
    }
    const result = await state.lists.result(receipt.id);
    if (/^(Recusado|Requer conferência)/.test(result)) throw new Error(result);
    await new Promise((resolve) => setTimeout(resolve, 2000));
  }
  throw new Error("O processamento continua pendente. Confira o dossiê para retomar na etapa registrada.");
}

function moveCheckpoint(doc) {
  const origin = doc.record.origem;
  const key = ["bim.piloto.movimento", state.project, origin.arquivo, origin.versao, doc.record.modelo.pastas.published].join(".");
  return {
    read() {
      try { return localStorage.getItem(key) || storage(key); }
      catch { return storage(key); }
    },
    write(version) {
      storage(key, version);
      try { localStorage.setItem(key, version); } catch { /* Mantém o recibo da aba. */ }
    },
  };
}

async function resumePilot(doc, reviewedVersion) {
  await confirmProject();
  doc = await state.lists.locate(doc.key, state.project);
  showDelivery(doc);
  if (["publicado", "devolvido"].includes(doc?.stage)) {
    status("Decisão e envio registrados no SharePoint.");
    return;
  }
  if (!["piloto_mover", "piloto_avisar"].includes(doc?.stage)) throw new Error("Consulte a etapa atual do piloto.");
  const { api, operadorTrimble } = await trimbleApi(dossierIdentities(doc));
  const origin = doc.record.origem;
  if (doc.stage === "piloto_mover") {
    status("Aprovação registrada. Conferindo a movimentação para Publicado.");
    let destination;
    try { destination = await pilotArrival(api, doc, moveCheckpoint(doc), reviewedVersion); }
    catch (error) {
      if (error instanceof PilotArrivalPending || reviewedVersion) byId("pilot-reconcile").hidden = false;
      throw error;
    }
    await confirmProject();
    const evidence = reviewedVersion ? "Chegada conferida no histórico do Trimble pelo operador." : "Chegada confirmada no Trimble.";
    await waitForPilot(pilotCommand(doc, "piloto_chegada", evidence, operadorTrimble, destination), true);
  } else {
    const expected = doc.record.publicado || origin;
    assertActual(expected, await api.file(expected.arquivo), origin.projeto, origin.nome);
    await waitForPilot(pilotCommand(doc, "piloto_avisar", "Localização reconfirmada no Trimble.", operadorTrimble), true);
  }
  status("Decisão e envio registrados no SharePoint.");
}

byId("pilot-decision").onsubmit = (event) => {
  event.preventDefault();
  const reject = event.submitter === byId("pilot-reject");
  const note = byId("pilot-note").value.trim();
  void perform(async () => {
    if (reject && !note) throw new Error("Informe o motivo da recusa.");
    const doc = state.delivery;
    const { api, operadorTrimble } = await trimbleApi(dossierIdentities(doc));
    const origin = doc.record.origem;
    assertActual(origin, await api.file(origin.arquivo), state.project, origin.nome);
    const next = await waitForPilot(pilotCommand(doc, reject ? "piloto_recusar" : "piloto_aprovar",
      note || "Revisão aprovada.", operadorTrimble), reject);
    if (reject) status("Recusa e envio registrados no SharePoint.");
    else await resumePilot(next);
  });
};

byId("pilot-resume").onclick = () => perform(() => resumePilot(state.delivery));
byId("pilot-reconcile").onsubmit = (event) => {
  event.preventDefault();
  const version = byId("pilot-version").value.trim();
  void perform(() => resumePilot(state.delivery, identifier(version)));
};

byId("verify").onclick = () => perform(async () => {
  await confirmProject();
  const { key, id } = state.receipt;
  const doc = await state.lists.locate(key, state.project);
  if (doc?.history.some((event) => String(event.comando) === id)) {
    showDelivery(doc);
    byId("receipt-info").textContent = "Aplicação confirmada pelo evento do pedido " + id + " no dossiê " + key + ".";
    status("Histórico oficial atualizado.");
  } else {
    const outcome = await state.lists.result(id);
    byId("receipt-info").textContent = "Pedido " + id + " ainda sem evento no dossiê. " + (outcome || "Aguardando processamento.");
    status("Consulte o parecer da fila e atualize novamente.");
  }
});
start().catch((error) => status(error.message));
