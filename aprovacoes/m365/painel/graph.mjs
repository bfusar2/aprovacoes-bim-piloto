import { identifier } from "../../trimble.mjs";
import { normalizar, trimbleEsperado, validarIdentidades } from "./identidade.mjs";

const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const keyPattern = /^[A-Z0-9_-]{1,100}\|[A-Z0-9_-]{1,24}$/;
const graph = "https://graph.microsoft.com/v1.0";
export const graphScope = "https://graph.microsoft.com/Lists.SelectedOperations.Selected";

// Keys correspond to the transition table frozen into the server-side flow.
export const steps = {
  qualidade: ["verificar", "qualidade", "Concluir verificação de qualidade"],
  revisao: ["revisar", "revisor", "Aprovar revisão técnica"],
  compartilhamento: ["compartilhar", "publicador", "Registrar cópia compartilhada"],
  compartilhado: ["entregar", "autor", "Encaminhar entrega formal"],
  autorizacao: ["autorizar", "lider", "Autorizar entrega"],
  aceite: ["aceitar", "contratante", "Aceitar entrega"],
  publicacao: ["publicar", "publicador", "Registrar cópia para conferência"],
  conferencia: ["conferir", "conferente", "Conferir publicação"],
};
export const returnable = ["qualidade", "revisao", "autorizacao", "aceite", "conferencia"];

export function graphSettings(settings) {
  const ids = settings.entra;
  if (!ids) throw new Error("A operação direta via Microsoft Graph ainda precisa ser configurada neste projeto. Use o formulário nativo do SharePoint pelo painel de submissão.");
  if (!["tenantId", "clientId", "siteCollectionId", "webId", "solicitacoesId", "entregasId"].every((key) =>
    typeof ids[key] === "string" && guidPattern.test(ids[key])) || ids.solicitacoesId.toLowerCase() === ids.entregasId.toLowerCase()) {
    throw new Error("Confira os identificadores do aplicativo Entra, do site e das listas.");
  }
  const base = graph + "/sites/" + new URL(settings.sharePointSite).hostname + "," + ids.siteCollectionId + "," + ids.webId + "/lists/";
  return { tenantId: ids.tenantId, clientId: ids.clientId,
    intakeItems: base + ids.solicitacoesId + "/items", deliveryItems: base + ids.entregasId + "/items" };
}

export function commandKey(value) {
  if (typeof value !== "string" || !keyPattern.test(value)) throw new Error("Informe a chave no formato CÓDIGO|REVISÃO.");
  return value;
}

// Confere se o membro ativo do projeto no Trimble corresponde à conta Microsoft conectada. Sem mapeamento
// administrado, os dois precisam ter o mesmo e-mail. Esta checagem protege contra engano do operador. Ela não
// prova identidade: o formulário nativo do SharePoint não passa por ela, e o fluxo decide só pela identidade
// "Criado por" do SharePoint.
export async function verifyOperator(trimbleRequest, projectId, username, identities = {}) {
  const person = await trimbleRequest("projects/" + identifier(projectId) + "/users/me?minimal=false");
  const email = normalizar(person?.email);
  const expected = trimbleEsperado(identities, username);
  if (person?.status !== "ACTIVE" || !person.id || !email.includes("@") || !normalizar(username) || !expected || email !== expected) {
    throw new Error("A conta Microsoft e o membro ativo deste projeto Trimble precisam usar o mesmo e-mail, ou o administrador do modelo deve registrar o mapeamento em identidades." +
      (normalizar(username) ? " Conta Microsoft: " + normalizar(username) + "." : ""));
  }
  return identifier(person.id);
}

// Mapa de identidades congelado no dossiê. Um mapa inválido bloqueia a decisão em vez de ser ignorado.
export function dossierIdentities(delivery) {
  try { return validarIdentidades(delivery?.record?.modelo?.identidades); }
  catch { throw new Error("O mapa de identidades congelado neste dossiê é inválido. Peça ao responsável pelo modelo para conferi-lo."); }
}

export function graphLists(config, acquire, fetcher = fetch) {
  async function request(url, fields) {
    const accessToken = await acquire();
    let response;
    const failure = "O envio ficou sem confirmação. Confira os pedidos no SharePoint antes de reenviar.";
    try {
      response = await fetcher(url, { method: fields ? "POST" : "GET", redirect: "error", cache: "no-store", credentials: "omit",
        headers: { Authorization: "Bearer " + accessToken, ...(fields && { "Content-Type": "application/json" }) },
        ...(fields && { body: JSON.stringify({ fields }) }), signal: AbortSignal.timeout(20000),
      });
    } catch { throw new Error(fields ? failure : "Falha ao consultar o Microsoft Graph."); }
    if (!response.ok) {
      if (fields && response.status >= 500) throw new Error(failure);
      throw new Error(response.status === 403 ? "O Microsoft Graph negou acesso. Confira as concessões das listas e a permissão desta conta." :
        "Microsoft Graph retornou " + response.status + ". Confira a conta e a configuração do site.");
    }
    try { return await response.json(); } catch { throw new Error(fields ? failure : "Resposta ilegível do Microsoft Graph."); }
  }
  function pageUrl(endpoint, cursor) {
    const list = new URL(endpoint);
    const url = cursor ? new URL(cursor) : list;
    if (url.origin !== list.origin || url.pathname !== list.pathname || url.username || url.password || url.hash) {
      throw new Error("O link de paginação não pertence à lista configurada.");
    }
    return url;
  }
  return {
    async browse(projectId, cursor) {
      const url = pageUrl(config.deliveryItems, cursor);
      if (!cursor) {
        url.searchParams.set("$top", "25");
        url.searchParams.set("$expand", "fields($select=Chave,Title,Estado,Origem)");
      }
      const page = await request(url.href);
      if (!Array.isArray(page.value)) throw new Error("Listagem de entregas indisponível.");
      return {
        rows: page.value.map((item) => item.fields).filter((row) => row?.Origem?.startsWith(projectId + "|") && keyPattern.test(row.Chave)),
        next: page["@odata.nextLink"] || null,
      };
    },
    async intake(cursor) {
      const url = pageUrl(config.intakeItems, cursor);
      if (!cursor) {
        url.searchParams.set("$top", "50");
        url.searchParams.set("$select", "id,createdDateTime");
        url.searchParams.set("$expand", "fields($select=Chave,Acao,Resultado)");
      }
      const page = await request(url.href);
      if (!Array.isArray(page.value)) throw new Error("Listagem de pedidos indisponível.");
      return {
        rows: page.value.flatMap((item) => /^\d+$/.test(String(item?.id)) ? [{
          id: String(item.id), created: item.createdDateTime, key: item.fields?.Chave,
          action: item.fields?.Acao, result: item.fields?.Resultado,
        }] : []),
        next: page["@odata.nextLink"] || null,
      };
    },
    async locate(key, projectId) {
      const url = new URL(config.deliveryItems);
      url.searchParams.set("$filter", "fields/Chave eq '" + commandKey(key) + "'");
      url.searchParams.set("$expand", "fields($select=Title,Chave,Estado,Geracao,DadosJson,HistoricoJson)");
      const result = await request(url.href);
      if (!Array.isArray(result.value) || result.value.length > 1) throw new Error("O registro da entrega precisa ser conferido no SharePoint.");
      if (!result.value.length) return null;
      const row = result.value[0].fields;
      const record = JSON.parse(row.DadosJson);
      if (row.Chave !== key || record.origem?.projeto !== projectId) throw new Error("A entrega pertence a outro projeto ou sua chave foi alterada.");
      const history = JSON.parse(row.HistoricoJson);
      if (!Array.isArray(history) || !Number.isSafeInteger(row.Geracao) || row.Geracao < 1 || !record.modelo?.papeis) {
        throw new Error("Dossiê inválido no SharePoint.");
      }
      return { key, stage: row.Estado, generation: row.Geracao, record, history };
    },
    async submit(command) {
      const { Title, Acao, Chave, Geracao, DadosJson } = command;
      const created = await request(config.intakeItems, { Title, Acao, Chave: commandKey(Chave), Geracao, DadosJson });
      if (!/^\d+$/.test(String(created?.id))) throw new Error("O SharePoint não retornou o ID do pedido. Consulte a lista antes de repetir.");
      return { id: String(created.id), key: Chave };
    },
    async result(id) {
      if (!/^\d+$/.test(String(id))) throw new Error("Número do pedido inválido.");
      const item = await request(config.intakeItems + "/" + id + "?$expand=fields($select=Resultado)");
      return item.fields?.Resultado || "";
    },
  };
}

export function reference(delivery) {
  return delivery.record.publicado || delivery.record.compartilhado || delivery.record.origem;
}

export async function findDestination(api, delivery, action, suppliedId = "") {
  if (action !== "compartilhar" && action !== "publicar") throw new Error("A etapa atual não registra cópia.");
  const target = delivery.record.modelo.pastas[action === "compartilhar" ? "shared" : "published"];
  const { projeto, nome } = delivery.record.origem;
  let item;
  if (suppliedId) item = await api.file(identifier(suppliedId));
  else {
    const files = await api.items(target);
    const matches = files.filter((file) => file.projectId === projeto && file.parentId === target && file.name === nome);
    if (matches.length !== 1) throw new Error("A busca precisa localizar uma única cópia com o mesmo nome na pasta esperada. Confira o documento no Trimble ou informe o ID da cópia.");
    const listed = matches[0];
    item = await api.file(identifier(listed.id));
    if (item.versionId !== listed.versionId) throw new Error("A versão de destino mudou durante a busca. Confira novamente.");
  }
  if (item.id === reference(delivery).arquivo || item.projectId !== projeto || item.parentId !== target || item.name !== nome) {
    throw new Error("Nome, identidade ou pasta da cópia divergente no Trimble.");
  }
  return item;
}

export function assertActual(expected, actual, projectId, name) {
  if (actual.id !== expected.arquivo || actual.versionId !== expected.versao || actual.parentId !== expected.pasta ||
      actual.projectId !== projectId || actual.name !== name) {
    throw new Error("O arquivo atual diverge da versão, do nome ou da pasta do dossiê. Registre a divergência antes de aprovar.");
  }
}

const hashPattern = /^[A-Za-z0-9+/=_-]{4,128}$/;

export function describeCopyEvidence(evidence) {
  if (evidence?.modo === "hash_igual") return "Os hashes informados pelo Trimble para as duas versões coincidem.";
  if (evidence?.modo === "conferencia_humana") return "A correspondência depende da conferência humana registrada.";
  return "";
}

export async function compareCopies(api, source, destination) {
  async function lookup(item) {
    let items;
    try { items = await api.items(item.pasta); }
    catch { return null; } // O acesso aos metadados do arquivo pode existir sem listagem da pasta.
    const matched = items.find((file) => file.id === item.arquivo && file.versionId === item.versao && file.parentId === item.pasta);
    const hash = typeof matched?.hash === "string" ? matched.hash.trim() : "";
    if (!hash) return null;
    return /^[a-f0-9]+$/i.test(hash) ? hash.toLowerCase() : hash;
  }
  const [a, b] = await Promise.all([lookup(source), lookup(destination)]);
  if (a && b && a !== b) throw new Error("Os hashes informados pelo Trimble para as duas versões divergem. Confira o conteúdo antes de prosseguir.");
  const hashOrigem = a && hashPattern.test(a) ? a : "";
  const hashDestino = b && hashPattern.test(b) ? b : "";
  return {
    modo: hashOrigem && hashDestino ? "hash_igual" : "conferencia_humana",
    hashOrigem,
    hashDestino,
  };
}

export function decision(delivery, action, note, checked, destination, evidence = {}) {
  const rule = steps[delivery.stage];
  if (!rule || !(action === rule[0] || action === "encerrar" || (action === "devolver" && returnable.includes(delivery.stage)))) throw new Error("Ação incompatível com a etapa.");
  if (!checked || typeof note !== "string" || !note.trim() || note.trim().length > 1000) throw new Error("Confirme a referência e registre seu parecer (até 1.000 caracteres).");
  const current = reference(delivery);
  const payload = { parecer: note.trim(), arquivo: identifier(current.arquivo), versao: identifier(current.versao), pasta: identifier(current.pasta) };
  if (action === "verificar") payload.qualidadeConferida = true;
  if (action === "revisar") payload.compatibilidadeConferida = true;
  if (evidence.operadorTrimble && !["devolver", "encerrar"].includes(action)) {
    payload.operadorTrimble = identifier(evidence.operadorTrimble);
  }
  if (action === "compartilhar" || action === "publicar") {
    const target = delivery.record.modelo.pastas[action === "compartilhar" ? "shared" : "published"];
    if (!destination || destination.id === current.arquivo || destination.parentId !== target ||
        destination.projectId !== delivery.record.origem.projeto || destination.name !== delivery.record.origem.nome) {
      throw new Error("Confira o ID da cópia e a pasta de destino no Trimble.");
    }
    payload.arquivoDestino = identifier(destination.id);
    payload.versaoDestino = identifier(destination.versionId);
    payload.pastaDestino = identifier(destination.parentId);
  }
  if (evidence.comparacao && ["compartilhar", "publicar", "conferir"].includes(action)) {
    payload.comparacao = evidence.comparacao;
  }
  return { Title: delivery.key, Acao: action, Chave: delivery.key, Geracao: delivery.generation, DadosJson: JSON.stringify(payload) };
}
