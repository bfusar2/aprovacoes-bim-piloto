import { inspect } from "./workflow.mjs";

const HOSTS = {
  northamerica: "app", europe: "app21", unitedkingdom: "app22", asia: "app31", australia: "app32",
};

export class Problem extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

export function identifier(value) {
  if (typeof value !== "string" || !/^[\w-]{1,96}$/.test(value)) {
    throw new Problem(400, "Identificador inválido.");
  }
  return value;
}

export function trimble(region, token, fetcher = fetch) {
  if (!HOSTS[region]) throw new Problem(500, "Região do projeto inválida.");
  const base = "https://" + HOSTS[region] + ".connect.trimble.com/tc/api/2.0/";
  return async (path, { method = "GET", body, ifMatch } = {}) => {
    const version = ifMatch === undefined ? undefined : identifier(ifMatch);
    let response;
    try {
      response = await fetcher(base + path, {
        method, redirect: "error", signal: AbortSignal.timeout(14000),
        headers: {
          Authorization: "Bearer " + token,
          ...(body && { "Content-Type": "application/json" }),
          ...(version && { "If-Match": version }),
        },
        body: body && JSON.stringify(body),
      });
    } catch {
      throw new Problem(502, "Não foi possível consultar o Trimble Connect. Atualize a tela e tente novamente.");
    }
    if (!response.ok) {
      const messages = {
        401: "Sessão Trimble expirada ou inválida.",
        403: "O Trimble Connect negou acesso a este recurso.",
        404: "Recurso indisponível no projeto Trimble.",
        409: "Já existe um arquivo com esse nome na pasta de destino.",
        412: "O Trimble Connect recusou a operação por conflito de versão ou arquivo existente.",
      };
      throw new Problem(messages[response.status] ? response.status : 502,
        messages[response.status] || "A API do Trimble Connect retornou erro (" + response.status + ").");
    }
    try { return await response.json(); }
    catch { throw new Problem(502, "Resposta inesperada do Trimble Connect."); }
  };
}

export function core(api, projectId, config) {
  function verifyProject(item) {
    if (item?.projectId !== projectId) throw new Problem(404, "Recurso fora do projeto.");
    return item;
  }
  async function file(id, versionId) {
    const item = verifyProject(await api("files/" + identifier(id) +
      (versionId ? "?versionId=" + identifier(versionId) + "&minimal=false" : "?minimal=false")));
    if (item.id !== id || (versionId && item.versionId !== versionId) || !item.versionId || !item.parentId ||
        typeof item.name !== "string") throw new Problem(409, "A identidade, a versão ou a pasta do arquivo não pôde ser confirmada.");
    return item;
  }
  async function items(folderId) {
    const folder = verifyProject(await api("folders/" + identifier(folderId)));
    if (folder.id !== folderId) throw new Problem(404, "Pasta indisponível no projeto.");
    const result = await api("folders/" + folderId + "/items");
    if (!Array.isArray(result)) throw new Problem(502, "Listagem de pasta indisponível.");
    return result.filter((item) => item.type === "FILE" && item.projectId === projectId &&
      item.parentId === folderId && typeof item.name === "string");
  }
  return {
    file, items,
    async actor() {
      const person = await api("projects/" + projectId + "/users/me?minimal=false");
      if (person?.status !== "ACTIVE" || !person.email || !person.id) {
        throw new Problem(403, "A conta precisa ser membro ativo deste projeto no Trimble Connect.");
      }
      const email = person.email.toLowerCase();
      const roles = Object.keys(config.roles).filter((role) => config.roles[role].includes(email));
      if (!roles.length) throw new Problem(403, "A conta ainda não tem um papel no fluxo deste projeto.");
      return { id: identifier(person.id), name: [person.firstName, person.lastName].filter(Boolean).join(" ") || person.email, roles };
    },
    async inspect(id) {
      const item = await file(id);
      return { file: item, inspection: inspect(item, config) };
    },
    async download(id, versionId) {
      await file(id, versionId);
      const info = await api("files/fs/" + identifier(id) + "/downloadurl?versionId=" + identifier(versionId));
      const url = new URL(info.url || info.downloadUrl || "about:blank");
      if (url.protocol !== "https:" || url.username || url.password) {
        throw new Problem(502, "Link de download da versão indisponível.");
      }
      return url.href;
    },
    copy(source, destination) {
      return api("files", {
        method: "POST",
        body: { parentId: destination, parentType: "FOLDER", fromFileVersionId: source.versionId, mergeExisting: false },
      });
    },
    async movePdf(source, destinationId, onPatched) {
      const { id, versionId, parentId, name } = source;
      identifier(versionId);
      identifier(destinationId);
      if (typeof name !== "string" || !/\.pdf$/i.test(name) || parentId === destinationId) {
        throw new Problem(400, "Selecione um PDF e uma pasta de destino diferente da origem.");
      }
      const current = await file(id);
      if (current.versionId !== versionId || current.parentId !== parentId || current.name !== name) {
        throw new Problem(409, "O PDF mudou desde a aprovação. Confira a pasta e a revisão atual.");
      }
      const destinationFiles = await items(destinationId);
      const normalizedName = name.normalize("NFC").toLowerCase();
      if (destinationFiles.some((item) => item.name.normalize("NFC").toLowerCase() === normalizedName)) {
        throw new Problem(409, "Já existe um arquivo com esse nome na pasta de destino.");
      }
      let moved;
      let confirmed;
      try {
        moved = await api("files/" + id, {
          method: "PATCH", ifMatch: versionId, body: { parentId: destinationId },
        });
      } catch (error) {
        if (error.status !== 502) throw error;
        throw new Problem(502, "O Trimble ficou sem confirmar a movimentação. Confira o PDF antes de tentar novamente.");
      }
      if (moved?.id === id && moved.versionId && moved.versionId !== versionId) onPatched?.(identifier(moved.versionId));
      try {
        confirmed = await file(id);
      } catch {
        throw new Problem(502, "O Trimble aceitou a movimentação. Confira a pasta e a versão antes de continuar.");
      }
      if (moved?.id !== id || confirmed.parentId !== destinationId || confirmed.name !== name ||
          confirmed.versionId === versionId || confirmed.versionId !== moved.versionId) {
        throw new Problem(409, "O Trimble retornou uma pasta ou versão divergente após a movimentação. Confira o PDF antes de continuar.");
      }
      return { before: current, after: confirmed };
    },
  };
}
