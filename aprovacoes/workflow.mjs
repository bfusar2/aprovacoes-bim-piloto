export const STATES = {
  qa: "Verificação de qualidade",
  review: "Revisão técnica",
  share: "Pronto para compartilhar",
  shared: "Compartilhado",
  authorize: "Aguardando autorização",
  accept: "Aguardando aceite",
  publish: "Pronto para publicar",
  published: "Publicado",
  returned: "Devolvido",
  superseded: "Substituído",
};

export const CHECKLIST = {
  qa: [
    ["metadata", "Identificação, revisão e adequação ao uso conferidas"],
    ["complete", "Arquivo abre e os entregáveis estão completos"],
  ],
  review: [
    ["requirements", "Conteúdo atende aos requisitos de informação aplicáveis"],
    ["coordination", "Compatibilidade com as referências e disciplinas verificada"],
  ],
};

export const ACTIONS = {
  verify: { from: "qa", to: "review", role: "qa", label: "Concluir verificação" },
  review: { from: "review", to: "share", role: "reviewer", label: "Aprovar compartilhamento" },
  share: { from: "share", to: "shared", role: "publisher", label: "Compartilhar revisão" },
  deliver: { from: "shared", to: "authorize", role: "submitter", label: "Submeter para entrega formal" },
  authorize: { from: "authorize", to: "accept", role: "authorizer", label: "Autorizar entrega" },
  accept: { from: "accept", to: "publish", role: "acceptor", label: "Aceitar entrega" },
  publish: { from: "publish", to: "published", role: "publisher", label: "Publicar revisão" },
};

export const ROLES = {
  submitter: "Autor", qa: "Verificador", reviewer: "Revisor técnico",
  authorizer: "Contratada líder", acceptor: "Contratante", publisher: "Publicador",
};

export function availableActions(doc, actor) {
  if (doc.transfer) {
    const action = doc.transfer.action;
    return actor.roles.includes(ACTIONS[action].role) ? [action] : [];
  }
  const permitted = Object.entries(ACTIONS)
    .filter(([, rule]) => rule.from === doc.state && actor.roles.includes(rule.role))
    .filter(([action]) => {
      if (action === "deliver") return doc.author.id === actor.id;
      if (action === "review" || action === "authorize" || action === "accept") {
        return doc.author.id !== actor.id && (action !== "accept" || doc.authorizer !== actor.id);
      }
      return true;
    })
    .map(([action]) => action);
  if (permitted.some((action) => ["verify", "review", "authorize", "accept"].includes(action))) {
    permitted.push("return");
  }
  if (doc.author.id === actor.id && !["published", "returned", "superseded"].includes(doc.state)) {
    permitted.push("withdraw");
  }
  return permitted;
}

export function inspect(file, config) {
  const parts = /^(.+)\.([A-Za-z0-9]+)$/.exec(file.name);
  const stem = parts?.[1] || "";
  const position = stem.lastIndexOf(config.revisionSeparator);
  const code = position > 0 ? stem.slice(0, position) : "";
  const revision = position > 0 ? stem.slice(position + config.revisionSeparator.length) : "";
  const expectedRevision = new RegExp("^" + config.revisionPrefix + "[0-9]{" + config.revisionDigits + "}$");
  const checks = [
    { label: "Pasta de origem", pass: file.parentId === config.folders.wip.id, detail: config.folders.wip.name },
    { label: "Identificação do documento", pass: /^[A-Z0-9]+(?:[-_][A-Z0-9]+)*$/.test(code) && code.startsWith(config.codePrefix),
      detail: "Código em maiúsculas, conforme a convenção do projeto." },
    { label: "Revisão documental", pass: expectedRevision.test(revision),
      detail: config.revisionSeparator + config.revisionPrefix + "0".repeat(config.revisionDigits) + " antes da extensão." },
    { label: "Formato permitido", pass: config.formats.includes(parts?.[2].toLowerCase()), detail: config.formats.join(", ").toUpperCase() },
    { label: "Versão no Trimble", pass: Boolean(file.versionId), detail: file.versionId || "Versão indisponível." },
  ];
  return { code, revision, checks, pass: checks.every((check) => check.pass) };
}
