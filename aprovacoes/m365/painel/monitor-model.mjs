import { steps } from "./graph.mjs";

export const stages = Object.freeze({
  qualidade: "Qualidade",
  revisao: "Revisão técnica",
  compartilhamento: "Aguardando cópia compartilhada",
  compartilhado: "Compartilhada",
  autorizacao: "Autorização",
  aceite: "Aceite",
  publicacao: "Aguardando cópia publicada",
  conferencia: "Conferência independente",
  publicado: "Publicada",
  devolvido: "Devolvida",
  encerrado: "Encerrada",
});

export function stageLabel(stage) {
  return Object.hasOwn(stages, stage) ? stages[stage] : "Etapa indisponível";
}

export function countDeliveries(rows) {
  return rows.reduce((counts, row) => {
    if (Object.hasOwn(steps, row.Estado)) counts.open++;
    if (row.Estado === "conferencia") counts.conference++;
    if (row.Estado === "publicado") counts.published++;
    return counts;
  }, { open: 0, conference: 0, published: 0 });
}

export function officialEvent(command, dossier) {
  if (!dossier || command.key !== dossier.key) return undefined;
  return dossier.history.find((event) => String(event?.comando) === command.id);
}

export function commandState(command, dossier, now = Date.now()) {
  if (officialEvent(command, dossier)) return "applied";
  const result = typeof command.result === "string" ? command.result.trim() : "";
  if (result) return "notice";
  const date = Date.parse(command.created);
  if (!Number.isFinite(date)) return "undated";
  return now - date > 15 * 60_000 ? "overdue" : "waiting";
}
