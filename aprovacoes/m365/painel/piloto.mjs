import { identifier } from "../../trimble.mjs";
import { assertActual, commandKey } from "./graph.mjs";

export class PilotArrivalPending extends Error {}

export const pilotBuild = "piloto-bim-2026.11.24.1";
export const pilotStages = {
  piloto_decisao: "Aguardando decisão",
  piloto_mover: "Aprovado · movimentação pendente",
  piloto_avisar: "Decisão aplicada · aviso pendente",
  piloto_envio_pendente: "Aviso reservado · aguardando confirmação do Outlook",
};

const byStage = {
  piloto_decisao: ["piloto_aprovar", "piloto_recusar"],
  piloto_mover: ["piloto_chegada"],
  piloto_avisar: ["piloto_avisar"],
};

export function pilotCommand(doc, action, note, operator, destination) {
  if (!byStage[doc.stage]?.includes(action)) throw new Error("Atualize a etapa do piloto.");
  const source = doc.record.origem;
  const message = note.trim();
  if (!message || message.length > 1000) throw new Error("Registre um comentário de até 1.000 caracteres.");
  const payload = {
    arquivo: identifier(source.arquivo), versao: identifier(source.versao), pasta: identifier(source.pasta),
    parecer: message, operadorTrimble: identifier(operator),
  };
  if (action === "piloto_chegada") {
    assertActual({ arquivo: source.arquivo, versao: destination?.versionId, pasta: doc.record.modelo.pastas.published },
      destination || {}, source.projeto, source.nome);
    if (destination.versionId === source.versao) throw new Error("A nova versão do PDF no destino precisa ser confirmada.");
    Object.assign(payload, {
      arquivoDestino: destination.id, versaoDestino: identifier(destination.versionId), pastaDestino: destination.parentId,
    });
  }
  return { Title: source.nome, Acao: action, Chave: commandKey(doc.key), Geracao: doc.generation, DadosJson: JSON.stringify(payload) };
}

export async function pilotArrival(api, doc, checkpoint, reviewedVersion) {
  if (doc.stage !== "piloto_mover") throw new Error("A movimentação exige aprovação registrada no SharePoint.");
  const source = doc.record.origem;
  const target = identifier(doc.record.modelo.pastas.published);
  const live = await api.file(source.arquivo);
  if (reviewedVersion) {
    const version = identifier(reviewedVersion);
    if (version === source.versao) throw new Error("Informe a nova versão registrada no histórico do Trimble.");
    assertActual({ arquivo: source.arquivo, versao: version, pasta: target }, live, source.projeto, source.nome);
    const historical = await api.file(source.arquivo, source.versao);
    assertActual(source, historical, source.projeto, source.nome);
    return live;
  }
  if (live.parentId === target) {
    const version = checkpoint.read();
    if (!version || version === source.versao) {
      throw new PilotArrivalPending("O PDF está em Publicado, mas falta a confirmação da versão criada pelo movimento. Confira o histórico no Trimble.");
    }
    assertActual({ arquivo: source.arquivo, versao: version, pasta: target }, live, source.projeto, source.nome);
    return live;
  }
  assertActual(source, live, source.projeto, source.nome);
  const { after } = await api.movePdf(live, target, checkpoint.write);
  return after;
}
