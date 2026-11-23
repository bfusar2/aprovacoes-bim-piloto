import { identifier } from "../../trimble.mjs";
import { inspect } from "../../workflow.mjs";
import { validarIdentidades } from "./identidade.mjs";

const regions = new Set(["northamerica", "europe", "unitedkingdom", "asia", "australia"]);

export function sharePointListUrl(siteAddress, segment, view) {
  let site;
  try { site = new URL(siteAddress); } catch { /* Tratado junto com os demais endereços inválidos. */ }
  if (!site || site.protocol !== "https:" || !/^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.sharepoint\.com$/.test(site.hostname) ||
      site.port || site.username || site.password || site.search || site.hash ||
      !/^\/(?:$|(?:sites|teams)\/[a-zA-Z0-9_-]+\/?$)/.test(site.pathname)) {
    throw new Error("O projeto precisa de uma URL HTTPS válida do site corporativo SharePoint.");
  }
  if (typeof segment !== "string" || !/^[\p{L}\d_-][\p{L}\d _-]{0,100}$/u.test(segment)) {
    throw new Error("Configure o segmento da URL da lista SharePoint.");
  }
  if (view !== "NewForm" && view !== "AllItems") throw new Error("Exibição da lista SharePoint inválida.");
  return site.origin + site.pathname.replace(/\/$/, "") + "/Lists/" + encodeURIComponent(segment) + "/" + view + ".aspx";
}

export function projectSettings(projects, id) {
  if (!Object.hasOwn(projects, identifier(id))) {
    throw new Error("Este projeto ainda precisa ser habilitado pelo administrador da extensão.");
  }
  const settings = projects[id];
  const formUrl = sharePointListUrl(settings.sharePointSite, settings.intakeList, "NewForm");
  const boardUrl = sharePointListUrl(settings.sharePointSite, settings.boardList, "AllItems");
  if (!regions.has(settings.region) || !["-", "_"].includes(settings.revisionSeparator) ||
      typeof settings.revisionPrefix !== "string" || !/^[A-Z]{1,3}$/.test(settings.revisionPrefix) ||
      !Number.isInteger(settings.revisionDigits) || settings.revisionDigits < 1 || settings.revisionDigits > 6 ||
      typeof settings.codePrefix !== "string" || !/^[A-Z0-9_-]*$/.test(settings.codePrefix) ||
      settings.codePrefix.length > 32 || !Array.isArray(settings.formats) || !settings.formats.length ||
      settings.formats.some((ext) => typeof ext !== "string" || !/^[a-z0-9]{1,10}$/.test(ext)) ||
      typeof settings.modelo !== "string" || !/^[A-Z0-9_-]{1,100}\|[A-Z0-9_-]{1,24}$/.test(settings.modelo)) {
    throw new Error("As regras de revisão ou o modelo de fluxo estão incompletos na configuração do painel.");
  }
  identifier(settings.folders?.wip?.id);
  if (typeof settings.folders.wip.name !== "string" || !settings.folders.wip.name.trim()) throw new Error("Configure o nome da pasta de elaboração.");
  // Identidades só valem para a conferência feita ao registrar uma submissão, que ainda não tem dossiê.
  // Nas decisões vale o mapa congelado no dossiê. Vale tanto para a configuração local quanto para a pública.
  const identidades = settings.identidades === undefined ? {} : { identidades: validarIdentidades(settings.identidades) };
  return { ...settings, ...identidades, formUrl, boardUrl };
}

export function capture(projectId, settings, selected, latest) {
  if (selected.projectId !== projectId || latest.projectId !== projectId || selected.id !== latest.id ||
      selected.versionId !== latest.versionId || selected.parentId !== latest.parentId ||
      selected.name !== latest.name) {
    throw new Error("O arquivo mudou de versão, nome ou pasta. Atualize a lista e confira a revisão antes de continuar.");
  }
  const inspection = inspect(latest, settings);
  if (!inspection.pass) {
    throw new Error(inspection.checks.filter((check) => !check.pass).map((check) => check.label + " — " + check.detail).join(" "));
  }
  const record = {
    modelo: settings.modelo, projeto: projectId, arquivo: identifier(latest.id),
    versao: identifier(latest.versionId), pasta: identifier(latest.parentId),
    nome: latest.name, codigo: inspection.code, revisao: inspection.revision,
  };
  const key = record.codigo + "|" + record.revisao;
  if (record.nome.length > 200 || record.codigo.length > 100 || [projectId, record.arquivo, record.versao].join("|").length > 255) {
    throw new Error("O identificador excede o tamanho aceito pelo fluxo. Limites de 200 caracteres no nome, 100 no código e 255 na chave de origem.");
  }
  return {
    Title: latest.name, Acao: "submeter", Chave: key,
    Geracao: 0, DadosJson: JSON.stringify(record),
  };
}
