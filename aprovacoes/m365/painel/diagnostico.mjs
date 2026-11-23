import { identifier, trimble } from "../../trimble.mjs";

export async function checkCore(request, projectId) {
  const member = await request("projects/" + identifier(projectId) + "/users/me?minimal=false");
  if (member?.status !== "ACTIVE" || !member.id) throw new Error("A Core API respondeu. A conta precisa constar como membro ativo deste projeto.");
}

function start() {
  const byId = (id) => document.getElementById(id);
  const show = (message) => { byId("result").textContent = message; };
  const run = byId("run");
  const region = byId("region");
  let sdk, project, request, version = 0;

  async function received(value, pending) {
    if (!pending || pending !== request || pending.processing) return;
    if (value === "pending" || value === "granted") return show("Aguardando a confirmação de acesso no Trimble.");
    if (value === "denied") {
      request = undefined;
      region.disabled = false;
      run.disabled = !region.value;
      return show("O Trimble recusou a permissão de acesso à sessão.");
    }
    if (typeof value !== "string" || !value) return;
    pending.processing = true;
    try {
      const current = await sdk.project.getProject();
      if (current.id !== project.id) throw new Error("O projeto ativo mudou. Reabra o diagnóstico no projeto de teste.");
      await checkCore(trimble(pending.region, value), project.id);
      if (version === pending.version) show("Workspace API conectada. A Core API 2.0 reconheceu a conta como membro ativo na região " + pending.region + ".");
    } catch (error) {
      if (version === pending.version) show((error.status ? "Core API HTTP " + error.status + " · " : "") + error.message);
    } finally {
      if (request === pending) {
        request = undefined;
        region.disabled = false;
        run.disabled = !region.value;
      }
    }
  }

  function onEvent(event, args) {
    if (event === "extension.accessToken") void received(args?.data, request);
    if (["project.onChanged", "extension.sessionInvalid", "extension.sessionLogOut"].includes(event)) {
      ++version;
      request = undefined;
      run.disabled = true;
      show("A sessão ou o projeto mudou. Reabra o diagnóstico.");
    }
  }

  async function connect() {
    if (parent === window) {
      show("Abra o diagnóstico pelo painel instalado no projeto do Trimble Connect.");
      return;
    }
    await new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://components.connect.trimble.com/trimble-connect-workspace-api/index.js";
      script.onload = resolve;
      script.onerror = () => reject(new Error("Falha ao carregar a biblioteca do Trimble."));
      document.head.append(script);
    });
    sdk = await window.TrimbleConnectWorkspace.connect(parent, onEvent, 30000);
    project = await sdk.project.getProject();
    identifier(project.id);
    byId("project").textContent = project.name || project.id;
    const suggested = String(project.location || project.region || "").toLowerCase().replace(/[^a-z]/g, "");
    if ([...region.options].some((option) => option.value === suggested)) region.value = suggested;
    run.disabled = !region.value;
    region.onchange = () => { run.disabled = !region.value || Boolean(request); };
    run.onclick = async () => {
      const pending = { region: region.value, version: ++version, processing: false };
      request = pending;
      run.disabled = true;
      region.disabled = true;
      show("Solicitando o token de leitura da sessão Trimble…");
      try {
        await received(await sdk.extension.requestPermission("accesstoken"), pending);
      } catch (error) {
        if (pending === request) {
          request = undefined;
          run.disabled = false;
          region.disabled = false;
          show(error.message);
        }
      }
    };
    show("Workspace API conectada. Confira a região e execute o teste.");
  }
  connect().catch((error) => show(error.message));
}

if (typeof document !== "undefined") start();
