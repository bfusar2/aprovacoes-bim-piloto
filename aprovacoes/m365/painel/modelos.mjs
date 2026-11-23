import { loadSettings } from "./configuracao.mjs";
import { normalizar } from "./identidade.mjs";
import { sharePointListUrl } from "./model.mjs";
import { limiteDoModelo, papeis, validarModelo } from "./validacao-modelo.mjs";

const tamanhoDeImportacao = 100000;
const chaveValida = /^[A-Z0-9_-]{1,100}\|[A-Z0-9_-]{1,24}$/;
const pastaDoModelo = ["wip", "shared", "published"];
const rotulos = {
  autor: "Autor da submissão",
  qualidade: "Verificação de qualidade",
  revisor: "Revisão técnica",
  publicador: "Compartilhamento e publicação",
  lider: "Autorização da contratada líder",
  contratante: "Aceite da contratante",
  conferente: "Conferência independente da publicação",
};

const objeto = (valor) => valor && typeof valor === "object" && !Array.isArray(valor);
const texto = (valor) => typeof valor === "string" ? valor : "";
const itens = (valor) => valor.split(/[\r\n,;]/).map(normalizar).filter(Boolean);

export function lerImportacao(textoJson) {
  if (typeof textoJson !== "string" || textoJson.length > tamanhoDeImportacao) {
    throw new Error("Use um JSON com até 100.000 caracteres para importar.");
  }
  let dados;
  try { dados = JSON.parse(textoJson.replace(/^\uFEFF/, "")); }
  catch { throw new Error("O arquivo ou trecho colado contém JSON inválido."); }
  if (!objeto(dados)) throw new Error("Selecione um objeto JSON de DadosJson ou um registro de BIMModelos.");
  if (!Object.hasOwn(dados, "DadosJson")) return { modelo: dados };

  let modelo;
  try { modelo = JSON.parse(dados.DadosJson); }
  catch { throw new Error("O campo DadosJson do registro precisa conter o texto JSON do modelo."); }
  if (!objeto(modelo)) throw new Error("DadosJson deve conter um objeto JSON.");
  return {
    modelo,
    metadados: { titulo: texto(dados.Title), chave: texto(dados.Chave), ativo: dados.Ativo === true },
  };
}

export function lerIdentidades(linhas) {
  const pares = [];
  const contas = new Set();
  for (const [indice, linha] of linhas.split(/\r?\n/).entries()) {
    if (!linha.trim()) continue;
    const partes = linha.split(/→|=>/).map(normalizar);
    if (partes.length !== 2 || partes.some((valor) => !valor)) {
      throw new Error("Linha " + (indice + 1) + " das correspondências precisa ter o formato conta Microsoft → e-mail Trimble.");
    }
    const [conta, endereco] = partes;
    if (contas.has(conta)) throw new Error("Conta Microsoft repetida nas correspondências: " + conta + ".");
    contas.add(conta);
    pares.push([conta, endereco]);
  }
  return Object.fromEntries(pares);
}

export function prepararRegistro({ titulo, chave, ativo }, modelo) {
  const { erros, avisos } = validarModelo(modelo);
  if (typeof titulo !== "string" || !titulo.trim() || titulo.trim().length > 255) {
    erros.unshift("Informe um título de 1 a 255 caracteres.");
  }
  if (typeof chave !== "string" || !chaveValida.test(chave)) {
    erros.unshift("Informe a chave no formato NOME|VERSÃO. Use letras maiúsculas, números, hífen ou sublinhado.");
  }
  if (typeof ativo !== "boolean") erros.unshift("Informe se o modelo está ativo.");
  const registro = erros.length ? null : { Title: titulo.trim(), Chave: chave, Ativo: ativo, DadosJson: JSON.stringify(modelo) };
  return { erros, avisos, registro };
}

function iniciarEditor() {
  const elemento = (id) => document.getElementById(id);
  const formulario = elemento("editor-form");
  const campo = (nome) => formulario.elements.namedItem(nome);
  const definir = (nome, valor) => { campo(nome).value = texto(valor); };
  const pegar = (nome) => campo(nome).value.trim();
  let base = {};
  let projetoAtivo;
  let configuracao;
  let registro;
  let houveEdicao = false;
  let leitura = 0;
  let verificandoContexto = window.parent !== window;

  for (const papel of papeis) {
    const rotulo = document.createElement("label");
    rotulo.textContent = rotulos[papel];
    const enderecos = document.createElement("textarea");
    enderecos.name = papel;
    enderecos.rows = 2;
    enderecos.spellcheck = false;
    enderecos.className = "code";
    enderecos.required = true;
    rotulo.append(enderecos);
    elemento("roles").append(rotulo);
  }

  function atualizarLinks() {
    const site = pegar("site");
    let lista, novo;
    let aviso = "";
    if (verificandoContexto) {
      aviso = "Aguardando a identificação do projeto para abrir a lista.";
    } else if (projetoAtivo && pegar("projeto") !== projetoAtivo.id) {
      aviso = "Para abrir a lista neste painel, confira se o rascunho pertence ao projeto ativo.";
    } else if (!site) {
      aviso = "Preencha o site para abrir a lista BIMModelos.";
    } else {
      try {
        lista = sharePointListUrl(site, "BIMModelos", "AllItems");
        novo = sharePointListUrl(site, "BIMModelos", "NewForm");
      } catch (falha) { aviso = falha.message; }
    }
    for (const [id, endereco] of [["models", lista], ["new-model", registro && novo]]) {
      const link = elemento(id);
      link.hidden = !endereco;
      if (endereco) link.href = endereco;
      else link.removeAttribute("href");
    }
    elemento("site-status").textContent = aviso;
  }

  function limparSaida() {
    registro = null;
    elemento("check").hidden = true;
    elemento("publication").hidden = true;
    elemento("copy-status").textContent = "";
    elemento("prepared-fields").replaceChildren();
    atualizarLinks();
  }

  function carregarImportacao({ modelo, metadados }) {
    houveEdicao = true;
    base = modelo;
    if (metadados) {
      definir("titulo", metadados.titulo);
      definir("chave", metadados.chave);
      campo("ativo").checked = metadados.ativo;
    }
    for (const nome of ["projeto", "prefixo", "separadorRevisao", "prefixoRevisao", "criterios"]) definir(nome, modelo[nome]);
    campo("digitosRevisao").value = Number.isFinite(modelo.digitosRevisao) ? modelo.digitosRevisao : "";
    definir("formatos", Array.isArray(modelo.formatos) ? modelo.formatos.map(texto).join(", ") : "");
    for (const nome of pastaDoModelo) definir(nome, modelo.pastas?.[nome]);
    for (const papel of papeis) {
      const enderecos = modelo.papeis?.[papel];
      definir(papel, Array.isArray(enderecos) ? enderecos.map(texto).join("\n") : "");
    }
    const mapa = objeto(modelo.identidades) ? Object.entries(modelo.identidades) : [];
    definir("identidades", mapa.map(([conta, email]) => conta + " → " + texto(email)).join("\n"));
    if (mapa.length) campo("identidades").closest("details").open = true;
    limparSaida();
  }

  function montarModelo() {
    const identidades = lerIdentidades(pegar("identidades"));
    const modelo = {
      ...base,
      projeto: pegar("projeto"),
      pastas: { ...(objeto(base.pastas) ? base.pastas : {}), ...Object.fromEntries(pastaDoModelo.map((nome) => [nome, pegar(nome)])) },
      papeis: { ...(objeto(base.papeis) ? base.papeis : {}), ...Object.fromEntries(papeis.map((papel) => [papel, itens(pegar(papel))])) },
      prefixo: pegar("prefixo"),
      separadorRevisao: pegar("separadorRevisao"),
      prefixoRevisao: pegar("prefixoRevisao"),
      digitosRevisao: campo("digitosRevisao").valueAsNumber,
      formatos: itens(pegar("formatos")),
      criterios: pegar("criterios"),
    };
    if (Object.keys(identidades).length) modelo.identidades = identidades;
    else delete modelo.identidades;
    return modelo;
  }

  function aplicarProjeto() {
    if (!projetoAtivo) return;
    definir("projeto", projetoAtivo.id);
    if (configuracao) {
      definir("site", configuracao.sharePointSite);
      definir("wip", configuracao.folders?.wip?.id);
      definir("prefixo", configuracao.codePrefix);
      definir("separadorRevisao", configuracao.revisionSeparator);
      definir("prefixoRevisao", configuracao.revisionPrefix);
      campo("digitosRevisao").value = configuracao.revisionDigits;
      definir("formatos", configuracao.formats.join(", "));
    }
    limparSaida();
  }

  async function projetoNaSessao(workspace) {
    const chamada = ++leitura;
    verificandoContexto = true;
    projetoAtivo = undefined;
    configuracao = undefined;
    elemento("use-project").hidden = true;
    elemento("context-status").textContent = "Consultando o projeto…";
    if (campo("site").readOnly) definir("site", "");
    campo("site").readOnly = false;
    limparSaida();
    let atual;
    try { atual = await workspace.project.getProject(); }
    catch (falha) {
      if (chamada === leitura) throw falha;
      return;
    }
    if (chamada !== leitura) return;
    const configurada = await loadSettings(atual.id).catch(() => null);
    if (chamada !== leitura) return;
    verificandoContexto = false;
    projetoAtivo = atual;
    configuracao = configurada;
    if (configuracao) {
      definir("site", configuracao.sharePointSite);
      campo("site").readOnly = true;
    }
    if (!houveEdicao) aplicarProjeto();
    elemento("context-status").textContent = "Projeto ativo · " + (atual.name || atual.id) + ". " +
      (configuracao ? "Site definido na configuração do painel." : "Informe o site SharePoint deste projeto.");
    elemento("use-project").hidden = false;
    atualizarLinks();
  }

  function prepararLista(id, mensagens) {
    elemento(id).replaceChildren(...mensagens.map((mensagem) => {
      const item = document.createElement("li");
      item.textContent = mensagem;
      return item;
    }));
  }

  async function copiar(valor, alvo, legenda) {
    try {
      await navigator.clipboard.writeText(valor);
      elemento("copy-status").textContent = legenda + " copiado.";
    } catch {
      if (!alvo) {
        elemento("copy-status").textContent = "A cópia do registro inteiro foi bloqueada. Copie cada campo acima.";
        return;
      }
      if (alvo instanceof HTMLTextAreaElement) {
        alvo.focus();
        alvo.select();
      } else {
        const selecao = window.getSelection();
        const intervalo = document.createRange();
        intervalo.selectNodeContents(alvo);
        selecao?.removeAllRanges();
        selecao?.addRange(intervalo);
      }
      elemento("copy-status").textContent = legenda + " selecionado. Use Ctrl+C ou Copiar no dispositivo.";
    }
  }

  function renderizarRegistro(campos) {
    const linhas = [];
    for (const [nome, conteudo] of Object.entries(campos)) {
      const linha = document.createElement("div");
      const titulo = document.createElement("dt");
      titulo.textContent = nome === "Title" ? "Title / Título" : nome;
      const detalhe = document.createElement("dd");
      const valor = nome === "Ativo" ? (conteudo ? "Sim" : "Não") : conteudo;
      let alvo = detalhe;
      if (nome === "DadosJson") {
        const entrada = document.createElement("textarea");
        entrada.rows = 5;
        entrada.readOnly = true;
        entrada.setAttribute("aria-label", "DadosJson");
        entrada.value = valor;
        detalhe.append(entrada);
        alvo = entrada;
      } else detalhe.textContent = valor;
      const botao = document.createElement("button");
      botao.type = "button";
      botao.className = "secondary";
      botao.textContent = "Copiar";
      botao.setAttribute("aria-label", "Copiar " + nome);
      botao.onclick = () => copiar(valor, alvo, nome);
      linha.append(titulo, detalhe, botao);
      linhas.push(linha);
    }
    elemento("prepared-fields").replaceChildren(...linhas);
    elemento("publication").hidden = false;
    atualizarLinks();
  }

  formulario.oninput = () => { houveEdicao = true; limparSaida(); };
  formulario.onsubmit = (evento) => {
    evento.preventDefault();
    houveEdicao = true;
    limparSaida();
    let resultado;
    try {
      resultado = prepararRegistro({ titulo: pegar("titulo"), chave: pegar("chave"), ativo: campo("ativo").checked }, montarModelo());
    } catch (falha) { resultado = { erros: [falha.message], avisos: [] }; }
    registro = resultado.registro;
    prepararLista("model-erros", resultado.erros);
    prepararLista("model-avisos", resultado.avisos);
    elemento("check").hidden = false;
    elemento("check-status").textContent = registro
      ? "Modelo validado. DadosJson usa " + registro.DadosJson.length + " dos " + limiteDoModelo + " caracteres disponíveis."
      : "Corrija os itens listados para preparar o registro.";
    if (registro) renderizarRegistro(registro);
    elemento("check").scrollIntoView({ block: "nearest" });
  };
  elemento("import-form").onsubmit = async (evento) => {
    evento.preventDefault();
    const arquivo = elemento("import-file").files[0];
    const botoes = evento.currentTarget.querySelectorAll("button");
    for (const botao of botoes) botao.disabled = true;
    try {
      const colado = elemento("import-json").value.trim();
      if (colado && arquivo) throw new Error("Escolha uma única origem para importar o modelo.");
      if (arquivo?.size > tamanhoDeImportacao * 4) throw new Error("O arquivo excede o tamanho permitido.");
      carregarImportacao(lerImportacao(arquivo ? await arquivo.text() : colado));
      elemento("import-status").textContent = "Rascunho importado. Confira as regras e valide para preparar os campos.";
    } catch (falha) { elemento("import-status").textContent = falha.message; }
    finally { for (const botao of botoes) botao.disabled = false; }
  };
  elemento("use-project").onclick = () => { houveEdicao = true; aplicarProjeto(); };
  elemento("copy-record").onclick = () => {
    if (registro) void copiar(JSON.stringify(registro, null, 2), undefined, "Registro JSON");
  };
  atualizarLinks();

  const mensagemOffline = "Editor disponível. Informe o projeto e o site para preparar o modelo.";
  if (window.parent === window) {
    elemento("context-status").textContent = mensagemOffline;
  } else {
    const script = document.createElement("script");
    script.src = "https://components.connect.trimble.com/trimble-connect-workspace-api/index.js";
    const indisponivel = () => {
      verificandoContexto = false;
      elemento("context-status").textContent = mensagemOffline;
      atualizarLinks();
    };
    script.onerror = indisponivel;
    script.onload = async () => {
      let workspace;
      try {
        workspace = await window.TrimbleConnectWorkspace.connect(window.parent, (evento) => {
          if (workspace && evento === "project.onChanged") void projetoNaSessao(workspace).catch(indisponivel);
        }, 10000);
        await projetoNaSessao(workspace);
      } catch { indisponivel(); }
    };
    document.head.append(script);
  }
}

if (typeof document !== "undefined") iniciarEditor();
