import { conflitosDePapeis, normalizar, validarIdentidades } from "./identidade.mjs";

// Valida o JSON de um modelo de aprovação antes de gravá-lo em BIMModelos. As regras repetem as que o
// fluxo aplica ao ler o modelo (gerar-fluxo.mjs), para que um modelo recusado seja encontrado aqui e não
// na primeira submissão. O editor e a linha de comando usam estas mesmas regras.

export const papeis = ["autor", "qualidade", "revisor", "publicador", "lider", "contratante", "conferente"];
export const limiteDoModelo = 16000;
const campos = new Set(["projeto", "pastas", "papeis", "prefixo", "separadorRevisao", "prefixoRevisao", "digitosRevisao",
  "formatos", "criterios", "identidades"]);
const identificador = /^[A-Za-z0-9_-]{1,96}$/;
const email = /^[^@\s|'"<>\\]{1,100}@[^@\s|'"<>\\]{1,59}$/;

export function validarModelo(modelo) {
  const erros = [], avisos = [];
  const erro = (texto) => erros.push(texto);
  if (!modelo || typeof modelo !== "object" || Array.isArray(modelo)) return { erros: ["O modelo deve ser um objeto JSON."], avisos };
  for (const nome of Object.keys(modelo)) {
    if (!campos.has(nome)) avisos.push("Campo desconhecido '" + nome + "': ele será congelado em todos os dossiês, mas nenhuma regra o usa.");
  }

  if (typeof modelo.projeto !== "string" || !identificador.test(modelo.projeto)) erro("projeto: informe o ID do projeto Trimble (letras, dígitos, '-' e '_').");
  const pastas = modelo.pastas;
  const ids = ["wip", "shared", "published"].map((nome) => pastas?.[nome]);
  if (!pastas || ids.some((id) => typeof id !== "string" || !identificador.test(id))) erro("pastas: informe os IDs de wip, shared e published.");
  else if (new Set(ids).size !== 3) erro("pastas: wip, shared e published devem ser três pastas diferentes.");

  if (!modelo.papeis || typeof modelo.papeis !== "object" || Array.isArray(modelo.papeis)) erro("papeis: informe os sete papéis.");
  else {
    for (const nome of Object.keys(modelo.papeis)) if (!papeis.includes(nome)) avisos.push("Papel desconhecido '" + nome + "'.");
    for (const papel of papeis) {
      const lista = modelo.papeis[papel];
      if (!Array.isArray(lista) || lista.length < 1 || lista.length > 60) { erro("papeis." + papel + ": informe de 1 a 60 e-mails."); continue; }
      const vistos = new Set();
      for (const valor of lista) {
        if (typeof valor !== "string" || valor !== normalizar(valor) || valor.length > 160 || !email.test(valor)) {
          erro("papeis." + papel + ": use e-mails em minúsculas, sem espaços (" + String(valor).slice(0, 60) + ").");
        } else if (vistos.has(valor)) erro("papeis." + papel + ": e-mail repetido (" + valor + ").");
        vistos.add(valor);
      }
    }
  }

  if (typeof modelo.prefixo !== "string" || modelo.prefixo.length > 32 || !/^[A-Z0-9_-]*$/.test(modelo.prefixo)) erro("prefixo: até 32 caracteres entre A-Z, 0-9, '-' e '_'.");
  if (!["-", "_"].includes(modelo.separadorRevisao)) erro("separadorRevisao: use '-' ou '_'.");
  if (typeof modelo.prefixoRevisao !== "string" || !/^[A-Z]{1,3}$/.test(modelo.prefixoRevisao)) erro("prefixoRevisao: de 1 a 3 letras maiúsculas.");
  if (!Number.isInteger(modelo.digitosRevisao) || modelo.digitosRevisao < 1 || modelo.digitosRevisao > 6) erro("digitosRevisao: inteiro de 1 a 6.");
  if (!Array.isArray(modelo.formatos) || !modelo.formatos.length || modelo.formatos.some((formato) => typeof formato !== "string" || !/^[a-z0-9]{1,10}$/.test(formato))) {
    erro("formatos: informe extensões em minúsculas, sem ponto (por exemplo, ifc, pdf).");
  }
  if (typeof modelo.criterios !== "string" || !modelo.criterios.trim() || modelo.criterios.length > 2000) erro("criterios: texto de 1 a 2.000 caracteres.");

  try { validarIdentidades(modelo.identidades); } catch (falha) { erro("identidades: " + falha.message); }
  if (!erros.length) {
    const conflitos = conflitosDePapeis(modelo);
    erros.push(...conflitos.erros);
    avisos.push(...conflitos.avisos);
  }

  const tamanho = JSON.stringify(modelo).length;
  if (tamanho > limiteDoModelo) erro("O modelo compactado tem " + tamanho + " caracteres e o fluxo aceita até " + limiteDoModelo + ".");
  else if (tamanho > limiteDoModelo * 0.75) avisos.push("O modelo compactado tem " + tamanho + " de " + limiteDoModelo + " caracteres permitidos.");
  return { erros, avisos };
}
