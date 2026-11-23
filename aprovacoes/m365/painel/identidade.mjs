// Identidades administradas. O Trimble identifica a pessoa pelo e-mail do membro do projeto; o Microsoft
// Entra, pelo nome de usuário da conta (UPN); e o SharePoint, pelo e-mail gravado em "Criado por". Esses
// três valores nem sempre coincidem (por exemplo, um convidado externo tem UPN
// "nome_dominio.com#EXT#@empresa.onmicrosoft.com"). Nenhuma equivalência é presumida: quando os valores
// diferem, o administrador do modelo registra o mapeamento em "identidades", que fica congelado em cada
// dossiê junto com os papéis. O fluxo no Power Automate não usa este mapa. Ele serve à conferência feita
// pelo painel e à validação do modelo, e jamais substitui a identidade "Criado por" do SharePoint.

const enderecoValido = /^[^@\s|'"<>\\]{1,128}@[^@\s|'"<>\\]{1,128}$/;
const limiteDeEntradas = 200;

export const normalizar = (valor) => (typeof valor === "string" ? valor.trim().toLowerCase() : "");

// Pares de papéis que o fluxo exige que sejam exercidos por pessoas diferentes em uma mesma entrega.
export const papeisIncompativeis = [
  ["autor", "revisor"], ["autor", "lider"], ["autor", "contratante"], ["lider", "contratante"], ["publicador", "conferente"],
];

// Aceita somente { "conta-microsoft": "email-no-trimble" } com ambos em minúsculas e sem espaços.
export function validarIdentidades(mapa) {
  if (mapa === undefined || mapa === null) return {};
  if (typeof mapa !== "object" || Array.isArray(mapa)) throw new Error("As identidades devem ser um objeto conta Microsoft → e-mail no Trimble.");
  const entradas = Object.entries(mapa);
  if (entradas.length > limiteDeEntradas) throw new Error("O mapa de identidades excede " + limiteDeEntradas + " entradas.");
  const resultado = {};
  for (const [conta, trimble] of entradas) {
    if (conta !== normalizar(conta) || !enderecoValido.test(conta)) {
      throw new Error("Conta Microsoft inválida no mapa de identidades: use minúsculas, sem espaços (" + conta.slice(0, 80) + ").");
    }
    if (typeof trimble !== "string" || trimble !== normalizar(trimble) || !enderecoValido.test(trimble)) {
      throw new Error("E-mail do Trimble inválido para " + conta.slice(0, 80) + ": use um e-mail em minúsculas, sem espaços.");
    }
    resultado[conta] = trimble;
  }
  return resultado;
}

// E-mail que o membro do Trimble deve ter para a conta Microsoft informada. Sem mapeamento, a conta é o próprio e-mail.
export function trimbleEsperado(identidades, conta) {
  const chave = normalizar(conta);
  return identidades && Object.hasOwn(identidades, chave) ? normalizar(identidades[chave]) : chave;
}

// Procura pessoas que exercem papéis incompatíveis. Com a mesma conta, o fluxo recusa a decisão no dossiê.
// Com contas diferentes ligadas à mesma pessoa pelo mapa, o fluxo não tem como perceber, então é erro.
export function conflitosDePapeis(modelo) {
  const identidades = modelo?.identidades ?? {};
  const pessoas = new Map();
  for (const [papel, contas] of Object.entries(modelo?.papeis ?? {})) {
    for (const conta of Array.isArray(contas) ? contas : []) {
      const login = normalizar(conta);
      const pessoa = trimbleEsperado(identidades, login);
      if (!pessoas.has(pessoa)) pessoas.set(pessoa, new Map());
      const papeis = pessoas.get(pessoa);
      if (!papeis.has(papel)) papeis.set(papel, new Set());
      papeis.get(papel).add(login);
    }
  }
  const erros = [], avisos = [];
  for (const [pessoa, papeis] of pessoas) {
    for (const [a, b] of papeisIncompativeis) {
      if (!papeis.has(a) || !papeis.has(b)) continue;
      const contasDiferentes = [...papeis.get(a)].some((x) => [...papeis.get(b)].some((y) => x !== y));
      if (contasDiferentes) {
        erros.push(pessoa + " exerce os papéis " + a + " e " + b + " por contas diferentes. O fluxo não consegue impedir que a mesma pessoa decida nas duas funções.");
      } else {
        avisos.push(pessoa + " aparece nos papéis " + a + " e " + b + ". O fluxo recusará a decisão nos dossiês em que ela for " + a + ".");
      }
    }
  }
  return { erros, avisos };
}
