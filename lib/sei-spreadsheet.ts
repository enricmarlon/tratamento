const COLUMN_DEFINITIONS = [
  { id: "matricula", label: "Matrícula", aliases: ["matricula"] },
  { id: "nome", label: "Nome do Aluno", aliases: ["nome do aluno"] },
  { id: "primeiroNome", label: "nome", aliases: [] },
  { id: "nascimento", label: "Data Nascimento", aliases: ["data nascimento", "data de nascimento"] },
  { id: "cpf", label: "CPF", aliases: ["cpf"] },
  { id: "telefoneResidencial", label: "Telefone Residencial", aliases: ["telefone residencial"] },
  { id: "celular", label: "Celular", aliases: ["celular"] },
  { id: "email", label: "Email", aliases: ["email", "e mail"] },
  { id: "rg", label: "RG", aliases: ["rg"] },
  { id: "orgaoEmissor", label: "Órgão Emissor", aliases: ["orgao emissor", "orgao", "emissor"] },
  { id: "emissaoRg", label: "Data Emissão RG", aliases: ["data emissao rg", "data emissao"] },
  { id: "cidade", label: "Cidade", aliases: ["cidade"] },
  { id: "endereco", label: "Endereço", aliases: ["endereco"] },
  { id: "setor", label: "Setor", aliases: ["setor"] },
  { id: "cep", label: "CEP", aliases: ["cep"] },
] as const;

export type SeiColumnId = (typeof COLUMN_DEFINITIONS)[number]["id"];
export type SeiColumn = { id: SeiColumnId; label: string };
export type SeiRow = Partial<Record<SeiColumnId, string>>;
export type SeiSpreadsheet = {
  columns: SeiColumn[];
  rows: SeiRow[];
  emptyRowsRemoved: number;
};

function normalizeHeader(value: string) {
  return value.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function firstName(value: string) {
  const first = value.trim().split(/\s+/)[0] ?? "";
  if (!first) return "";
  const lowercase = first.toLocaleLowerCase("pt-BR");
  return lowercase.charAt(0).toLocaleUpperCase("pt-BR") + lowercase.slice(1);
}

function normalizeBrazilianMobile(value: string) {
  const digits = value.replace(/\D/g, "");
  if (!digits) return "";

  const hasCountryCode = digits.startsWith("55") && digits.length >= 12;
  const nationalNumber = hasCountryCode ? digits.slice(2) : digits;
  let normalized = hasCountryCode ? digits : `55${digits}`;

  if (nationalNumber.length === 10) {
    normalized = `55${nationalNumber.slice(0, 2)}9${nationalNumber.slice(2)}`;
  }

  return normalized;
}

function findHeader(row: string[]) {
  const normalized = row.map(normalizeHeader);
  const indexes = new Map<SeiColumnId, number>();

  for (const column of COLUMN_DEFINITIONS) {
    const index = normalized.findIndex((cell) => column.aliases.some((alias) => cell === alias || cell.startsWith(`${alias} `)));
    if (index >= 0) indexes.set(column.id, index);
  }

  if (!indexes.has("matricula") || !indexes.has("nome")) return null;
  return indexes;
}

function findDataColumnOffset(row: string[], header: Map<SeiColumnId, number>) {
  let bestOffset = 0;
  let bestScore = -1;

  for (const offset of [0, -1, 1]) {
    const score = [...header.values()].filter((index) => Boolean(row[index + offset])).length;
    if (score > bestScore) {
      bestOffset = offset;
      bestScore = score;
    }
  }

  return bestOffset;
}

export async function readSeiSpreadsheet(file: File): Promise<SeiSpreadsheet> {
  const XLSX = await import("@e965/xlsx");
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  let emptyRowsRemoved = 0;
  const records: SeiRow[] = [];

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const sourceRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, raw: false, defval: "", blankrows: true });
    let currentHeader: Map<SeiColumnId, number> | null = null;
    let dataColumnOffset = 0;
    let offsetResolved = false;

    for (const sourceRow of sourceRows) {
      const row = sourceRow.map((cell) => String(cell ?? "").trim());
      const header = findHeader(row);
      if (header) {
        currentHeader = header;
        dataColumnOffset = 0;
        offsetResolved = false;
        continue;
      }
      if (!currentHeader) continue;
      if (!row.some(Boolean)) {
        emptyRowsRemoved += 1;
        continue;
      }
      if (!offsetResolved) {
        dataColumnOffset = findDataColumnOffset(row, currentHeader);
        offsetResolved = true;
      }

      const record: SeiRow = {};
      for (const [columnId, index] of currentHeader) record[columnId] = row[index + dataColumnOffset] ?? "";
      const values = Object.values(record);
      if (values.every((value) => !value)) {
        emptyRowsRemoved += 1;
        continue;
      }

      const studentName = record.nome ?? "";
      if (!studentName || normalizeHeader(studentName).startsWith("total matriculados")) continue;
      record.primeiroNome = firstName(studentName);
      record.celular = normalizeBrazilianMobile(record.celular ?? "");
      records.push(record);
    }
  }

  const columns = COLUMN_DEFINITIONS
    .filter((column) => records.some((record) => Boolean(record[column.id]?.trim())))
    .map(({ id, label }) => ({ id, label }));

  return { columns, rows: records, emptyRowsRemoved };
}