export type ContactHeaderMapping = {
  headers: string[];
  nameIndex: number;
  phoneIndex: number;
};

function normalizeColumnHeader(header: string) {
  return header.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

export function mapContactHeaders(sourceHeaders: string[]): ContactHeaderMapping {
  const normalized = sourceHeaders.map(normalizeColumnHeader);
  const directPhoneIndex = normalized.findIndex((header) => header === "telefone" || header.startsWith("telefone "));
  const phoneIndex = directPhoneIndex >= 0
    ? directPhoneIndex
    : normalized.findIndex((header) => header === "celular" || header === "contato");
  const nameIndex = normalized.findIndex((header) => header === "nome" || header.startsWith("nome "));
  const headers = sourceHeaders.map((header, index) => index === phoneIndex ? "telefone" : header);

  return { headers, nameIndex, phoneIndex };
}