export type ContactHeaderMapping = {
  headers: string[];
  nameIndex: number;
  phoneIndex: number;
};

export function decodeCsv(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2));

  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

function normalizeColumnHeader(header: string) {
  return header.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

export function mapContactHeaders(sourceHeaders: string[]): ContactHeaderMapping {
  const normalized = sourceHeaders.map(normalizeColumnHeader);
  const directPhoneIndex = normalized.findIndex((header) => header === "telefone" || header.startsWith("telefone "));
  const phoneIndex = directPhoneIndex >= 0
    ? directPhoneIndex
    : normalized.findIndex((header) => ["celular", "contato", "phone", "mobile"].includes(header));
  const specificNameIndex = normalized.findIndex((header) =>
    header === "first name" || header === "firstname" || header === "nome" || header.startsWith("nome "),
  );
  const nameIndex = specificNameIndex >= 0 ? specificNameIndex : normalized.findIndex((header) => header === "name");
  const emailIndex = normalized.findIndex((header) => header === "email" || header === "e mail");
  const headers = sourceHeaders.map((header, index) => {
    if (index === nameIndex) return "nome";
    if (index === phoneIndex) return "telefone";
    if (index === emailIndex) return "email";
    return header;
  });

  return { headers, nameIndex, phoneIndex };
}