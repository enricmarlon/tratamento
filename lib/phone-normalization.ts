export function cleanPhone(phone: string) {
  const digits = phone.replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("55") && (digits.length === 12 || digits.length === 13)) return digits;
  return `55${digits}`;
}

export function parseSuppressionList(text: string) {
  const phones = text.replace(/^\uFEFF/, "").split(/[\r\n,;\t|]+/).map(cleanPhone).filter(Boolean);
  return [...new Set(phones)];
}