"use client";

import ExcelJS from "exceljs";
import { ArrowDownToLine, ArrowRight, Check, FileSpreadsheet, LockKeyhole, LogOut, RotateCcw, Upload, X } from "lucide-react";
import { useRef, useState } from "react";
import { signOut } from "./actions/auth";

type ContactSheet = {
  headers: string[];
  rows: string[][];
  sourcePhones: string[];
  sourceNames: string[];
  reviewReasons: string[][];
  phoneIndex: number;
  nameIndex: number;
  fileName: string;
};

type CourseMessagePart = {
  text: string;
  bold?: boolean;
  italic?: boolean;
};

function parseCourseMessage(message: string): CourseMessagePart[] {
  const parts: CourseMessagePart[] = [];
  const markerPattern = /(\*[^*]+\*|_[^_]+_)/g;
  let lastIndex = 0;

  for (const match of message.matchAll(markerPattern)) {
    const index = match.index ?? 0;
    if (index > lastIndex) parts.push({ text: message.slice(lastIndex, index) });
    const marker = match[0];
    parts.push({ text: marker.slice(1, -1), ...(marker.startsWith("*") ? { bold: true } : { italic: true }) });
    lastIndex = index + marker.length;
  }

  if (lastIndex < message.length) parts.push({ text: message.slice(lastIndex) });
  return parts;
}

function delimiterFor(text: string) {
  const line = text.split(/\r?\n/, 1)[0] ?? "";
  const counts = new Map([[",", 0], [";", 0], ["\t", 0]]);
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    if (line[index] === '"') {
      if (quoted && line[index + 1] === '"') index += 1;
      else quoted = !quoted;
    } else if (!quoted && counts.has(line[index])) {
      counts.set(line[index], (counts.get(line[index]) ?? 0) + 1);
    }
  }
  return [...counts].sort((left, right) => right[1] - left[1])[0][0];
}

function parseCsv(text: string, delimiter: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const content = text.replace(/^\uFEFF/, "");

  for (let index = 0; index < content.length; index += 1) {
    const char = content[index];
    if (char === '"') {
      if (quoted && content[index + 1] === '"') {
        field += '"';
        index += 1;
      } else quoted = !quoted;
    } else if (!quoted && char === delimiter) {
      row.push(field);
      field = "";
    } else if (!quoted && (char === "\r" || char === "\n")) {
      if (char === "\r" && content[index + 1] === "\n") index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += char;
  }
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function cleanHeader(header: string) {
  return header.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function cleanPhone(phone: string) {
  const digits = phone.replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("55") && (digits.length === 12 || digits.length === 13)) return digits;
  return `55${digits}`;
}

const lowercaseNameParticles = new Set(["da", "das", "de", "do", "dos", "e"]);
const nameAbbreviations = new Set(["jr", "sc", "mr", "dr", "sr", "sra", "eng", "arq", "prof", "profa"]);
const joinedNamePrefixes = ["maria", "jose", "ana", "joao", "luiz", "luis", "pedro", "paulo", "antonio", "francisco"];
const joinedNameSuffixes = new Set([
  "julia", "juliana", "eduarda", "clara", "luiza", "helena", "beatriz", "isabel", "gabriela", "laura",
  "neto", "neta", "filho", "filha", "miguel", "gabriel", "lucas", "rafael", "vitoria", "manuel", "maria",
  "joao", "pedro", "paulo", "antonio", "augusto", "felipe", "henrique", "mariana", "sofia", "luana", "lara",
]);
const professionAndBusinessTerms = new Set([
  "engenheiro", "engenheira", "engenharia", "eng", "psicologo", "psicologa", "psicologia", "pisco",
  "arq", "arquiteta", "arquitetura", "arquiteto", "loja", "studio", "estudio", "medico", "medica",
  "medicina", "dentista", "odontologia", "advogado", "advogada", "advocacia", "contador", "contadora",
  "contabilidade", "corretor", "corretora", "imobiliaria", "consultoria", "consultor", "consultora", "clinica",
  "farmacia", "restaurante", "comercio", "comercial", "empresa", "oficina", "construtora", "mercado", "servicos",
  "salon", "salao", "barbearia", "pizzaria", "academia", "escritorio", "arquiteto", "arquiteta",
]);

function normalizeNameCase(name: string) {
  const lowercase = name.trim().toLocaleLowerCase("pt-BR");
  return lowercase.replace(/(^|[\s'-])(\p{L})(\p{L}*)/gu, (match, separator: string, initial: string, rest: string, offset: number) => {
    const word = `${initial}${rest}`;
    if (offset > 0 && lowercaseNameParticles.has(word)) return `${separator}${word}`;
    return `${separator}${initial.toLocaleUpperCase("pt-BR")}${rest}`;
  });
}

function nameReviewReasons(name: string) {
  const reasons: string[] = [];
  const normalized = cleanHeader(name);
  const words = normalized.match(/[a-z0-9]+/g) ?? [];
  const decomposedName = name.normalize("NFD").toLocaleLowerCase("pt-BR");

  if (words.some((word) => nameAbbreviations.has(word))) reasons.push("Possível abreviação");
  if (/[^\p{L}\p{M}\s'-]/u.test(name)) reasons.push("Caractere fora do padrão");
  if (/(^|[^\p{L}])\p{M}/u.test(decomposedName) || /(^|[^aeiou])\u0301/u.test(decomposedName) || /(?:Ã|Â)[^\u0000-\u007f]|�/u.test(name)) {
    reasons.push("Possível problema de acentuação/arquivo");
  }

  const joined = normalized.replace(/[^a-z]/g, "");
  if (!/[\s'-]/u.test(name) && joinedNamePrefixes.some((prefix) =>
    joined.startsWith(prefix) && joinedNameSuffixes.has(joined.slice(prefix.length)),
  )) {
    reasons.push("Possível nome composto sem espaço");
  }
  if (words.some((word) => professionAndBusinessTerms.has(word))) {
    reasons.push("Possível profissão ou negócio");
  }

  return reasons;
}

export default function HomePanel({ userEmail }: { userEmail: string }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const courseDialogRef = useRef<HTMLDialogElement>(null);
  const [sheet, setSheet] = useState<ContactSheet | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [courseMessage, setCourseMessage] = useState("");

  async function importCsv(file?: File) {
    if (!file) return;
    setError("");
    if (!file.name.toLowerCase().endsWith(".csv")) {
      setError("Selecione um arquivo CSV para continuar.");
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      setError("O arquivo excede o limite de 20 MB.");
      return;
    }

    setLoading(true);
    try {
      const text = await file.text();
      const parsed = parseCsv(text, delimiterFor(text));
      const headers = parsed[0]?.map((header) => header.trim()) ?? [];
      const normalized = headers.map(cleanHeader);
      const missing = ["nome", "sobrenome", "telefone", "email"].filter((name) => !normalized.includes(name));
      if (missing.length) {
        setSheet(null);
        setError(`Não encontrei as colunas: ${missing.join(", ")}.`);
        return;
      }

      const phoneIndex = normalized.indexOf("telefone");
      const nameIndex = normalized.indexOf("nome");
      const sourceRows = parsed.slice(1).filter((row) => row.some((cell) => cell.trim()));
      if (!sourceRows.length) {
        setSheet(null);
        setError("O CSV não contém contatos para tratar.");
        return;
      }

      const records = sourceRows.map((source, sourceIndex) => {
        const row = headers.map((_, index) => source[index] ?? "");
        const sourceName = row[nameIndex];
        const reasons = nameReviewReasons(sourceName);
        row[nameIndex] = normalizeNameCase(sourceName);
        row[phoneIndex] = cleanPhone(row[phoneIndex]);
        return {
          row,
          sourceName,
          sourcePhone: source[phoneIndex] ?? "",
          reasons,
          sourceIndex,
        };
      });
      records.sort((first, second) => Number(second.reasons.length > 0) - Number(first.reasons.length > 0) || first.sourceIndex - second.sourceIndex);
      setSheet({
        headers,
        rows: records.map((record) => record.row),
        sourceNames: records.map((record) => record.sourceName),
        sourcePhones: records.map((record) => record.sourcePhone),
        reviewReasons: records.map((record) => record.reasons),
        phoneIndex,
        nameIndex,
        fileName: file.name,
      });
    } catch {
      setSheet(null);
      setError("Não foi possível ler esse arquivo. Confira se ele está em formato CSV.");
    } finally {
      setLoading(false);
    }
  }

  async function downloadExcel(courseText?: string) {
    if (!sheet) return;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Contatos tratados");
    const headers = [...sheet.headers];
    const rows = sheet.rows.map((row) => [...row]);

    if (courseText !== undefined) {
      const existingCourseIndex = headers.findIndex((header) => cleanHeader(header) === "curso");
      if (existingCourseIndex < 0) {
        headers.splice(4, 0, "curso");
        rows.forEach((row) => row.splice(4, 0, courseText));
      } else if (existingCourseIndex !== 4) {
        const [courseHeader] = headers.splice(existingCourseIndex, 1);
        headers.splice(4, 0, courseHeader);
        rows.forEach((row) => {
          const [courseValue] = row.splice(existingCourseIndex, 1);
          row.splice(4, 0, courseValue ?? "");
        });
      }
      rows.forEach((row) => { row[4] = courseText; });
    }

    worksheet.addRow(headers);
    rows.forEach((row) => worksheet.addRow(row));
    if (courseText !== undefined) {
      rows.forEach((_, index) => {
        worksheet.getCell(index + 2, 5).alignment = { vertical: "top", wrapText: true };
      });
    }
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.columns = headers.map((header) => ({ width: cleanHeader(header) === "curso" && courseText !== undefined ? 42 : Math.min(Math.max(header.length + 4, 16), 32) }));

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${sheet.fileName.replace(/\.csv$/i, "")}_tratado.xlsx`;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function openCourseEditor() {
    setCourseMessage("");
    courseDialogRef.current?.showModal();
  }

  async function downloadWithCourse(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!courseMessage.trim()) return;
    await downloadExcel(courseMessage);
    courseDialogRef.current?.close();
  }

  function reset() {
    setSheet(null);
    setError("");
    if (inputRef.current) inputRef.current.value = "";
  }

  const changedPhones = sheet?.sourcePhones.filter((phone) => cleanPhone(phone) !== phone).length ?? 0;
  const adjustedNames = sheet?.sourceNames.filter((name, index) => name !== sheet.rows[index][sheet.nameIndex]).length ?? 0;
  const namesToReview = sheet?.reviewReasons.filter((reasons) => reasons.length > 0).length ?? 0;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#inicio" aria-label="Plan, início"><span className="brand-mark"><FileSpreadsheet size={20} /></span><span>Plan</span></a>
        <div className="sidebar-caption">FERRAMENTAS</div>
        <div className="nav-item nav-item-active"><span className="nav-indicator" /><FileSpreadsheet size={17} /><span>Tratar contatos</span></div>
        <div className="sidebar-bottom"><LockKeyhole size={16} /><div><strong>Seus dados, protegidos</strong><p>Os arquivos ficam no seu dispositivo.</p></div></div>
      </aside>

      <main className="main-content" id="inicio">
        <header className="topbar"><span className="breadcrumb">Ferramentas <span>/</span> Tratar contatos</span><div className="topbar-actions"><span className="local-badge"><span /> Processamento local</span><span className="auth-email" title={userEmail}>{userEmail}</span><form action={signOut}><button className="logout-button" type="submit"><LogOut size={15} /><span>Sair</span></button></form></div></header>
        <div className="content-wrap">
          <section className="page-heading">
            <div><div className="eyebrow">ORGANIZE SUA BASE</div><h1>Contatos prontos<br className="mobile-break" /> para disparar.</h1><p>Padronize nomes e telefones, confira os cadastros sinalizados e exporte em Excel.</p></div>
            <div className="format-note"><span>CSV</span><ArrowRight size={15} /><span className="format-xlsx">XLSX</span></div>
          </section>

          <section className="workflow" aria-label="Tratamento da planilha">
            <div className="section-topline"><div className="step-label"><span>01</span><h2>Importe sua planilha</h2></div>{sheet && <button className="text-button" onClick={reset} type="button"><RotateCcw size={15} /> Trocar arquivo</button>}</div>
            {!sheet ? (
              <label className={`upload-zone${loading ? " is-reading" : ""}`} htmlFor="csv-file" onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); void importCsv(event.dataTransfer.files[0]); }}>
                <input ref={inputRef} id="csv-file" type="file" accept=".csv,text/csv" onChange={(event) => void importCsv(event.target.files?.[0])} />
                <span className="upload-icon"><Upload size={21} /></span><span className="upload-title">{loading ? "Lendo sua planilha..." : "Solte o arquivo CSV aqui"}</span>
                <span className="upload-description">ou <span className="upload-link">escolha no computador</span></span><span className="upload-meta">CSV · até 20 MB</span>
              </label>
            ) : (
              <div className="file-loaded"><div className="loaded-file-icon"><FileSpreadsheet size={20} /></div><div className="loaded-file-info"><strong>{sheet.fileName}</strong><span>{sheet.rows.length.toLocaleString("pt-BR")} contatos identificados</span></div><span className="loaded-check"><Check size={16} /></span></div>
            )}
            {error && <p className="error-message" role="alert">{error}</p>}
            <div className="column-hint"><span className="hint-label">COLUNAS ESPERADAS</span><div className="column-list">{["nome", "sobrenome", "telefone", "email"].map((column) => <span className="column-chip" key={column}>{column}</span>)}</div></div>
          </section>

          {sheet && <>
            <section className="result-section">
              <div className="section-topline"><div className="step-label"><span>02</span><h2>Confira o resultado</h2></div><span className="ready-label"><Check size={14} /> Pronto para exportar</span></div>
              <div className="stats-row">
                <div className="stat-block"><span className="stat-label">CONTATOS NA PLANILHA</span><strong>{sheet.rows.length.toLocaleString("pt-BR")}</strong></div>
                <div className="stat-block"><span className="stat-label">NOMES AJUSTADOS</span><strong>{adjustedNames.toLocaleString("pt-BR")}</strong></div>
                <div className="stat-block stat-highlight"><span className="stat-label">NOMES PARA REVISAR</span><strong>{namesToReview.toLocaleString("pt-BR")}</strong></div>
                <div className="stat-block"><span className="stat-label">TELEFONES ATUALIZADOS</span><strong>{changedPhones.toLocaleString("pt-BR")}</strong></div>
              </div>
              <div className="table-frame"><div className="table-scroll"><table><thead><tr><th>triagem</th>{sheet.headers.map((header, index) => <th key={`${header}-${index}`}>{header}</th>)}</tr></thead><tbody>
                {sheet.rows.slice(0, 6).map((row, rowIndex) => <tr className={sheet.reviewReasons[rowIndex].length ? "review-row" : ""} key={`${rowIndex}-${row[0]}`}>
                  <td>{sheet.reviewReasons[rowIndex].length ? <><span className="review-flag">Revisar</span><span className="review-detail">{sheet.reviewReasons[rowIndex].join(" · ")}</span></> : <span className="review-clear">OK</span>}</td>
                  {row.map((cell, cellIndex) => <td className={cellIndex === sheet.phoneIndex ? "phone-cell" : ""} key={`${rowIndex}-${cellIndex}`}>
                    {cellIndex === sheet.nameIndex && sheet.sourceNames[rowIndex] !== cell ? <span className="phone-change"><span className="old-phone">{sheet.sourceNames[rowIndex] || "vazio"}</span><ArrowRight size={13} /><span>{cell || "vazio"}</span></span> : cellIndex === sheet.phoneIndex && sheet.sourcePhones[rowIndex] !== cell ? <span className="phone-change"><span className="old-phone">{sheet.sourcePhones[rowIndex] || "vazio"}</span><ArrowRight size={13} /><span>{cell || "vazio"}</span></span> : cell || <span className="empty-cell">—</span>}
                  </td>)}
                </tr>)}
              </tbody></table></div>{sheet.rows.length > 6 && <div className="table-footer">Exibindo 6 de {sheet.rows.length.toLocaleString("pt-BR")} contatos</div>}</div>
            </section>
            <section className="export-bar"><div className="export-copy"><span className="export-icon"><FileSpreadsheet size={19} /></span><div><strong>Planilha tratada</strong><span>Arquivo Excel · .xlsx</span></div></div><div className="export-actions"><button className="download-button course-button" onClick={openCourseEditor} type="button"><FileSpreadsheet size={16} /> Adicionar curso e baixar</button><button className="download-button" onClick={() => void downloadExcel()} type="button"><ArrowDownToLine size={17} /> Baixar Excel</button></div></section>
            <dialog ref={courseDialogRef} className="course-dialog" aria-labelledby="course-dialog-title" onClick={(event) => { if (event.target === event.currentTarget) courseDialogRef.current?.close(); }}>
              <form className="course-form" onSubmit={(event) => void downloadWithCourse(event)}>
                <div className="course-dialog-heading"><div><h2 id="course-dialog-title">Adicionar curso à planilha</h2><p>A mensagem será preenchida na coluna E para todos os contatos.</p></div><button className="course-close" type="button" aria-label="Fechar editor" onClick={() => courseDialogRef.current?.close()}><X size={17} /></button></div>
                <label className="course-field-label" htmlFor="course-message">Mensagem do curso</label>
                <textarea id="course-message" className="course-textarea" value={courseMessage} onChange={(event) => setCourseMessage(event.target.value)} maxLength={32767} placeholder="Ex.: Inscrições abertas para o curso *intensivo* com certificado _incluso_." required />
                <p className="course-format-help">Use *asteriscos* para negrito e _underscores_ para itálico.</p>
                <div className="course-preview-wrap"><span className="course-field-label">Prévia</span><div className="course-preview" aria-live="polite">{courseMessage ? parseCourseMessage(courseMessage).map((part, index) => part.bold ? <strong key={index}>{part.text}</strong> : part.italic ? <em key={index}>{part.text}</em> : <span key={index}>{part.text}</span>) : <span className="course-preview-empty">A prévia da mensagem aparecerá aqui.</span>}</div></div>
                <div className="course-dialog-actions"><button className="course-cancel" type="button" onClick={() => courseDialogRef.current?.close()}>Cancelar</button><button className="download-button" type="submit" disabled={!courseMessage.trim()}><ArrowDownToLine size={16} /> Gerar Excel com curso</button></div>
              </form>
            </dialog>
          </>}

          <footer className="page-footer"><span>Feito para deixar sua base em ordem.</span><span><LockKeyhole size={13} /> Seus dados não saem do navegador</span></footer>
        </div>
      </main>
    </div>
  );
}