"use client";

import ExcelJS from "exceljs";
import { ArrowDownToLine, ArrowRight, Check, FileSpreadsheet, LockKeyhole, RotateCcw, Upload } from "lucide-react";
import { useRef, useState } from "react";

type ContactSheet = {
  headers: string[];
  rows: string[][];
  sourcePhones: string[];
  phoneIndex: number;
  fileName: string;
};

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

export default function HomePanel() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [sheet, setSheet] = useState<ContactSheet | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

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
      const sourceRows = parsed.slice(1).filter((row) => row.some((cell) => cell.trim()));
      if (!sourceRows.length) {
        setSheet(null);
        setError("O CSV não contém contatos para tratar.");
        return;
      }

      const sourcePhones = sourceRows.map((row) => row[phoneIndex] ?? "");
      const rows = sourceRows.map((source) => {
        const row = headers.map((_, index) => source[index] ?? "");
        row[phoneIndex] = cleanPhone(row[phoneIndex]);
        return row;
      });
      setSheet({ headers, rows, sourcePhones, phoneIndex, fileName: file.name });
    } catch {
      setSheet(null);
      setError("Não foi possível ler esse arquivo. Confira se ele está em formato CSV.");
    } finally {
      setLoading(false);
    }
  }

  async function downloadExcel() {
    if (!sheet) return;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Contatos tratados");
    worksheet.addRow(sheet.headers);
    sheet.rows.forEach((row) => worksheet.addRow(row));
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.columns = sheet.headers.map((header) => ({ width: Math.min(Math.max(header.length + 4, 16), 32) }));

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${sheet.fileName.replace(/\.csv$/i, "")}_tratado.xlsx`;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function reset() {
    setSheet(null);
    setError("");
    if (inputRef.current) inputRef.current.value = "";
  }

  const identifiedPhones = sheet?.sourcePhones.filter((phone) => phone.replace(/\D/g, "")).length ?? 0;
  const changedPhones = sheet?.sourcePhones.filter((phone) => cleanPhone(phone) !== phone).length ?? 0;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#inicio" aria-label="Folha, início"><span className="brand-mark"><FileSpreadsheet size={20} /></span><span>folha<span className="brand-period">.</span></span></a>
        <div className="sidebar-caption">FERRAMENTAS</div>
        <div className="nav-item nav-item-active"><span className="nav-indicator" /><FileSpreadsheet size={17} /><span>Tratar contatos</span></div>
        <div className="sidebar-bottom"><LockKeyhole size={16} /><div><strong>Seus dados, protegidos</strong><p>Os arquivos ficam no seu dispositivo.</p></div></div>
      </aside>

      <main className="main-content" id="inicio">
        <header className="topbar"><span className="breadcrumb">Ferramentas <span>/</span> Tratar contatos</span><span className="local-badge"><span /> Processamento local</span></header>
        <div className="content-wrap">
          <section className="page-heading">
            <div><div className="eyebrow">ORGANIZE SUA BASE</div><h1>Contatos prontos<br className="mobile-break" /> para conversar.</h1><p>Padronize os telefones da sua planilha e exporte tudo em Excel.</p></div>
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
                <div className="stat-block"><span className="stat-label">TELEFONES IDENTIFICADOS</span><strong>{identifiedPhones.toLocaleString("pt-BR")}</strong></div>
                <div className="stat-block stat-highlight"><span className="stat-label">NÚMEROS ATUALIZADOS</span><strong>{changedPhones.toLocaleString("pt-BR")}</strong></div>
              </div>
              <div className="table-frame"><div className="table-scroll"><table><thead><tr>{sheet.headers.map((header, index) => <th key={`${header}-${index}`}>{header}</th>)}</tr></thead><tbody>
                {sheet.rows.slice(0, 6).map((row, rowIndex) => <tr key={`${rowIndex}-${row[0]}`}>{row.map((cell, cellIndex) => <td className={cellIndex === sheet.phoneIndex ? "phone-cell" : ""} key={`${rowIndex}-${cellIndex}`}>
                  {cellIndex === sheet.phoneIndex && sheet.sourcePhones[rowIndex] !== cell ? <span className="phone-change"><span className="old-phone">{sheet.sourcePhones[rowIndex] || "vazio"}</span><ArrowRight size={13} /><span>{cell || "vazio"}</span></span> : cell || <span className="empty-cell">—</span>}
                </td>)}</tr>)}
              </tbody></table></div>{sheet.rows.length > 6 && <div className="table-footer">Exibindo 6 de {sheet.rows.length.toLocaleString("pt-BR")} contatos</div>}</div>
            </section>
            <section className="export-bar"><div className="export-copy"><span className="export-icon"><FileSpreadsheet size={19} /></span><div><strong>Planilha tratada</strong><span>Arquivo Excel · .xlsx</span></div></div><button className="download-button" onClick={() => void downloadExcel()} type="button"><ArrowDownToLine size={17} /> Baixar Excel</button></section>
          </>}

          <footer className="page-footer"><span>Feito para deixar sua base em ordem.</span><span><LockKeyhole size={13} /> Seus dados não saem do navegador</span></footer>
        </div>
      </main>
    </div>
  );
}