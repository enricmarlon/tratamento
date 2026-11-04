"use client";

import { ArrowDownToLine, Check, ClipboardList, FileSpreadsheet, LockKeyhole, LogOut, RotateCcw, Split, Upload } from "lucide-react";
import Link from "next/link";
import { useRef, useState } from "react";
import { signOut } from "./actions/auth";
import { readSeiSpreadsheet, type SeiColumn, type SeiColumnId, type SeiSpreadsheet } from "../lib/sei-spreadsheet";
import { createExcelBlob, safeExcelFileName } from "../lib/excel-export";

const MAX_FILE_SIZE = 30 * 1024 * 1024;
const ACCEPTED_EXTENSIONS = /\.(xlsx|xls|ods)$/i;
const ACCEPTED_MIME_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.oasis.opendocument.spreadsheet",
]);

function cleanCell(value: string, reduceSpaces: boolean) {
  const trimmed = value.trim();
  return reduceSpaces ? trimmed.replace(/\s+/g, " ") : trimmed;
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function SeiTreatmentPanel({ userEmail }: { userEmail: string }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [sheet, setSheet] = useState<SeiSpreadsheet | null>(null);
  const [selectedColumns, setSelectedColumns] = useState<SeiColumnId[]>([]);
  const [columnNames, setColumnNames] = useState<Partial<Record<SeiColumnId, string>>>({});
  const [fileName, setFileName] = useState("");
  const [reduceSpaces, setReduceSpaces] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");

  async function importFile(file?: File) {
    if (!file) return;
    setError("");
    if (!ACCEPTED_EXTENSIONS.test(file.name) && !ACCEPTED_MIME_TYPES.has(file.type)) {
      setError("Selecione um arquivo XLS, XLSX ou ODS.");
      if (inputRef.current) inputRef.current.value = "";
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      setError("O arquivo excede o limite de 30 MB.");
      if (inputRef.current) inputRef.current.value = "";
      return;
    }

    setLoading(true);
    try {
      const imported = await readSeiSpreadsheet(file);
      if (!imported.rows.length || !imported.columns.length) {
        setSheet(null);
        setError("Não encontrei uma tabela com as colunas Matrícula e Nome do Aluno em nenhuma aba.");
        return;
      }

      setSheet(imported);
      setSelectedColumns(imported.columns.map((column) => column.id));
      setColumnNames(Object.fromEntries(imported.columns.map((column) => [column.id, column.id === "celular" ? "telefone" : column.label])));
      setFileName(`${file.name.replace(ACCEPTED_EXTENSIONS, "")}_organizado`);
      setReduceSpaces(false);
    } catch {
      setSheet(null);
      setError("Não foi possível ler a planilha. Confira se o arquivo está íntegro.");
    } finally {
      setLoading(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function reset() {
    setSheet(null);
    setSelectedColumns([]);
    setColumnNames({});
    setFileName("");
    setReduceSpaces(false);
    setError("");
    if (inputRef.current) inputRef.current.value = "";
  }

  function toggleColumn(columnId: SeiColumnId) {
    setSelectedColumns((current) => current.includes(columnId)
      ? current.filter((id) => id !== columnId)
      : [...current, columnId]);
  }

  const visibleColumns = sheet?.columns.filter((column) => selectedColumns.includes(column.id)) ?? [];

  async function exportSpreadsheet() {
    if (!sheet || !visibleColumns.length || exporting) return;
    setExporting(true);
    setError("");
    try {
      const headers = visibleColumns.map((column) => columnNames[column.id]?.trim() || column.label);
      const rows = sheet.rows.map((record) => visibleColumns.map((column) => cleanCell(record[column.id] ?? "", reduceSpaces)));
      const blob = await createExcelBlob(headers, rows);
      const baseName = fileName.replace(ACCEPTED_EXTENSIONS, "").trim() || "tratamento_sei";
      downloadBlob(blob, safeExcelFileName(baseName));
    } catch {
      setError("Não foi possível gerar o arquivo Excel. Tente novamente.");
    } finally {
      setExporting(false);
    }
  }

  function updateColumnName(column: SeiColumn, value: string) {
    setColumnNames((current) => ({ ...current, [column.id]: value }));
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <Link className="brand" href="/" aria-label="Plan, início"><span className="brand-mark"><FileSpreadsheet size={20} /></span><span>Plan</span></Link>
        <div className="sidebar-caption">FERRAMENTAS</div>
        <Link className="nav-item tool-nav-link" href="/"><FileSpreadsheet size={17} /><span>Tratar contatos</span></Link>
        <Link className="nav-item tool-nav-link" href="/dividir"><Split size={17} /><span>Dividir</span></Link>
        <Link className="nav-item nav-item-active tool-nav-link" href="/tratamento-sei"><span className="nav-indicator" /><ClipboardList size={17} /><span>Tratamento SEI</span></Link>
        <div className="sidebar-bottom"><LockKeyhole size={16} /><div><strong>Seus dados, protegidos</strong><p>Os arquivos ficam no seu dispositivo.</p></div></div>
      </aside>

      <main className="main-content">
        <header className="topbar"><span className="breadcrumb">Ferramentas <span>/</span> Tratamento SEI</span><div className="topbar-actions"><span className="local-badge"><span /> Processamento local</span><span className="auth-email" title={userEmail}>{userEmail}</span><form action={signOut}><button className="logout-button" type="submit"><LogOut size={15} /><span>Sair</span></button></form></div></header>
        <div className="content-wrap">
          <section className="page-heading">
            <div><div className="eyebrow">ORGANIZAÇÃO DE RELATÓRIOS</div><h1>Tratamento SEI.</h1><p>Consolide os dados dos alunos e prepare uma planilha limpa para exportação.</p></div>
            <div className="format-note"><span>RELATÓRIO</span><ArrowDownToLine size={15} /><span className="format-xlsx">XLSX</span></div>
          </section>

          <section className="workflow" aria-label="Importar relatório SEI">
            <div className="section-topline"><div className="step-label"><span>01</span><h2>Importe o relatório</h2></div>{sheet && <button className="text-button" onClick={reset} type="button"><RotateCcw size={15} /> Trocar arquivo</button>}</div>
            <input ref={inputRef} id="sei-file" className="sei-file-input" type="file" accept=".xlsx,.xls,.ods,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.oasis.opendocument.spreadsheet" onChange={(event) => void importFile(event.target.files?.[0])} />
            {!sheet ? (
              <label className={`upload-zone${loading ? " is-reading" : ""}`} htmlFor="sei-file" onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); void importFile(event.dataTransfer.files[0]); }}>
                <span className="upload-icon"><Upload size={21} /></span><span className="upload-title">{loading ? "Lendo sua planilha..." : "Solte o relatório Excel ou ODS aqui"}</span>
                <span className="upload-description">ou <span className="upload-link">escolha no computador</span></span><span className="upload-meta">XLS · XLSX · ODS · até 30 MB</span>
              </label>
            ) : (
              <div className="file-loaded"><div className="loaded-file-icon"><FileSpreadsheet size={20} /></div><div className="loaded-file-info"><strong>{sheet.rows.length.toLocaleString("pt-BR")} registros encontrados</strong><span>{sheet.columns.length} colunas com dados · {sheet.emptyRowsRemoved.toLocaleString("pt-BR")} linhas vazias ignoradas</span></div><button className="text-button sei-replace-button" onClick={() => inputRef.current?.click()} type="button"><Upload size={14} /> Trocar</button><span className="loaded-check"><Check size={16} /></span></div>
            )}
            {error && <p className="error-message" role="alert">{error}</p>}
          </section>

          {sheet && <>
            <section className="workflow sei-workflow" aria-label="Organizar colunas">
              <div className="section-topline"><div className="step-label"><span>02</span><h2>Escolha e renomeie as colunas</h2></div><span className="sei-selected-count">{visibleColumns.length} de {sheet.columns.length} selecionadas</span></div>
              <div className="sei-options-row">
                <label className="sei-space-toggle"><input type="checkbox" checked={reduceSpaces} onChange={(event) => setReduceSpaces(event.target.checked)} /><span><strong>Reduzir espaços internos repetidos</strong><small>Espaços no início e no fim são sempre removidos.</small></span></label>
                <div className="sei-column-actions"><button className="text-button" type="button" onClick={() => setSelectedColumns(sheet.columns.map((column) => column.id))}>Todas</button><button className="text-button" type="button" onClick={() => setSelectedColumns([])}>Nenhuma</button></div>
              </div>
              <div className="sei-column-list">
                <div className="sei-column-header"><span>Manter</span><span>Coluna original</span><span>Nome no Excel</span></div>
                {sheet.columns.map((column) => {
                  const selected = selectedColumns.includes(column.id);
                  return <div className={`sei-column-row${selected ? "" : " is-excluded"}`} key={column.id}>
                    <label className="sei-checkbox-cell"><input type="checkbox" checked={selected} onChange={() => toggleColumn(column.id)} aria-label={`Manter ${column.label}`} /></label>
                    <span className="sei-original-name">{column.label}</span>
                    <input className="sei-rename-input" type="text" value={columnNames[column.id] ?? column.label} onChange={(event) => updateColumnName(column, event.target.value)} aria-label={`Renomear coluna ${column.label}`} disabled={!selected} />
                  </div>;
                })}
              </div>
            </section>

            <section className="result-section sei-result" aria-label="Prévia da planilha organizada">
              <div className="section-topline"><div className="step-label"><span>03</span><h2>Prévia da planilha</h2></div><span className="sei-preview-count">{sheet.rows.length.toLocaleString("pt-BR")} linhas</span></div>
              {visibleColumns.length ? (
                <div className="table-frame"><div className="table-scroll"><table className="sei-preview-table"><thead><tr>{visibleColumns.map((column) => <th key={column.id}>{columnNames[column.id]?.trim() || column.label}</th>)}</tr></thead><tbody>
                  {sheet.rows.slice(0, 8).map((record, rowIndex) => <tr key={`${record.matricula ?? "registro"}-${rowIndex}`}>{visibleColumns.map((column) => <td key={column.id}>{cleanCell(record[column.id] ?? "", reduceSpaces) || <span className="empty-cell">—</span>}</td>)}</tr>)}
                </tbody></table></div>{sheet.rows.length > 8 && <div className="table-footer">Exibindo 8 de {sheet.rows.length.toLocaleString("pt-BR")} linhas</div>}</div>
              ) : <p className="sei-no-columns">Selecione pelo menos uma coluna para ver a prévia.</p>}
            </section>

            <section className="sei-export-bar" aria-label="Exportar arquivo">
              <label className="sei-filename-field"><span>Nome do arquivo</span><div><input type="text" value={fileName} onChange={(event) => setFileName(event.target.value)} maxLength={150} /><span>.xlsx</span></div></label>
              <button className="download-button" type="button" onClick={() => void exportSpreadsheet()} disabled={!visibleColumns.length || exporting}><ArrowDownToLine size={16} /> {exporting ? "Gerando Excel..." : "Baixar Excel organizado"}</button>
            </section>
          </>}

          <footer className="page-footer"><span>Feito para deixar sua base em ordem.</span><span><LockKeyhole size={13} /> Seus dados não saem do navegador</span></footer>
        </div>
      </main>
    </div>
  );
}