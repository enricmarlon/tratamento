"use client";

import { ArrowDownToLine, ArrowRight, Check, FileSpreadsheet, LockKeyhole, LogOut, Split } from "lucide-react";
import Link from "next/link";
import { useState } from "react";
import { signOut } from "./actions/auth";
import { useWorkflow } from "./workflow-context";
import { createExcelBlob, safeExcelFileName } from "../lib/excel-export";

type PreparedFile = { fileName: string; blob: Blob; contacts: number };

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function SplitPanel({ userEmail }: { userEmail: string }) {
  const { sheet, splitCourseText } = useWorkflow();
  const [splitParts, setSplitParts] = useState<2 | 3>(2);
  const [splitPercentages, setSplitPercentages] = useState(["70", "30"]);
  const [splitFileNames, setSplitFileNames] = useState<string[]>([]);
  const [splitFiles, setSplitFiles] = useState<PreparedFile[]>([]);
  const [splitError, setSplitError] = useState("");
  const [preparing, setPreparing] = useState(false);

  function updateSplitParts(parts: 2 | 3) {
    setSplitParts(parts);
    setSplitPercentages(parts === 2 ? ["70", "30"] : ["50", "30", "20"]);
    setSplitFileNames([]);
    setSplitFiles([]);
    setSplitError("");
  }

  async function prepareSplitFiles() {
    if (!sheet || preparing) return;
    const percentages = splitPercentages.slice(0, splitParts).map(Number);
    const totalPercentage = percentages.reduce((total, percentage) => total + percentage, 0);
    if (percentages.some((percentage) => !Number.isFinite(percentage) || percentage <= 0) || Math.abs(totalPercentage - 100) > 0.001) {
      setSplitError("Informe percentuais maiores que zero cuja soma seja 100%.");
      return;
    }
    if (sheet.rows.length < splitParts) {
      setSplitError(`A base precisa ter pelo menos ${splitParts} contatos para gerar arquivos não vazios.`);
      return;
    }

    setPreparing(true);
    setSplitError("");
    setSplitFiles([]);
    const baseName = sheet.fileName.replace(/\.csv$/i, "");
    const preparedFiles: PreparedFile[] = [];
    let start = 0;

    try {
      for (let index = 0; index < splitParts; index += 1) {
        const remainingParts = splitParts - index - 1;
        const partSize = remainingParts === 0
          ? sheet.rows.length - start
          : Math.min(Math.max(Math.round(sheet.rows.length * percentages[index] / 100), 1), sheet.rows.length - start - remainingParts);
        const rows = sheet.rows.slice(start, start + partSize);
        const blob = await createExcelBlob(sheet.headers, rows, splitCourseText);
        preparedFiles.push({
          fileName: safeExcelFileName(splitFileNames[index] || `${baseName}_parte_${index + 1}`),
          blob,
          contacts: rows.length,
        });
        start += partSize;
      }

      if (new Set(preparedFiles.map((file) => file.fileName.toLocaleLowerCase("pt-BR"))).size !== preparedFiles.length) {
        setSplitError("Cada arquivo precisa ter um nome diferente.");
        return;
      }
      setSplitFiles(preparedFiles);
    } catch {
      setSplitError("Não foi possível preparar os arquivos Excel. Tente novamente.");
    } finally {
      setPreparing(false);
    }
  }

  const splitPercentageTotal = splitPercentages.slice(0, splitParts).reduce((total, percentage) => total + (Number(percentage) || 0), 0);
  const validSplitConfiguration = splitPercentages.slice(0, splitParts).every((percentage) => Number(percentage) > 0)
    && Math.abs(splitPercentageTotal - 100) <= 0.001
    && (sheet?.rows.length ?? 0) >= splitParts;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <Link className="brand" href="/" aria-label="Plan, início"><span className="brand-mark"><FileSpreadsheet size={20} /></span><span>Plan</span></Link>
        <div className="sidebar-caption">FERRAMENTAS</div>
        <Link className="nav-item tool-nav-link" href="/"><FileSpreadsheet size={17} /><span>Tratar contatos</span></Link>
        <Link className="nav-item nav-item-active split-nav-item tool-nav-link" href="/dividir"><span className="nav-indicator" /><Split size={17} /><span>Dividir</span></Link>
        <div className="sidebar-bottom"><LockKeyhole size={16} /><div><strong>Seus dados, protegidos</strong><p>Os arquivos ficam no seu dispositivo.</p></div></div>
      </aside>

      <main className="main-content">
        <header className="topbar"><span className="breadcrumb">Ferramentas <span>/</span> Dividir base</span><div className="topbar-actions"><span className="local-badge"><span /> Processamento local</span><span className="auth-email" title={userEmail}>{userEmail}</span><form action={signOut}><button className="logout-button" type="submit"><LogOut size={15} /><span>Sair</span></button></form></div></header>
        <div className="content-wrap">
          <section className="page-heading">
            <div><div className="eyebrow">FERRAMENTAS DE PLANILHA</div><h1>Dividir sua base.</h1><p>Separe os contatos tratados em arquivos Excel independentes.</p></div>
            <div className="format-note"><span>BASE</span><ArrowRight size={15} /><span className="format-xlsx">XLSX</span></div>
          </section>

          {sheet ? (
            <>
              <div className="file-loaded split-source-file"><div className="loaded-file-icon"><FileSpreadsheet size={20} /></div><div className="loaded-file-info"><strong>{sheet.fileName}</strong><span>{sheet.rows.length.toLocaleString("pt-BR")} contatos disponíveis</span></div><span className="loaded-check"><Check size={16} /></span></div>
              <section className="workflow split-workflow" aria-label="Dividir base">
                <div className="section-topline"><div className="step-label"><span>01</span><h2>Configure os arquivos</h2></div>{splitCourseText !== undefined && <span className="split-course-note"><Check size={14} /> Curso incluído</span>}</div>
                <div className="split-controls">
                  <div className="split-control-group"><span className="course-field-label">Quantidade de arquivos</span><div className="split-segment" role="group" aria-label="Quantidade de arquivos">
                    <button className={splitParts === 2 ? "is-selected" : ""} type="button" aria-pressed={splitParts === 2} onClick={() => updateSplitParts(2)}>2 arquivos</button>
                    <button className={splitParts === 3 ? "is-selected" : ""} type="button" aria-pressed={splitParts === 3} onClick={() => updateSplitParts(3)}>3 arquivos</button>
                  </div></div>
                  <div className="split-control-group split-total"><span className="course-field-label">Distribuição</span><span className={Math.abs(splitPercentageTotal - 100) <= 0.001 ? "split-total-valid" : "split-total-invalid"}>Soma: {splitPercentageTotal.toLocaleString("pt-BR", { maximumFractionDigits: 2 })}%</span></div>
                </div>
                <div className="split-fields">{Array.from({ length: splitParts }, (_, index) => {
                  const defaultName = `${sheet.fileName.replace(/\.csv$/i, "")}_parte_${index + 1}`;
                  return <div className="split-file-field" key={index}>
                    <label className="course-field-label" htmlFor={`split-percent-${index}`}>Arquivo {index + 1}</label>
                    <div className="split-input-row"><input id={`split-percent-${index}`} className="split-percent-input" type="number" min="0.1" max="99.9" step="0.1" value={splitPercentages[index] ?? ""} onChange={(event) => { const next = [...splitPercentages]; next[index] = event.target.value; setSplitPercentages(next); setSplitFiles([]); setSplitError(""); }} aria-label={`Porcentagem do arquivo ${index + 1}`} /><span className="split-percent-sign">%</span><input className="split-name-input" type="text" value={splitFileNames[index] ?? defaultName} onChange={(event) => { const next = [...splitFileNames]; next[index] = event.target.value; setSplitFileNames(next); setSplitFiles([]); setSplitError(""); }} aria-label={`Nome do arquivo ${index + 1}`} /><span className="split-extension">.xlsx</span></div>
                  </div>;
                })}</div>
                {splitError && <p className="error-message" role="alert">{splitError}</p>}
                <div className="split-actions"><button className="download-button" type="button" onClick={() => void prepareSplitFiles()} disabled={!validSplitConfiguration || preparing}><Split size={16} /> {preparing ? "Preparando..." : "Preparar arquivos"}</button><span>Os arquivos ficam prontos para baixar individualmente.</span></div>
                {splitFiles.length > 0 && <div className="split-download-list" aria-live="polite">{splitFiles.map((file) => <div className="split-download-item" key={file.fileName}><div><strong>{file.fileName}</strong><span>{file.contacts.toLocaleString("pt-BR")} contatos</span></div><button className="download-button" type="button" onClick={() => downloadBlob(file.blob, file.fileName)}><ArrowDownToLine size={15} /> Baixar</button></div>)}</div>}
              </section>
            </>
          ) : (
            <section className="split-empty-state">
              <span className="export-icon"><Split size={20} /></span>
              <div><h2>Nenhuma base pronta para dividir</h2><p>Importe e trate um CSV primeiro. Depois, volte aqui pela coluna de ferramentas.</p></div>
              <Link className="download-button" href="/"><FileSpreadsheet size={16} /> Ir para Tratar contatos</Link>
            </section>
          )}

          <footer className="page-footer"><span>Feito para deixar sua base em ordem.</span><span><LockKeyhole size={13} /> Seus dados não saem do navegador</span></footer>
        </div>
      </main>
    </div>
  );
}