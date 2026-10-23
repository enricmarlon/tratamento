"use client";

import { createContext, useContext, useState, type Dispatch, type ReactNode, type SetStateAction } from "react";

export type ContactRecord = {
  row: string[];
  sourcePhones: string;
  sourceNames: string;
  reviewReasons: string[];
};

export type ContactSheetBase = {
  headers: string[];
  allRecords: ContactRecord[];
  phoneIndex: number;
  nameIndex: number;
  fileName: string;
};

export type ContactSheet = ContactSheetBase & {
  rows: string[][];
  sourcePhones: string[];
  sourceNames: string[];
  reviewReasons: string[][];
  removedCount: number;
};

type WorkflowContextValue = {
  sheet: ContactSheet | null;
  setSheet: Dispatch<SetStateAction<ContactSheet | null>>;
  splitCourseText: string | undefined;
  setSplitCourseText: Dispatch<SetStateAction<string | undefined>>;
};

const WorkflowContext = createContext<WorkflowContextValue | null>(null);

export function WorkflowProvider({ children }: { children: ReactNode }) {
  const [sheet, setSheet] = useState<ContactSheet | null>(null);
  const [splitCourseText, setSplitCourseText] = useState<string | undefined>();

  return (
    <WorkflowContext.Provider value={{ sheet, setSheet, splitCourseText, setSplitCourseText }}>
      {children}
    </WorkflowContext.Provider>
  );
}

export function useWorkflow() {
  const context = useContext(WorkflowContext);
  if (!context) throw new Error("useWorkflow precisa ser usado dentro de WorkflowProvider.");
  return context;
}