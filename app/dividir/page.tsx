import { Suspense } from "react";
import { redirect } from "next/navigation";
import { createClient } from "../../lib/supabase/server";
import { getSupabaseConfig } from "../../lib/supabase/config";
import SplitPanel from "../split-panel";

export default function SplitPage() {
  return (
    <Suspense fallback={<main className="auth-loading">Verificando acesso...</main>}>
      <ProtectedSplitPage />
    </Suspense>
  );
}

async function ProtectedSplitPage() {
  if (!getSupabaseConfig()) redirect("/login");

  const supabase = await createClient();
  const { data, error } = await supabase.auth.getClaims();
  if (error || !data?.claims) redirect("/login");

  const email = typeof data.claims.email === "string" ? data.claims.email : "";
  return <SplitPanel userEmail={email} />;
}