import { Suspense } from "react";
import { redirect } from "next/navigation";
import { createClient } from "../../lib/supabase/server";
import { getSupabaseConfig } from "../../lib/supabase/config";
import SeiTreatmentPanel from "../sei-treatment-panel";

export default function SeiTreatmentPage() {
  return (
    <Suspense fallback={<main className="auth-loading">Verificando acesso...</main>}>
      <ProtectedSeiTreatmentPage />
    </Suspense>
  );
}

async function ProtectedSeiTreatmentPage() {
  if (!getSupabaseConfig()) redirect("/login");

  const supabase = await createClient();
  const { data, error } = await supabase.auth.getClaims();
  if (error || !data?.claims) redirect("/login");

  const email = typeof data.claims.email === "string" ? data.claims.email : "";
  return <SeiTreatmentPanel userEmail={email} />;
}