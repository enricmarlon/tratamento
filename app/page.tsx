import { Suspense } from "react";
import HomePanel from "./home-panel";
import { createClient } from "../lib/supabase/server";
import { getSupabaseConfig } from "../lib/supabase/config";
import { redirect } from "next/navigation";

export default function Home() {
  return (
    <Suspense fallback={<main className="auth-loading">Verificando acesso...</main>}>
      <ProtectedHome />
    </Suspense>
  );
}

async function ProtectedHome() {
  if (!getSupabaseConfig()) redirect("/login");

  const supabase = await createClient();
  const { data, error } = await supabase.auth.getClaims();
  if (error || !data?.claims) redirect("/login");

  const email = typeof data.claims.email === "string" ? data.claims.email : "";
  return <HomePanel userEmail={email} />;
}