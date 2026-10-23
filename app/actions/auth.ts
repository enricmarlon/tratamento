"use server";

import { redirect } from "next/navigation";
import { getSupabaseConfig } from "../../lib/supabase/config";
import { createClient } from "../../lib/supabase/server";

export type LoginState = {
  error?: string;
};

export async function signIn(_previousState: LoginState, formData: FormData): Promise<LoginState> {
  const emailValue = formData.get("email");
  const passwordValue = formData.get("password");

  if (typeof emailValue !== "string" || typeof passwordValue !== "string") {
    return { error: "Informe seu e-mail e sua senha." };
  }

  const email = emailValue.trim();
  if (!email || email.length > 320 || !email.includes("@") || !passwordValue || passwordValue.length > 1024) {
    return { error: "E-mail ou senha inválidos." };
  }

  if (!getSupabaseConfig()) {
    return { error: "A autenticação ainda não foi configurada neste ambiente." };
  }

  try {
    const supabase = await createClient();
    const { error } = await supabase.auth.signInWithPassword({ email, password: passwordValue });
    if (error) return { error: "Não foi possível entrar. Confira seus dados ou fale com o administrador." };
  } catch {
    return { error: "Não foi possível conectar ao serviço de autenticação. Tente novamente." };
  }

  return redirect("/");
}

export async function signOut() {
  if (getSupabaseConfig()) {
    try {
      const supabase = await createClient();
      await supabase.auth.signOut();
    } catch {
      return redirect("/login");
    }
  }

  return redirect("/login");
}