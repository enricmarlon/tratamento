"use client";

import { LockKeyhole, LogIn } from "lucide-react";
import Link from "next/link";
import { useActionState } from "react";
import { useFormStatus } from "react-dom";
import { signIn } from "../actions/auth";

function SubmitButton({ disabled }: { disabled: boolean }) {
  const { pending } = useFormStatus();

  return (
    <button className="login-submit" disabled={pending || disabled} type="submit">
      <LogIn size={17} />
      {pending ? "Validando acesso..." : "Entrar"}
    </button>
  );
}

export default function LoginForm({ isConfigured }: { isConfigured: boolean }) {
  const [state, formAction] = useActionState(signIn, {});

  return (
    <main className="login-page">
      <section className="login-panel" aria-labelledby="login-title">
        <Link className="brand login-brand" href="/" aria-label="Plan">
          <span className="brand-mark"><LockKeyhole size={19} /></span>
          <span>Plan</span>
        </Link>
        <div className="login-heading">
          <span className="eyebrow">ACESSO RESTRITO</span>
          <h1 id="login-title">Entre na sua conta</h1>
          <p>Use seu e-mail e senha autorizados para continuar.</p>
        </div>

        {!isConfigured && (
          <p className="login-notice" role="status">
            O Supabase ainda não foi configurado. Adicione as variáveis de ambiente para habilitar o acesso.
          </p>
        )}
        {state.error && <p className="login-error" role="alert">{state.error}</p>}

        <form action={formAction} className="login-form">
          <label htmlFor="email">E-mail</label>
          <input autoComplete="username" autoCapitalize="none" id="email" name="email" placeholder="voce@empresa.com" required type="email" />
          <label htmlFor="password">Senha</label>
          <input autoComplete="current-password" id="password" name="password" placeholder="Sua senha" required type="password" />
          <SubmitButton disabled={!isConfigured} />
        </form>
        <p className="login-footnote"><LockKeyhole size={13} /> Acesso somente para usuários autorizados.</p>
      </section>
    </main>
  );
}