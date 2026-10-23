import { getSupabaseConfig } from "../../lib/supabase/config";
import LoginForm from "./login-form";

export default function LoginPage() {
  return <LoginForm isConfigured={Boolean(getSupabaseConfig())} />;
}