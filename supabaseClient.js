import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.error(
    "VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY არ არის მითითებული — შეამოწმეთ .env ფაილი ან Vercel-ის Environment Variables."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
