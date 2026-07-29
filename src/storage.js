import { supabase } from "./supabaseClient";

export async function storageGet(key) {
  const { data, error } = await supabase
    .from("kv_store")
    .select("value")
    .eq("key", key)
    .maybeSingle();

  if (error) throw error;
  if (!data) return null;
  return { key, value: data.value };
}

export async function storageSet(key, value) {
  const { error } = await supabase
    .from("kv_store")
    .upsert({ key, value, updated_at: new Date().toISOString() });

  if (error) throw error;
  return { key, value };
}
