import { supabase } from "./supabaseClient";

// Every record set is stored as one JSON value in kv_store, so a save rewrites the whole
// set. Two people with the app open will each save the copy they loaded, and the second
// write silently erases the first - no error, no trace of the lost entries.
//
// storageSetGuarded closes that by writing only if the row still carries the updated_at
// the caller last read. If someone else has saved in the meantime the write is refused
// and reported as a conflict, so the caller can reload rather than overwrite.

export async function storageGet(key) {
  const { data, error } = await supabase
    .from("kv_store")
    .select("value, updated_at")
    .eq("key", key)
    .maybeSingle();

  if (error) throw error;
  if (!data) return null;
  // `rev` is the value to hand back to storageSetGuarded on the next write.
  return { key, value: data.value, rev: data.updated_at ?? null };
}

// Unconditional write. Fine for settings a single person edits (rates, free-day rules);
// use storageSetGuarded for anything several people add records to.
export async function storageSet(key, value) {
  const { data, error } = await supabase
    .from("kv_store")
    .upsert({ key, value, updated_at: new Date().toISOString() })
    .select("updated_at");

  if (error) throw error;
  return { key, value, rev: data && data[0] ? data[0].updated_at : null };
}

/**
 * Writes `value` only if the stored row is still at revision `rev`.
 *
 * @param rev  the `rev` from the last storageGet/storageSetGuarded for this key, or
 *             null/undefined if the caller believes the key does not exist yet.
 * @returns    { ok: true, rev } on success, or { ok: false, conflict: true, current }
 *             where `current` is the row that is actually there now.
 */
export async function storageSetGuarded(key, value, rev) {
  const now = new Date().toISOString();

  // No revision means "this key should not exist yet". Insert, and let the primary key
  // reject it if another session created it first.
  if (rev === null || rev === undefined) {
    const { data, error } = await supabase
      .from("kv_store")
      .insert({ key, value, updated_at: now })
      .select("updated_at");

    if (!error) return { ok: true, key, value, rev: data && data[0] ? data[0].updated_at : now };
    // 23505 = unique violation: someone else created this key between our read and write.
    if (error.code === "23505" || /duplicate key/i.test(error.message || "")) {
      const current = await storageGet(key);
      return { ok: false, conflict: true, current };
    }
    throw error;
  }

  const { data, error } = await supabase
    .from("kv_store")
    .update({ value, updated_at: now })
    .eq("key", key)
    .eq("updated_at", rev)
    .select("updated_at");

  if (error) throw error;
  if (data && data.length > 0) return { ok: true, key, value, rev: data[0].updated_at };

  // Nothing matched: either the row moved on, or it was deleted entirely.
  const current = await storageGet(key);
  return { ok: false, conflict: true, current };
}
