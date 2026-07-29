create table if not exists kv_store (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

alter table kv_store enable row level security;

create policy "Allow all reads" on kv_store
  for select using (true);

create policy "Allow all inserts" on kv_store
  for insert with check (true);

create policy "Allow all updates" on kv_store
  for update using (true);
