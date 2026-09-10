-- გაუშვით ეს მთლიანად Supabase-ს პროექტში: SQL Editor → New query → ჩასვით → Run

create table if not exists kv_store (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

-- ინდექსი, რომ "პრეფიქსით ძებნა" (company:, courier:, parcel:...) სწრაფი იყოს
create index if not exists kv_store_key_prefix_idx on kv_store (key text_pattern_ops);

-- ავტომატურად ანახლებს updated_at-ს ყოველ ჩაწერაზე
create or replace function kv_store_set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists kv_store_updated_at on kv_store;
create trigger kv_store_updated_at
  before update on kv_store
  for each row execute function kv_store_set_updated_at();

-- Row Level Security: ჩართულია, მაგრამ წესები საჯაროდ ღიაა (ანუ ნებისმიერს, ვისაც
-- საიტის ბმული აქვს, შეუძლია წაკითხვა/ჩაწერა) — ეს ზუსტად იმეორებს იმ დონის
-- (არა-კრიპტოგრაფიულ) დაცვას, რაც Claude-ის არტიფაქტში იყო. საჭიროების
-- შემთხვევაში მოგვიანებით შეგიძლიათ ეს წესები გაამკაცროთ Supabase Auth-ით.
alter table kv_store enable row level security;

drop policy if exists "public read" on kv_store;
create policy "public read" on kv_store for select using (true);

drop policy if exists "public insert" on kv_store;
create policy "public insert" on kv_store for insert with check (true);

drop policy if exists "public update" on kv_store;
create policy "public update" on kv_store for update using (true);

drop policy if exists "public delete" on kv_store;
create policy "public delete" on kv_store for delete using (true);
