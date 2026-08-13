-- =============================================
-- MT 設備系統 — 頭像
-- 取代「頭像資料」工作表（原本把 base64 字串直接塞在格子裡）
-- 執行方式：Supabase 後台 → SQL Editor → 貼上 → Run
-- =============================================

create table if not exists public.avatars (
  name       text        primary key,   -- 使用者姓名，與 Keeper 名稱一致
  image_data text        not null,      -- 完整的 data:image/jpeg;base64,... 字串
  updated_at timestamptz default now()
);

-- RLS：與現有資料表相同的開放政策（前端用 publishable key 直接讀寫）
alter table public.avatars enable row level security;

drop policy if exists "avatars public read"  on public.avatars;
drop policy if exists "avatars public write" on public.avatars;

create policy "avatars public read"  on public.avatars for select using (true);
create policy "avatars public write" on public.avatars for all    using (true) with check (true);
