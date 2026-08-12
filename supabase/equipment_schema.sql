-- =============================================
-- MT 設備系統 — Supabase 設備清單表
-- 用途：取代 GAS 讀取「工作表 1」+「網站新增設備」
-- 執行方式：Supabase 後台 → SQL Editor → 貼上 → Run
-- =============================================

create table if not exists public.equipment (
  fix_no            text primary key,              -- 設備編號（唯一，來自 Sheet B 欄）
  fix_type          text        default '',        -- 類型
  device_name       text        default '',        -- 設備名稱
  qty_asset         text        default '1',       -- 數量
  keeper            text        default '',        -- 保管人（「我的設備」用這欄篩）
  status            text        default 'available',
  borrower          text        default '',
  dt_borrow         text        default '',        -- 日期一律存字串 yyyy-MM-dd，與 GAS formatDate 輸出一致
  dt_due            text        default '',
  dt_return         text        default '',
  return_confirmed  boolean     default false,
  source            text        default 'sheet1',  -- sheet1 | web：記錄原本在哪張工作表，GAS 回寫時要用
  updated_at        timestamptz default now()
);

-- 「我的設備」與首頁篩選會用到
create index if not exists equipment_keeper_idx on public.equipment (keeper);
create index if not exists equipment_status_idx on public.equipment (status);

-- RLS：與 station_bookings 相同的開放政策（前端用 publishable key 直接讀寫）
alter table public.equipment enable row level security;

drop policy if exists "equipment public read"   on public.equipment;
drop policy if exists "equipment public insert" on public.equipment;
drop policy if exists "equipment public update" on public.equipment;
drop policy if exists "equipment public delete" on public.equipment;

create policy "equipment public read"   on public.equipment for select using (true);
create policy "equipment public insert" on public.equipment for insert with check (true);
create policy "equipment public update" on public.equipment for update using (true) with check (true);
create policy "equipment public delete" on public.equipment for delete using (true);
