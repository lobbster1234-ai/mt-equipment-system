-- =============================================
-- MT 設備系統 — 「🏢 手動輸入設備」搬遷用資料表
-- 執行方式：Supabase 後台 → SQL Editor → 貼上 → Run
-- =============================================

-- ---------------------------------------------
-- 1. 部門儀器借用紀錄（取代「MT部門儀器」工作表）
-- ---------------------------------------------
create table if not exists public.dept_equipment (
  id             text primary key,           -- 沿用 GAS 原本的 8 碼 ID，舊資料才對得上
  device_name    text        not null default '',
  borrower       text        not null default '',
  borrower_email text        not null default '',
  dt_borrow      text        default '',      -- 存字串：有 yyyy-MM-dd 也有 yyyy-MM-ddTHH:mm 兩種格式
  dt_due         text        default '',
  dt_return      text        default '',      -- 空字串 = 還沒歸還
  status         text        default '借用中',
  created_at     timestamptz default now()
);

-- 列表預設只看未歸還的，逾期提醒也靠這個
create index if not exists dept_equipment_open_idx on public.dept_equipment (dt_return, dt_due);


-- ---------------------------------------------
-- 2. Keeper 聯絡資訊
-- 注意：Google Sheet 仍是正本（還有 8 個 GAS 功能在讀它），
-- 這裡是 GAS 定時同步過來的副本，只給「手動輸入設備」用。
-- 要新增/修改 Keeper 請去 Sheet 改，不要直接改這張表。
-- ---------------------------------------------
create table if not exists public.keepers (
  name       text primary key,
  email      text        not null default '',
  updated_at timestamptz default now()
);


-- ---------------------------------------------
-- 3. 寄信佇列
-- 使用者操作當下寫一筆進來，GAS 定時觸發器每 5 分鐘撈出來寄。
--
-- 刻意「不」讓前端指定收件人與信件內容，只傳事件類型 + 資料，
-- 由 GAS 套自己的範本組信。否則這張表等於一個公開的轉信站，
-- 任何人都能用你的 Google 帳號寄任意內容給任意人。
-- ---------------------------------------------
create table if not exists public.mail_queue (
  id         bigserial   primary key,
  mail_type  text        not null,            -- dept_borrow | dept_return
  payload    jsonb       not null default '{}',
  status     text        not null default 'pending',  -- pending | sent | failed
  attempts   int         not null default 0,
  last_error text        default '',
  created_at timestamptz default now(),
  sent_at    timestamptz
);

-- 觸發器每次撈「還沒寄成功的、最舊的優先」
create index if not exists mail_queue_pending_idx on public.mail_queue (status, created_at);


-- ---------------------------------------------
-- RLS：與現有 station_bookings / equipment 相同的開放政策
-- ---------------------------------------------
alter table public.dept_equipment enable row level security;
alter table public.keepers        enable row level security;
alter table public.mail_queue     enable row level security;

drop policy if exists "dept_equipment public all" on public.dept_equipment;
create policy "dept_equipment public all" on public.dept_equipment
  for all using (true) with check (true);

-- Keeper 只讓前端讀，寫入由 GAS 同步負責
drop policy if exists "keepers public read"  on public.keepers;
drop policy if exists "keepers public write" on public.keepers;
create policy "keepers public read"  on public.keepers for select using (true);
create policy "keepers public write" on public.keepers for all    using (true) with check (true);

-- 前端只需要「投遞」，寄送狀態由 GAS 更新
drop policy if exists "mail_queue public all" on public.mail_queue;
create policy "mail_queue public all" on public.mail_queue
  for all using (true) with check (true);
