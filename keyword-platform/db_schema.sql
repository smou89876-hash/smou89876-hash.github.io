-- 重点词数据 · 云数据库 schema（Postgres，兼容 Supabase / Neon）
-- 用法：在 Supabase/Neon 的 SQL Editor 里整段执行一次即可；db_ingest.py 也会自动建表。

-- 每日关键词明细表：每个 (日期, 关键词, 平台) 一行，重复导入会 UPSERT 覆盖（幂等）
-- platform：'小程序' / 'APP'，区分双端数据；历史数据均为小程序
CREATE TABLE IF NOT EXISTS keyword_daily (
    date            DATE        NOT NULL,                 -- 日期
    keyword         TEXT        NOT NULL,                 -- 关键词
    platform        TEXT        NOT NULL DEFAULT '小程序',  -- 平台：小程序 / APP
    division        TEXT,                                 -- 关键词 Division
    search_users    INTEGER,                              -- 搜索人数
    search_count    INTEGER,                              -- 搜索次数
    gmv             NUMERIC(16,2),                        -- GMV
    add_cart_rate   NUMERIC(10,6),                        -- 加购率（小数，0.57 = 57%）
    pay_rate        NUMERIC(10,6),                        -- 支付率（小数）
    no_result_rate  NUMERIC(10,6),                        -- 搜索无结果率（小数）
    source_file     TEXT,                                 -- 来源 xlsx 文件名（溯源）
    ingested_at     TIMESTAMPTZ DEFAULT now(),            -- 入库时间
    PRIMARY KEY (date, keyword, platform)
);

CREATE INDEX IF NOT EXISTS idx_kw_daily_keyword ON keyword_daily (keyword);
CREATE INDEX IF NOT EXISTS idx_kw_daily_division ON keyword_daily (division);
CREATE INDEX IF NOT EXISTS idx_kw_daily_platform ON keyword_daily (platform);

-- 导入批次日志：每导入一个文件记一行，便于核对"哪一周/哪个平台的表已入库"
CREATE TABLE IF NOT EXISTS ingest_log (
    id           BIGSERIAL PRIMARY KEY,
    source_file  TEXT NOT NULL,
    platform     TEXT,
    row_count    INTEGER,
    date_min     DATE,
    date_max     DATE,
    keyword_cnt  INTEGER,
    ingested_at  TIMESTAMPTZ DEFAULT now()
);
