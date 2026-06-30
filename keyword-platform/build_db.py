#!/usr/bin/env python3
"""
基于数据（数据库优先，否则本地 xlsx）生成双端（小程序/APP）看板数据，嵌入 HTML。

- 有 DATABASE_URL（环境变量或 .env）：从 keyword_daily 表读取。
- 没有连接串：回退读取本目录 *.xlsx（平台按文件名自动判定），数据等同入库内容。

输出 JSON 结构（合计=小程序+APP，在前端按需相加，不落库以省体积）：
  { meta:{start,end,days,nKw,platforms:[...] },
    kw:[ {n,d, mp:{u,c,r,g,p}?, app:{u,c,r,g,p}? }, ... ] }

用法：python3 build_db.py
依赖：openpyxl（回退读 xlsx 时）/ psycopg2-binary（读库时）
"""
import os, re, json, glob
from collections import defaultdict

DIR = os.path.dirname(os.path.abspath(__file__))
PLAT_KEY = {'小程序': 'mp', 'APP': 'app'}   # 平台 -> JSON 短键


def _row_iter_db(conn):
    cur = conn.cursor(name='kw_stream')
    cur.itersize = 20000
    cur.execute("""
        SELECT date, keyword, platform, division,
               search_users, search_count, gmv, pay_rate, no_result_rate
        FROM keyword_daily
    """)
    for date, kw, plat, div, su, sc, gmv, payrate, nores in cur:
        yield (str(date)[:10], kw, plat, div,
               int(su or 0), int(sc or 0), int(round(float(gmv or 0))),
               float(payrate) if payrate is not None else None,
               float(nores) if nores is not None else None)
    cur.close()


def _row_iter_xlsx():
    import db_ingest as di
    for path in sorted(glob.glob(os.path.join(DIR, '*.xlsx'))):
        plat = di.platform_from_name(path)
        rows, _ = di.parse_xlsx(path, plat)
        # parse_xlsx 元组: date,kw,plat,div,su,sc,gmv,addcart,payrate,nores,fname
        for r in rows:
            yield (r[0], r[1], r[2], r[3],
                   int(r[4] or 0), int(r[5] or 0), int(round(float(r[6] or 0))),
                   r[8], r[9])


def get_rows(force_xlsx=False):
    """返回 (来源说明, 行迭代器物化为 list)。force_xlsx=True 时跳过数据库。"""
    # 尝试数据库
    url = None
    envf = os.path.join(DIR, '.env')
    if os.path.exists(envf):
        for line in open(envf, encoding='utf-8'):
            line = line.strip()
            if line.startswith('DATABASE_URL') and '=' in line:
                url = line.split('=', 1)[1].strip().strip('"').strip("'")
    url = os.environ.get('DATABASE_URL', url)
    if url and not force_xlsx:
        import psycopg2
        conn = psycopg2.connect(url)
        rows = list(_row_iter_db(conn))
        conn.close()
        return f"数据库（{len(rows)} 行）", rows
    rows = list(_row_iter_xlsx())
    return f"本地 xlsx（无 DATABASE_URL，{len(rows)} 行）", rows


def build_payload(rows):
    # 统一日期轴 = 两端所有日期的并集
    all_dates = sorted({r[0] for r in rows})
    didx = {d: i for i, d in enumerate(all_dates)}
    n = len(all_dates)

    # kw -> {div, platform -> {u,c,r,g,p arrays}}
    kwdata = {}
    kwdiv = {}
    for date, kw, plat, div, su, sc, gmv, payrate, nores in rows:
        pk = PLAT_KEY.get(plat)
        if pk is None:
            continue
        if div and kw not in kwdiv:
            kwdiv[kw] = div
        kwdiv.setdefault(kw, div or '')
        d = kwdata.setdefault(kw, {})
        arr = d.get(pk)
        if arr is None:
            arr = {k: [0] * n for k in 'ucrgp'}
            d[pk] = arr
        i = didx[date]
        arr['u'][i] = su
        arr['c'][i] = sc
        arr['g'][i] = gmv
        # r=无结果绝对件数(按次数加权)，p=支付绝对人数(按人数加权)，与看板 nrr/pr 聚合口径一致
        arr['r'][i] = int(round(nores * sc)) if (nores is not None) else 0
        arr['p'][i] = int(round(payrate * su)) if (payrate is not None) else 0

    kw_list = []
    for kw in sorted(kwdata.keys()):
        entry = {'n': kw, 'd': kwdiv.get(kw, '')}
        for pk in ('mp', 'app'):
            if pk in kwdata[kw]:
                entry[pk] = kwdata[kw][pk]
        kw_list.append(entry)

    return {
        'meta': {
            'start': all_dates[0], 'end': all_dates[-1], 'days': n,
            'nKw': len(kw_list), 'platforms': ['小程序', 'APP', '合计'],
        },
        'kw': kw_list,
    }


DATA_JS = '重点词_分析数据.js'   # 外部数据文件（<script src> 引入，本地 file:// 也能用）


def write_outputs(payload, source):
    json_str = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))

    # 1) 外部数据文件：挂到 window.__KW_DATA__，HTML 用 <script src> 引入
    js_path = os.path.join(DIR, DATA_JS)
    with open(js_path, 'w', encoding='utf-8') as f:
        f.write('window.__KW_DATA__=' + json_str + ';\n')
    print(f"已写入 {DATA_JS}（{len(json_str)//1024} KB，外部加载）")

    # 2) 瘦身 HTML：把内嵌的 22MB 数据标签替换成外部引入（幂等）
    html_path = os.path.join(DIR, '重点词分析.html')
    with open(html_path, encoding='utf-8') as f:
        html = f.read()
    include = f'<script src="{DATA_JS}"></script>'
    html2 = re.sub(r'<script id="data" type="application/json">.*?</script>',
                   include, html, flags=re.DOTALL)
    if html2 == html:
        if include in html:
            print('HTML 已是外部加载，无需改动')
        else:
            print('警告：未找到数据标签，未改动 HTML')
    else:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html2)
        print('已瘦身 重点词分析.html（数据改为外部加载）')


if __name__ == '__main__':
    import sys
    force_xlsx = '--xlsx' in sys.argv[1:]
    source, rows = get_rows(force_xlsx=force_xlsx)
    print(f"数据来源：{source}")
    payload = build_payload(rows)
    m = payload['meta']
    plat_kw = {'小程序': 0, 'APP': 0}
    for k in payload['kw']:
        if 'mp' in k: plat_kw['小程序'] += 1
        if 'app' in k: plat_kw['APP'] += 1
    print(f"日期：{m['start']} ~ {m['end']}（{m['days']} 天）  关键词并集：{m['nKw']}")
    print(f"  小程序词数 {plat_kw['小程序']} · APP词数 {plat_kw['APP']}")
    write_outputs(payload, source)
    print('完成。')
