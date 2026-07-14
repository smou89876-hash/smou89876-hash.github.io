#!/usr/bin/env python3
"""每周增量更新：把本目录 xlsx 直接 UPSERT 进现有 重点词_分析数据.js（秒级，零数据库读取）。

每周流程：
    1) 把新一周两份 xlsx 放进 keyword-platform/（命名含 APP → APP，否则小程序）
    2) python3 merge_weekly.py          # 体检 + 合并 + 尽力入库（库睡着不阻塞）
    3) git add/commit/push               # 你的 push 会自动触发 Pages 部署

说明：
    - 合并是 UPSERT 语义：同 (日期, 关键词, 平台) 覆盖，重复运行幂等。
    - Division 原样透传，取最新一周的分类；出现此前没见过的取值只提示不拦截。
      （归一类需求属偶发，按用户当次说明单独处理，不在本脚本内置规则。）
    - 数据体检不过（率值异常/列错位嫌疑）会停下，不写出任何文件。
    - Supabase 入库为尽力而为：失败只警告不阻塞（之后可 python3 db_ingest.py 补录，
      或用云端 rebuild-keyword-data workflow 从库全量重建对账）。
"""
import os, sys, glob, json, subprocess, datetime
import db_ingest as di

DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JS = os.path.join(DIR, '重点词_分析数据.js')
PLAT_KEY = {'小程序': 'mp', 'APP': 'app'}


def load_payload():
    raw = open(DATA_JS, encoding='utf-8').read()
    return json.loads(raw[raw.index('=') + 1:].rstrip().rstrip(';'))


def validate(rows, fname):
    """体检两档：结构性异常（列错位：率值>5 或超标行占比>1%）→ 停止；
    孤立小超标（小词跨期归因，率值 1~5）→ 仅警告。"""
    fatal, warn = [], []
    for r in rows:
        # parse_xlsx 元组: date,kw,plat,div,su,sc,gmv,addcart,payrate,nores,fname
        for name, v, lim in (('加购率', r[7], 1), ('支付率', r[8], 1), ('无结果率', r[9], 1)):
            if v is None or 0 <= v <= lim:
                continue
            (fatal if (v < 0 or v > 5) else warn).append(f"{r[0]} {r[1]} {name}={v}")
    if fatal or len(warn) > max(5, len(rows) * 0.01):
        print(f"✗ {fname} 数据体检不通过（疑似列错位/脏数据）：严重 {len(fatal)} 行，超标 {len(warn)} 行，示例：", file=sys.stderr)
        for b in (fatal + warn)[:5]:
            print("   " + b, file=sys.stderr)
        sys.exit("停止：请核对数据源后重导（未改动任何文件）。")
    for b in warn:
        print(f"  ⚠ {fname} 孤立超标（放行）：{b}")


def main():
    payload = load_payload()
    meta = payload['meta']
    old_end, old_n = meta['end'], meta['days']
    print(f"现有数据：{meta['start']} ~ {old_end}（{old_n} 天，{meta['nKw']} 词）")

    files = sorted(glob.glob(os.path.join(DIR, '*.xlsx')))
    if not files:
        sys.exit("目录里没有 xlsx 文件")

    # 1) 解析 + 体检（division 原样透传；新取值只提示，归一类需求按用户说明单独处理）
    known_divs = {k.get('d') for k in payload['kw'] if k.get('d')}
    all_rows = []
    for path in files:
        plat = di.platform_from_name(path)
        rows, st = di.parse_xlsx(path, plat)
        validate(rows, st['file'])
        all_rows.extend((r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[8], r[9]) for r in rows)
        new_divs = {r[3] for r in rows if r[3] and r[0] > old_end} - known_divs
        if new_divs:
            print(f"  ℹ {st['file']} 出现新 Division 取值：{sorted(new_divs)}（原样保留，如需归一请说明）")
        print(f"  ✓ {st['file']}: {st['n']} 行 [{st['platform']}] {st['dmin']}~{st['dmax']}")

    # 2) 扩展日期轴（.js 不存日期轴明细，按 start + days 连续推算）
    d0 = datetime.date.fromisoformat(meta['start'])
    old_dates = [(d0 + datetime.timedelta(days=i)).isoformat() for i in range(old_n)]
    new_dates = sorted({r[0] for r in all_rows})
    axis = sorted(set(old_dates) | set(new_dates))
    # 日期轴必须连续（看板按天索引），补齐中间缺口
    dA = datetime.date.fromisoformat(axis[0]); dB = datetime.date.fromisoformat(axis[-1])
    axis = [(dA + datetime.timedelta(days=i)).isoformat() for i in range((dB - dA).days + 1)]
    didx = {d: i for i, d in enumerate(axis)}
    n = len(axis)
    pad = n - old_n
    if axis[:old_n] != old_dates:
        sys.exit("停止：新日期早于现有起始日，需用云端全量重建而非增量合并。")

    # 3) 现有词全部补零到新长度
    kw_index = {}
    for k in payload['kw']:
        kw_index[k['n']] = k
        if pad > 0:
            for pk in ('mp', 'app'):
                if pk in k:
                    for key in 'ucrgp':
                        k[pk][key].extend([0] * pad)

    # 4) UPSERT 新数据
    created = 0
    for date, kw, plat, div, su, sc, gmv, payrate, nores in all_rows:
        pk = PLAT_KEY.get(plat)
        if pk is None:
            continue
        k = kw_index.get(kw)
        if k is None:
            k = {'n': kw, 'd': div or ''}
            kw_index[kw] = k
            payload['kw'].append(k)
            created += 1
        elif div and date > old_end:
            # division 只由比现有数据更新的日期覆盖：目录里滞留的历史周文件
            # 只能补数值，不能把旧分类（如已归一前的取值）倒灌回来
            k['d'] = div
        arr = k.get(pk)
        if arr is None:
            arr = {key: [0] * n for key in 'ucrgp'}
            k[pk] = arr
        i = didx[date]
        arr['u'][i] = int(su or 0)
        arr['c'][i] = int(sc or 0)
        arr['g'][i] = int(round(float(gmv or 0)))
        arr['r'][i] = int(round(nores * (sc or 0))) if nores is not None else 0
        arr['p'][i] = int(round(payrate * (su or 0))) if payrate is not None else 0

    payload['kw'].sort(key=lambda k: k['n'])
    meta.update(start=axis[0], end=axis[-1], days=n, nKw=len(payload['kw']))

    # 5) 写出
    js = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    with open(DATA_JS, 'w', encoding='utf-8') as f:
        f.write('window.__KW_DATA__=' + js + ';\n')
    print(f"✓ 已写入 重点词_分析数据.js：{meta['start']} ~ {meta['end']}（{n} 天，{meta['nKw']} 词，新词 {created}）")

    # 6) 尽力入库（失败不阻塞发布）
    print("→ 尝试入库 Supabase（失败不影响发布）…")
    try:
        subprocess.run([sys.executable, '-u', 'db_ingest.py'], cwd=DIR, check=True, timeout=300)
    except Exception as e:
        print(f"⚠ 入库未完成（{e}）。看板不受影响；稍后可 python3 db_ingest.py 补录。")

    print("\n下一步：git add 重点词_分析数据.js && git commit && git push（push 自动部署）")


if __name__ == '__main__':
    main()
