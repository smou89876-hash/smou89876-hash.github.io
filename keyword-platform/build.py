#!/usr/bin/env python3
"""
读取 keyword-platform/ 下所有 xlsx 文件，合并数据，生成 JSON，嵌入 HTML。
用法：python3 build.py
"""

import json
import re
import glob
import os
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("请先安装 openpyxl：pip install openpyxl")
    raise

DIR = os.path.dirname(os.path.abspath(__file__))

# 字段列索引（0-based），对应表头：
# 日期 关键词Division 关键词 搜索人数 搜索次数 GMV 加购率% 支付率% 搜索无结果率%
COL_DATE   = 0
COL_DIV    = 1
COL_KW     = 2
COL_U      = 3  # 搜索人数
COL_C      = 4  # 搜索次数
COL_GMV    = 5  # GMV
COL_PAY    = 7  # 支付率%（可选列）
COL_NORES  = 8  # 搜索无结果率%（可选列，旧文件可能没有）


def load_xlsx(path):
    """返回 {(date_str, kw): {div, u, c, g, r}} 字典"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = next(rows)  # 跳过表头

    # 动态检测可选列
    nores_col = None
    pay_col = None
    for i, h in enumerate(header):
        if h and '无结果' in str(h):
            nores_col = i
        if h and '支付率' in str(h):
            pay_col = i

    data = {}
    for row in rows:
        if not row[COL_DATE] or not row[COL_KW]:
            continue
        date_str = str(row[COL_DATE])[:10]  # 取 YYYY-MM-DD
        kw  = str(row[COL_KW]).strip()
        div = str(row[COL_DIV]).strip() if row[COL_DIV] else ''
        u   = int(row[COL_U] or 0)
        c   = int(row[COL_C] or 0)
        g   = int(round(float(row[COL_GMV] or 0)))
        # r = 无结果搜索绝对件数（按搜索次数加权，便于聚合 sum(r)/sum(c)）
        if nores_col is not None and row[nores_col] is not None:
            r = int(round(float(row[nores_col]) * c))
        else:
            r = 0
        # p = 支付绝对人数（按搜索人数加权，便于聚合 sum(p)/sum(u)）
        if pay_col is not None and row[pay_col] is not None:
            p = int(round(float(row[pay_col]) * u))
        else:
            p = 0
        data[(date_str, kw)] = {'div': div, 'u': u, 'c': c, 'g': g, 'r': r, 'p': p}
    wb.close()
    return data


def build_json():
    xlsx_files = sorted(glob.glob(os.path.join(DIR, '*.xlsx')))
    if not xlsx_files:
        raise FileNotFoundError("未找到任何 xlsx 文件")

    print(f"读取 {len(xlsx_files)} 个文件：")
    merged = {}   # (date, kw) -> row
    kw_div = {}   # kw -> division

    for path in xlsx_files:
        print(f"  {os.path.basename(path)}")
        rows = load_xlsx(path)
        for (date, kw), row in rows.items():
            merged[(date, kw)] = row
            kw_div[kw] = row['div']

    all_dates = sorted({d for d, _ in merged})
    all_kws   = sorted(kw_div.keys())
    print(f"日期范围：{all_dates[0]} ~ {all_dates[-1]}（{len(all_dates)} 天）")
    print(f"关键词数：{len(all_kws)}")

    kw_list = []
    for kw in all_kws:
        u_arr, c_arr, g_arr, r_arr, p_arr = [], [], [], [], []
        for date in all_dates:
            row = merged.get((date, kw))
            if row:
                u_arr.append(row['u'])
                c_arr.append(row['c'])
                g_arr.append(row['g'])
                r_arr.append(row['r'])
                p_arr.append(row['p'])
            else:
                u_arr.append(0)
                c_arr.append(0)
                g_arr.append(0)
                r_arr.append(0)
                p_arr.append(0)
        kw_list.append({
            'n': kw,
            'd': kw_div[kw],
            'u': u_arr,
            'c': c_arr,
            'r': r_arr,
            'g': g_arr,
            'p': p_arr,
        })

    result = {
        'meta': {
            'start': all_dates[0],
            'end':   all_dates[-1],
            'days':  len(all_dates),
            'nKw':   len(all_kws),
        },
        'kw': kw_list,
    }
    return result


def embed_into_html(data):
    html_path = os.path.join(DIR, '重点词分析.html')
    json_path = os.path.join(DIR, '重点词_分析数据.json')

    json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    # 写独立 JSON 文件
    with open(json_path, 'w', encoding='utf-8') as f:
        f.write(json_str)
    print(f"已写入 {os.path.basename(json_path)}（{len(json_str)//1024} KB）")

    # 替换 HTML 内嵌数据
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    new_tag = f'<script id="data" type="application/json">{json_str}</script>'
    html_new = re.sub(
        r'<script id="data" type="application/json">.*?</script>',
        new_tag,
        html,
        flags=re.DOTALL,
    )

    if html_new == html:
        print("警告：HTML 中未找到 <script id=\"data\"> 标签，跳过嵌入")
    else:
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_new)
        print(f"已更新 {os.path.basename(html_path)}")


if __name__ == '__main__':
    data = build_json()
    embed_into_html(data)
    print("完成。")
