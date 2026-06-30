#!/usr/bin/env python3
"""
一键更新：把本目录新增的 xlsx 导入数据库，并用数据库数据重建看板。

日常用法：
    1) 把新一周的 xlsx 放进 keyword-platform/ 目录
       （文件名含 APP → 自动按 APP 入库，否则按小程序）
    2) 运行：  python3 update.py
    3) 浏览器刷新 重点词分析.html 即可看到新数据

说明：
    - 已入库的文件会自动跳过，只导入新文件（按文件名判断）。
    - 导入后会从数据库重读全部数据重建看板（约 1-3 分钟，属正常）。
    - 想强制重导某个文件：python3 db_ingest.py --force 文件.xlsx
"""
import os, sys, glob, subprocess
import db_ingest as di

DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    # 1) 找出尚未入库的 xlsx
    conn = di.get_conn(); cur = conn.cursor()
    cur.execute(di.SCHEMA); conn.commit()
    cur.execute("SELECT DISTINCT source_file FROM ingest_log")
    done = {r[0] for r in cur.fetchall()}
    conn.close()

    xlsx = sorted(glob.glob(os.path.join(DIR, '*.xlsx')))
    new = [p for p in xlsx if os.path.basename(p) not in done]

    if not new:
        print("没有检测到新的 xlsx 文件（目录里的都已入库）。")
        print("如需强制重建看板：python3 build_db.py")
        return

    print(f"发现 {len(new)} 个新文件，准备导入：")
    for p in new:
        print(f"  + {os.path.basename(p)}  →  {di.platform_from_name(p)}")

    # 2) 导入数据库（db_ingest 会自动跳过已入库文件）
    print("\n→ 第 1 步：导入数据库 …")
    subprocess.run([sys.executable, '-u', 'db_ingest.py'], cwd=DIR, check=True)

    # 3) 用数据库数据重建看板（外部数据文件 + 瘦身 HTML）
    print("\n→ 第 2 步：用数据库数据重建看板（约 1-3 分钟）…")
    subprocess.run([sys.executable, '-u', 'build_db.py'], cwd=DIR, check=True)

    print("\n✅ 完成：新数据已入库，看板已更新。浏览器刷新 重点词分析.html 即可。")
    print("   （xlsx 已在数据库，确认无误后可自行删除）")


if __name__ == '__main__':
    main()
