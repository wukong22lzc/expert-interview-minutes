#!/usr/bin/env python3
"""按分节标题把一个多场访谈合并的 docx 切成单场 txt，供逐场核对 / 改写。

用法:
    python3 split_sections.py 汇总.docx outdir --mode 逐字稿
    python3 split_sections.py 汇总.docx outdir --mode 纪要
    python3 split_sections.py 汇总.docx outdir --heading '^(.+?)（逐字稿）$'

输出 outdir/s01_<标题>.txt …，每行 "行号|[样式名] 文本"；行号与 read_file 显示一致，
子代理引用证据可直接写 L123。同时打印每场起止行，便于按行区间切对侧文件。
"""
import argparse
import os
import re

import docx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx_path")
    ap.add_argument("outdir")
    ap.add_argument("--mode", default="逐字稿", help="分节标记，如 逐字稿 / 纪要")
    ap.add_argument("--heading", default=None, help="自定义标题正则")
    ap.add_argument("--prefix", default=None, help="输出文件名前缀，默认按 mode 取 s/j")
    a = ap.parse_args()

    pat = re.compile(a.heading or r"^(.+?)（%s）$" % re.escape(a.mode))
    d = docx.Document(a.docx_path)
    os.makedirs(a.outdir, exist_ok=True)

    rows, secs = [], []
    for i, p in enumerate(d.paragraphs, 1):
        t = p.text.strip()
        rows.append("%d|[%s] %s" % (i, p.style.name, t))
        m = pat.match(t)
        if m:
            secs.append({"name": m.group(1), "start": i})

    if not secs:
        raise SystemExit("没匹配到分节标题，检查 --mode / --heading")

    prefix = a.prefix or ("s" if a.mode == "逐字稿" else "j")
    for k, s in enumerate(secs):
        end = secs[k + 1]["start"] - 1 if k + 1 < len(secs) else len(d.paragraphs)
        slug = re.sub(r"[^0-9A-Za-z\u4e00-\u9fa5]", "", s["name"])[:12] or "sec"
        fn = os.path.join(a.outdir, "%s%02d_%s.txt" % (prefix, k + 1, slug))
        with open(fn, "w") as f:
            f.write("\n".join(rows[s["start"] - 1:end]))
        print("%s\tL%d-L%d\t%s" % (fn, s["start"], end, s["name"]))


if __name__ == "__main__":
    main()
