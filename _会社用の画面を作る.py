r"""docs\app.html（会社用の画面の元）から、会社用の画面 gas\index.html を作る。
使い方：python _会社用の画面を作る.py
- config.js は読まない（会社用の画面では google.script.run で通信する）
- gantt.js・gantt.css は中に埋め込む（GAS は1つのファイルで配るため）
- favicon は外す
- GAS の HtmlService は <script> の中の // から行末を消すので、文字の中に // があれば止める
作ったら gas\ で clasp push → clasp deploy -i <会社用の公開ID>（作業記録を参照）
"""
import re
import sys
from pathlib import Path

here = Path(__file__).resolve().parent
src = (here / "docs" / "app.html").read_text(encoding="utf-8")
gjs = (here / "docs" / "gantt.js").read_text(encoding="utf-8")
gcss = (here / "docs" / "gantt.css").read_text(encoding="utf-8")


def rep(s, a, b):
    assert s.count(a) == 1, a
    return s.replace(a, b)


s = src
s = rep(s, '<link rel="icon" href="favicon.svg" type="image/svg+xml">\n', '')
s = rep(s, '<script src="config.js"></script>\n', '')
s = rep(s, '<script src="gantt.js"></script>\n', '<script>\n' + gjs + '\n</script>\n')
s = rep(s, '<link rel="stylesheet" href="gantt.css">\n', '<style>\n' + gcss + '\n</style>\n')

# <script> の中で、行の頭の説明書き以外に // が出てくる行を探す（行末の説明書きは消えても困らない）
bad = []
for m in re.finditer(r"<script>(.*?)</script>", s, re.S):
    for line in m.group(1).split("\n"):
        i = line.find("//")
        if i < 0:
            continue
        before = line[:i]
        # 行末の説明書き：// の前に引用符が閉じていない文字が無ければよい
        if before.count("'") % 2 == 0 and before.count('"') % 2 == 0 and before.count("`") % 2 == 0:
            continue
        bad.append(line.strip())
if bad:
    print("文字の中に // があります（GAS で画面が止まります）:")
    for b in bad:
        print("  " + b)
    sys.exit(1)

(here / "gas" / "index.html").write_text(s, encoding="utf-8")
print("gas\\index.html を作りました")
