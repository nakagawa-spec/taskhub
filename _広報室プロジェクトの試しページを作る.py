"""実際の表に接続しない試しページを作る（タスクハブの会社用の画面＋広報室タブ）。引数で出力先フォルダを指定できる。
先に python _会社用の画面を作る.py で gas/index.html を作っておく。"""
import json
import sys
from pathlib import Path

here = Path(__file__).resolve().parent
out = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else here / '_試しページ'
out.mkdir(parents=True, exist_ok=True)
mock = (here / '_広報室の試し環境.js').read_text(encoding='utf-8')
server = '\n'.join((here / 'gas' / name).read_text(encoding='utf-8') for name in ['Code.gs', 'Kouhou.gs'])
# 説明書きを含む元の処理は文字列として渡す。連続した斜線は生成した画面にも入れない。
source = json.dumps(server + '\nreturn {kp:kp, th:th};', ensure_ascii=True).replace('/', r'\u002f').replace('<', r'\u003c')
bridge = """
const server = new Function(SOURCE)();
window.google={script:{get run(){
  const handlers={};
  const runner={withSuccessHandler(fn){handlers.ok=fn;return runner;},withFailureHandler(fn){handlers.fail=fn;return runner;},kp(action,req){setTimeout(()=>{try{handlers.ok(server.kp(action,req));}catch(e){if(handlers.fail)handlers.fail(e);}},30);},
    th(action,req){setTimeout(()=>{try{handlers.ok(server.th(action,req));}catch(e){if(handlers.fail)handlers.fail(e);}},30);}};
  return runner;
},url:{getLocation(fn){setTimeout(()=>fn({parameter:{view:'kouhou'}}),50);}}}};
""".replace('SOURCE', source)
html = (here / 'gas' / 'index.html').read_text(encoding='utf-8')
html = html.replace('<script>', '<script>\n' + mock + '\n' + bridge + '\n</script>\n<script>', 1)
html = html.replace('<body>', '<body><p style="margin:12px;text-align:center;color:#825e22">試しページです。本物の表には接続しません。変更は開き直すと元に戻ります。</p>', 1)
(out / 'taskhub.html').write_text(html, encoding='utf-8')
print(str(out / 'taskhub.html'))
