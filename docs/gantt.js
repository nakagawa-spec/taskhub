/**
 * タスクハブ ガントチャート(タスクハブ本体と共有ページの両方で使う)
 *
 * TaskhubGantt.html({tasks, today, checkable, clickable, narrow, width}) で HTML を返す。
 * width(表示できる横幅px)を渡すと、その幅いっぱいまで日付を並べる。
 * addRow を渡すと、一番下に空の入力行を付ける(日付の欄をドラッグして期間を選び、名前を入れて追加)。
 *   addRow = {s, e, title}(選択中の期間と入力中の名前。未選択なら空)
 * - 行には data-gtask="<タスクID>"、完了チェックには data-gcheck="<タスクID>" が付くので、
 *   クリック等は呼び出し側で拾う。
 * - タスクの期間は startDate〜dueDate。開始日が無ければ期限日の1日だけのバーにする。
 */
(function(){
  "use strict";
  var WD = ['日','月','火','水','木','金','土'];

  function pad(n){ return String(n).padStart(2,'0'); }
  function parse(s){ var p = String(s).split('-'); return new Date(+p[0], +p[1]-1, +p[2]); }
  function fmt(d){ return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate()); }
  function md(s){ var d = parse(s); return (d.getMonth()+1)+'/'+d.getDate(); }
  function mdw(s){ var d = parse(s); return (d.getMonth()+1)+'/'+d.getDate()+'('+WD[d.getDay()]+')'; }
  function addDays(d, n){ var x = new Date(d); x.setDate(x.getDate()+n); return x; }
  function diffDays(a, b){ return Math.round((b - a) / 86400000); }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }

  /** タスクの期間 {s, e}。日付が無ければ null */
  function span(t){
    var e = t.dueDate || '';
    var s = t.startDate || e;
    if(!s) return null;
    if(!e) e = s;
    if(s > e){ var x = s; s = e; e = x; }
    return {s: s, e: e};
  }

  function summary(list, today){
    var r = {total: list.length, done: 0, overdue: 0, start: '', end: '', pct: 0};
    list.forEach(function(t){
      if(t.completed) r.done++;
      else if(t.dueDate && t.dueDate < today) r.overdue++;
      var p = span(t);
      if(!p) return;
      if(!r.start || p.s < r.start) r.start = p.s;
      if(!r.end || p.e > r.end) r.end = p.e;
    });
    r.pct = r.total ? Math.round(r.done * 100 / r.total) : 0;
    return r;
  }

  function sortRows(list){
    return list.slice().sort(function(a, b){
      var pa = span(a), pb = span(b);
      if(!pa !== !pb) return pa ? -1 : 1;
      if(pa && pb){
        var c = pa.s.localeCompare(pb.s); if(c) return c;
        c = pa.e.localeCompare(pb.e); if(c) return c;
      }
      return (a.order||0) - (b.order||0);
    });
  }

  function html(opts){
    var list = sortRows(opts.tasks || []);
    var today = opts.today;
    if(!list.length && !opts.addRow){
      return '<div class="g-empty">'+esc(opts.emptyText || 'タスクがまだありません。')+'</div>';
    }
    var sm = summary(list, today);
    var draft = opts.addRow || null;
    if(draft && draft.s){
      if(!sm.start || draft.s < sm.start) sm.start = draft.s;
      if(!sm.end || draft.e > sm.end) sm.end = draft.e;
    }
    var DW = opts.narrow ? 24 : 30;
    var t0 = parse(today);
    var first = sm.start ? parse(sm.start) : t0;
    var last = sm.end ? parse(sm.end) : addDays(t0, 13);
    first = addDays(first, -2);
    last = addDays(last, 3);
    // 今日が近くにあれば表示範囲に入れる(遠すぎるときは入れない)
    if(t0 < first && diffDays(t0, first) <= 21) first = t0;
    if(t0 > last && diffDays(last, t0) <= 21) last = addDays(t0, 2);
    var n = diffDays(first, last) + 1;
    if(n < 21){ n = 21; }
    // 表示幅が分かれば、右に空白が残らないよう日付を延ばす
    if(opts.width){
      var fit = Math.floor((opts.width - (opts.narrow ? 128 : 220)) / DW);
      if(n < fit) n = fit;
    }
    if(n > 400){ n = 400; }

    var months = '', days = '', cols = '';
    var segStart = 0;
    for(var i = 0; i <= n; i++){
      var d = addDays(first, i);
      var monthEnd = i === n || (i > 0 && d.getDate() === 1);
      if(monthEnd){
        var md0 = addDays(first, segStart);
        var w = (i - segStart) * DW;
        months += '<div class="g-month" style="left:'+(segStart*DW)+'px;width:'+w+'px"><span>'+(w >= 60 ? md0.getFullYear()+'年' : '')+(md0.getMonth()+1)+'月</span></div>';
        segStart = i;
      }
      if(i === n) break;
      var ds = fmt(d);
      var dow = d.getDay();
      var cls = (dow === 0 ? ' sun' : dow === 6 ? ' sat' : '') + (ds === today ? ' today' : '');
      days += '<div class="g-day'+cls+'" style="width:'+DW+'px"><b>'+d.getDate()+'</b><small>'+WD[dow]+'</small></div>';
      if(dow === 0 || dow === 6 || ds === today){
        cols += '<div class="g-col'+(ds === today ? ' today' : ' we')+'" style="left:'+(i*DW)+'px;width:'+DW+'px"></div>';
      }
    }
    var ti = diffDays(first, t0);
    var nowLine = (ti >= 0 && ti < n) ? '<div class="g-now" style="left:'+(ti*DW + DW/2)+'px"></div>' : '';

    var rows = list.map(function(t){
      var p = span(t);
      var cls = 'g-bar';
      if(t.completed) cls += ' done';
      else if(t.dueDate && t.dueDate < today) cls += ' overdue';
      else if(t.waiting) cls += ' waiting';
      var bar;
      if(p){
        var a = Math.max(0, diffDays(first, parse(p.s)));
        var b = Math.min(n - 1, diffDays(first, parse(p.e)));
        var bw = (b - a + 1) * DW - 4;
        var label = p.s === p.e ? md(p.e) : md(p.s)+'–'+md(p.e);
        var tip = (p.s === p.e ? mdw(p.e) : mdw(p.s)+' 〜 '+mdw(p.e)) + (t.completed ? '（完了）' : '');
        bar = b < a ? '' :
          '<div class="'+cls+'" data-gtask="'+esc(t.id)+'" style="left:'+(a*DW+2)+'px;width:'+bw+'px" title="'+esc(t.title+'\n'+tip)+'">' +
            (bw >= 64 ? '<span>'+label+'</span>' : '') +
          '</div>' +
          (bw < 64 ? '<div class="g-side" style="left:'+((b+1)*DW+4)+'px">'+label+'</div>' : '');
      } else {
        bar = '<div class="g-nodate">日程未定</div>';
      }
      var check = opts.checkable
        ? '<input type="checkbox" class="g-check" data-gcheck="'+esc(t.id)+'"'+(t.completed ? ' checked' : '')+' aria-label="完了">'
        : '<span class="g-dot'+(t.completed ? ' done' : '')+'"></span>';
      var title = opts.clickable
        ? '<button type="button" class="g-title" data-gtask="'+esc(t.id)+'">'+esc(t.title)+'</button>'
        : '<span class="g-title">'+esc(t.title)+'</span>';
      return '<div class="g-row'+(t.completed ? ' is-done' : '')+'">' +
        '<div class="g-lab">'+check+'<div class="g-lab-body">'+title +
          (t.owner ? '<span class="g-owner">@'+esc(t.owner)+'</span>' : '') +
        '</div></div>' +
        '<div class="g-track" style="width:'+(n*DW)+'px">'+bar+'</div>' +
      '</div>';
    }).join('');

    if(draft){
      var firstStr = fmt(first);
      rows += '<div class="g-row g-addrow">' +
        '<div class="g-lab"><span class="g-plus">＋</span>' +
          '<input type="text" class="g-addinput" data-gadd placeholder="新しいタスク名" value="'+esc(draft.title || '')+'" maxlength="100" aria-label="新しいタスク名">' +
          (draft.s ? '<button type="button" class="g-addclear" data-gclear title="選んだ期間を取り消す">×</button>' : '') +
        '</div>' +
        '<div class="g-track g-addtrack" data-gfirst="'+firstStr+'" data-gdw="'+DW+'" data-gn="'+n+'" style="width:'+(n*DW)+'px">' +
          (draft.s ? draftBar(firstStr, DW, n, draft.s, draft.e) : '<div class="g-addhint">'+(opts.narrow ? '日付を2回タップして期間を選ぶ' : '日付の欄をドラッグして期間を選ぶ')+'</div>') +
        '</div>' +
      '</div>';
    }

    return '<div class="gantt'+(opts.narrow ? ' narrow' : '')+'">' +
      '<div class="g-scroll">' +
        '<div class="g-inner">' +
          '<div class="g-head">' +
            '<div class="g-lab g-corner">タスク</div>' +
            '<div class="g-track" style="width:'+(n*DW)+'px">' +
              '<div class="g-months">'+months+'</div>' +
              '<div class="g-days">'+days+'</div>' +
            '</div>' +
          '</div>' +
          '<div class="g-body">' +
            '<div class="g-bg" style="width:'+(n*DW)+'px">'+cols+nowLine+'</div>' +
            rows +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  /** 入力行で選んでいる期間のバー */
  function draftBar(firstStr, DW, n, s, e){
    var first = parse(firstStr);
    var a = Math.max(0, diffDays(first, parse(s)));
    var b = Math.min(n - 1, diffDays(first, parse(e)));
    var days = b - a + 1;
    var label = s === e ? mdw(s) : md(s)+'–'+md(e)+'（'+days+'日）';
    return '<div class="g-bar draft" style="left:'+(a*DW+2)+'px;width:'+(days*DW-4)+'px"><span>'+label+'</span></div>';
  }

  /** 入力行の track 上の x 座標 → 日付文字列 */
  function dateAt(track, clientX){
    var r = track.getBoundingClientRect();
    var dw = +track.dataset.gdw, n = +track.dataset.gn;
    var i = Math.max(0, Math.min(n - 1, Math.floor((clientX - r.left) / dw)));
    return fmt(addDays(parse(track.dataset.gfirst), i));
  }

  /** 画面全体を描き直さずに、選択中のバーだけ差し替える */
  function setDraft(track, s, e){
    var old = track.querySelector('.g-bar.draft, .g-addhint');
    if(old) old.remove();
    track.insertAdjacentHTML('beforeend', draftBar(track.dataset.gfirst, +track.dataset.gdw, +track.dataset.gn, s, e));
  }

  /** 今日の線が見える位置まで横にスクロールする */
  function scrollToToday(root){
    var sc = root && root.querySelector('.g-scroll');
    var now = root && root.querySelector('.g-now');
    var lab = root && root.querySelector('.g-corner');
    if(!sc || !now || !lab) return;
    var visible = sc.clientWidth - lab.offsetWidth;
    var x = now.offsetLeft;
    if(x > visible * 0.7) sc.scrollLeft = x - visible * 0.3;
  }

  window.TaskhubGantt = {html: html, summary: summary, span: span, scrollToToday: scrollToToday, md: md, mdw: mdw, dateAt: dateAt, setDraft: setDraft};
})();
