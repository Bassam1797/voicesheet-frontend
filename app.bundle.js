// === Dark mode persistence ===
(window)&&window.addEventListener('load', function(){
  try{
    var saved = localStorage.getItem('vs_dark');
    if(saved === '1'){ document.body.classList.add('dark'); }
    var btn = document.getElementById('btn-dark');
    if(btn){
      btn.textContent = document.body.classList.contains('dark') ? 'Light' : 'Dark';
      (btn)&&btn.addEventListener('click', function(){
        var on = document.body.classList.toggle('dark');
        localStorage.setItem('vs_dark', on ? '1' : '0');
        btn.textContent = on ? 'Light' : 'Dark';
      });
    }
  }catch(e){ console.error(e); }
});

// === Sheet rename + persistence (event delegation; no render function surgery) ===
(window)&&window.addEventListener('load', function(){
  try{
    // Restore saved names (if any) before first renderTabs call happens later
    var savedNames = JSON.parse(localStorage.getItem('vs_sheet_names') || 'null');
    if (Array.isArray(savedNames) && typeof sheets !== 'undefined') {
      for (var i=0; i<Math.min(savedNames.length, sheets.length); i++) {
        if (savedNames[i]) sheets[i].name = savedNames[i];
      }
    }
  }catch(e){ console.warn('Sheet names restore failed', e); }

  var tabs = document.getElementById('sheets-tabs');
  if(!tabs) return;
  (tabs)&&tabs.addEventListener('dblclick', function(e){
    var btn = e.target.closest('button');
    if(!btn) return;
    // Find index by text match; fallback to position
    var all = Array.prototype.slice.call(tabs.querySelectorAll('button'));
    var idx = all.indexOf(btn);
    if(idx < 0) return;
    var current = (typeof sheets !== 'undefined' && sheets[idx]) ? sheets[idx].name : btn.textContent;
    var nn = prompt('Rename sheet:', current);
    if(nn && nn.trim()){
      if(typeof sheets !== 'undefined' && sheets[idx]){ sheets[idx].name = nn.trim(); }
      // Update button text immediately
      btn.textContent = nn.trim();
      // Persist all names
      try{
        var names = (typeof sheets !== 'undefined') ? sheets.map(function(s){ return s.name; }) : all.map(function(b){return b.textContent;});
        localStorage.setItem('vs_sheet_names', JSON.stringify(names));
      }catch(e){ console.warn('Persist names failed', e); }
    }
  });
});

(function(){
  try {
    if (!document.body.dataset) document.body.dataset = {};
    if (!document.body.dataset.theme) document.body.dataset.theme = 'light';
    if (!document.body.dataset.compact) document.body.dataset.compact = 'false';

    var btnDark = document.getElementById('v22-toggle-dark');
    var btnCompact = document.getElementById('v22-toggle-compact');
    var micSelect = document.getElementById('v22-mic-mode');
    var micMode = micSelect ? (micSelect.value || 'numbers') : 'numbers';

    if (btnDark) (btnDark)&&btnDark.addEventListener('click', function(){
      document.body.dataset.theme = document.body.dataset.theme === 'light' ? 'dark' : 'light';
    });
    if (btnCompact) (btnCompact)&&btnCompact.addEventListener('click', function(){
      document.body.dataset.compact = document.body.dataset.compact === 'false' ? 'true' : 'false';
    });
    if (micSelect) (micSelect)&&micSelect.addEventListener('change', function(){
      micMode = this.value;
    });

    function getActiveCellElement(){
      var ae = document.activeElement;
      if (ae && (ae.isContentEditable || ae.getAttribute('contenteditable') === 'true')) return ae;
      var cell = document.querySelector('td.active,[data-active="true"], td[data-row][data-col].active');
      if (cell) return cell;
      var input = document.querySelector('.grid input:focus, table input:focus, textarea:focus');
      if (input) return input;
      return null;
    }

    function patchRecognizer(rec){
      if (!rec || rec.__v22Patched) return;
      rec.__v22Patched = true;
      var prevOnResult = rec.onresult;
      rec.onresult = function(event){
  window._dictationActive = true;
        try {
          var res = event && event.results && event.results[event.results.length - 1];
          var transcript = res && res[0] ? (res[0].transcript || '') : '';
          if (micMode === 'numbers') {
            var filtered = transcript.replace(/[^0-9+\-.,]/g, '');
            if (filtered.indexOf(',') !== -1 && filtered.indexOf('.') !== -1) {
              filtered = filtered.replace(/,/g, '');
            } else {
              filtered = filtered.replace(/,/g, '.');
            }
            var el = getActiveCellElement();
            if (el) {
              if ('value' in el) el.value = filtered;
              else el.textContent = filtered;
            }
          }
        } catch(e) { /* no-op */ }
        if (typeof prevOnResult === 'function') return prevOnResult.call(this, event);
      };
    }

    var recognizerPoll = setInterval(function(){
      var rec = window.recognition || window.SpeechRecognitionInstance || window.webkitSpeechRecognitionInstance;
      if (rec) {
        try { patchRecognizer(rec); } catch(err){ console.error(err); }
      }
    }, 500);

    var SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (SR && !SR.__v22CtorPatched) {
      SR.__v22CtorPatched = true;
      var Orig = SR;
      var Wrapped = function(){
        var instance = new Orig();
        try { patchRecognizer(instance); } catch(err){ console.error(err); }
        return instance;
      };
      Wrapped.prototype = Orig.prototype;
      window.SpeechRecognition = Wrapped;
      if ('webkitSpeechRecognition' in window) window.webkitSpeechRecognition = Wrapped;
    }
  } catch (err) {
    console.warn('V2.2 inject error:', err);
  }
})();

(function(){
  var Stack = [];
  var Redo = [];
  var session = null;
  var GRID_SEL = 'td[contenteditable], td[contenteditable="true"], td[contenteditable=""], input, textarea';

  function startSession(el){
    endSession(true);
    session = { el: el, start: getVal(el), last: getVal(el) };
  }
  function getVal(el){ return ('value' in el) ? el.value : el.textContent; }
  function setVal(el, v){ if ('value' in el) el.value = v; else el.textContent = v; }

  function endSession(push){
    if (!session) return;
    var current = getVal(session.el);
    if (push && current !== session.start) {
      Stack.push({ el: session.el, before: session.start, after: current });
      Redo.length = 0;
    }
    session = null;
  }

  (document)&&document.addEventListener('focusin', function(e){
    var el = e.target.closest && e.target.closest(GRID_SEL);
    if (el) startSession(el);
  });
  (document)&&document.addEventListener('input', function(e){
    var el = e.target.closest && e.target.closest(GRID_SEL);
    if (el && session && el === session.el) session.last = getVal(el);
  });
  (document)&&document.addEventListener('focusout', function(e){
    var el = e.target.closest && e.target.closest(GRID_SEL);
    if (el && session && el === session.el) endSession(true);
  });
  (document)&&document.addEventListener('keydown', function(e){
    if (!session) return;
    if (e.key === 'Enter' || e.key === 'Tab' || e.key.startsWith('Arrow')) {
      endSession(true);
    }
  }, true);

  (document)&&document.addEventListener('keydown', function(e){
    var k = e.key.toLowerCase();
    var z = (k === 'z') && (e.ctrlKey || e.metaKey);
    var y = ((k === 'y') && (e.ctrlKey || e.metaKey)) || (e.shiftKey && (k === 'z') && (e.ctrlKey || e.metaKey));
    if (!z && !y) return;
    if (session) endSession(true);
    if (z) {
      var op = Stack.pop();
      if (op) {
        setVal(op.el, op.before);
        try { op.el.focus(); } catch(err){ console.error(err); }
        Redo.push(op);
        e.preventDefault(); e.stopPropagation(); return;
      }
    } else if (y) {
      var op = Redo.pop();
      if (op) {
        setVal(op.el, op.after);
        try { op.el.focus(); } catch(err){ console.error(err); }
        Stack.push(op);
        e.preventDefault(); e.stopPropagation(); return;
      }
    }
  }, true);
})();

(function(){
  function $(sel){return document.querySelector(sel);}
  var slider = $('#grid-hscroll');
  if (!slider) return;
  var container = document.querySelector('.grid-container, .grid, .sheet-container, .table-container, main, body');
  var tbl = document.querySelector('table');
  if (tbl) {
    var p = tbl.parentElement;
    while (p && p !== document.body) {
      if (p.scrollWidth > p.clientWidth) { container = p; break; }
      p = p.parentElement;
    }
  }
  function syncMax(){
    if (!container) return;
    var max = Math.max(0, container.scrollWidth - container.clientWidth);
    slider.max = String(max);
    slider.value = String(container.scrollLeft);
  }
  function onScroll(){ slider.value = String(container.scrollLeft); }
  function onInput(){ container.scrollLeft = parseInt(slider.value||'0',10); }
  if (container) {
    (container)&&container.addEventListener('scroll', onScroll, {passive:true});
    (slider)&&slider.addEventListener('input', onInput);
    (window)&&window.addEventListener('resize', syncMax);
    setTimeout(syncMax, 0);
    setTimeout(syncMax, 250);
  }
})();

/* ===== r2e2 augmentation: non-destructive fixes ===== */
(function(){
  function $(sel){ return document.querySelector(sel); }
  function $all(sel){ return Array.from(document.querySelectorAll(sel)); }

  /* ---- Horizontal slider wiring ---- */
  function getGridScroller(){
    return document.querySelector('.grid-container') || document.getElementById('grid-container') || document.querySelector('main') || document.body;
  }
  function syncHScroll(){
    var sc = getGridScroller();
    var r = $('#hscroll');
    if(!sc || !r) return;
    var max = Math.max(0, sc.scrollWidth - sc.clientWidth);
    r.max = String(max);
    r.value = String(sc.scrollLeft);
  }
  function wireHScroll(){
    var sc = getGridScroller();
    var r = $('#hscroll');
    if(!sc || !r) return;
    (r)&&r.addEventListener('input', function(){ sc.scrollLeft = Number(r.value||0); });
    (sc)&&sc.addEventListener('scroll', function(){ r.value = String(sc.scrollLeft); }, { passive:true });
    (window)&&window.addEventListener('resize', syncHScroll);
    setTimeout(syncHScroll, 60);
  }

  /* ---- Whole-cell Undo/Redo (commit on blur/Enter/Tab) ---- */
  window.undoStack = window.undoStack || [];
  window.redoStack = window.redoStack || [];
  function cellInput(el){
    return el && el.tagName==='INPUT' && (el.closest('td, .cell, .grid-cell'));
  }
  function cellPos(inp){
    var td = inp.closest('td, .cell, .grid-cell'); if(!td) return null;
    var r = Number(td.getAttribute('data-row')||td.getAttribute('data-r')||td.dataset.row||0);
    var c = Number(td.getAttribute('data-col')||td.getAttribute('data-c')||td.dataset.col||0);
    return {r:r||1, c:c||1, td:td};
  }
  function setCellValue(r,c,val){
    var td = document.querySelector('[data-row="'+r+'"][data-col="'+c+'"], .cell[data-r="'+r+'"][data-c="'+c+'"], .grid-cell[data-row="'+r+'"][data-col="'+c+'"]');
    var inp = td ? td.querySelector('input') : null;
    if(inp){ inp.value = val; }
    try{
      if(window.sheets && window.sheets[window.currentSheet||0]){
        var s = window.sheets[window.currentSheet||0];
        if(s.data){ if(s.data[r-1] && s.data[r-1][c-1] !== undefined) s.data[r-1][c-1]=val; }
        if(s.cells){ if(s.cells[r-1] && s.cells[r-1][c-1]) s.cells[r-1][c-1].value=val; }
      }
    }catch(err){ console.error(err); }
  }
  var preValMap = new WeakMap();
  (document)&&document.addEventListener('focusin', function(e){
    if(cellInput(e.target)){ preValMap.set(e.target, e.target.value); }
  });
  function commitIfChanged(inp){
    if(!cellInput(inp)) return;
    var prev = preValMap.get(inp);
    var now = inp.value;
    var pos = cellPos(inp);
    if(prev !== undefined && prev !== now){
      window.undoStack.push({type:'set', r:pos.r, c:pos.c, prev:prev, next:now});
      window.redoStack.length = 0;
      setCellValue(pos.r,pos.c,now);
    }
    preValMap.delete(inp);
  }
  (document)&&document.addEventListener('focusout', function(e){ if(cellInput(e.target)) commitIfChanged(e.target); });
  (document)&&document.addEventListener('keydown', function(e){
    if((e.key==='Enter' || e.key==='Tab') && cellInput(document.activeElement)){
      commitIfChanged(document.activeElement);
    }
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); doUndo(); }
    if(((e.ctrlKey||e.metaKey) && (e.key.toLowerCase()==='y' || (e.shiftKey && e.key.toLowerCase()==='z')))){ e.preventDefault(); doRedo(); }
  });
  function doUndo(){
    var it = window.undoStack.pop(); if(!it) return;
    if(it.type==='set'){ setCellValue(it.r,it.c, it.prev||''); window.redoStack.push({type:'set', r:it.r,c:it.c, prev:(it.next||''), next:(it.prev||'')}); }
  }
  function doRedo(){
    var it = window.redoStack.pop(); if(!it) return;
    if(it.type==='set'){ setCellValue(it.r,it.c, it.next||''); window.undoStack.push({type:'set', r:it.r,c:it.c, prev:(it.prev||''), next:(it.next||'')}); }
  }
  window.doUndo = window.doUndo || doUndo;
  window.doRedo = window.doRedo || doRedo;

  /* ---- Freeze Top Row / First Col with active styling ---- */
  function addToggle(btn, onToggle){
    if(!btn) return;
    (btn)&&btn.addEventListener('click', function(){
      btn.classList.toggle('toggled');
      if(btn.classList.contains('toggled')){
        btn.style.filter = 'brightness(0.85)';
      }else{
        btn.style.filter = '';
      }
      try{ onToggle(btn.classList.contains('toggled')); }catch(err){ console.error(err); }
    });
  }
  function toggleStickyTopRow(enabled){
    var thead = document.querySelector('table thead');
    if(thead){
      thead.style.position = enabled ? 'sticky' : '';
      thead.style.top = enabled ? '0' : '';
      thead.style.zIndex = enabled ? '2' : '';
    }
  }
  function toggleStickyFirstCol(enabled){
    var firstColCells = document.querySelectorAll('table tr > *:first-child');
    firstColCells.forEach(function(el){
      el.style.position = enabled ? 'sticky' : '';
      el.style.left = enabled ? '0' : '';
      el.style.zIndex = enabled ? '1' : '';
      if(el.tagName==='TH'){
        el.style.background = enabled ? 'var(--toolbar-bg, #f8f9fa)' : '';
      }
    });
  }
  (function setupFreezeButtons(){
    var btnTop = document.getElementById('freeze-top-row') || document.getElementById('btn-freeze-top');
    var btnFirst = document.getElementById('freeze-first-col') || document.getElementById('btn-freeze-firstcol');
    addToggle(btnTop, toggleStickyTopRow);
    addToggle(btnFirst, toggleStickyFirstCol);
  })();

  /* ---- Dictation with Advance limit & wrap ---- */
  function setActive(r,c){
    if(typeof window.setActive === 'function'){ window.setActive(r,c); return; }
    var td = document.querySelector('[data-row="'+r+'"][data-col="'+c+'"], .cell[data-r="'+r+'"][data-c="'+c+'"], .grid-cell[data-row="'+r+'"][data-col="'+c+'"]');
    var inp = td && td.querySelector('input'); if(inp) inp.focus({preventScroll:true});
  }
  function getActive(){
    if(window.active && window.active.row && window.active.col){ return {r:window.active.row, c:window.active.col}; }
    if(window.active && window.active.r && window.active.c){ return {r:window.active.r, c:window.active.c}; }
    var inp = document.activeElement && document.activeElement.tagName==='INPUT' ? document.activeElement : null;
    if(!inp) return {r:1,c:1};
    var td = inp.closest('td, .cell, .grid-cell');
    return { r:Number(td?.getAttribute('data-row')||td?.dataset.row||td?.getAttribute('data-r')||1),
             c:Number(td?.getAttribute('data-col')||td?.dataset.col||td?.getAttribute('data-c')||1) };
  }
  function gridSize(){
    try{
      if(window.sheets){
        var s = window.sheets[window.currentSheet||0];
        if(s && s.data){ return {rows:s.data.length, cols:(s.data[0]||[]).length}; }
        if(s && s.cells){ return {rows:s.cells.length, cols:(s.cells[0]||[]).length}; }
      }
    }catch(err){ console.error(err); }
    var lastRow = Math.max.apply(null, $all('td[data-row], .cell[data-r], .grid-cell[data-row]').map(td=>Number(td.getAttribute('data-row')||td.dataset.row||td.getAttribute('data-r')||1)));
    var lastCol = Math.max.apply(null, $all('td[data-col], .cell[data-c], .grid-cell[data-col]').map(td=>Number(td.getAttribute('data-col')||td.dataset.col||td.getAttribute('data-c')||1)));
    return {rows:lastRow||100, cols:lastCol||26};
  }
  function move(dir){
    var a=getActive(), size=gridSize();
    var limitInput = document.getElementById('advanceLimit'); var N = Math.max(1, Number(limitInput && limitInput.value) || 3);
    var start = window._blockStart || {r:a.r, c:a.c};
    if(window._blockStart == null){ window._blockStart = {r:a.r, c:a.c}; start = window._blockStart; }
    var idx = (window._idxInBlock==null) ? 0 : window._idxInBlock;
    if(dir==='right'){
      if(idx < N-1){ idx++; a.c = Math.min(size.cols, a.c+1); }
      else { idx=0; a.r = Math.min(size.rows, a.r+1); a.c = start.c; }
    } else if(dir==='down'){
      if(idx < N-1){ idx++; a.r = Math.min(size.rows, a.r+1); }
      else { idx=0; a.c = Math.min(size.cols, a.c+1); a.r = start.r; }
    } else if(dir==='left'){
      if(idx < N-1){ idx++; a.c = Math.max(1, a.c-1); }
      else { idx=0; a.r = Math.max(1, a.r-1); a.c = start.c; }
    } else if(dir==='up'){
      if(idx < N-1){ idx++; a.r = Math.max(1, a.r-1); }
      else { idx=0; a.c = Math.max(1, a.c-1); a.r = start.r; }
    }
    window._idxInBlock = idx;
    setActive(a.r,a.c);
  }
  function normalizeNumbers(text){
    return String(text||'').replace(/[^0-9\.\-:\/\s,]/g,'').replace(/,/g,'.').trim();
  }
  function handleDictation(text){
    var modeSel = document.getElementById('mic-mode') || document.getElementById('micMode') || document.getElementById('v22-mic-mode');
    var mode = modeSel ? modeSel.value : 'numbers';
    var cleaned = (mode==='numbers') ? normalizeNumbers(text) : String(text||'').trim();

    var a=getActive();
    setCellValue(a.r, a.c, cleaned);
    window.undoStack.push({type:'set', r:a.r, c:a.c, prev:'', next:cleaned});
    window.redoStack.length = 0;

    var auto = document.getElementById('auto-advance') || document.getElementById('autoAdvance');
    if(auto && auto.checked){
      var dirSel = document.getElementById('voice-direction') || document.getElementById('direction');
      var dir = dirSel ? dirSel.value : 'right';
      move(dir);
    }
  }
  function setupMic(){
    var modeSel = document.getElementById('mic-mode') || document.getElementById('micMode') || document.getElementById('v22-mic-mode');
    if(modeSel){ modeSel.value = 'numbers'; }
    var Rec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if(!Rec) return;
    var rec = window._r2e2_recognition || (window._r2e2_recognition = new Rec());
    rec.continuous = true; rec.interimResults = false; rec.lang = 'en-US';
    rec.onresult = function(ev){
  window._dictationActive = true;
      for(var i=ev.resultIndex;i<ev.results.length;i++){
        var res = ev.results[i];
        if(res.isFinal){ handleDictation(res[0].transcript || ''); }
      }
    };
    var micBtn = document.getElementById('start-mic') || document.getElementById('micBtn');
    if(micBtn){
      (micBtn)&&micBtn.addEventListener('click', function(){
        if(window._r2e2_listening){ rec.stop(); window._r2e2_listening=false; setMicAriaPressed(micBtn, false); micBtn.textContent = '🎤 Start Mic'; }
        else { try{ rec.start(); window._r2e2_listening=true; setMicAriaPressed(micBtn, true); micBtn.textContent = '🛑 Stop Mic'; }catch(err){ console.error(err); } }
      });
    }
  }

  function init(){
    wireHScroll();
    setupMic();
    var sc = getGridScroller();
    if(sc){
      new MutationObserver(function(){ setTimeout(function(){ syncHScroll(); }, 50); }).observe(sc, {childList:true, subtree:true});
    }
  }
  if(document.readyState==='loading') (document)&&document.addEventListener('DOMContentLoaded', init);
  else init();
})();

// === r2e7: Non-destructive fixes appended on top of r2e2 ===
(function(){
  // ---------- Whole-cell Undo/Redo buffering ----------
  var _origPushUndo = typeof window.pushUndo === 'function' ? window.pushUndo : function(){};
  var _sessions = {}; window._editSessions = _sessions;
  window.pushUndo = function(entry){
    try{
      if(entry && entry.type === 'set'){
        var key = entry.r + '|' + entry.c;
        if(_sessions[key] && _sessions[key].active){
          // swallow per-keystroke entries during an active edit session
          return;
        }
      }
    }catch(err){ console.error(err); }
    return _origPushUndo(entry);
  };
  function beginSession(r,c){
    var key = r + '|' + c;
    var startVal = '';
    try{
      startVal = (window.cells && window.cells[r-1] && window.cells[r-1][c-1] ? window.cells[r-1][c-1].value : '');
    }catch(err){ console.error(err); }
    _sessions[key] = { active:true, start:startVal };
  }
  function commitSession(r,c){
    var key = r + '|' + c;
    var sess = _sessions[key];
    if(!sess || !sess.active) return;
    var endVal = '';
    try{
      endVal = (window.cells && window.cells[r-1] && window.cells[r-1][c-1] ? window.cells[r-1][c-1].value : '');
    }catch(err){ console.error(err); }
    if(endVal !== sess.start){
      try{ _origPushUndo({type:'set', r:r, c:c, prev:sess.start, next:endVal}); }catch(err){ console.error(err); }
      if(window.redoStack) window.redoStack = [];
    }
    sess.active = false;
  }
  // Start on focus, commit on blur / Enter / Tab (capture so we run before default handlers)
  (document)&&document.addEventListener('focusin', function(e){
    var inp = e.target;
    if(inp && inp.tagName === 'INPUT' && inp.closest('.cell')){
      var td = inp.closest('.cell');
      var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
      beginSession(r,c);
    }
  }, true);
  (document)&&document.addEventListener('focusout', function(e){
    var inp = e.target;
    if(inp && inp.tagName === 'INPUT' && inp.closest('.cell')){
      var td = inp.closest('.cell');
      var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
      commitSession(r,c);
    }
  }, true);
  (document)&&document.addEventListener('keydown', function(e){
    if((e.key === 'Enter' || e.key === 'Tab') && document.activeElement && document.activeElement.tagName === 'INPUT' && document.activeElement.closest('.cell')){
      var td = document.activeElement.closest('.cell');
      var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
      commitSession(r,c);
    }
  }, true);

  // ---------- Freeze toggles: pressed state + proper toggle ----------
  function updateFreezeButtons(){
    try{
      var topBtn = document.getElementById('btn-freeze-top');
      var colBtn = document.getElementById('btn-freeze-firstcol');
      if(topBtn){ topBtn.classList.add('btn-toggle'); topBtn.classList.toggle('pressed', !!(window.frozenPane && window.frozenPane.row)); }
      if(colBtn){ colBtn.classList.add('btn-toggle'); colBtn.classList.toggle('pressed', !!(window.frozenPane && window.frozenPane.col)); }
    }catch(err){ console.error(err); }
  }
  function toggleTop(e){
    e.preventDefault(); e.stopPropagation();
    window.frozenPane = window.frozenPane || {row:0,col:0};
    window.frozenPane.row = window.frozenPane.row ? 0 : 1;
    if(typeof window.rebuild === 'function') window.rebuild();
    updateFreezeButtons();
  }
  function toggleFirst(e){
    e.preventDefault(); e.stopPropagation();
    window.frozenPane = window.frozenPane || {row:0,col:0};
    window.frozenPane.col = window.frozenPane.col ? 0 : 1;
    if(typeof window.rebuild === 'function') window.rebuild();
    updateFreezeButtons();
  }
  (function bindFreeze(){
    try{
      var topBtn = document.getElementById('btn-freeze-top');
      var colBtn = document.getElementById('btn-freeze-firstcol');
      if(topBtn){ (topBtn)&&topBtn.addEventListener('click', toggleTop, true); }
      if(colBtn){ (colBtn)&&colBtn.addEventListener('click', toggleFirst, true); }
      updateFreezeButtons();
    }catch(err){ console.error(err); }
  })();

  // ---------- +Sheet robustness ----------
  (function ensureAddSheet(){
    var btn = document.getElementById('btn-add-sheet');
    if(!btn) return;
    (btn)&&btn.addEventListener('click', function(e){
      e.preventDefault(); e.stopPropagation();
      try{ if(typeof window.commitToSheet === 'function') window.commitToSheet(); }catch(err){ console.error(err); }
      if(!window.sheets) window.sheets = [{name:'Sheet 1', data:null}];
      window.sheets.push({ name:'Sheet ' + (window.sheets.length+1), data:null });
      try{
        if(typeof window.loadFromSheet === 'function'){
          window.loadFromSheet(window.sheets.length-1);
        }else{
          window.currentSheet = window.sheets.length-1;
          var R = window.ROWS || 100, C = window.COLS || 26;
          window.cells = Array.from({length:R}, ()=>Array.from({length:C}, ()=>({value:'', style:{}})));
          if(typeof window.rebuild === 'function') window.rebuild();
        }
        if(typeof window.renderTabs === 'function') window.renderTabs();
      }catch(err){ console.error(err); }
    }, true);
  })();

  // ---------- Horizontal slider: ensure wired & visible ----------
  (function ensureHSlider(){
    var slider = document.getElementById('hscroll');
    var main = document.querySelector('main');
    if(!slider || !main) return;
    function sync(){
      var w = main.scrollWidth - main.clientWidth;
      if(w < 0) w = 0;
      slider.max = String(w);
      slider.value = String(main.scrollLeft);
    }
    slider.style.display = 'block';
    (slider)&&slider.addEventListener('input', function(){ main.scrollLeft = parseInt(slider.value||'0',10) || 0; }, {passive:true});
    (main)&&main.addEventListener('scroll', function(){ slider.value = String(main.scrollLeft); }, {passive:true});
    (window)&&window.addEventListener('resize', sync);
    setTimeout(sync, 150);
  })();

  // ---------- Dictation: Advance limit (wrap logic) & Numbers default ----------
  (function overrideDictation(){
    var dirSel = document.getElementById('voice-direction');
    var autoChk = document.getElementById('autoAdvance') || document.getElementById('auto-advance');
    var limitInput = document.getElementById('advanceLimit');
    var micModeSel = document.getElementById('v22-mic-mode') || document.getElementById('mic-mode') || document.getElementById('micMode');

    var state = { dir:null, startCol:null, startRow:null, count:0 };

    function normalizeNumbers(t){ return (t||'').replace(/[^0-9\.\-]/g,'').trim(); }

    function applyAdvance(){
      var N = Math.max(1, parseInt(limitInput && limitInput.value ? limitInput.value : 3, 10) || 3);
      var dir = (dirSel && dirSel.value) ? dirSel.value : 'right';
      if(dir === 'right'){
        state.count++;
        if(state.count < N){
          if(typeof moveActive === 'function') moveActive('right',1,{wrapRows:false});
        }else{
          state.count = 0;
          var nextR = Math.min((window.ROWS||9999), (state.startRow||window.active.r)+1);
          if(typeof setActive === 'function') setActive(nextR, state.startCol || window.active.c);
        }
      }else if(dir === 'down'){
        state.count++;
        if(state.count < N){
          if(typeof moveActive === 'function') moveActive('down',1,{wrapRows:false});
        }else{
          state.count = 0;
          var nextC = Math.min((window.COLS||9999), (state.startCol||window.active.c)+1);
          if(typeof setActive === 'function') setActive(state.startRow || window.active.r, nextC);
        }
      }else{
        if(typeof moveActive === 'function') moveActive(dir,1,{wrapRows:true});
      }
      var actInput = document.querySelector('.cell.active input'); if(actInput) setTimeout(function(){ actInput.focus(); }, 10);
    }

    function bindWhenReady(){
      var rec = window.recognition || window.webkitSpeechRecognitionInstance || null;
      if(!rec){ setTimeout(bindWhenReady, 300); return; }

      // default mic mode to Numbers
      if(micModeSel) micModeSel.value = 'numbers';

      rec.onresult = function(ev){
  window._dictationActive = true;
        try{
          var res = ev.results[ev.results.length-1][0];
          var raw = (res && res.transcript ? res.transcript : '').trim();
          var mode = micModeSel ? micModeSel.value : 'numbers';
          var text = (mode === 'numbers') ? normalizeNumbers(raw) : raw;
          if(!text) return;

          // ensure edit session is active to group undo
          var el = document.activeElement;
          if(el && el.tagName==='INPUT' && el.closest('.cell')){
            var td = el.closest('.cell');
            var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
            if(!window._editSessions || !window._editSessions[r+'|'+c] || !window._editSessions[r+'|'+c].active){
              // trigger focusin path if needed
              try{ el.dispatchEvent(new FocusEvent('focus', {bubbles:true})); }catch(err){ console.error(err); }
            }
          }

          // write cell
          if(typeof setCell === 'function'){
            setCell(window.active.r, window.active.c, String(text), { skipUndo:false });
          }else{
            if(window.cells && window.cells[window.active.r-1] && window.cells[window.active.r-1][window.active.c-1]){
              window.cells[window.active.r-1][window.active.c-1].value = String(text);
            }
          }

          // initialize wrapping state when direction changes or new block
          var dir = (dirSel && dirSel.value) ? dirSel.value : 'right';
          if(state.dir !== dir || state.startCol === null || state.startRow === null){
            state.dir = dir; state.startCol = window.active.c; state.startRow = window.active.r; state.count = 0;
          }

          if(autoChk && autoChk.checked){ applyAdvance(); }
        }catch(err){ console.error(err); }
      };
    }
    bindWhenReady();
  })();

  // ---------- Remove Dark/Compact if any DOM remnants ----------
  (function rmViewRemnants(){
    ['v22-toggle-dark','v22-toggle-compact'].forEach(function(id){
      var el = document.getElementById(id); if(el && el.parentNode) try{ el.parentNode.removeChild(el);}catch(err){ console.error(err); }
    });
  })();

})(); // end r2e7 overrides

(function(){
  /* ================== Robust Zoom (Chromium-friendly) ================== */
  (function(){
    var zoom = document.getElementById('zoom-slider');
    var main = document.querySelector('main');
    if(!zoom || !main) return;
    // Choose a sensible target inside main to scale
    var target = main.querySelector('#gridHost') || main.querySelector('#grid') || main.firstElementChild || main;
    function applyZoom2(){
      var z = parseFloat(zoom.value || '1') || 1;
      // Keep original variable for any CSS that references it
      try { document.documentElement.style.setProperty('--zoom', String(z)); } catch(err){ console.error(err); }
      // Use CSS zoom for Chromium (no layout jumps like transform)
      try { target.style.zoom = z; } catch(err){ console.error(err); }
    }
    (zoom)&&zoom.addEventListener('input', applyZoom2);
    // Initialize
    setTimeout(applyZoom2, 50);
  })();

  /* ================== Whole-cell Undo/Redo consolidation ================== */
  (function(){
    var _origPushUndo = window.pushUndo ? window.pushUndo.bind(window) : function(){};
    window._editSessions = window._editSessions || {};
    function key(r,c){ return r + '|' + c; }

    window.pushUndo = function(entry){
      // swallow keystroke-level 'set' entries during an active edit session for that cell
      try{
        if(entry && entry.type === 'set' && entry.r && entry.c){
          var s = window._editSessions[key(entry.r, entry.c)];
          if(s && s.active) return;
        }
      }catch(err){ console.error(err); }
      return _origPushUndo(entry);
    };

    function begin(r,c){
      var k = key(r,c);
      var prev = '';
      try{ prev = (window.cells && window.cells[r-1] && window.cells[r-1][c-1] ? (window.cells[r-1][c-1].value||'') : ''); }catch(err){ console.error(err); }
      window._editSessions[k] = { active:true, prev:prev };
    }
    function commit(r,c){
      var k = key(r,c);
      var sess = window._editSessions[k];
      if(!sess || !sess.active) return;
      var next = '';
      try{ next = (window.cells && window.cells[r-1] && window.cells[r-1][c-1] ? (window.cells[r-1][c-1].value||'') : ''); }catch(err){ console.error(err); }
      if(next !== sess.prev){
        try{ _origPushUndo({ type:'set', r:r, c:c, prev:sess.prev, next:next }); if(window.redoStack) window.redoStack=[]; }catch(err){ console.error(err); }
      }
      sess.active = false;
    }

    // Capture at document level: focus begins session, blur/Enter/Tab commits
    (document)&&document.addEventListener('focusin', function(e){
      var inp = e.target;
      if(inp && inp.tagName === 'INPUT' && inp.closest('.cell')){
        var td = inp.closest('.cell'); var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
        begin(r,c);
      }
    }, true);
    (document)&&document.addEventListener('focusout', function(e){
      var inp = e.target;
      if(inp && inp.tagName === 'INPUT' && inp.closest('.cell')){
        var td = inp.closest('.cell'); var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
        commit(r,c);
      }
    }, true);
    (document)&&document.addEventListener('keydown', function(e){
      if((e.key === 'Enter' || e.key === 'Tab') && document.activeElement && document.activeElement.tagName === 'INPUT' && document.activeElement.closest('.cell')){
        var td = document.activeElement.closest('.cell'); var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
        commit(r,c);
      }
    }, true);

    // After any full rebuild, pressed states/handlers reattach
    var _origRebuild = window.rebuild;
    window.rebuild = function(){
      if(typeof _origRebuild === 'function') _origRebuild();
      // nothing special here for undo; sessions are per-input focus
    };
  })();

  /* ================== Dictation Advance Limit with proper wrapping ================== */
  (function(){
    var dirSel = document.getElementById('voice-direction');
    var auto = document.getElementById('autoAdvance') || document.getElementById('auto-advance');
    var limitEl = document.getElementById('advanceLimit');
    var micMode = document.getElementById('v22-mic-mode') || document.getElementById('mic-mode') || document.getElementById('micMode');
    // default Numbers
    if(micMode) micMode.value = 'numbers';

    var state = { dir: null, baseRow: null, baseCol: null, count: 0 };

    function normalizeNumbers(t){ return (t||'').replace(/[^0-9\.\-]/g,'').trim(); }

    function ensureBaseFromActive(){
      if(!window.active) return;
      if(state.baseRow == null || state.baseCol == null || state.dir !== (dirSel && dirSel.value)){
        state.dir = dirSel ? dirSel.value : 'right';
        state.baseRow = window.active.r;
        state.baseCol = window.active.c;
        state.count = 0;
      }
    }

    function advanceOnce(){
      var N = Math.max(1, parseInt(limitEl && limitEl.value ? limitEl.value : '3', 10) || 3);
      var dir = (dirSel && dirSel.value) ? dirSel.value : 'right';
      if(dir === 'right'){
        state.count++;
        if(state.count < N){
          if(typeof moveActive === 'function') moveActive('right',1,{wrapRows:false});
        }else{
          state.count = 0;
          // move to next row, same starting column (wrap)
          state.baseRow = Math.min((window.ROWS||9999), (window.active.r + 1));
          if(typeof setActive === 'function') setActive(state.baseRow, state.baseCol);
        }
      }else if(dir === 'down'){
        state.count++;
        if(state.count < N){
          if(typeof moveActive === 'function') moveActive('down',1,{wrapRows:false});
        }else{
          state.count = 0;
          // move to next column, same starting row block
          state.baseCol = Math.min((window.COLS||9999), (window.active.c + 1));
          if(typeof setActive === 'function') setActive(state.baseRow, state.baseCol);
        }
      }else{
        // left/up: fallback to normal move
        if(typeof moveActive === 'function') moveActive(dir,1,{wrapRows:true});
      }
      // keep focus in cell
      var inp = document.querySelector('.cell.active input'); if(inp) setTimeout(function(){ try{ inp.focus(); }catch(err){ console.error(err); } }, 10);
    }

    function bindRec(){
      var rec = window.recognition || window.webkitSpeechRecognitionInstance || null;
      if(!rec){ setTimeout(bindRec, 300); return; }
      rec.onresult = function(e){
  window._dictationActive = true;
        try{
          var r = e.results[e.results.length-1][0]; var raw = (r && r.transcript) ? r.transcript.trim() : '';
          var mode = micMode ? micMode.value : 'numbers';
          var text = (mode === 'numbers') ? normalizeNumbers(raw) : raw;
          if(!text) return;
          // write into active cell
          if(typeof setCell === 'function'){ setCell(window.active.r, window.active.c, String(text), { skipUndo:false }); }
          else if(window.cells && window.active){ window.cells[window.active.r-1][window.active.c-1].value = String(text); var inp = document.querySelector('.cell.active input'); if(inp) inp.value = String(text); }
          ensureBaseFromActive();
          if(auto && auto.checked) advanceOnce();
        }catch(err){ console.error(err); }
      };
    }
    bindRec();

    // Reset base block when user manually changes direction or active
    (function(){
      var reset = function(){ state.baseRow=null; state.baseCol=null; state.count=0; state.dir=null; };
      if(dirSel) (dirSel)&&dirSel.addEventListener('change', reset);
      (document)&&document.addEventListener('selectionchange', function(){ /* best-effort */ });
      (document)&&document.addEventListener('click', function(e){
        if(e.target && e.target.closest && e.target.closest('.cell')){ reset(); }
      }, true);
    })();
  })();

  /* ================== Delete Sheet button & behavior ================== */
  (function(){
    var tabs = document.getElementById('sheet-tabs');
    if(!tabs) return;
    // Inject a Delete Sheet button right after + Sheet (or at end if not found)
    var addBtn = document.getElementById('btn-add-sheet') || tabs.querySelector('[data-action="add-sheet"]');
    var delBtn = document.getElementById('btn-del-sheet');
    if(!delBtn){
      delBtn = document.createElement('button');
      delBtn.id = 'btn-del-sheet';
      delBtn.title = 'Delete current sheet';
      delBtn.textContent = '− Sheet';
      if(addBtn && addBtn.parentNode){
        addBtn.parentNode.insertBefore(delBtn, addBtn.nextSibling);
      }else{
        tabs.appendChild(delBtn);
      }
    }
    (delBtn)&&delBtn.addEventListener('click', function(e){
      e.preventDefault(); e.stopPropagation();
      try{ if(typeof window.commitToSheet === 'function') window.commitToSheet(); }catch(err){ console.error(err); }
      if(!window.sheets || !window.sheets.length){ return; }
      // Choose index to delete = currentSheet or last
      var idx = (typeof window.currentSheet === 'number') ? window.currentSheet : (window.sheets.length-1);
      // Prevent deleting last remaining sheet: recreate blank after
      window.sheets.splice(idx, 1);
      if(window.sheets.length === 0){
        window.sheets.push({ name:'Sheet 1', data:null });
        window.currentSheet = 0;
      }else{
        if(idx >= window.sheets.length) idx = window.sheets.length - 1;
        window.currentSheet = idx;
      }
      if(typeof window.loadFromSheet === 'function') window.loadFromSheet(window.currentSheet);
      if(typeof window.renderTabs === 'function') window.renderTabs();
    }, true);
  })();
})(); 

// Delete current sheet handler (keeps at least one sheet)
(function(){
  var delBtn = document.getElementById('btn-del-sheet');
  if(!delBtn) return;
  (delBtn)&&delBtn.addEventListener('click', function(){
    try { commitToSheet(); } catch(err){ console.error(err); }
    if (typeof sheets === 'undefined' || typeof currentSheet === 'undefined') return;
    if (sheets.length <= 1){
      sheets[0] = { name: 'Sheet 1', data: null };
      currentSheet = 0;
      try { loadFromSheet(0); } catch(err){ console.error(err); }
    } else {
      sheets.splice(currentSheet, 1);
      if (currentSheet >= sheets.length) currentSheet = sheets.length - 1;
      try { loadFromSheet(currentSheet); } catch(err){ console.error(err); }
    }
    try { renderTabs(); } catch(err){ console.error(err); }
  });
})();

// Robust Zoom binding (Chromium-friendly) without breaking layout
(function(){
  var slider = document.getElementById('zoom-slider');
  if(!slider) return;
  var host = document.getElementById('gridHost') || document.getElementById('grid-container') || document.getElementById('grid') || document.body;
  function applyZ(){
    var z = parseFloat(slider.value || '1') || 1;
    try { document.documentElement.style.setProperty('--zoom', String(z)); } catch(err){ console.error(err); }
    try { host.style.zoom = z; } catch(err){ console.error(err); }
  }
  (slider)&&slider.addEventListener('input', applyZ);
  setTimeout(applyZ, 50);
})();

// Freeze buttons behave like ribbon toggles (pressed class synced with state)
(function(){
  var btnTop = document.getElementById('freezeTopRow');
  var btnCol = document.getElementById('freezeFirstCol');
  function sync(){
    if (typeof frozenTopRow !== 'undefined'){
      if (btnTop) btnTop.classList.toggle('pressed', !!frozenTopRow);
    }
    if (typeof frozenFirstCol !== 'undefined'){
      if (btnCol) btnCol.classList.toggle('pressed', !!frozenFirstCol);
    }
  }
  if (btnTop) (btnTop)&&btnTop.addEventListener('click', function(){ setTimeout(sync, 0); });
  if (btnCol) (btnCol)&&btnCol.addEventListener('click', function(){ setTimeout(sync, 0); });
  (document)&&document.addEventListener('DOMContentLoaded', sync);
})();

(function(){
  // Make toNumbers available (keeps decimals, trims stray trailing dot like "5.")
  window.toNumbers = function(txt){
    var s = String(txt||'').replace(/[^0-9.\-]/g,'').trim();
    if (/^\-?\d+\.$/.test(s)) s = s.slice(0, -1);
    return s;
  };
  // Ensure normalizeNumbers delegates to toNumbers if used elsewhere
  var _oldNorm = window.normalizeNumbers;
  window.normalizeNumbers = function(t){
    return window.toNumbers(t);
  };
})();

/* ================== Whole-cell Undo/Redo consolidation — FINAL ================== */
(function(){
  var _origPushUndo = (typeof window.pushUndo === 'function') ? window.pushUndo.bind(window) : function(){};
  var session = { active:false, r:null, c:null, prev:'', justPushed:false };

  function readCell(r,c){
    try{
      if (window.cells && window.cells[r-1] && window.cells[r-1][c-1]) return String(window.cells[r-1][c-1].value ?? '');
      if (typeof getCell === 'function') return String(getCell(r,c) ?? '');
    }catch(err){ console.error(err); }
    return '';
  }

  function startSession(r,c){
    session.active = true;
    session.r = r; session.c = c;
    session.prev = readCell(r,c);
    session.justPushed = false;
  }

  function endSession(push){
    if(!session.active) return;
    var r = session.r, c = session.c;
    var before = session.prev, after = readCell(r,c);
    session.active = false;
    if (push && after !== before){
      try{
        _origPushUndo({ type:'set', r:r, c:c, prev:before, next:after });
        session.justPushed = true;
        setTimeout(function(){ session.justPushed = false; }, 0);
      }catch(err){ console.error(err); }
    }
  }

  window.pushUndo = function(entry){
    try{
      if (!entry || !entry.type) return _origPushUndo(entry);
      if (session.active && entry.type === 'set' && entry.r === session.r && entry.c === session.c) return;
      if (session.active && (entry.type === 'cursor' || entry.type === 'select')) return;
      if (!session.active && session.justPushed && entry.type === 'set' && entry.r === session.r && entry.c === session.c){
        session.justPushed = false;
        return;
      }
    }catch(err){ console.error(err); }
    return _origPushUndo(entry);
  };

  (document)&&document.addEventListener('focusin', function(e){
    var inp = e.target;
    if (inp && inp.tagName === 'INPUT' && inp.closest('.cell')){
      var td = inp.closest('.cell');
      var r = +td.getAttribute('data-r'), c = +td.getAttribute('data-c');
      startSession(r,c);
    }
  }, true);

  function maybeCommitFrom(el, push){
    if (el && el.tagName === 'INPUT' && el.closest('.cell')) endSession(push);
  }
  (document)&&document.addEventListener('focusout', function(e){ maybeCommitFrom(e.target, true); }, true);
  (document)&&document.addEventListener('keydown', function(e){
    if (e.key === 'Enter' || e.key === 'Tab')  maybeCommitFrom(document.activeElement, true);
    if (e.key === 'Escape')                    maybeCommitFrom(document.activeElement, false);
  }, true);
})();

(document)&&document.addEventListener('DOMContentLoaded', () => {
  wireMicrophoneControls();
});

function wireMicrophoneControls() {
  const micBtn = document.getElementById('micBtn');
  let recognizing = false;
  let recognition;

  // Create mic status span if missing
  let micStatus = document.getElementById('micStatus');
  if (!micStatus) {
    micStatus = document.createElement('span');
    micStatus.id = 'micStatus';
    micStatus.textContent = 'Mic off';
    micBtn.insertAdjacentElement('afterend', micStatus);
  }

  if (!('webkitSpeechRecognition' in window || 'SpeechRecognition' in window)) {
    micStatus.textContent = 'Speech recognition not supported in this browser.';
    micBtn.disabled = true;
    return;
  }

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.lang = 'en-US';

  recognition.onstart = () => {
    recognizing = true;
    micStatus.textContent = 'Listening...';
    micBtn.classList.add('listening');
  };

  recognition.onend = () => {
    recognizing = false;
    micStatus.textContent = 'Mic off';
    micBtn.classList.remove('listening');
  };

  recognition.onerror = (event) => {
    console.error('Recognition error:', event.error);
    micStatus.textContent = 'Error: ' + event.error;
  };

  recognition.onresult = (event) => {
    let finalTranscript = '';
    for (let i = event.resultIndex; i < event.results.length; ++i) {
      if (event.results[i].isFinal) {
        finalTranscript += event.results[i][0].transcript;
      }
    }
    if (finalTranscript.trim()) {
      console.log('Recognized:', finalTranscript);
      // Future: insert into grid
    }
  };

  micBtn.onclick = () => {
    if (recognizing) {
      recognition.stop();
    } else {
      recognition.start();
    }
  };
}

(window)&&window.addEventListener('load', () => {
  console.log("🎤 Mic script initializing");

  const micBtn = document.getElementById('micBtn');
  if (!micBtn) {
    console.warn("Mic button not found.");
    return;
  }

  let micStatus = document.getElementById('micStatus');
  if (!micStatus) {
    micStatus = document.createElement('span');
    micStatus.id = 'micStatus';
    micStatus.textContent = 'Mic off';
    micBtn.insertAdjacentElement('afterend', micStatus);
  }

  if (!('webkitSpeechRecognition' in window || 'SpeechRecognition' in window)) {
    micStatus.textContent = 'Speech recognition not supported.';
    micBtn.disabled = true;
    return;
  }

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  const recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.lang = 'en-US';

  let recognizing = false;

  recognition.onstart = () => {
    recognizing = true;
    micStatus.textContent = 'Listening...';
    micBtn.classList.add('listening');
  };

  recognition.onend = () => {
    recognizing = false;
    micStatus.textContent = 'Mic off';
    micBtn.classList.remove('listening');
  };

  recognition.onerror = (event) => {
    console.error('Recognition error:', event.error);
    micStatus.textContent = 'Error: ' + event.error;
  };

  recognition.onresult = (event) => {
    let finalTranscript = '';
    for (let i = event.resultIndex; i < event.results.length; ++i) {
      if (event.results[i].isFinal) {
        finalTranscript += event.results[i][0].transcript;
      }
    }
    if (finalTranscript.trim()) {
      console.log('🎤 Final Recognized:', finalTranscript);
    }
  };

  (micBtn)&&micBtn.addEventListener('click', () => {
    if (recognizing) {
      recognition.stop();
    } else {
      try {
        recognition.start();
      } catch (e) {
        console.error('Failed to start recognition:', e);
      }
    }
  });
});

// === Dictation-aware normalization ===
function normalizeDigitsTo0xxx(val){
  try{
    var cb = document.getElementById('optScaleDigits');
    var enabled = cb ? cb.checked : true;
    if (!enabled) return val;
    var t = (val==null ? '' : String(val)).trim();
    if (/^[-+]?\d+$/.test(t) && !/\./.test(t)){
      var sign = t.charAt(0) === '-' ? '-' : (t.charAt(0) === '+' ? '+' : '');
      var digits = t.replace(/[^0-9]/g,'');
      if (digits.length) return sign + '0.' + digits;
    }
    return t;
  } catch(e){ return val; }
}
(function(){
  // Global flag to mark dictation commits
  window._dictationActive = false;
  // Monkey-patch setCell: only adjust when dictation is active
  if (typeof window.setCell === 'function'){
    var _origSetCell = window.setCell;
    window.setCell = function(r,c,val,opts){
      try{
        if (window._dictationActive) val = normalizeDigitsTo0xxx(val);
      }catch(err){ console.error(err); }
      return _origSetCell.call(this, r,c,val,opts);
    };
  }
})();

(function(){
  function getSel(){ return typeof sel!=='undefined' ? normSel(sel) : null; }
  function getCellVal(r,c){
    try{ return (cells[r-1] && cells[r-1][c-1] ? (cells[r-1][c-1].value||'') : ''); }catch(e){ return ''; }
  }

  function collectSeedValsDown(sr, c){
    var vals = [];
    // Take up to first 2 non-empty seeds from the top of the selection to infer a step
    for (var r=sr.r1; r<=sr.r2 && vals.length<2; r++){
      var v = getCellVal(r,c);
      if (String(v).trim() !== '') vals.push(v);
    }
    // If none found, allow single seed from active cell
    if (vals.length===0 && typeof active==='object'){ vals.push(getCellVal(active.r, c)); }
    return vals;
  }

  function collectSeedValsRight(sr, r){
    var vals = [];
    for (var c=sr.c1; c<=sr.c2 && vals.length<2; c++){
      var v = getCellVal(r,c);
      if (String(v).trim() !== '') vals.push(v);
    }
    if (vals.length===0 && typeof active==='object'){ vals.push(getCellVal(r, active.c)); }
    return vals;
  }

  function fillDown(){
    var sr = getSel(); if(!sr){ alert('Select a range first.'); return; }
    var c = (typeof active==='object' && active.c>=sr.c1 && active.c<=sr.c2) ? active.c : sr.c1;
    var seeds = collectSeedValsDown(sr, c);
    if (seeds.length===0){ alert('Provide at least one seed value at the top of the selection.'); return; }

    // Determine start row after seeds
    var seedRows = Math.max(1, seeds.length);
    var startRow = sr.r1 + seedRows;
    var endRow = sr.r2;
    if (endRow < startRow){
      var more = parseInt(prompt('Rows to fill below?', '10')||'0',10);
      if (more>0){ endRow = startRow + more - 1; } else { return; }
    }
    var count = (endRow - startRow + 1);
    if (count <= 0) return;

    // Build the series
    var series = (seeds.length>=2) ? detectSeries(seeds) : null;
    var out = buildSeries(seeds.length?seeds:[''], series, count);

    // Apply to each column in selection
    for (var cc=sr.c1; cc<=sr.c2; cc++){
      for (var j=0; j<count; j++){
        setCell(startRow + j, cc, out[j]);
      }
    }
  }

  function fillRight(){
    var sr = getSel(); if(!sr){ alert('Select a range first.'); return; }
    var r = (typeof active==='object' && active.r>=sr.r1 && active.r<=sr.r2) ? active.r : sr.r1;
    var seeds = collectSeedValsRight(sr, r);
    if (seeds.length===0){ alert('Provide at least one seed value at the left of the selection.'); return; }

    var seedCols = Math.max(1, seeds.length);
    var startCol = sr.c1 + seedCols;
    var endCol = sr.c2;
    if (endCol < startCol){
      var more = parseInt(prompt('Columns to fill to the right?', '10')||'0',10);
      if (more>0){ endCol = startCol + more - 1; } else { return; }
    }
    var count = (endCol - startCol + 1);
    if (count <= 0) return;

    var series = (seeds.length>=2) ? detectSeries(seeds) : null;
    var out = buildSeries(seeds.length?seeds:[''], series, count);

    for (var rr=sr.r1; rr<=sr.r2; rr++){
      for (var i=0; i<count; i++){
        setCell(rr, startCol + i, out[i]);
      }
    }
  }

  // Wire buttons
  var btnDown = document.getElementById('btn-fill-down');
  var btnRight = document.getElementById('btn-fill-right');
  if (btnDown) btnDown.addEventListener('click', fillDown);
  if (btnRight) btnRight.addEventListener('click', fillRight);

  // Keyboard shortcuts: Alt+Shift+D / Alt+Shift+R
  window.addEventListener('keydown', function(ev){
    if (ev.altKey && ev.shiftKey && !ev.ctrlKey && !ev.metaKey){
      if (ev.key.toLowerCase()==='d'){ ev.preventDefault(); fillDown(); }
      if (ev.key.toLowerCase()==='r'){ ev.preventDefault(); fillRight(); }
    }
  });
})();

/* ===== Minimal Autosave (localStorage only; no OneDrive) ===== */
(function(){
  function sanitize(name){
    name = String(name||'Sheet 1').trim().replace(/[\\/:*?"<>|]/g,'-').slice(0,64);
    return name || 'Sheet 1';
  }
  function getSheetName(){
    try { return sanitize(localStorage.getItem('sheetName') || document.title.split(' – ')[0] || 'Sheet 1'); }
    catch { return 'Sheet 1'; }
  }
  function keyMain(){ return 'vs:autosave:' + getSheetName(); }
  function keyBackup(ts){ return `vs:backup:${getSheetName()}:${ts}`; }
  function nowISO(){ return new Date().toISOString().replace('T',' ').slice(0,19); }
  function nowLocal(){ return new Date().toLocaleTimeString('en-GB', {hour12:false, timeZone:'Europe/London'}); }

  // Build a sparse snapshot of the grid
  function snapshot(){
    try{
      var out = [];
      var rows = (typeof ROWS !== 'undefined' ? ROWS : 1200);
      var cols = (typeof COLS !== 'undefined' ? COLS : 52);
      for (var r=1; r<=rows; r++){
        for (var c=1; c<=cols; c++){
          var v = (cells[r-1] && cells[r-1][c-1] ? (cells[r-1][c-1].value ?? '') : '');
          if (v !== '') out.push([r,c,String(v)]);
        }
      }
      return { rows: rows, cols: cols, data: out };
    } catch(e){
      console.warn('snapshot() failed', e);
      return { rows: (typeof ROWS!=='undefined'?ROWS:1200), cols: (typeof COLS!=='undefined'?COLS:52), data: [] };
    }
  }

  // Apply a snapshot back into the grid (rebuilds if size differs)
  function applySnapshot(snap){
    if (!snap || !snap.rows || !snap.cols) return false;
    try{
      var host = document.getElementById('grid-container') || document.querySelector('.grid-host');
      if (typeof createGrid === 'function' && host){
        if (typeof ROWS !== 'undefined') ROWS = snap.rows;
        if (typeof COLS !== 'undefined') COLS = snap.cols;
        var _defR=1200,_defC=52; var _r=Math.max(_defR, snap.rows||0), _c=Math.max(_defC, snap.cols||0);
        createGrid(host, _r, _c);
      }
      if (Array.isArray(snap.data)){
        for (var i=0;i<snap.data.length;i++){
          var t = snap.data[i], r=t[0], c=t[1], v=t[2];
          if (typeof setCell === 'function') setCell(r,c,v,{skipUndo:true});
        }
      }
      return true;
    }catch(e){
      console.warn('applySnapshot() failed', e);
      return false;
    }
  }

  var Autosave = {
    enabled: true,
    INTERVAL_MS: 5000,
    DEBOUNCE_MS: 800,
    _deb: null,
    init: function(){
      var cb = document.getElementById('optAutosave');
      this.enabled = cb ? cb.checked : true;
      if (cb) cb.addEventListener('change', ()=>{ this.enabled = cb.checked; if (this.enabled) this.save('toggle-on'); });

      var btnSave = document.getElementById('btnSaveNow');
      var btnRestore = document.getElementById('btnRestore');
      if (btnSave) btnSave.addEventListener('click', ()=> this.save('manual'));
      if (btnRestore) btnRestore.addEventListener('click', ()=> this.loadLatest());

      // Try restoring on first load
      this.loadLatest();

      // Periodic save
      setInterval(()=>{ if (this.enabled) this.save('interval'); }, this.INTERVAL_MS);

      // Save on tab hide/close
      document.addEventListener('visibilitychange', ()=>{
        if (document.visibilityState === 'hidden' && this.enabled) this.save('hidden');
      });
      window.addEventListener('beforeunload', ()=>{ if (this.enabled) try{ this.save('unload'); }catch{}; });
    },
    touch: function(){
      if (!this.enabled) return;
      clearTimeout(this._deb);
      this._deb = setTimeout(()=> this.save('debounce'), this.DEBOUNCE_MS);
    },
    save: function(why){
      var status = document.getElementById('autosaveStatus');
      try{
        status && (status.textContent = 'Saving…');
        var snap = snapshot();
        var json = JSON.stringify(snap);
        localStorage.setItem(keyMain(), json);

        // Keep a few rolling backups
        var ts = nowISO().replace(/[: ]/g,'-');
        localStorage.setItem(keyBackup(ts), json);
        var keys = Object.keys(localStorage).filter(k => k.startsWith('vs:backup:'+getSheetName()+':')).sort();
        while (keys.length > 5) { try{ localStorage.removeItem(keys.shift()); }catch{} }

        status && (status.textContent = 'Saved ' + nowLocal());
        return true;
      }catch(e){
        console.warn('Autosave failed', e);
        status && (status.textContent = 'Save failed');
        return false;
      }
    },
    loadLatest: function(){
      try{
        var raw = localStorage.getItem(keyMain());
        if (!raw) return false;
        var ok = applySnapshot(JSON.parse(raw));
        var status = document.getElementById('autosaveStatus');
        status && (status.textContent = ok ? ('Restored ' + nowLocal()) : 'Restore failed');
        return ok;
      }catch(e){
        console.warn('Autosave restore failed', e);
        var status = document.getElementById('autosaveStatus');
        status && (status.textContent = 'Restore failed');
        return false;
      }
    }
  };

  // Patch setCell once to trigger debounced autosave
  if (typeof window.setCell === 'function' && !window.setCell.__autosavePatched){
    var _orig = window.setCell;
    window.setCell = function(r,c,val,opts){
      var res = _orig.apply(this, arguments);
      try { Autosave.touch(); } catch(err){ console.error(err); }
      return res;
    };
    window.setCell.__autosavePatched = true;
  } else {
    // If setCell isn't defined yet, try after DOM loaded
    document.addEventListener('DOMContentLoaded', function(){
      if (typeof window.setCell === 'function' && !window.setCell.__autosavePatched){
        var _o = window.setCell;
        window.setCell = function(r,c,val,opts){
          var res = _o.apply(this, arguments);
          try { Autosave.touch(); } catch(err){ console.error(err); }
          return res;
        };
        window.setCell.__autosavePatched = true;
      }
    });
  }

  // Expose in case you want manual calls
  window.Autosave = Autosave;

  // Boot
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ()=> Autosave.init());
  } else {
    Autosave.init();
  }
})();

(function(){
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  document.addEventListener('DOMContentLoaded', function(){
    const micBtnOrig = document.getElementById('micBtn');
    if(!micBtnOrig || !SR) return;

    // Remove any previously attached listeners by cloning the button
    const clone = micBtnOrig.cloneNode(true);
    micBtnOrig.parentNode.replaceChild(clone, micBtnOrig);
    const micBtn = document.getElementById('micBtn');

    // Ensure mic status element exists
    let micStatus = document.getElementById('micStatus');
    if (!micStatus){
      micStatus = document.createElement('span');
      micStatus.id = 'micStatus';
      micStatus.textContent = 'Mic off';
      micBtn.insertAdjacentElement('afterend', micStatus);
    }

    // External controls (if present)
    const modeSel   = document.getElementById('v22-mic-mode') || document.getElementById('micMode') || document.getElementById('mic-mode');
    const dirSel    = document.getElementById('direction') || document.getElementById('voice-direction');
    const autoChk   = document.getElementById('autoAdvance') || document.getElementById('auto-advance');
    const limitInp  = document.getElementById('advanceLimit');

    // Single recognizer
    const rec = new SR();
    rec.continuous = true;
    rec.interimResults = false;
    try { rec.lang = 'en-US'; } catch(err){ console.error(err); }

    let keepAlive = false;
    let running = false;
    let backoff = 300; // ms (exp backoff up to 3s)
    let restartTimer = null;

    // Optional: keep AudioContext awake + Wake Lock
    let ac = null;
    try {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (AC) ac = new AC();
    } catch(err){ console.error(err); }
    async function resumePipelines(){
      try { if (ac && ac.state !== 'running') await ac.resume(); } catch(err){ console.error(err); }
      if ('wakeLock' in navigator && document.visibilityState === 'visible'){
        try { window.__wakeLock = window.__wakeLock || await navigator.wakeLock.request('screen'); } catch(err){ console.error(err); }
      }
    }

    function scheduleRestart(reason){
      if(!keepAlive) return;
      clearTimeout(restartTimer);
      const delay = Math.min(backoff, 3000);
      restartTimer = setTimeout(()=>{ try { rec.start(); } catch(err){ console.error(err); } }, delay);
      backoff = Math.min(delay * 1.6, 3000);
    }

    function setUI(listening){
      if (listening){
        micBtn.classList.add('listening');
        micBtn.classList.remove('btn-green');
        micBtn.classList.add('btn-red');
        micBtn.textContent = '🛑 Stop Mic';
        micStatus.textContent = 'Listening...';
      }else{
        micBtn.classList.remove('listening');
        micBtn.classList.remove('btn-red');
        micBtn.classList.add('btn-green');
        micBtn.textContent = '🎤 Start Mic';
        micStatus.textContent = 'Mic off';
      }
    }

    function toNumberish(text){
      if(!text) return null;
      let s = String(text).toLowerCase().trim();
      s = s.replace(/comma/g,'.').replace(/\s+(point|dot)\s+/g,'.');
      if(/[0-9]/.test(s)){
        // STRICT: keep digits and dots only; ignore hyphens entirely
        s = s.replace(/[^0-9.]/g,'');
        const parts = s.split('.'); if (parts.length>2) s = parts[0]+'.'+parts.slice(1).join('');
        if (/^\d+\.$/.test(s)) s = s.slice(0,-1);
        if(!s || s==='.') return null;
        const n = Number(s); return isFinite(n) ? String(n) : null;
      }
      return String(text||'').trim();
    }
return String(text||'').trim();
    }

    function writeActive(next){
      try{
        if (typeof setCell === 'function' && window.active){
          window._dictationActive = true; // allow 0.xxx normalization if your code checks this flag
          setCell(window.active.r, window.active.c, next);
          return;
        }
        const el = document.activeElement;
        if (el && el.tagName==='INPUT' && el.closest('.cell')){
          el.value = next; el.dispatchEvent(new Event('input',{bubbles:true}));
        }
      }catch(e){ console.warn('writeActive failed', e); }
    }

    function advanceWithLimitCompat(){
      if (!(autoChk && autoChk.checked)) return;
      const dir = (dirSel && dirSel.value) ? dirSel.value : 'right';
      const N = Math.max(1, parseInt(limitInp && limitInp.value ? limitInp.value : '3', 10));

      window.__dictBlock = window.__dictBlock || { startR: window.active ? window.active.r : 1, startC: window.active ? window.active.c : 1, count: 0 };

      const a = window.active ? { r: window.active.r, c: window.active.c } : { r:1, c:1 };
      const size = { rows: window.ROWS || 100, cols: window.COLS || 26 };
      let idx = window.__dictBlock.count;

      if (dir === 'right'){
        if (idx < N-1){ idx++; a.c = Math.min(size.cols, a.c+1); }
        else { idx=0; a.r = Math.min(size.rows, a.r+1); a.c = window.__dictBlock.startC; }
      } else if (dir === 'down'){
        if (idx < N-1){ idx++; a.r = Math.min(size.rows, a.r+1); }
        else { idx=0; a.c = Math.min(size.cols, a.c+1); a.r = window.__dictBlock.startR; }
      } else if (dir === 'left'){
        if (idx < N-1){ idx++; a.c = Math.max(1, a.c-1); }
        else { idx=0; a.r = Math.max(1, a.r-1); a.c = window.__dictBlock.startC; }
      } else if (dir === 'up'){
        if (idx < N-1){ idx++; a.r = Math.max(1, a.r-1); }
        else { idx=0; a.c = Math.max(1, a.c-1); a.r = window.__dictBlock.startR; }
      }
      window.__dictBlock.count = idx;
      if (typeof setActive === 'function') setActive(a.r, a.c, true);
      const actInput = document.querySelector('.cell.active input'); if (actInput) setTimeout(()=>{ try{ actInput.focus(); }catch(err){ console.error(err); } }, 10);
    }

    rec.onstart = function(){ running = true; backoff = 300; setUI(true); };
    rec.onend   = function(){ running = false; setUI(false); if (keepAlive) scheduleRestart('onend'); };
    rec.onerror = function(e){
      const err = (e && e.error) || '';
      if (err === 'not-allowed' || err === 'service-not-allowed'){ keepAlive=false; setUI(false); return; }
      if (keepAlive) scheduleRestart('onerror:'+err);
    };
    rec.onresult = function(ev){
      try{
        for (let i=ev.resultIndex; i<ev.results.length; i++){
          const item = ev.results[i]; if (!item.isFinal) continue;
          const raw = (item[0] && item[0].transcript ? item[0].transcript : '').trim();
          
          const mode = modeSel ? modeSel.value : 'numbers';
          let val = (mode === 'numbers') ? toNumberish(raw) : raw;
          if (val==null || val==='') continue;

          // Inline 0.xxx scaling (strict): when toggle ON and value is pure digits, force positive 0.<digits>
          let doScale = false;
          try { doScale = !!(document.getElementById('optScaleDigits') && document.getElementById('optScaleDigits').checked); } catch(err){ console.error(err); }
          if (doScale) {
            const s = String(val);
            const onlyDigits = /^\d+$/.test(s);
            const alreadyDec = /^-?0\.\d+$/.test(s);
            if (onlyDigits && !alreadyDec) {
              const digits = s.replace(/\D/g,'');
              val = '0.' + digits;
            }
          }

          window.__dictScaledInMic = true;
          writeActive(String(val));
          window.__dictScaledInMic = false;
          advanceWithLimitCompat();
    
        }
      } finally { window._dictationActive = false; }
    };

    async function startMic(){ keepAlive = true; await resumePipelines(); try { rec.start(); } catch(err){ console.error(err); } }
    function stopMic(){ keepAlive = false; clearTimeout(restartTimer); try { rec.stop(); } catch(err){ console.error(err); } setUI(false); if (window.__wakeLock){ try{ window.__wakeLock.release(); }catch(err){ console.error(err); } window.__wakeLock=null; } }

    window.startMicAlways = startMic;
    window.stopMicAlways  = stopMic;

    micBtn.addEventListener('click', function(){
      if (!running && !keepAlive) startMic(); else stopMic();
    });

    document.addEventListener('visibilitychange', function(){
      if (document.visibilityState === 'visible'){
        resumePipelines();
        if (keepAlive && !running) scheduleRestart('visibility');
      }
    });
    window.addEventListener('click', ()=>{ resumePipelines(); }, { once:true });
  });
})();

(function(){
  function needsScaling(str){
    if (!str) return false;
    const s = String(str).toLowerCase().trim();
    if (/^0\.\d+$/.test(s)) return false;  // already decimal, positive
    if (/^-0\.\d+$/.test(s)) return false; // already decimal, negative
    return /^\d+$/.test(s);                 // pure digits only
  }
  function scaleDigits(raw, makeNegative){
    const digits = String(raw).replace(/\D/g, '');
    if (!digits) return raw;
    return (makeNegative ? '-' : '') + '0.' + digits;
  }
  function wantScale(){
    const cb = document.getElementById('optScaleDigits'); return !!(cb && cb.checked);
  }
  if (typeof window.setCell === 'function' && !window.setCell.__scalePatchStrict){
    const _set = window.setCell;
    window.setCell = function(r, c, val, opts){
      try{
        if (window._dictationActive && wantScale()){ if (window.__dictScaledInMic===true) return _set.call(this, r, c, val, opts);
          const raw = (val==null ? '' : String(val));
          if (needsScaling(raw)){
            const unsigned = raw.replace(/^\s*-\s*/, ''); // ignore sanitized minus
            const neg = (window.__dictSaidMinus === true);  // trust transcript flag only
            val = scaleDigits(unsigned, neg);
          }
        }
      }catch(e){ /* ignore */ }
      return _set.call(this, r, c, val, opts);
    };
    window.setCell.__scalePatchStrict = true;
  }
})();

document.addEventListener('DOMContentLoaded', function(){
  var tgl = document.getElementById('toolbar-toggle-min');
  var ribbon = document.querySelector('.ribbon');
  if(tgl && ribbon){
    tgl.addEventListener('click', function(){
      try {
        var isCollapsed = ribbon.classList.toggle('collapsed');
        var header = document.querySelector('header');
        if(header){ header.classList.toggle('compact', isCollapsed); }
        tgl.setAttribute('aria-expanded', isCollapsed ? 'false' : 'true');
      } catch(err){ console.error(err); }
    });
  }
});


/* ================== Utilities ================== */
var MIN_COL_WIDTH = 20;
var ROW_HEIGHT = 22;
function colLabel(n){ var s=''; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }
function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
function normSel(a){ return { r1:Math.min(a.r1,a.r2), c1:Math.min(a.c1,a.c2), r2:Math.max(a.r1,a.r2), c2:Math.max(a.c1,a.c2) }; }
function parseMaybeNumber(s){ var t=(s==null?'':String(s)).trim().replace(',', '.'); var n=Number(t); return isFinite(n)? n : null; }
function parseMaybeDate(s){ var t=(s==null?'':String(s)).trim(); var d=new Date(t); return isNaN(+d)? null : d; }
function isoDate(d){ return d.toISOString().slice(0,10); }

/* ================== State ================== */
var ROWS=100, COLS=26;
var container=null;
var cells=[];   // grid data
var active={r:1,c:1};
var sel={r1:1,c1:1,r2:1,c2:1};
var undoStack=[], redoStack=[];
var colWidths = (function(){ try{ return JSON.parse(localStorage.getItem('colWidths') || '[]'); }catch(e){ return []; } })();
var rowHeights = (function(){ try{ return JSON.parse(localStorage.getItem('rowHeights') || '[]'); }catch(e){ return []; } })();
var merges = [];
var micWired=false;
function setMicAriaPressed(btn, state){ try{ if(btn){ btn.setAttribute("aria-pressed", state ? "true" : "false"); } }catch(e){ console.error(e); } }

var frozenPane = { row: 0, col: 0 };

/* ================== DOM Helpers ================== */
function cellTd(r,c){ return container ? container.querySelector('.cell[data-r="'+r+'"][data-c="'+c+'"]') : null; }
function cellInput(r,c){ var td = cellTd(r,c); return td ? td.querySelector('input') : null; }
function updateStatusSelection(){
  var R = normSel(sel);
  var single = R.r1 === R.r2 && R.c1 === R.c2;
  var el = document.getElementById('status-selection');
  if (el) el.textContent = single ? (colLabel(R.c1)+R.r1) : (colLabel(R.c1)+R.r1+':'+colLabel(R.c2)+R.r2);
}

/* ================== Rendering ================== */
function applyStyleToInput(inp, st){
  inp.style.fontWeight = (st && st.bold) ? '700' : '';
  inp.style.fontStyle  = (st && st.italic) ? 'italic' : '';
  inp.style.fontSize   = (st && st.size) || '';
  inp.style.color      = (st && st.color) || '';
  inp.style.backgroundColor = (st && st.bg) || '';
  inp.style.whiteSpace = (st && st.wrap) ? 'normal' : 'nowrap';
  inp.style.textAlign  = (st && st.align) || 'left';
}

function mergeCoverAt(r,c){ 
  for (var i=0;i<merges.length;i++){ var m=merges[i]; if(r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2) return m; }
  return null;
}

function render(){
  if(!container) return;
  var table=document.createElement('table'); table.className='grid'; table.setAttribute('role','grid'); table.setAttribute('aria-rowcount', String(ROWS)); table.setAttribute('aria-colcount', String(COLS));

  // THEAD
  var thead=document.createElement('thead');
  var trh=document.createElement('tr');
  var corner=document.createElement('th'); corner.setAttribute('role','columnheader');
  (corner)&&corner.addEventListener('click', function(){ sel={r1:1,c1:1,r2:ROWS,c2:COLS}; setActive(1,1,true); drawSelection(); updateStatusSelection(); });
  trh.appendChild(corner);
  for(var c=1;c<=COLS;c++){
    var th=document.createElement('th'); th.setAttribute('role','columnheader'); th.textContent=colLabel(c); th.setAttribute('data-c', c);
    (th)&&th.addEventListener('click', (function(cc){ return function(){ sel={r1:1,c1:cc,r2:ROWS,c2:cc}; setActive(1, cc, true); drawSelection(); updateStatusSelection(); }; })(c));
    if (colWidths[c]) th.style.width = colWidths[c] + 'px';
    /* header sticky set later */
    var rez=document.createElement('div'); rez.className='col-resizer';
    (rez)&&rez.addEventListener('mousedown', (function(captureC, captureTh){ return function(e){ startColResize(e, captureC, captureTh); };})(c, th));
    th.appendChild(rez);
    trh.appendChild(th);
  }
  thead.appendChild(trh); table.appendChild(thead);

  // TBODY
  var tbody=document.createElement('tbody');
  for(var r=1;r<=ROWS;r++){
    var tr=document.createElement('tr');
    if(rowHeights[r]){ tr.style.height = rowHeights[r] + 'px'; }
    var rh=document.createElement('th'); rh.className='row-hdr'; rh.setAttribute('role','rowheader'); rh.setAttribute('scope','row'); rh.textContent=r; 
    (rh)&&rh.addEventListener('click', (function(rr){ return function(){ sel={r1:rr,c1:1,r2:rr,c2:COLS}; setActive(rr, 1, true); drawSelection(); updateStatusSelection(); }; })(r)); 
    /* row header sticky set later */
    tr.appendChild(rh);
    
    for(var c2=1;c2<=COLS;c2++){
      var td=document.createElement('td'); td.className='cell'; td.setAttribute('role','gridcell'); td.setAttribute('data-r', r); td.setAttribute('data-c', c2); td.setAttribute('aria-rowindex', String(r)); td.setAttribute('aria-colindex', String(c2));
      /* frozen styling applied after render */

      var m=mergeCoverAt(r,c2);
      if(m && !(m.r1===r && m.c1===c2)){ td.hidden=true; td.setAttribute("aria-hidden","true"); tr.appendChild(td); continue; }
      if(m && m.r1===r && m.c1===c2){ td.rowSpan=(m.r2-m.r1+1); td.colSpan=(m.c2-m.c1+1); }

      var input=document.createElement('input');
      var cellData = cells[r-1] && cells[r-1][c2-1] ? cells[r-1][c2-1] : {value:'', style:{}};
      input.value = cellData.value || '';
      input.setAttribute('aria-label', 'Cell '+colLabel(c2)+r);
      applyStyleToInput(input, cellData.style);
      (function(rr,cc){
        (input)&&input.addEventListener('focus', function(){ setActive(rr,cc,true); });
        
        // Whole‑cell undo/redo: live model on input, but snapshot/commit on blur or Enter/Tab
        (input)&&input.addEventListener('input', function(e){
          var next=e.target.value;
          cells[rr-1][cc-1].value=next;
        });
        (input)&&input.addEventListener('focus', function(){
          input.dataset.snapshot = (cells[rr-1][cc-1].value || '');
        });
        function commitCellEdit(){
          var prev = input.dataset.snapshot || '';
          var next = cells[rr-1][cc-1].value || '';
          if(prev !== next){
            pushUndo({type:'set', r:rr, c:cc, prev:prev, next:next});
            redoStack = [];
            input.dataset.snapshot = next;
          }
        }
        (input)&&input.addEventListener('blur', commitCellEdit);
        (input)&&input.addEventListener('keydown', function(ev){
          if(ev.key === 'Enter' || ev.key === 'Tab'){
            commitCellEdit();
          }
        });
(input)&&input.addEventListener('mousedown', function(e){
  if (e.button === 2) {
    var within = rr >= sel.r1 && rr <= sel.r2 && cc >= sel.c1 && cc <= sel.c2;
    if (within) { return; }
  }
  startSelection(e, rr, cc);
});
      })(r,c2);
      td.appendChild(input);

      var fh=document.createElement('div'); fh.className='fill-handle';
      (function(rr,cc){ (fh)&&fh.addEventListener('mousedown', function(e){ startFillDrag(e, rr, cc); }); })(r,c2);
      td.appendChild(fh);

      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  container.innerHTML=''; container.appendChild(table);
  applyFreezeOffsets();

// v20: hard-compact columns on phones
(function(){
  try{
    if(window.innerWidth <= 600){
      var table = container && container.querySelector('table.grid'); if(!table) return;
      for(var c=1;c<=COLS;c++){
        var th = table.querySelector('thead th[data-c="'+c+'"]');
        var current = colWidths[c] || (th ? Math.round(th.getBoundingClientRect().width) : 0);
        if(!colWidths[c] || current > 60){
          colWidths[c] = 36;
          if(th) th.style.width = '36px';
        }
      }
    }
  }catch(err){ console.error(err); }
})();

// Set compact default column widths on phones (first load only)
(function(){
  try{
    if(window.innerWidth <= 600){
      var table = container && container.querySelector('table.grid'); if(!table) return;
      for(var c=1; c<=COLS; c++){
        if(!colWidths[c]){
          var base = th ? th.getBoundingClientRect().width : 80; colWidths[c] = Math.max(36, Math.floor(base*0.2));
          var th = table.querySelector('thead th[data-c="'+c+'"]');
          if(th) th.style.width = colWidths[c] + 'px';
        }
      }
    }
  }catch(err){ console.error(err); }
})();

}

function rebuild(){ var p={r:active.r, c:active.c}; render(); applyFreezeOffsets();

// v20: hard-compact columns on phones
(function(){
  try{
    if(window.innerWidth <= 600){
      var table = container && container.querySelector('table.grid'); if(!table) return;
      for(var c=1;c<=COLS;c++){
        var th = table.querySelector('thead th[data-c="'+c+'"]');
        var current = colWidths[c] || (th ? Math.round(th.getBoundingClientRect().width) : 0);
        if(!colWidths[c] || current > 60){
          colWidths[c] = 36;
          if(th) th.style.width = '36px';
        }
      }
    }
  }catch(err){ console.error(err); }
})();

// Set compact default column widths on phones (first load only)
(function(){
  try{
    if(window.innerWidth <= 600){
      var table = container && container.querySelector('table.grid'); if(!table) return;
      for(var c=1; c<=COLS; c++){
        if(!colWidths[c]){
          var base = th ? th.getBoundingClientRect().width : 80; colWidths[c] = Math.max(36, Math.floor(base*0.2));
          var th = table.querySelector('thead th[data-c="'+c+'"]');
          if(th) th.style.width = colWidths[c] + 'px';
        }
      }
    }
  }catch(err){ console.error(err); }
})();
 setActive(p.r,p.c); drawSelection(); }
function applyFreezeOffsets(){
  if(!container) return;
  var table = container.querySelector('table.grid'); if(!table) return;
  var thead = table.querySelector('thead'); var cornerTh = thead ? thead.querySelector('th:first-child') : null;
  // Measure after layout
  var headerHeight = thead ? thead.getBoundingClientRect().height : 0;
  var cornerWidth = cornerTh ? cornerTh.getBoundingClientRect().width : 0;

  // Clear old inline styles
  var allFrozen = table.querySelectorAll('.frozen');
  for(var i=0;i<allFrozen.length;i++){ allFrozen[i].classList.remove('frozen'); allFrozen[i].style.top=''; allFrozen[i].style.left=''; allFrozen[i].style.zIndex=''; allFrozen[i].style.position=''; }
  var headThs = thead ? thead.querySelectorAll('th[data-c]') : [];
  for(var j=0;j<headThs.length;j++){ headThs[j].style.left=''; headThs[j].style.position=''; headThs[j].style.zIndex=''; }

  // Always keep column headers sticky at top 0
  if (thead){ var ths = thead.querySelectorAll('th'); for(var i=0;i<ths.length;i++){ ths[i].style.position='sticky'; ths[i].style.top='0px'; ths[i].style.zIndex='9'; ths[i].style.background='#fff'; } }
  // Apply header stickies for frozen columns
  if(frozenPane.col > 0){
    for(var c=1;c<=frozenPane.col;c++){
      var h = table.querySelector('thead th[data-c="'+c+'"]');
      if(h){ h.style.position='sticky'; h.style.left=cornerWidth+'px'; h.style.zIndex='8'; h.style.background='#fff'; }
      // Column cells
      var colCells = table.querySelectorAll('td[data-c="'+c+'"]');
      for(var k=0;k<colCells.length;k++){ var td=colCells[k]; td.classList.add('frozen'); td.style.position='sticky'; td.style.left=cornerWidth+'px'; td.style.zIndex='7'; td.style.background='#fff'; }
    }
  }

  // Apply row header stickies for frozen rows
  if(frozenPane.row > 0){
    for(var r=1;r<=frozenPane.row;r++){
      var rowHdr = table.querySelector('tbody tr:nth-child('+r+') th.row-hdr');
      if(rowHdr){ rowHdr.style.position='sticky'; rowHdr.style.top=headerHeight+'px'; rowHdr.style.zIndex='8'; rowHdr.style.background='#fff'; }
      // Row cells
      var rowCells = table.querySelectorAll('td[data-r="'+r+'"]');
      for(var m=0;m<rowCells.length;m++){ var tdr=rowCells[m]; tdr.classList.add('frozen'); tdr.style.position='sticky'; tdr.style.top=headerHeight+'px'; tdr.style.zIndex='7'; tdr.style.background='#fff'; }
    }
  }

  // Intersection: ensure both offsets and highest z-index
  if(frozenPane.row > 0 && frozenPane.col > 0){
    for(var r2=1;r2<=frozenPane.row;r2++){
      for(var c2=1;c2<=frozenPane.col;c2++){
        var inter = table.querySelector('td[data-r="'+r2+'"][data-c="'+c2+'"]');
        if(inter){ inter.style.left=cornerWidth+'px'; inter.style.top=headerHeight+'px'; inter.style.zIndex='9'; inter.style.background='#fff'; }
      }
    }
  }
}

/* ================== Selection ================== */
function startSelection(e, r, c){
  sel={ r1:r, c1:c, r2:r, c2:c };
  drawSelection(); updateStatusSelection();

  function move(ev){
    var target = ev.target && ev.target.closest ? ev.target.closest('.cell') : null;
    if(!target) return;
    sel.r2 = parseInt(target.getAttribute('data-r'), 10);
    sel.c2 = parseInt(target.getAttribute('data-c'), 10);
    drawSelection(); updateStatusSelection();
  }
  function up(){
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
  }
  (window)&&window.addEventListener('mousemove', move);
  (window)&&window.addEventListener('mouseup', up);
}

function drawSelection(){
  if(!container) return;
  var olds = container.querySelectorAll('.sel-rect');
  for (var i=0;i<olds.length;i++) olds[i].classList.remove('sel-rect');
  var R=normSel(sel);
  for(var r=R.r1;r<=R.r2;r++) for(var c=R.c1;c<=R.c2;c++){
    var td = cellTd(r,c); if (td) td.classList.add('sel-rect');
  }
}

/* ================== Fill & Series ================== */
function startFillDrag(e, r, c){
  e.stopPropagation(); e.preventDefault();
  var sr=normSel(sel);
  var isRowSeed=(sr.r1===sr.r2 && sr.c1!==sr.c2);
  var isColSeed=(sr.c1===sr.c2 && sr.r1!==sr.r2);
  var seeds=collectSeeds(sr,isRowSeed,isColSeed,r,c);
  var lastSig='';

  function move(ev){
    var target = ev.target && ev.target.closest ? ev.target.closest('.cell') : null; if(!target) return;
    var rr=parseInt(target.getAttribute('data-r'),10), cc=parseInt(target.getAttribute('data-c'),10);
    var dRow=Math.abs(rr-r), dCol=Math.abs(cc-c);
    var horizontal=dCol>dRow;
    var targetRect=computeFillTarget(sr,horizontal,rr,cc);
    var sig=targetRect.r1+','+targetRect.c1+','+targetRect.r2+','+targetRect.c2+','+horizontal;
    if(sig===lastSig) return; lastSig=sig;
    applyFill(targetRect,seeds,horizontal);
    sel=targetRect; drawSelection(); updateStatusSelection();
  }
  function up(){ window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); }
  (window)&&window.addEventListener('mousemove', move);
  (window)&&window.addEventListener('mouseup', up);
}

function collectSeeds(rect,isRowSeed,isColSeed,fr,fc){
  var vals, r, c;
  if(isRowSeed){ vals=[]; for(c=rect.c1;c<=rect.c2;c++) vals.push(cells[rect.r1-1][c-1].value||''); return {type:'row', values:vals}; }
  if(isColSeed){ vals=[]; for(r=rect.r1;r<=rect.r2;r++) vals.push(cells[r-1][rect.c1-1].value||''); return {type:'col', values:vals}; }
  return { type:'single', values:[(cells[fr-1][fc-1].value||'')] };
}

function computeFillTarget(seedRect,horizontal,rr,cc){
  var sr=normSel(seedRect);
  if(horizontal){
    var right=cc>=sr.c2; var c1=right? sr.c2+1:cc; var c2=right? cc:sr.c1-1;
    return { r1:sr.r1, c1:Math.min(c1,c2), r2:sr.r2, c2:Math.max(c1,c2) };
  } else {
    var down=rr>=sr.r2; var r1=down? sr.r2+1:rr; var r2=down? rr:sr.r1-1;
    return { r1:Math.min(r1,r2), c1:sr.c1, r2:Math.max(r1,r2), c2:sr.c2 };
  }
}

function detectSeries(vals){
  var i, nums=vals.map(parseMaybeNumber), numsOK=true;
  for(i=0;i<nums.length;i++){ if(!isFinite(nums[i])){ numsOK=false; break; } }
  if(numsOK && vals.length>=2){ var step=nums[nums.length-1]-nums[nums.length-2]; return { kind:'number', step:isFinite(step)?step:0 }; }
  var dates=vals.map(parseMaybeDate), datesOK=true;
  for(i=0;i<dates.length;i++){ if(!(dates[i] instanceof Date)){ datesOK=false; break; } }
  if(datesOK && vals.length>=2){ var stepMs=(+dates[dates.length-1])-(+dates[dates.length-2]); var stepDays=Math.round(stepMs/86400000)||1; return { kind:'date', stepDays: stepDays }; }
  return null;
}

function buildSeries(seedVals, series, count){
  var i, out, last;
  if(!series){ last=(seedVals[seedVals.length-1]||''); out=[]; for(i=0;i<count;i++) out.push(String(last)); return out; }
  if(series.kind==='number'){ var start=parseMaybeNumber(seedVals[seedVals.length-1])||0; out=[]; for(i=0;i<count;i++) out.push(String(start+series.step*(i+1))); return out; }
  if(series.kind==='date'){ var startDate=parseMaybeDate(seedVals[seedVals.length-1])||new Date(); out=[]; for(i=0;i<count;i++){ var d=new Date(+startDate+(series.stepDays*(i+1)*86400000)); out.push(isoDate(d)); } return out; }
  last=(seedVals[seedVals.length-1]||''); out=[]; for(i=0;i<count;i++) out.push(String(last)); return out;
}

function applyFill(target,seeds,horizontal){
  if(target.r2<target.r1 || target.c2<target.c1) return;
  var series=(seeds.values.length>=2)?detectSeries(seeds.values):null;
  var i,j, count, out;
  if(horizontal){
    for(var r=target.r1;r<=target.r2;r++){
      count=target.c2-target.c1+1; out=buildSeries(seeds.values.length?seeds.values:[''],series,count);
      for(i=0;i<count;i++) setCell(r, target.c1+i, out[i]);
    }
  } else {
    for(var c=target.c1;c<=target.c2;c++){
      count=target.r2-target.r1+1; out=buildSeries(seeds.values.length?seeds.values:[''],series,count);
      for(j=0;j<count;j++) setCell(target.r1+j, c, out[j]);
    }
  }
}

/* ================== Column Resize ================== */
function startColResize(e, col, th){
  e.preventDefault(); e.stopPropagation();
  var startX=e.clientX, startW=th.offsetWidth;
  function move(ev){ var w=Math.max(MIN_COL_WIDTH, startW+(ev.clientX-startX)); th.style.width=w+'px'; }
  function up(ev){
    var w=Math.max(MIN_COL_WIDTH, startW+(ev.clientX-startX)); 
    colWidths[col]=w; localStorage.setItem('colWidths', JSON.stringify(colWidths)); 
    window.removeEventListener('mousemove',move);
    window.removeEventListener('mouseup',up);
  }
  (window)&&window.addEventListener('mousemove', move);
  (window)&&window.addEventListener('mouseup', up);
}

/* ================== Clipboard & Context ================== */
function copySelectionTSV(){
  var R=normSel(sel), lines=[], r, c, row;
  for(r=R.r1;r<=R.r2;r++){ row=[]; for(c=R.c1;c<=R.c2;c++) row.push((cells[r-1][c-1].value||'')); lines.push(row.join('\t')); }
  var tsv=lines.join('\n');
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).catch(function(){ fallbackCopy(tsv); });
  } else { fallbackCopy(tsv); }
  function fallbackCopy(text){
    var ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
    document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta);
  }
}

function pasteTSVAtTopLeft(tsv){
  var R=normSel(sel);
  var matrix=tsv.replace(/\r/g,'').split('\n').map(function(line){ return line.split('\t'); });
  for(var i=0;i<matrix.length;i++) for(var j=0;j<matrix[i].length;j++){
    var r=R.r1+i, c=R.c1+j; if(r<=ROWS && c<=COLS) setCell(r,c,matrix[i][j]);
  }
  drawSelection();
}

/* ================== Undo/Redo ================== */
function pushUndo(entry){ undoStack.push(entry); }
function undo(){
  var last=undoStack.pop(); if(!last) return;
  if(last.type==='set'){
    var cur=cells[last.r-1][last.c-1].value||'';
    cells[last.r-1][last.c-1].value=last.prev||'';
    var inp=cellInput(last.r,last.c); if(inp) inp.value=last.prev||'';
    redoStack.push({ type:'set', r:last.r, c:last.c, prev:cur, next:last.prev||'' });
  } else if (last.type === 'style') {
    var curSt = JSON.parse(JSON.stringify(cells[last.r-1][last.c-1].style||{}));
    cells[last.r-1][last.c-1].style = JSON.parse(JSON.stringify(last.prev||{}));
    var inp2 = cellInput(last.r, last.c); if (inp2) applyStyleToInput(inp2, last.prev||{});
    redoStack.push({ type:'style', r:last.r, c:last.c, prev:curSt, next:last.prev||{} });
  }
  drawSelection();
}
function redo(){
  var last=redoStack.pop(); if(!last) return;
  if(last.type==='set'){
    var cur=cells[last.r-1][last.c-1].value||'';
    cells[last.r-1][last.c-1].value=last.next||'';
    var inp=cellInput(last.r,last.c); if(inp) inp.value=last.next||'';
    undoStack.push({ type:'set', r:last.r, c:last.c, prev:cur, next:last.next||'' });
  } else if (last.type === 'style') {
    var curSt = JSON.parse(JSON.stringify(cells[last.r-1][last.c-1].style||{}));
    cells[last.r-1][last.c-1].style = JSON.parse(JSON.stringify(last.next||{}));
    var inp2 = cellInput(last.r, last.c); if (inp2) applyStyleToInput(inp2, last.next||{});
    undoStack.push({ type:'style', r:last.r, c:last.c, prev:curSt, next:last.next||{} });
  }
  drawSelection();
}

/* ================== Shortcuts (incl. arrow nav while input focused) ================== */
function attachGlobalShortcuts(){
  (document)&&document.addEventListener('keydown', function(e){
    var meta=e.ctrlKey||e.metaKey;
    if(meta && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if(meta && (e.key.toLowerCase()==='y' || (e.shiftKey && e.key.toLowerCase()==='z'))){ e.preventDefault(); redo(); }
    if(meta && e.key.toLowerCase()==='c'){ e.preventDefault(); copySelectionTSV(); }
    if(meta && e.key.toLowerCase()==='x'){ e.preventDefault(); copySelectionTSV(); deleteSelection(); }
    if(meta && e.key.toLowerCase()==='v'){ 
      e.preventDefault(); 
      if (navigator.clipboard && navigator.clipboard.readText){
        navigator.clipboard.readText().then(function(t){ if(t){ pasteTSVAtTopLeft(t);} }).catch(function(){});
      }
    }
    if(meta && e.key.toLowerCase()==='b'){ e.preventDefault(); toggleStyle('bold'); }
    if(meta && e.key.toLowerCase()==='i'){ e.preventDefault(); toggleStyle('italic'); }
    if(meta && e.key.toLowerCase()==='s'){ e.preventDefault(); saveGrid(); }
    if(meta && e.key.toLowerCase()==='o'){ e.preventDefault(); loadGrid(); }

    // Delete multi-cell selection even if input is focused; otherwise delete when not in a cell
    if((e.key==='Delete' || e.key==='Backspace') && !meta){
      var isMulti = (sel.r1!==sel.r2) || (sel.c1!==sel.c2);
      if (isMulti) { e.preventDefault(); deleteSelection(); return; }
      var inCell = document.activeElement && document.activeElement.closest ? document.activeElement.closest('.cell') : null;
      if(!inCell){ e.preventDefault(); deleteSelection(); }
    }

    // Tab/Enter movement from within an input
    if(e.key==='Tab' || e.key==='Enter'){
      var inInput = document.activeElement && document.activeElement.tagName==='INPUT' && document.activeElement.closest('.cell');
      if(inInput){
        e.preventDefault();
        var dir=(e.key==='Tab') ? (e.shiftKey?'left':'right') : (e.shiftKey?'up':'down');
        moveActive(dir,1,{wrapRows:false});
        return;
      }
    }

    // Arrow keys always move selection, even when input focused
    if(e.key==='ArrowRight' || e.key==='ArrowLeft' || e.key==='ArrowDown' || e.key==='ArrowUp'){
      e.preventDefault();
      if(e.key==='ArrowRight'){ setActive(active.r, clamp(active.c+1,1,COLS), false); }
      if(e.key==='ArrowLeft'){  setActive(active.r, clamp(active.c-1,1,COLS), false); }
      if(e.key==='ArrowDown'){  setActive(clamp(active.r+1,1,ROWS), active.c, false); }
      if(e.key==='ArrowUp'){    setActive(clamp(active.r-1,1,ROWS), active.c, false); }
      return;
    }
  });
}

/* ================== Public-ish API ================== */
function createGrid(hostEl, rows, cols){
  container=hostEl; ROWS=rows||1200; COLS=cols||52;
  var r,c;
  cells=[];
  for(r=0;r<ROWS;r++){ var row=[]; for(c=0;c<COLS;c++) row.push({value:'', style:{}}); cells.push(row); }
  merges=[]; undoStack=[]; redoStack=[];
  render(); attachGlobalShortcuts(); setActive(1,1); drawSelection();
  wireToolbar(); wireContextMenu(); wireExportImport(); wireMicrophoneControls(); wireHelpModal();
}

function clearGrid(){
  for(var r=1;r<=ROWS;r++) for(var c=1;c<=COLS;c++){ cells[r-1][c-1].value=''; var inp=cellInput(r,c); if(inp) inp.value=''; }
  undoStack=[]; redoStack=[]; merges=[]; drawSelection();
}

function getData(){ 
  var out=[], r, c; 
  for(r=0;r<ROWS;r++){ var row=[]; for(c=0;c<COLS;c++) row.push(cells[r][c].value||''); out.push(row); }
  return out; 
}

function setCell(r,c,val,opts){
  if(r<1||c<1||r>ROWS||c>COLS) return;
  opts = opts || {};
  var prev=cells[r-1][c-1].value||''; var next=(val==null?'':String(val));
  if(prev===next) return;
  cells[r-1][c-1].value=next;
  var inp=cellInput(r,c); if(inp && inp.value!==next) inp.value=next;
  if(!opts.skipUndo){ pushUndo({type:'set', r:r, c:c, prev:prev, next:next}); redoStack=[]; }
}

function getActive(){ return colLabel(active.c)+active.r; }

function setActive(r,c,keepSelection){
  active.r=clamp(r,1,ROWS); active.c=clamp(c,1,COLS);
  var actives = container ? container.querySelectorAll('.cell.active') : [];
  for (var i=0;i<actives.length;i++) actives[i].classList.remove('active');
  var td=cellTd(active.r,active.c); if (td) td.classList.add('active');
  if(!keepSelection){ sel={ r1:active.r, c1:active.c, r2:active.r, c2:active.c }; }
  var inp=td ? td.querySelector('input') : null; if(inp) inp.focus({preventScroll:true});
  var s=document.getElementById('status-cell'); if(s) s.textContent=getActive();
  drawSelection(); updateStatusSelection();
}

function moveActive(dir, steps, opts){
  steps = steps || 1; opts = opts || {wrapRows:true, blockWidth:26, blockHeight:100};
  dir=String(dir||'').toLowerCase(); var r=active.r, c=active.c;
  var BW=clamp(opts.blockWidth||COLS,1,COLS), startBC=Math.floor((c-1)/BW)*BW+1, endBC=Math.min(startBC+BW-1,COLS);
  var BH=clamp(opts.blockHeight||ROWS,1,ROWS), startBR=Math.floor((r-1)/BH)*BH+1, endBR=Math.min(startBR+BH-1,ROWS);
  if(dir==='right'){ c+=steps; if(opts.wrapRows) while(c>endBC){ c=startBC+(c-endBC-1); r=Math.min(ROWS,r+1);} else c=Math.min(COLS,c); }
  else if(dir==='left'){ c-=steps; if(opts.wrapRows) while(c<startBC){ c=endBC-(startBC-c-1); r=Math.max(1,r-1);} else c=Math.max(1,c); }
  else if(dir==='down'){ r+=steps; if(opts.wrapRows) while(r>endBR){ r=startBR+(r-endBR-1); c=Math.min(COLS,c+1);} else r=Math.min(ROWS,r); }
  else if(dir==='up'){ r-=steps; if(opts.wrapRows) while(r<startBR){ r=endBR-(startBR-r-1); c=Math.max(1,c-1);} else r=Math.max(1,r); }
  setActive(r,c);
}

function deleteSelection(){
  var R=normSel(sel);
  for(var r=R.r1;r<=R.r2;r++) for(var c=R.c1;c<=R.c2;c++) setCell(r,c,'');
}

/* ================== Row/Col & Merge ================== */
function insertRow(where){
  where = where || 'below';
  var r=active.r;
  var row=[]; for(var i=0;i<COLS;i++) row.push({value:'',style:{}});
  var idx=(where==='above')?(r-1):r;
  cells.splice(idx,0,row); ROWS=cells.length; rebuild(); setActive((where==='above'?r:r+1), active.c, false);
}
function insertCol(where){
  where = where || 'right';
  var c=active.c, idx=(where==='left')?(c-1):c;
  for(var r=0;r<ROWS;r++) cells[r].splice(idx,0,{value:'',style:{}});
  COLS=cells[0].length; rebuild(); setActive(active.r,(where==='left'?c:c+1), false);
}
function deleteRow(){
  if(ROWS<=1) return;
  var del=active.r;
  cells.splice(del-1,1); ROWS=cells.length;
  merges = merges.filter(function(m){ return !(del>=m.r1 && del<=m.r2); });
  rebuild(); setActive(Math.min(del,ROWS), active.c, false);
}
function deleteCol(){
  if(COLS<=1) return;
  var del=active.c;
  for(var r=0;r<ROWS;r++) cells[r].splice(del-1,1);
  COLS=cells[0].length;
  merges = merges.filter(function(m){ return !(del>=m.c1 && del<=m.c2); });
  rebuild(); setActive(active.r, Math.min(del,COLS), false);
}
function moveRow(dir){
  var r=active.r, to=(dir==='up')?Math.max(1,r-1):Math.min(ROWS,r+1); if(to===r) return;
  var row=cells.splice(r-1,1)[0]; cells.splice(to-1,0,row); rebuild(); setActive(to,active.c, false);
}
function moveCol(dir){
  var c=active.c, to=(dir==='left')?Math.max(1,c-1):Math.min(COLS,c+1); if(to===c) return;
  for(var r=0;r<ROWS;r++){ var val=cells[r].splice(c-1,1)[0]; cells[r].splice(to-1,0,val); }
  rebuild(); setActive(active.r,to, false);
}
function mergeSelection(){
  var R=normSel(sel); var filled=0, r, c;
  for(r=R.r1;r<=R.r2;r++) for(c=R.c1;c<=R.c2;c++) if(cells[r-1][c-1].value) filled++;
  if(filled>1 && !confirm('Merging keeps only the top-left value. Continue?')) return;
  for(r=R.r1;r<=R.r2;r++) for(c=R.c1;c<=R.c2;c++) if(!(r===R.r1&&c===R.c1)) setCell(r,c,'');
  merges.push(R); rebuild(); setActive(R.r1,R.c1, true);
}
/* Unmerge any merged range that intersects the selection */
function unmergeSelection(){ 
  var R=normSel(sel); 
  var changed=false; 
  merges = merges.filter(function(m){ 
    var noOverlap = (R.r2 < m.r1) || (R.r1 > m.r2) || (R.c2 < m.c1) || (R.c1 > m.c2);
    if (!noOverlap) { changed=true; return false; }
    return true;
  }); 
  if(changed){ rebuild(); }
}

/* ================== Cell Formatting ================== */
function toggleStyle(prop){
  var R = normSel(sel);
  var current = (cells[R.r1-1][R.c1-1].style && cells[R.r1-1][R.c1-1].style[prop]) || false;
  for(var r=R.r1; r<=R.r2; r++) for(var c=R.c1; c<=R.c2; c++){
    if(!cells[r-1][c-1].style) cells[r-1][c-1].style = {};
    var prev = JSON.parse(JSON.stringify(cells[r-1][c-1].style));
    cells[r-1][c-1].style[prop] = !current;
    var inp = cellInput(r, c); if(inp) applyStyleToInput(inp, cells[r-1][c-1].style);
    pushUndo({type: 'style', r:r, c:c, prev: prev, next: JSON.parse(JSON.stringify(cells[r-1][c-1].style))});
  }
  redoStack = [];
}
function applyStyleToSelection(prop, value){
  var R = normSel(sel);
  for(var r=R.r1; r<=R.r2; r++) for(var c=R.c1; c<=R.c2; c++){
    if(!cells[r-1][c-1].style) cells[r-1][c-1].style = {};
    var prev = JSON.parse(JSON.stringify(cells[r-1][c-1].style));
    cells[r-1][c-1].style[prop] = value;
    var inp = cellInput(r, c); if(inp) applyStyleToInput(inp, cells[r-1][c-1].style);
    pushUndo({type: 'style', r:r, c:c, prev: prev, next: JSON.parse(JSON.stringify(cells[r-1][c-1].style))});
  }
  redoStack = [];
}

/* ================== Freeze Panes ================== */

function freezeTopRow(){
  frozenPane.row = 1;
  var mode = document.getElementById('status-mode'); if (mode) mode.textContent = 'Frozen: Top row';
  rebuild();
}
function freezeFirstCol(){
  frozenPane.col = 1;
  var mode = document.getElementById('status-mode'); if (mode) mode.textContent = 'Frozen: First column';
  rebuild();
}

function toggleFreezePanes(){
  if(frozenPane.row === active.r-1 && frozenPane.col === active.c-1) {
    frozenPane = {row: 0, col: 0};
    var mode = document.getElementById('status-mode'); if (mode) mode.textContent = 'Ready';
  } else {
    frozenPane = {row: active.r-1, col: active.c-1};
    var mode2 = document.getElementById('status-mode'); if (mode2) mode2.textContent = 'Frozen at '+colLabel(active.c)+active.r;
  }
  rebuild();
}

/* ================== Toolbar, Context, Import/Export, Mic ================== */
function wireToolbar(){
  function $(id){ return document.getElementById(id); }
  var el;
  if (el=$('btn-insert-row-above')) (el)&&el.addEventListener('click',function(){insertRow('above');});
  if (el=$('btn-insert-row-below')) (el)&&el.addEventListener('click',function(){insertRow('below');});
  if (el=$('btn-delete-row')) (el)&&el.addEventListener('click',deleteRow);
  if (el=$('btn-move-row-up')) (el)&&el.addEventListener('click',function(){moveRow('up');});
  if (el=$('btn-move-row-down')) (el)&&el.addEventListener('click',function(){moveRow('down');});

  if (el=$('btn-insert-col-left')) (el)&&el.addEventListener('click',function(){insertCol('left');});
  if (el=$('btn-insert-col-right')) (el)&&el.addEventListener('click',function(){insertCol('right');});
  if (el=$('btn-delete-col')) (el)&&el.addEventListener('click',deleteCol);
  if (el=$('btn-move-col-left')) (el)&&el.addEventListener('click',function(){moveCol('left');});
  if (el=$('btn-move-col-right')) (el)&&el.addEventListener('click',function(){moveCol('right');});

  if (el=$('btn-merge')) (el)&&el.addEventListener('click',mergeSelection);
  if (el=$('btn-unmerge')) (el)&&el.addEventListener('click',unmergeSelection);
  if (el=$('btn-autofit')) (el)&&el.addEventListener('click',function(){autoWidthFitSelected();});
  if (el=$('btn-freeze')) (el)&&el.addEventListener('click',toggleFreezePanes);
  if (el=$('btn-freeze-top')) (el)&&el.addEventListener('click',freezeTopRow);
  if (el=$('btn-freeze-firstcol')) (el)&&el.addEventListener('click',freezeFirstCol);

  if (el=$('btn-new-grid')) (el)&&el.addEventListener('click',function(){createGrid(container, ROWS, COLS);});
  if (el=$('btn-clear-grid')) (el)&&el.addEventListener('click',function(){ if(confirm('Clear ALL cells?')) clearGrid(); });

  if (el=$('btn-undo')) (el)&&el.addEventListener('click',undo);
  if (el=$('btn-redo')) (el)&&el.addEventListener('click',redo);
  if (el=$('btn-bold')) (el)&&el.addEventListener('click',function(){toggleStyle('bold');});
  if (el=$('btn-italic')) (el)&&el.addEventListener('click',function(){toggleStyle('italic');});
  if (el=$('btn-text-color')) { (el)&&el.addEventListener('change',function(e){applyStyleToSelection('color', e.target.value); drawSelection();}); (el)&&el.addEventListener('input',function(e){applyStyleToSelection('color', e.target.value); drawSelection();}); }
  if (el=$('btn-save')) (el)&&el.addEventListener('click',saveGrid);
  if (el=$('btn-load')) (el)&&el.addEventListener('click',loadGrid);

  // Ribbon toggle
  var toggle=document.getElementById('btn-toggle-ribbon');
  if(toggle){
    (toggle)&&toggle.addEventListener('click', function(){
      var rb=document.querySelector('.ribbon'); var hd=document.querySelector('header');
      if(!rb) return;
      if(rb.classList.contains('collapsed')){ rb.classList.remove('collapsed'); if(hd) hd.classList.remove('compact'); toggle.textContent='Hide Toolbar'; }
      else { rb.classList.add('collapsed'); if(hd) hd.classList.add('compact'); toggle.textContent='Show Toolbar'; }
    });
  }

}

function wireContextMenu(){
  var menu=document.getElementById('context-menu'); if(!menu) return;
  (document)&&document.addEventListener('contextmenu',function(e){
    var el = e.target && e.target.closest ? e.target.closest('.cell') : null;
    if (el) {
      var r = parseInt(el.getAttribute('data-r'),10);
      var c = parseInt(el.getAttribute('data-c'),10);
      var withinSel = r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
      if (!withinSel) { sel = {r1:r, c1:c, r2:r, c2:c}; setActive(r, c, false); }
      else { setActive(r, c, true); } // keep selection
      drawSelection();
      updateStatusSelection();
    }
    var cell = e.target && e.target.closest ? e.target.closest('.cell') : null; if(!cell) return;
    e.preventDefault(); menu.style.display='block';
    var pad=6; var W=menu.offsetWidth||180; var H=menu.offsetHeight||150;
    var x=Math.min(e.clientX, window.innerWidth-W-pad);
    var y=Math.min(e.clientY, window.innerHeight-H-pad);
    menu.style.left=x+'px'; menu.style.top=y+'px';
  });
  (document)&&document.addEventListener('click',function(e){ if(menu && !menu.contains(e.target)) menu.style.display='none'; });
  var cpy=document.getElementById('ctx-copy'); if(cpy) (cpy)&&cpy.addEventListener('click', function(){ copySelectionTSV(); menu.style.display='none'; });
  var cut=document.getElementById('ctx-cut'); if(cut) (cut)&&cut.addEventListener('click', function(){ copySelectionTSV(); deleteSelection(); menu.style.display='none'; });
  var pst=document.getElementById('ctx-paste'); if(pst) (pst)&&pst.addEventListener('click', function(){ 
    if (navigator.clipboard && navigator.clipboard.readText){
      navigator.clipboard.readText().then(function(t){ if(t) pasteTSVAtTopLeft(t); }).catch(function(){});
    }
    menu.style.display='none';
  });
  var del=document.getElementById('ctx-delete'); if(del) (del)&&del.addEventListener('click', function(){ deleteSelection(); menu.style.display='none'; });
  var bld=document.getElementById('ctx-bold'); if(bld) (bld)&&bld.addEventListener('click', function(){ toggleStyle('bold'); menu.style.display='none'; });
  var itc=document.getElementById('ctx-italic'); if(itc) (itc)&&itc.addEventListener('click', function(){ toggleStyle('italic'); menu.style.display='none'; });
}

function filename(base, ext){
  var d=new Date(), p=function(n){return String(n).padStart(2,'0');};
  return base+'_'+d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+'_'+p(d.getHours())+'-'+p(d.getMinutes())+'.'+ext;
}

function matrixForExport(){
  var all=getData(); var R=normSel(sel);
  if(R.r1===R.r2 && R.c1===R.c2) return all;
  var out=[], r; for(r=R.r1;r<=R.r2;r++) out.push(all[r-1].slice(R.c1-1, R.c2));
  return out;
}

function wireExportImport(){
  var exX=document.getElementById('export-xlsx'); if(exX) (exX)&&exX.addEventListener('click',function(e){ e.preventDefault();
    if(typeof XLSX==='undefined'){ alert('XLSX library missing'); return; }
    var ws=XLSX.utils.aoa_to_sheet(matrixForExport()); var wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1'); XLSX.writeFile(wb, filename('VoiceSheet','xlsx'));
  });
  var exC=document.getElementById('export-csv'); if(exC) (exC)&&exC.addEventListener('click',function(e){ e.preventDefault();
    if(typeof XLSX==='undefined'){ alert('XLSX library missing'); return; }
    var ws=XLSX.utils.aoa_to_sheet(matrixForExport()); var csv=XLSX.utils.sheet_to_csv(ws);
    var blob=new Blob([csv],{type:'text/csv;charset=utf-8'}); var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','csv'); a.click();
  });
  var exT=document.getElementById('export-tsv'); if(exT) (exT)&&exT.addEventListener('click',function(e){ e.preventDefault();
    if(typeof XLSX==='undefined'){ alert('XLSX library missing'); return; }
    var ws=XLSX.utils.aoa_to_sheet(matrixForExport()); var tsv=XLSX.utils.sheet_to_csv(ws,{FS:'\t'});
    var blob=new Blob([tsv],{type:'text/tab-separated-values;charset=utf-8'}); var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','tsv'); a.click();
  });

  function pickFile(accept){ return new Promise(function(res){ var inp=document.createElement('input'); inp.type='file'; inp.accept=accept; inp.onchange=function(){ res(inp.files && inp.files[0]); }; inp.click(); }); }
  function activeRC(){ return {r:active.r, c:active.c}; }
  function fillFromMatrixAt(matrix,r0,c0){ for(var i=0;i<matrix.length;i++) for(var j=0;j<matrix[i].length;j++) setCell(r0+i,c0+j,matrix[i][j]); }

  var imX=document.getElementById('import-xlsx'); if(imX) (imX)&&imX.addEventListener('click', function(e){
    e.preventDefault(); pickFile('.xlsx').then(function(f){ if(!f) return;
      f.arrayBuffer().then(function(buf){
        var wb=XLSX.read(buf,{type:'array'}); var ws=wb.Sheets[wb.SheetNames[0]];
        var matrix=XLSX.utils.sheet_to_json(ws,{header:1});
        var pos=activeRC(); fillFromMatrixAt(matrix,pos.r,pos.c);
      });
    });
  });
  var imC=document.getElementById('import-csv'); if(imC) (imC)&&imC.addEventListener('click', function(e){
    e.preventDefault(); pickFile('.csv').then(function(f){ if(!f) return;
      f.text().then(function(text){
        var matrix=text.replace(/\r/g,'').split('\n').filter(Boolean).map(function(l){return l.split(',');});
        var pos=activeRC(); fillFromMatrixAt(matrix,pos.r,pos.c);
      });
    });
  });
  var imT=document.getElementById('import-tsv'); if(imT) (imT)&&imT.addEventListener('click', function(e){
    e.preventDefault(); pickFile('.tsv,.txt').then(function(f){ if(!f) return;
      f.text().then(function(text){
        var matrix=text.replace(/\r/g,'').split('\n').filter(Boolean).map(function(l){return l.split('\t');});
        var pos=activeRC(); fillFromMatrixAt(matrix,pos.r,pos.c);
      });
    });
  });
}

function saveGrid(){
  var data = { cells:cells, merges:merges, colWidths:colWidths, frozenPane:frozenPane, active:active, sel:sel };
  localStorage.setItem('voiceSheetData', JSON.stringify(data));
  var mode = document.getElementById('status-mode'); if (mode) mode.textContent = 'Saved';
  setTimeout(function(){ var m=document.getElementById('status-mode'); if(m) m.textContent='Ready'; }, 2000);
}

function loadGrid(){
  var saved = localStorage.getItem('voiceSheetData');
  if (!saved) { alert('No saved data found'); return; }
  try {
    var data = JSON.parse(saved);
    cells = data.cells || cells;
    merges = data.merges || [];
    colWidths = data.colWidths || [];
    frozenPane = data.frozenPane || {row: 0, col: 0};
    active = data.active || {r:1,c:1};
    sel = data.sel || {r1:1,c1:1,r2:1,c2:1};
    ROWS = cells.length;
    COLS = (cells[0] && cells[0].length) ? cells[0].length : 26;
    rebuild();
    var mode = document.getElementById('status-mode'); if (mode) mode.textContent = 'Loaded';
    setTimeout(function(){ var m=document.getElementById('status-mode'); if(m) m.textContent='Ready'; }, 2000);
  } catch (e) { alert('Error loading saved data'); console.error(e); }
}

/* ================== Microphone (onresult-only auto-advance) ================== */


/* wireMicrophoneControls removed (kept newer version) */

;

/* ================== Auto-fit ================== */
function autoWidthFitSelected(col){
  var c=col||active.c, ctx=document.createElement('canvas').getContext('2d'), body=getComputedStyle(document.body);
  ctx.font=body.fontSize+' '+body.fontFamily;
  var max=MIN_COL_WIDTH;
  for(var r=1;r<=ROWS;r++){
    var v=((cells[r-1][c-1] && cells[r-1][c-1].value) ? cells[r-1][c-1].value : '')+'  ';
    var w=ctx.measureText(v).width + 18;
    if (w>max) max=w;
  }
  var th=container ? container.querySelector('thead th[data-c="'+c+'"]') : null; if(th) th.style.width=max+'px';
  colWidths[c]=max; localStorage.setItem('colWidths', JSON.stringify(colWidths));
}

/* ================== Help Modal ================== */
function wireHelpModal() {
  var modal = document.getElementById('help-modal');
  var btn = document.getElementById('btn-help');
  var span = document.getElementById('help-close');
  if (btn) (btn)&&btn.addEventListener('click', function(){ if (modal) modal.style.display = 'block'; });
  if (span) (span)&&span.addEventListener('click', function(){ if (modal) modal.style.display = 'none'; });
  (window)&&window.addEventListener('click', function(event){ if (modal && event.target === modal) modal.style.display = 'none'; });
}

/* ================== Boot ================== */
createGrid(document.getElementById('grid-container'), 1200, 52);
var scLeft=document.getElementById('scroll-left'); if (scLeft) (scLeft)&&scLeft.addEventListener('click', function(){ var m=document.querySelector('main'); if(m) m.scrollBy({ left: -200, behavior: 'smooth' }); });
var scRight=document.getElementById('scroll-right'); if (scRight) (scRight)&&scRight.addEventListener('click', function(){ var m=document.querySelector('main'); if(m) m.scrollBy({ left: 200, behavior: 'smooth' }); });

// Long-press context menu for touch
(function(){
  var timer=null, startX=0, startY=0;
  function onTouchStart(e){
    var cell = e.target.closest ? e.target.closest('.cell') : null;
    if(!cell) return;
    startX = e.touches[0].clientX; startY = e.touches[0].clientY;
    timer = setTimeout(function(){
      // simulate context menu open
      var evt = new MouseEvent('contextmenu', {clientX:startX, clientY:startY, bubbles:true, cancelable:true});
      cell.dispatchEvent(evt);
    }, 550);
  }
  function onTouchMove(e){
    if(!timer) return;
    var dx = Math.abs(e.touches[0].clientX - startX);
    var dy = Math.abs(e.touches[0].clientY - startY);
    if(dx > 10 || dy > 10){ clearTimeout(timer); timer=null; }
  }
  function onTouchEnd(){ if(timer){ clearTimeout(timer); timer=null; } }
  (document)&&document.addEventListener('touchstart', onTouchStart, {passive:true});
  (document)&&document.addEventListener('touchmove', onTouchMove, {passive:true});
  (document)&&document.addEventListener('touchend', onTouchEnd, {passive:true});
})();

// iOS mode mic tip
(function(){
  var SR=window.SpeechRecognition||window.webkitSpeechRecognition;
  if(!SR){
    var micBtn=document.getElementById('micBtn');
    if(micBtn){ micBtn.textContent='🎤 iOS: use keyboard dictation'; micBtn.disabled=true; micBtn.title='Web Speech recognition is not available on iOS Safari/Chrome'; }
    var mode=document.getElementById('status-mode');
    if(mode){ mode.textContent='Tip: On iPhone, tap a cell and use the keyboard mic to dictate.'; setTimeout(function(){ mode.textContent='Ready'; }, 6000); }
  }
})();

// Auto-collapse toolbar on small screens
(function(){
  function collapseIfSmall(){
    try{
      if(window.innerWidth <= 600){
        var rb=document.querySelector('.ribbon'); var hd=document.querySelector('header');
        if(rb && !rb.classList.contains('collapsed')){ rb.classList.add('collapsed'); if(hd) hd.classList.add('compact'); }
        var t=document.getElementById('btn-toggle-ribbon'); if(t) t.textContent='Show Toolbar';
      }
    }catch(err){ console.error(err); }
  }
  (window)&&window.addEventListener('load', collapseIfSmall);
  (window)&&window.addEventListener('orientationchange', collapseIfSmall);
})();

// Always-visible header toggle
(function(){
  var btn = document.getElementById('toolbar-toggle-min');
  if(!btn) return;
  function setLabel(){
    var rb=document.querySelector('.ribbon');
    if(rb && rb.classList.contains('collapsed')) btn.textContent='☰ Show';
    else btn.textContent='☰ Hide';
  }
  (btn)&&btn.addEventListener('click', function(){
    var rb=document.querySelector('.ribbon'); var hd=document.querySelector('header');
    if(!rb) return;
    if(rb.classList.contains('collapsed')){ rb.classList.remove('collapsed'); if(hd) hd.classList.remove('compact'); }
    else { rb.classList.add('collapsed'); if(hd) hd.classList.add('compact'); }
    setLabel();
  });
  (window)&&window.addEventListener('load', setLabel);
})();

// Touch long-press & drag to extend selection; release to open context menu
(function(){
  var pressTimer=null, selecting=false, origin=null, lastClient={x:0,y:0};
  function cellFromTouch(t){
    var el = document.elementFromPoint(t.clientX, t.clientY);
    return el && el.closest ? el.closest('.cell') : null;
  }
  function toRC(cell){
    return cell ? { r: parseInt(cell.getAttribute('data-r'),10), c: parseInt(cell.getAttribute('data-c'),10) } : null;
  }
  function onStart(e){
    var t = e.touches && e.touches[0]; if(!t) return;
    var cell = cellFromTouch(t); if(!cell) return;
    lastClient = {x:t.clientX, y:t.clientY};
    origin = toRC(cell);
    pressTimer = setTimeout(function(){
      selecting = true;
      sel = { r1:origin.r, c1:origin.c, r2:origin.r, c2:origin.c };
      setActive(origin.r, origin.c, true);
      drawSelection(); updateStatusSelection();
      try{ e.preventDefault(); }catch(err){ console.error(err); }
    }, 450);
  }
  function onMove(e){
    var t = e.touches && e.touches[0]; if(!t) return;
    lastClient = {x:t.clientX, y:t.clientY};
    if(selecting){
      var cell = cellFromTouch(t); if(!cell) return;
      var rc = toRC(cell);
      sel.r2 = rc.r; sel.c2 = rc.c;
      drawSelection(); updateStatusSelection();
      try{ e.preventDefault(); }catch(err){ console.error(err); }
    }else if(pressTimer){
      // cancel long-press if we move too far
      var dx = Math.abs(t.clientX - lastClient.x), dy = Math.abs(t.clientY - lastClient.y);
      if(dx>12 || dy>12){ clearTimeout(pressTimer); pressTimer=null; }
    }
  }
  function onEnd(e){
    clearTimeout(pressTimer); pressTimer=null;
    if(selecting){
      selecting=false;
      // Open context menu at touch end
      var menuEvt = new MouseEvent('contextmenu', {clientX:lastClient.x, clientY:lastClient.y, bubbles:true, cancelable:true});
      var cell = cellFromTouch({clientX:lastClient.x, clientY:lastClient.y});
      if(cell){ cell.dispatchEvent(menuEvt); }
    }
  }
  (document)&&document.addEventListener('touchstart', onStart, {passive:false});
  (document)&&document.addEventListener('touchmove', onMove, {passive:false});
  (document)&&document.addEventListener('touchend', onEnd, {passive:false});
})();

// Long-press header to resize column/row (mobile friendly)
(function(){
  var panel = document.getElementById('resize-panel');
  var range = document.getElementById('resize-range');
  var label = document.getElementById('resize-label');
  var btnApply = document.getElementById('resize-apply');
  var btnCancel = document.getElementById('resize-cancel');
  var context = null; // {type:'col'|'row', index:number, th|tr|rh}
  function openPanel(cfg){
    context = cfg;
    if(cfg.type==='col'){
      label.textContent = 'Column '+colLabel(cfg.index)+' width';
      var th = document.querySelector('thead th[data-c="'+cfg.index+'"]');
      var w = th ? th.offsetWidth : (colWidths[cfg.index]||80);
      range.min = 32; range.max = 320; range.value = Math.max(32, Math.min(320, w));
    } else {
      label.textContent = 'Row '+cfg.index+' height';
      var tr = document.querySelector('tbody tr:nth-child('+cfg.index+')');
      var h = tr ? tr.offsetHeight : (rowHeights[cfg.index]||ROW_HEIGHT);
      range.min = 20; range.max = 80; range.value = Math.max(20, Math.min(80, h));
    }
    panel.style.display='block';
  }
  function closePanel(){ panel.style.display='none'; context=null; }
  (btnCancel)&&btnCancel.addEventListener('click', closePanel);
  (btnApply)&&btnApply.addEventListener('click', function(){
    if(!context) return; var v = parseInt(range.value,10);
    if(context.type==='col'){
      var th = document.querySelector('thead th[data-c="'+context.index+'"]');
      if(th){ th.style.width = v+'px'; }
      colWidths[context.index] = v; localStorage.setItem('colWidths', JSON.stringify(colWidths));
    } else {
      var tr = document.querySelector('tbody tr:nth-child('+context.index+')');
      if(tr){ tr.style.height = v+'px'; }
      rowHeights[context.index] = v; localStorage.setItem('rowHeights', JSON.stringify(rowHeights));
    }
    closePanel();
  });

  var pressTimer=null, startXY=null;
  function bindHeaderLongPress(){
    var thead = document.querySelector('thead'); if(!thead) return;
    (thead)&&thead.addEventListener('touchstart', function(e){
      var th = e.target.closest && e.target.closest('th[data-c]'); if(!th) return;
      var c = parseInt(th.getAttribute('data-c'),10);
      startXY = {x:e.touches[0].clientX, y:e.touches[0].clientY};
      pressTimer = setTimeout(function(){ openPanel({type:'col', index:c}); }, 450);
    }, {passive:true});
    (thead)&&thead.addEventListener('touchmove', function(e){
      if(!pressTimer) return;
      var dx = Math.abs(e.touches[0].clientX - startXY.x);
      var dy = Math.abs(e.touches[0].clientY - startXY.y);
      if(dx>12 || dy>12){ clearTimeout(pressTimer); pressTimer=null; }
    }, {passive:true});
    (thead)&&thead.addEventListener('touchend', function(){ if(pressTimer){ clearTimeout(pressTimer); pressTimer=null; } }, {passive:true});

    var tbody = document.querySelector('tbody'); if(!tbody) return;
    (tbody)&&tbody.addEventListener('touchstart', function(e){
      var rh = e.target.closest && e.target.closest('th.row-hdr'); if(!rh) return;
      var rr = parseInt(rh.textContent,10);
      startXY = {x:e.touches[0].clientX, y:e.touches[0].clientY};
      pressTimer = setTimeout(function(){ openPanel({type:'row', index:rr}); }, 450);
    }, {passive:true});
    (tbody)&&tbody.addEventListener('touchmove', function(e){
      if(!pressTimer) return;
      var dx = Math.abs(e.touches[0].clientX - startXY.x);
      var dy = Math.abs(e.touches[0].clientY - startXY.y);
      if(dx>12 || dy>12){ clearTimeout(pressTimer); pressTimer=null; }
    }, {passive:true});
    (tbody)&&tbody.addEventListener('touchend', function(){ if(pressTimer){ clearTimeout(pressTimer); pressTimer=null; } }, {passive:true});
  }
  (window)&&window.addEventListener('load', bindHeaderLongPress);
  (document)&&document.addEventListener('readystatechange', function(){ if(document.readyState==='complete') bindHeaderLongPress(); });
})();

// View controls
(function(){
  var fontSel = document.getElementById('font-size');
  var zoom = document.getElementById('zoom-slider');
  var darkBtn = document.getElementById('btn-dark');
  function applyFont(){ document.documentElement.style.setProperty('--app-font-size', fontSel.value); var inputs = container ? container.querySelectorAll('.cell input') : []; for(var i=0;i<inputs.length;i++){ inputs[i].style.fontSize = fontSel.value; } }
  function applyZoom(){ var z = parseFloat(zoom.value)||1; document.documentElement.style.setProperty('--zoom', z); var rows = container? container.querySelectorAll('tbody tr'):[]; for(var i=0;i<rows.length;i++){ var rh = (rowHeights[i+1]||ROW_HEIGHT)*z; rows[i].style.height = rh+'px'; } }
  (fontSel)&&fontSel.addEventListener('change', function(){ applyFont(); });
  (zoom)&&zoom.addEventListener('input', function(){ applyZoom(); });
  (darkBtn)&&darkBtn.addEventListener('click', function(){ document.body.classList.toggle('dark'); });
  // initialize
  (function(){ document.documentElement.style.setProperty('--app-font-size', fontSel.value); document.documentElement.style.setProperty('--zoom','1'); })();
})();

// Two-finger scroll lock (pan-x or pan-y)
(function(){
  var el = document.querySelector('main'); if(!el) return;
  var active=false, axis=null, last={x:0,y:0};
  (el)&&el.addEventListener('touchstart', function(e){
    if(e.touches.length===2){ active=true; axis=null; last={x:e.touches[0].clientX, y:e.touches[0].clientY}; }
  }, {passive:true});
  (el)&&el.addEventListener('touchmove', function(e){
    if(!active) return;
    var t=e.touches[0]; var dx=t.clientX-last.x, dy=t.clientY-last.y;
    if(axis===null){ axis = Math.abs(dx)>Math.abs(dy) ? 'x' : 'y'; }
    if(axis==='x'){ el.scrollLeft -= dx; } else { el.scrollTop -= dy; }
    last={x:t.clientX, y:t.clientY};
    e.preventDefault();
  }, {passive:false});
  (el)&&el.addEventListener('touchend', function(e){ if(e.touches.length===0){ active=false; axis=null; } }, {passive:true});
})();

// Touch support for fill-handle
(function(){
  (document)&&document.addEventListener('touchstart', function(e){
    var fh = e.target.closest ? e.target.closest('.fill-handle') : null; if(!fh) return;
    var cell = fh.parentElement; if(!cell) return;
    var r = parseInt(cell.getAttribute('data-r'),10), c = parseInt(cell.getAttribute('data-c'),10);
    startFillDrag({ preventDefault:function(){}, stopPropagation:function(){}, target:cell }, r, c);
  }, {passive:true});
})();

// Mini-toolbar behavior
(function(){
  var mt = document.getElementById('mini-toolbar'); if(!mt) return;
  function showAt(x,y){ mt.style.left=(x+6)+'px'; mt.style.top=(y-10)+'px'; mt.style.display='flex'; }
  function hide(){ mt.style.display='none'; }
  (document)&&document.addEventListener('click', function(e){ if(!mt.contains(e.target) && !e.target.closest('.cell')) hide(); });
  // Show on selection change
  (document)&&document.addEventListener('mouseup', function(e){ var td=e.target.closest?e.target.closest('.cell'):null; if(td){ showAt(e.clientX, e.clientY); } });
  (document)&&document.addEventListener('touchend', function(e){ var t=e.changedTouches && e.changedTouches[0]; if(t){ showAt(t.clientX, t.clientY); } }, {passive:true});
  document.getElementById('mini-copy').addEventListener('click', function(){ copySelectionTSV(); hide(); });
  document.getElementById('mini-cut').addEventListener('click', function(){ copySelectionTSV(); deleteSelection(); hide(); });
  document.getElementById('mini-delete').addEventListener('click', function(){ deleteSelection(); hide(); });
  document.getElementById('mini-paste').addEventListener('click', async function(){ try{ var t=await navigator.clipboard.readText(); if(t) pasteTSVAtTopLeft(t);}catch(err){ console.error(err); } hide(); });
  document.getElementById('mini-color').addEventListener('input', function(e){ applyStyleToSelection('color', e.target.value); });
})();

// Simple multi-sheet support
var sheets = [{ name:'Sheet 1', data:null }]; var currentSheet = 0;
function commitToSheet(){
  sheets[currentSheet].data = JSON.stringify({cells:cells, merges:merges, colWidths:colWidths, rowHeights:rowHeights, frozenPane:frozenPane});
}
function loadFromSheet(idx){
  try{
    currentSheet = idx;
    var d = sheets[idx].data ? JSON.parse(sheets[idx].data) : null;
    cells = d && d.cells ? d.cells : Array.from({length:ROWS},()=>Array.from({length:COLS},()=>({value:'',style:{}})));
    merges = d && d.merges ? d.merges : [];
    colWidths = d && d.colWidths ? d.colWidths : [];
    rowHeights = d && d.rowHeights ? d.rowHeights : [];
    frozenPane = d && d.frozenPane ? d.frozenPane : {row:0,col:0};
    rebuild();
  }catch(e){ console.error(e); }
}
function renderTabs(){
  var tabs=document.getElementById('sheets-tabs'); if(!tabs) return; tabs.innerHTML='';
  for(var i=0;i<sheets.length;i++){ (function(i){
    var b=document.createElement('button'); b.textContent=sheets[i].name; b.style.border='1px solid var(--border)'; b.style.borderRadius='8px'; b.style.padding='6px 10px';
    if(i===currentSheet){ b.style.background='#fff'; b.style.fontWeight='600'; }
    (b)&&b.addEventListener('click', function(){ commitToSheet(); loadFromSheet(i); renderTabs(); });
    tabs.appendChild(b);
  })(i); }
}
var addBtn = document.getElementById('btn-add-sheet'); if(addBtn){ (addBtn)&&addBtn.addEventListener('click', function(){ commitToSheet(); sheets.push({name:'Sheet '+(sheets.length+1), data:null}); renderTabs(); }); }
(window)&&window.addEventListener('load', renderTabs);



(function(){
  try{
    // Unified horizontal scroll: keep both #hscroll and #grid-hscroll in sync with the grid container.
    var grid = document.querySelector('#grid-container') || document.querySelector('#grid') || document.body;
    var slider1 = document.getElementById('hscroll');
    var slider2 = document.getElementById('grid-hscroll');

    function syncSliderMax(){
      if(!grid) return;
      var scrollEl = grid.querySelector('.grid') || grid;
      var max = Math.max(0, (scrollEl.scrollWidth||0) - (scrollEl.clientWidth||0));
      if(slider1){ slider1.max = String(max); }
      if(slider2){ slider2.max = String(max); }
    }
    function setScrollLeft(v){
      if(!grid) return;
      var scrollEl = grid.querySelector('.grid') || grid;
      scrollEl.scrollLeft = v;
      if(slider1 && slider1.value !== String(v)) slider1.value = String(v);
      if(slider2 && slider2.value !== String(v)) slider2.value = String(v);
    }
    function wireHScrollUnified(){
      var scrollEl = grid.querySelector('.grid') || grid;
      if(!scrollEl) return;
      syncSliderMax();
      if(slider1) slider1.addEventListener('input', function(){ setScrollLeft(Number(slider1.value||0)); });
      if(slider2) slider2.addEventListener('input', function(){ setScrollLeft(Number(slider2.value||0)); });
      scrollEl.addEventListener('scroll', function(){ 
        var v = scrollEl.scrollLeft||0;
        if(slider1 && slider1.value !== String(v)) slider1.value = String(v);
        if(slider2 && slider2.value !== String(v)) slider2.value = String(v);
      });
      window.addEventListener('resize', syncSliderMax);
    }
    // If a previous wireHScroll exists, let it run; then ensure ours also runs
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wireHScrollUnified);
    else wireHScrollUnified();

    // Microphone: if wireMicrophoneControls is defined, ensure it's called once.
    function safeWireMic(){
      if (typeof window.wireMicrophoneControls === 'function') {
        try{ window.wireMicrophoneControls(); }catch(e){ /* no-op */ }
      }
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', safeWireMic, { once:true });
    else safeWireMic();
  }catch(e){ /* swallow */ }
})();

(function(){
  function safetyBoot(){
    try{
      var host = document.getElementById('grid-container');
      if (!host) return;
      if (host.children && host.children.length > 0) return; // already rendered
      if (typeof createGrid === 'function'){
        createGrid(host, 1200, 52);
      }
    }catch(e){ /* swallow */ }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', safetyBoot, { once:true });
  else safetyBoot();
})();
