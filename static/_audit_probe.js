/* 体验审查探针：在真实浏览器里量尺寸、对比度、标签、溢出，不猜。
   结果 POST 给同源的 /__audit（5099 上的临时接收端）。 */
(function () {
  'use strict';

  function vis(el) {
    var st = getComputedStyle(el);
    if (st.display === 'none' || st.visibility === 'hidden' || parseFloat(st.opacity) === 0) return false;
    var r = el.getBoundingClientRect();
    return r.width > 0.5 && r.height > 0.5;
  }

  function desc(el) {
    var s = el.tagName.toLowerCase();
    if (el.id) s += '#' + el.id;
    var cls = (el.className || '').toString().trim().split(/\s+/).filter(Boolean).slice(0, 2);
    if (cls.length) s += '.' + cls.join('.');
    return s;
  }

  function txt(el) {
    var t = '';
    for (var i = 0; i < el.childNodes.length; i++) {
      var n = el.childNodes[i];
      if (n.nodeType === 3) t += n.nodeValue;
    }
    return t.replace(/\s+/g, ' ').trim();
  }

  function parseColor(c) {
    var m = /rgba?\(([^)]+)\)/.exec(c || '');
    if (!m) return null;
    var p = m[1].split(',').map(parseFloat);
    return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 };
  }

  function over(fg, bg) {
    var a = fg.a;
    return { r: fg.r * a + bg.r * (1 - a), g: fg.g * a + bg.g * (1 - a), b: fg.b * a + bg.b * (1 - a), a: 1 };
  }

  function effBg(el) {
    var stack = [], n = el;
    while (n && n.nodeType === 1) {
      var c = parseColor(getComputedStyle(n).backgroundColor);
      if (c && c.a > 0) stack.push(c);
      n = n.parentElement;
    }
    var acc = { r: 255, g: 255, b: 255, a: 1 };
    for (var i = stack.length - 1; i >= 0; i--) acc = over(stack[i], acc);
    return acc;
  }

  function lum(c) {
    function f(v) { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }
    return 0.2126 * f(c.r) + 0.7152 * f(c.g) + 0.0722 * f(c.b);
  }

  function ratio(a, b) {
    var l1 = lum(a), l2 = lum(b);
    if (l1 < l2) { var t = l1; l1 = l2; l2 = t; }
    return (l1 + 0.05) / (l2 + 0.05);
  }

  function run() {
    var out = {
      page: location.pathname,
      viewport: { w: innerWidth, h: innerHeight },
      scroll: { w: document.documentElement.scrollWidth, h: document.documentElement.scrollHeight },
      overflowX: document.documentElement.scrollWidth > innerWidth + 1,
    };

    var all = [].slice.call(document.querySelectorAll('body *')).filter(vis);

    /* ---- 1. 文字：字号分布 + 对比度 ---- */
    var sizes = {}, small = [], lowContrast = [];
    all.forEach(function (el) {
      var t = txt(el);
      if (!t || /^(SCRIPT|STYLE|NOSCRIPT)$/.test(el.tagName)) return;
      var st = getComputedStyle(el);
      var fs = parseFloat(st.fontSize);
      if (!fs) return;
      sizes[fs] = (sizes[fs] || 0) + 1;
      var color = parseColor(st.color);
      if (!color) return;
      var bg = effBg(el);
      var bgc = color.a < 1 ? over(color, bg) : color;
      var cr = ratio(bgc, bg);
      var bold = parseInt(st.fontWeight, 10) >= 600;
      // WCAG AA：正文 4.5，大字（>=18.66px 或 14px 加粗）3.0
      var need = (fs >= 18.66 || (bold && fs >= 14)) ? 3.0 : 4.5;
      var rec = { el: desc(el), text: t.slice(0, 24), fs: fs, cr: Math.round(cr * 100) / 100, need: need };
      if (cr < need) lowContrast.push(rec);
      if (fs < 12) small.push(rec);
    });
    out.fontSizes = sizes;
    out.smallText = small.slice(0, 40);
    out.smallTextCount = small.length;
    out.lowContrast = lowContrast.sort(function (a, b) { return a.cr - b.cr; }).slice(0, 40);
    out.lowContrastCount = lowContrast.length;

    /* ---- 2. 可点击目标尺寸 ---- */
    var interactive = [].slice.call(document.querySelectorAll(
      'a[href],button,select,input,textarea,summary,[role=button],[onclick]')).filter(vis);
    var tiny = [], smallT = [];
    interactive.forEach(function (el) {
      var r = el.getBoundingClientRect();
      var rec = { el: desc(el), text: (txt(el) || el.value || '').slice(0, 18), w: Math.round(r.width), h: Math.round(r.height) };
      if (r.height < 24 || r.width < 24) tiny.push(rec);
      else if (r.height < 32 || r.width < 32) smallT.push(rec);
    });
    out.interactiveCount = interactive.length;
    out.tapUnder24 = tiny.slice(0, 40);
    out.tapUnder24Count = tiny.length;
    out.tapUnder32 = smallT.slice(0, 40);
    out.tapUnder32Count = smallT.length;

    /* ---- 3. 表单可访问名 ---- */
    var unlabeled = [];
    [].slice.call(document.querySelectorAll('input,select,textarea')).filter(vis).forEach(function (el) {
      if (/hidden/i.test(el.type || '')) return;
      var name = el.getAttribute('aria-label') || el.getAttribute('title') || el.getAttribute('placeholder')
        || (el.labels && el.labels.length ? 'label' : '');
      if (!name) unlabeled.push(desc(el));
    });
    out.unlabeledFields = unlabeled;

    /* ---- 4. 图片 alt ---- */
    var noAlt = [].slice.call(document.querySelectorAll('img')).filter(vis)
      .filter(function (i) { return !i.hasAttribute('alt'); }).map(function (i) { return i.getAttribute('src'); });
    out.imgNoAlt = noAlt;

    /* ---- 5. 首屏认知负荷 ---- */
    var fold = all.filter(function (el) {
      var r = el.getBoundingClientRect();
      return r.top < innerHeight * 0.92 && r.bottom > 0;
    });
    var foldText = fold.map(txt).join(' ').replace(/\s+/g, ' ');
    out.foldChars = foldText.length;
    out.foldInteractive = fold.filter(function (el) {
      return el.matches('a[href],button,select,input,textarea,summary,[role=button],[onclick]');
    }).length;

    /* ---- 6. 结构信号 ---- */
    out.headings = [].slice.call(document.querySelectorAll('h1,h2,h3,h4')).filter(vis)
      .map(function (h) { return h.tagName + ':' + txt(h).slice(0, 20); });
    out.tabStops = [].slice.call(document.querySelectorAll(
      'a[href],button:not([disabled]),select:not([disabled]),input:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])'
    )).filter(vis).length;
    out.ariaLive = document.querySelectorAll('[role=status],[role=alert],[aria-live]').length;
    out.dialogs = document.querySelectorAll('dialog,[role=dialog]').length;

    /* ---- 7. 焦点样式是否声明 ---- */
    var focusRule = false;
    try {
      [].slice.call(document.styleSheets).forEach(function (ss) {
        var rules;
        try { rules = ss.cssRules; } catch (e) { return; }
        [].slice.call(rules || []).forEach(function (r) {
          if (r.selectorText && /:focus/.test(r.selectorText)) focusRule = true;
        });
      });
    } catch (e) { }
    out.hasFocusStyle = focusRule;

    /* ---- 8. 按钮清单（看 CTA 层级） ---- */
    out.buttons = [].slice.call(document.querySelectorAll('button,.btn,[role=button]')).filter(vis)
      .map(function (b) {
        var st = getComputedStyle(b);
        var r = b.getBoundingClientRect();
        return {
          el: desc(b), text: txt(b).slice(0, 22),
          w: Math.round(r.width), h: Math.round(r.height),
          bg: st.backgroundColor, color: st.color, fs: parseFloat(st.fontSize)
        };
      }).slice(0, 40);

    fetch('/__audit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(out)
    }).catch(function () { });
  }

  if (document.readyState === 'complete') setTimeout(run, 1500);
  else addEventListener('load', function () { setTimeout(run, 1500); });
})();
