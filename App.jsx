aimport { useState, useEffect, useRef, useCallback, useMemo, memo } from "react";
import * as XLSX from "xlsx";

// ===== KaTeX CDN loader =====
let katexLoaded = false;
function loadKatex(cb) {
  if (katexLoaded) { cb(); return; }
  if (window.katex) { katexLoaded = true; cb(); return; }
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = "https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.css";
  document.head.appendChild(link);
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.js";
  s.onload = () => { katexLoaded = true; cb(); };
  document.head.appendChild(s);
}
// Bir xil formula qayta-qayta hisoblanib, scroll paytida "qotish"ga sabab bo'lmasligi uchun natijani keshlaymiz
const katexRenderCache = new Map();
function renderKatexCached(content, displayMode) {
  const key = (displayMode ? "1:" : "0:") + content;
  if (katexRenderCache.has(key)) return katexRenderCache.get(key);
  let html = null;
  try { html = window.katex.renderToString(content, { throwOnError: false, displayMode }); }
  catch { html = null; }
  if (katexRenderCache.size > 2000) katexRenderCache.clear(); // haddan tashqari o'sib ketmasligi uchun
  katexRenderCache.set(key, html);
  return html;
}

// ===== Computer Modern (CMU Serif) web font loader =====
// LaTeX hujjat matnini asl LaTeX ko'rinishidagi CMU Serif shriftida chizish uchun
let cmuFontLoaded = false;
// Matn va formula BIR XIL shriftda ko'rinishi uchun KaTeX_Main'ni ustuvor qilamiz —
// u allaqachon katex.min.css orqali yuklanadi va formulalarda ishlatiladi, shu sababli
// matnda ham aynan shuni ishlatish ikkalasini bir xil qiladi (tashqi shrift zaxira sifatida qoladi)
const CMU_FONT_STACK = "KaTeX_Main, 'Computer Modern', 'CMU Serif', Georgia, serif";
function loadCmuFont() {
  if (cmuFontLoaded) return;
  cmuFontLoaded = true;
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = "https://fonts.cdnfonts.com/css/computer-modern";
  document.head.appendChild(link);
}

// ===== PDF.js loader =====
let pdfjsLoaded = false;
function loadPdfJs(cb) {
  if (pdfjsLoaded && window.pdfjsLib) { cb(); return; }
  if (window.pdfjsLib) { pdfjsLoaded = true; cb(); return; }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
  s.onload = () => {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    pdfjsLoaded = true;
    cb();
  };
  s.onerror = () => cb();
  document.head.appendChild(s);
}

function PdfViewer({ url, persistKey }) {
  const containerRef = useRef(null);
  const stateRef = useRef({ pdf:null, pages:[], zoom:1, cancelled:false, blobUrl:null, scrollTop:0 });
  const scrollRef = useRef(null);
  const [uiState, setUiState] = useState({ loading:true, error:null, zoom:1 });

  // Convert data: URL → Blob URL (works everywhere: Claude, CodeSandbox, Vercel)
  function dataUrlToBytes(dataUrl) {
    const b64 = dataUrl.split(",")[1] || "";
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  useEffect(() => {
    if (!url) return;
    const s = stateRef.current;
    s.cancelled = false;
    setUiState(u => ({...u, loading:true, error:null}));

    // Revoke previous blob URL to avoid memory leaks
    if (s.blobUrl) { URL.revokeObjectURL(s.blobUrl); s.blobUrl = null; }

    loadPdfJs(async () => {
      if (s.cancelled) return;
      try {
        let src;
        if (url.startsWith("data:")) {
          // Method 1: pass raw bytes directly to PDF.js (no fetch needed)
          const bytes = dataUrlToBytes(url);
          src = { data: bytes };
        } else {
          src = { url };
        }
        const pdf = await window.pdfjsLib.getDocument(src).promise;
        if (s.cancelled) return;
        s.pdf = pdf; s.pages = []; s.zoom = 1;
        await doRender(1);
        if (!s.cancelled) setUiState(u => ({...u, loading:false, zoom:1}));
      } catch(err) {
        console.error("[PDF error]", err);
        // Fallback: try creating a blob URL
        try {
          if (url.startsWith("data:")) {
            const bytes = dataUrlToBytes(url);
            const blob = new Blob([bytes], { type:"application/pdf" });
            const blobUrl = URL.createObjectURL(blob);
            s.blobUrl = blobUrl;
            const pdf2 = await window.pdfjsLib.getDocument({ url: blobUrl }).promise;
            if (s.cancelled) return;
            s.pdf = pdf2; s.pages = []; s.zoom = 1;
            await doRender(1);
            if (!s.cancelled) setUiState(u => ({...u, loading:false, zoom:1}));
          } else throw err;
        } catch(err2) {
          console.error("[PDF fallback error]", err2);
          if (!s.cancelled) setUiState(u => ({...u, loading:false, error:err2.message||"Xatolik"}));
        }
      }
    });

    return () => {
      s.cancelled = true;
      if (s.blobUrl) { URL.revokeObjectURL(s.blobUrl); s.blobUrl = null; }
    };
  }, [url]);

  async function doRender(zoomLevel) {
    const s = stateRef.current;
    const container = containerRef.current;
    if (!s.pdf || !container) return;
    // Save current scroll position before re-render
    if (scrollRef.current) {
      s.scrollTop = scrollRef.current.scrollTop;
    }

    const dpr = window.devicePixelRatio || 1;
    // Fit page exactly to container width for "full view" feel
    // containerWidth = scroll div inner width
    const containerWidth = (scrollRef.current?.clientWidth || window.innerWidth) - 16;

    const canvases = [];
    for (let p = 1; p <= s.pdf.numPages; p++) {
      if (s.cancelled) return;
      if (!s.pages[p-1]) s.pages[p-1] = await s.pdf.getPage(p);
      const page = s.pages[p-1];

      // Scale page to fit container width exactly, then multiply by DPR for sharpness
      const naturalVp = page.getViewport({ scale: 1 });
      const fitScale = (containerWidth / naturalVp.width) * zoomLevel;
      const renderScale = fitScale * dpr; // high-res canvas

      const MAX = 8000;
      const safeRenderScale = renderScale * Math.min(1, MAX / Math.max(
        naturalVp.width * renderScale, naturalVp.height * renderScale
      ));
      const renderVp = page.getViewport({ scale: safeRenderScale });
      const displayVp = page.getViewport({ scale: fitScale });

      const cv = document.createElement("canvas");
      cv.width  = Math.round(renderVp.width);
      cv.height = Math.round(renderVp.height);
      // Display at fitScale size — fills full width, auto height
      cv.style.cssText = `width:${Math.round(displayVp.width)}px;max-width:100%;height:auto;display:block;margin:0 auto 8px;background:white;box-shadow:0 1px 6px rgba(0,0,0,0.3);`;
      const ctx2d = cv.getContext("2d", { alpha: false });
      await page.render({ canvasContext: ctx2d, viewport: renderVp }).promise;
      if (s.cancelled) return;
      canvases.push(cv);
    }

    const c = containerRef.current;
    if (!c) return;
    c.innerHTML = "";
    canvases.forEach(cv => c.appendChild(cv));

    // Restore scroll
    if (s.scrollTop > 0 && scrollRef.current) {
      requestAnimationFrame(() => {
        if (scrollRef.current) scrollRef.current.scrollTop = s.scrollTop;
      });
    }
  }

  function changeZoom(delta) {
    const s = stateRef.current;
    const next = Math.min(3, Math.max(0.5, +(s.zoom + delta).toFixed(2)));
    s.zoom = next;
    setUiState(u => ({...u, zoom:next}));
    doRender(next);
  }

  // Hooks must be before any early return (React rules)
  const { loading, error, zoom } = uiState;

  return (
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",background:"#2a2a2a"}}>
      {loading && (
        <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",color:"white"}}>
          <div style={{fontSize:36,marginBottom:12}}>⏳</div>
          <p style={{color:"rgba(255,255,255,0.7)",fontSize:14}}>PDF yuklanmoqda...</p>
        </div>
      )}

      <div
        ref={el => {
          scrollRef.current = el;
          if (el && persistKey) {
            const saved = sessionStorage.getItem(persistKey);
            if (saved) requestAnimationFrame(() => { el.scrollTop = +saved; });
          }
        }}
        onScroll={e => {
          stateRef.current.scrollTop = e.currentTarget.scrollTop;
          if (persistKey) {
            if (stateRef.current.scrollThrottle) return;
            const val = e.currentTarget.scrollTop;
            stateRef.current.scrollThrottle = requestAnimationFrame(() => {
              sessionStorage.setItem(persistKey, val);
              stateRef.current.scrollThrottle = null;
            });
          }
        }}
        style={{
          flex:1, overflowY:"auto", overflowX:"hidden",
          padding: loading ? "0" : "0",
          WebkitOverflowScrolling:"touch",
          boxSizing:"border-box",
        }}
      >
        <div ref={containerRef} style={{width:"100%", margin:"0 auto"}} />
      </div>
    </div>
  );
}



// ===== LatexDocViewer — renders .tex source as a beautiful document =====
// Splits LaTeX source into text and math segments, renders math via KaTeX,
// so the raw \frac{}{} commands never show to the student/admin.
// itemize/enumerate ro'yxatlarini (ICHMA-ICH bo'lganlarini ham) bitta o'tishda to'g'ri
// raqamlaydi/harflaydi. Masalan: tashqi \begin{enumerate} -> 1. 2. 3. ...,
// ichki \begin{enumerate}[A)] yoki [label=\Alph*)] -> A) B) C) D) ...
function processLatexLists(s) {
  const tokenRe = /\\begin\{(itemize|enumerate)\}(\[[^\]]*\])?|\\end\{(itemize|enumerate)\}|\\item\b\s*/g;
  const stack = [];
  let result = "";
  let lastIndex = 0;
  let m;
  let safety = 0;
  while ((m = tokenRe.exec(s)) !== null && safety++ < 20000) {
    result += s.slice(lastIndex, m.index);
    lastIndex = tokenRe.lastIndex;
    if (m[1]) {
      // \begin{itemize|enumerate}[...]
      const type = m[1];
      const opt = m[2] || "";
      let style = "bullet";
      if (type === "enumerate") {
        if (/\\Alph|\\alph|^\[[A-Za-z]\)\]$/.test(opt.trim())) style = "alpha";
        // Aniq uslub ko'rsatilmagan bo'lsa: eng tashqi (0-daraja) ro'yxat odatda savollar
        // ro'yxati (1. 2. 3.), ichma-ich (nested) ro'yxat esa odatda javob variantlari (A) B) C))
        else style = stack.length > 0 ? "alpha" : "numeric";
      }
      stack.push({ type, style, counter: 0 });
    } else if (m[3]) {
      // \end{itemize|enumerate}
      if (stack.length) stack.pop();
    } else {
      // \item
      const top = stack[stack.length - 1];
      if (!top) { result += "\n• "; continue; }
      top.counter++;
      if (top.type === "itemize") {
        result += "\n• ";
      } else if (top.style === "alpha") {
        const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        result += `\n${letters[top.counter - 1] || top.counter}) `;
      } else {
        result += `\n${top.counter}. `;
      }
    }
  }
  result += s.slice(lastIndex);
  return result;
}

function parseLatexDocument(source) {
  if (!source) return [];
  // Safety: truncate very large documents to avoid infinite loop
  let s = source.slice(0, 200000);

  // \pgfplotsset{...} odatda \begin{document}dan OLDIN (preambulada) bo'ladi —
  // pastda preambula butunlay tashlab yuboriladi, shuning uchun grafik chizish uchun
  // kerak bo'lgan bu sozlamani hujjat kesilishidan oldin saqlab qolamiz.
  const pgfSetMatch = s.match(/\\pgfplotsset\{([^{}]*)\}/);
  const pgfplotsSet = pgfSetMatch ? `\\pgfplotsset{${pgfSetMatch[1]}}` : "\\pgfplotsset{compat=1.18}";
  const extraTikzLibs = [...s.matchAll(/\\usetikzlibrary\{([^{}]*)\}/g)].map(m => `\\usetikzlibrary{${m[1]}}`).join("\n");

  // Strip LaTeX preamble/document wrapper if present
  const docMatch = s.match(/\\begin\{document\}([\s\S]*)\\end\{document\}/);
  if (docMatch) s = docMatch[1];

  // Remove common LaTeX commands that don't affect content rendering
  s = s.replace(/\\documentclass(\[[^\]]*\])?\{[^}]*\}/g, "");
  s = s.replace(/\\usepackage(\[[^\]]*\])?\{[^}]*\}/g, "");
  s = s.replace(/\\pgfplotsset\{[^{}]*\}/g, "");
  s = s.replace(/\\maketitle/g, "");
  s = s.replace(/%.*$/gm, ""); // LaTeX comments

  // Convert sectioning to headers
  s = s.replace(/\\section\*?\{([^}]*)\}/g, "\n\n## $1\n\n");
  s = s.replace(/\\subsection\*?\{([^}]*)\}/g, "\n\n### $1\n\n");
  s = s.replace(/\\title\{([^}]*)\}/g, "\n\n# $1\n\n");

  // Markazlash/figure kabi vizual bo'lmagan (kontentga ta'sir qilmaydigan) muhitlarni olib tashlaymiz,
  // ichidagi kontent joyida qoladi
  s = s.replace(/\\begin\{center\}/g, "").replace(/\\end\{center\}/g, "");
  s = s.replace(/\\begin\{figure\}(\[[^\]]*\])?/g, "").replace(/\\end\{figure\}/g, "");
  s = s.replace(/\\caption\{([^}]*)\}/g, "\n$1\n");
  s = s.replace(/\\label\{[^}]*\}/g, "");
  s = s.replace(/\\centering/g, "");

  // Ro'yxatlarni (itemize/enumerate) qayta ishlaymiz — ICHMA-ICH joylashganini ham
  // to'g'ri tushunadi: masalan tashqi \begin{enumerate} savollarni 1. 2. 3. deb raqamlaydi,
  // ichki \begin{enumerate}[A)] esa javob variantlarini A) B) C) D) deb harflaydi.
  s = processLatexLists(s);

  // Extract TikZ chizmalarni va rasmlarni alohida bloklar sifatida ажратиб оламиз —
  // shunda ular ichidagi maxsus belgilar ($ va h.k.) matematik formula parserini chalg'itmaydi.
  // \begin{axis}/\addplot (pgfplots) ishlatilgan bo'lsa, kerakli paket va sozlamalarni
  // chizma kodining o'ziga qo'shib yuboramiz — shunda TikZJax uni to'g'ri kompilyatsiya qiladi.
  const tikzBlocks = [];
  s = s.replace(/\\begin\{tikzpicture\}([\s\S]*?)\\end\{tikzpicture\}/g, (m) => {
    const needsPgfplots = /\\begin\{axis\}|\\addplot/.test(m);
    const preamble = [needsPgfplots ? "\\usepackage{pgfplots}" : "", needsPgfplots ? pgfplotsSet : "", extraTikzLibs]
      .filter(Boolean).join("\n");
    tikzBlocks.push(preamble ? `${preamble}\n${m}` : m);
    return `\u0000TIKZ${tikzBlocks.length - 1}\u0000`;
  });
  s = s.replace(/\\includegraphics(?:\[[^\]]*\])?\{([^}]*)\}/g, (m, key) => `\u0000IMG:${key.trim()}\u0000`);

  // Xavfsizlik to'ri: tanilmagan/qoldiq \begin{...} yoki \end{...} teglari (ixtiyoriy [...] bilan)
  // xom matn sifatida ekranga chiqib ketmasligi uchun olib tashlanadi (ichidagi kontent joyida qoladi)
  s = s.replace(/\\(?:begin|end)\{[a-zA-Z*]+\}(\[[^\]]*\])?/g, "");

  // Split into segments: text vs math ($...$, $$...$$, \[...\], \(...\)) vs tikz vs image
  const segments = [];
  let i = 0;
  let safetyParse = 0;
  while (i < s.length && safetyParse++ < 100000) {
    if (s[i] === "\u0000") {
      const end = s.indexOf("\u0000", i + 1);
      if (end !== -1) {
        const token = s.slice(i + 1, end);
        if (token.startsWith("TIKZ")) {
          segments.push({ type: "tikz", content: tikzBlocks[+token.slice(4)] || "" });
        } else if (token.startsWith("IMG:")) {
          segments.push({ type: "image", key: token.slice(4) });
        }
        i = end + 1; continue;
      }
    }
    if (s.startsWith("$$", i)) {
      const end = s.indexOf("$$", i + 2);
      if (end !== -1) {
        segments.push({ type: "math", block: true, content: s.slice(i + 2, end) });
        i = end + 2; continue;
      }
    }
    if (s.startsWith("\\[", i)) {
      const end = s.indexOf("\\]", i + 2);
      if (end !== -1) {
        segments.push({ type: "math", block: true, content: s.slice(i + 2, end) });
        i = end + 2; continue;
      }
    }
    if (s.startsWith("\\(", i)) {
      const end = s.indexOf("\\)", i + 2);
      if (end !== -1) {
        segments.push({ type: "math", block: false, content: s.slice(i + 2, end) });
        i = end + 2; continue;
      }
    }
    if (s[i] === "$") {
      const end = s.indexOf("$", i + 1);
      if (end !== -1) {
        segments.push({ type: "math", block: false, content: s.slice(i + 1, end) });
        i = end + 1; continue;
      }
    }
    // Accumulate plain text until next math delimiter or maxsus blok
    let j = i + 1;
    let jSafety = 0;
    while (j < s.length && jSafety++ < 100000 && s[j] !== "$" && s[j] !== "\u0000" && !s.startsWith("\\[", j) && !s.startsWith("\\(", j)) j++;
    const text = s.slice(i, j);
    if (text) segments.push({ type: "text", content: text });
    i = j;
  }
  return segments;
}

// ===== TikZJax loader — brauzerda haqiqiy TikZ chizmalarini render qiladi =====
let tikzJaxLoaded = false, tikzJaxLoading = false;
function loadTikZJax(cb) {
  if (tikzJaxLoaded || window.tikzjax) { tikzJaxLoaded = true; cb(); return; }
  if (tikzJaxLoading) { setTimeout(() => loadTikZJax(cb), 300); return; }
  tikzJaxLoading = true;
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = "https://tikzjax.com/v1/fonts.css";
  document.head.appendChild(link);
  const s = document.createElement("script");
  s.src = "https://tikzjax.com/v1/tikzjax.js";
  s.onload = () => { tikzJaxLoaded = true; cb(); };
  s.onerror = () => { console.error("[TikZJax] yuklanmadi"); tikzJaxLoading = false; };
  document.head.appendChild(s);
}

// TikZ kodini haqiqiy chizmaga aylantiruvchi komponent
function TikzBlock({ code }) {
  const hostRef = useRef(null);
  const [status, setStatus] = useState(tikzJaxLoaded || window.tikzjax ? "loaded" : "loading"); // loading | loaded | failed
  useEffect(() => {
    let cancelled = false;
    const timeout = setTimeout(() => { if (!cancelled) setStatus(s => s === "loading" ? "failed" : s); }, 7000);
    loadTikZJax(() => { if (!cancelled) { clearTimeout(timeout); setStatus("loaded"); } });
    return () => { cancelled = true; clearTimeout(timeout); };
  }, []);
  useEffect(() => {
    if (status !== "loaded" || !hostRef.current) return;
    hostRef.current.innerHTML = "";
    const scriptEl = document.createElement("script");
    scriptEl.type = "text/tikz";
    scriptEl.textContent = code;
    hostRef.current.appendChild(scriptEl);
  }, [status, code]);
  if (status === "failed") {
    return (
      <div style={{ textAlign: "center", margin: "14pt 0" }}>
        <div style={{ display: "inline-block", border: `1px dashed ${C.warning}`, borderRadius: 8, padding: "10pt 14pt", background: "#FFFBEB", maxWidth: "100%", textAlign: "left" }}>
          <p style={{ margin: "0 0 6pt", color: "#92400E", fontSize: "0.8em", fontWeight: 700 }}>⚠ Chizma yuklanmadi</p>
          <p style={{ margin: "0 0 6pt", color: "#92400E", fontSize: "0.72em" }}>Internet aloqasi yo'q yoki bu ko'rinish tashqi skriptlarni cheklagan bo'lishi mumkin. Asl (yuklangan) saytda tekshirib ko'ring.</p>
          <pre style={{ margin: 0, fontSize: "0.68em", color: "#78716C", whiteSpace: "pre-wrap", fontFamily: "monospace", maxHeight: 120, overflowY: "auto" }}>{code}</pre>
        </div>
      </div>
    );
  }
  return (
    <div style={{ textAlign: "center", margin: "14pt 0", overflowX: "auto" }}>
      <div ref={hostRef} />
      {status === "loading" && <span style={{ color: "#94A3B8", fontSize: "0.85em" }}>Chizma yuklanmoqda...</span>}
    </div>
  );
}

// ===== Tasdiqlash oynasi (o'zimizniki) =====
// window.confirm() ko'plab embedded muhitlarda (masalan Telegram Mini App WebView)
// ishlamaydi yoki bloklanadi. Shu sababli o'zimizning UI orqali tasdiqlash oynasini
// ko'rsatamiz — bu HAR QANDAY muhitda ishonchli ishlaydi.
function ConfirmModal({ message, confirmLabel, danger, onConfirm, onCancel }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 99999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onCancel}>
      <div style={{ background: "white", borderRadius: 16, padding: 22, maxWidth: 340, width: "100%", boxShadow: "0 10px 40px rgba(0,0,0,0.25)" }} onClick={e => e.stopPropagation()}>
        <p style={{ margin: "0 0 20px", fontSize: 15, color: "#1a1a1a", lineHeight: 1.5, fontWeight: 600 }}>{message}</p>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onCancel} style={{ flex: 1, padding: "11px", borderRadius: 10, border: "1.5px solid #E2E8F0", background: "white", fontWeight: 700, fontSize: 14, cursor: "pointer", color: "#334155" }}>Bekor qilish</button>
          <button onClick={onConfirm} style={{ flex: 1, padding: "11px", borderRadius: 10, border: "none", background: danger === false ? "#4F6EF7" : "#EF4444", color: "white", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>{confirmLabel || "O'chirish"}</button>
        </div>
      </div>
    </div>
  );
}


// LaTeX/PDF hujjat ko'rish oynalarida ishlatiladi — foydalanuvchi qayerda
// to'xtagan bo'lsa, oynani qayta ochganda o'sha yerdan davom etadi.
function ScrollPersistDiv({ persistKey, style, children }) {
  const throttleRef = useRef(null);
  return (
    <div
      ref={el => {
        if (el && persistKey) {
          const saved = sessionStorage.getItem(persistKey);
          if (saved) requestAnimationFrame(() => { el.scrollTop = +saved; });
        }
      }}
      onScroll={e => {
        if (!persistKey) return;
        // Scroll paytida sessionStorage'ga har freymda yozish "qotish"ga sabab bo'lardi —
        // endi faqat requestAnimationFrame bilan cheklab yozamiz (throttling)
        if (throttleRef.current) return;
        const val = e.currentTarget.scrollTop;
        throttleRef.current = requestAnimationFrame(() => {
          sessionStorage.setItem(persistKey, val);
          throttleRef.current = null;
        });
      }}
      style={{ overscrollBehaviorX: "none", touchAction: "pan-y", ...style, overflowX: "hidden" }}
    >
      {children}
    </div>
  );
}

function LatexDocViewerImpl({ source, images }) {
  const [katexReady, setKatexReady] = useState(!!window.katex);
  useEffect(() => { if (!window.katex) loadKatex(() => setKatexReady(true)); else setKatexReady(true); }, []);
  useEffect(() => { loadCmuFont(); }, []);

  // Og'ir amal (parsing + KaTeX render) — faqat manba haqiqatan o'zgarganda qayta hisoblanadi,
  // taymer kabi tez-tez tiklanadigan holatlar tufayli qayta-qayta ishlamaydi
  const segments = useMemo(() => parseLatexDocument(source || ""), [source]);
  const imgMap = images || {};

  // A4 sahifa: 210mm x 297mm, Word'dagi kabi 12pt shrift, chetlar ekranga moslashadi
  // (kichik ekranda tor emas, kata ekranda/chop etishda haqiqiy A4 chegarasi kabi)
  return (
    <div style={{ background: "#E9ECF2", padding: "12px 4px", minHeight: "100%", width: "100%", maxWidth: "100%", boxSizing: "border-box", overflowX: "hidden", display: "flex", justifyContent: "center" }}>
      <div style={{
        width: "210mm", maxWidth: "100%", minHeight: "297mm", boxSizing: "border-box",
        padding: "clamp(14px, 6vw, 25mm) clamp(10px, 4vw, 20mm)", background: "white", boxShadow: "0 1px 3px rgba(0,0,0,0.1), 0 8px 28px rgba(0,0,0,0.08)",
        fontFamily: CMU_FONT_STACK, fontSize: "11pt", lineHeight: 1.45, color: "#1a1a1a",
      }}>
        {segments.map((seg, idx) => {
          if (seg.type === "tikz") return <TikzBlock key={idx} code={seg.content} />;
          if (seg.type === "image") {
            const src = imgMap[seg.key];
            return (
              <div key={idx} style={{ textAlign: "center", margin: "14pt 0" }}>
                {src
                  ? <img src={src} alt={seg.key} style={{ maxWidth: "100%", maxHeight: "90mm" }} />
                  : <span style={{ color: C.danger, fontSize: "0.85em", border: `1px dashed ${C.danger}`, padding: "6pt 10pt", borderRadius: 6, display: "inline-block" }}>⚠ Rasm topilmadi: {seg.key}</span>}
              </div>
            );
          }
          if (seg.type === "math") {
            const html = (katexReady && window.katex) ? renderKatexCached(seg.content, seg.block) : null;
            if (seg.block) {
              return (
                <div key={idx} style={{ textAlign: "center", margin: "12pt 0", overflowX: "auto" }}>
                  {html ? <span dangerouslySetInnerHTML={{ __html: html }} style={{ fontSize: "1em" }} /> : <span style={{ color: "#999" }}>...</span>}
                </div>
              );
            }
            return html
              ? <span key={idx} dangerouslySetInnerHTML={{ __html: html }} style={{ fontSize: "1em", marginLeft: "0.15em", marginRight: "0.15em" }} />
              : <span key={idx} style={{ color: "#999" }}>...</span>;
          }
          // Text segment — handle headers and line breaks
          const lines = seg.content.split("\n");
          return lines.map((line, li) => {
            const trimmed = line.trim();
            if (!trimmed) return <br key={idx + "-" + li} />;
            if (trimmed.startsWith("# ")) return <h1 key={idx + "-" + li} style={{ fontSize: "1.6em", fontWeight: 700, margin: "18pt 0 10pt", color: "#1a1a1a", fontFamily: "inherit" }}>{trimmed.slice(2)}</h1>;
            if (trimmed.startsWith("## ")) return <h2 key={idx + "-" + li} style={{ fontSize: "1.3em", fontWeight: 700, margin: "16pt 0 8pt", color: "#1a1a1a", fontFamily: "inherit" }}>{trimmed.slice(3)}</h2>;
            if (trimmed.startsWith("### ")) return <h3 key={idx + "-" + li} style={{ fontSize: "1.1em", fontWeight: 700, margin: "12pt 0 6pt", color: "#1a1a1a", fontFamily: "inherit" }}>{trimmed.slice(4)}</h3>;
            if (trimmed.startsWith("• ")) return <p key={idx + "-" + li} style={{ margin: "3pt 0 3pt 16pt", fontSize: "1em", color: "#1a1a1a", fontFamily: "inherit" }}>{trimmed}</p>;
            if (/^[A-Z]\)\s/.test(trimmed)) return <p key={idx + "-" + li} style={{ margin: "5pt 0 5pt 16pt", fontSize: "1em", color: "#1a1a1a", fontFamily: "inherit", fontWeight: 500 }}>{trimmed}</p>;
            return <span key={idx + "-" + li} style={{ fontSize: "1em", color: "#1a1a1a", lineHeight: 1.45, fontFamily: "inherit" }}>{trimmed} </span>;
          });
        })}
        {!source && <p style={{ color: "#94A3B8", textAlign: "center" }}>LaTeX hujjat bo'sh</p>}
      </div>
    </div>
  );
}
// Props (source/images) o'zgarmagan bo'lsa qayta render qilinmaydi — taymer kabi tez-tez
// yangilanadigan holatlar LaTeX/KaTeX'ni qayta hisoblashga majburlamaydi
const LatexDocViewer = memo(LatexDocViewerImpl, (prev, next) => prev.source === next.source && prev.images === next.images);

// Test uchun mavjud tillardagi hujjatlarni (PDF/LaTeX) ko'rsatadigan tugmalar qatori.
// Har bir til alohida tugma — bosilganda o'sha tildagi hujjat ochiladi.
function DocLangButtons({ test, onOpen, small }) {
  const langs = availableDocLangs(test);
  if (langs.length === 0) return null;
  return (
    <>
      {langs.map(l => {
        const d = getLangDoc(test, l.code);
        const isPdf = !!d.pdfUrl;
        return (
          <button key={l.code} onClick={() => onOpen(isPdf
            ? { type:"pdf", url:d.pdfUrl, id:test.id+"_"+l.code, name:`${test.name} — ${l.full}` }
            : { type:"latex", source:d.latexSource, id:test.id+"_"+l.code, images:d.latexImages, name:`${test.name} — ${l.full}` }
          )} style={{...S.badge,background:isPdf?"#FEF3C7":C.primaryLight,color:isPdf?"#92400E":C.primary,border:"none",cursor:"pointer",fontSize:small?12:13,display:"inline-flex",alignItems:"center",gap:4}}>
            <span><LangFlag lang={l}/></span><span>{isPdf?"📄":"∑"} {l.label}</span>
          </button>
        );
      })}
    </>
  );
}

function KatexSpan({ latex, block, fontSize }) {
  const ref = useRef(null);
  const [ready, setReady] = useState(!!window.katex);
  useEffect(() => { if (!window.katex) loadKatex(() => setReady(true)); }, []);
  useEffect(() => {
    if (!ready || !ref.current) return;
    try {
      window.katex.render(latex || "", ref.current, { throwOnError: false, displayMode: !!block });
    } catch {}
  }, [latex, ready, block]);
  if (!latex) return null;
  return <span ref={ref} style={{ fontFamily: "KaTeX_Main, serif", fontSize: fontSize || (block ? 20 : 16) }} />;
}

// Convert display string to LaTeX for render
function toLatex(s) {
  if (!s) return "";

  // Qavslar ichma-ich bo'lganda ham to'g'ri ishlashi uchun (masalan FRAC(√(3/2),5))
  // — mos yopiluvchi qavsni topamiz (chuqurlikni hisoblab)
  function findMatchClose(str, openIdx) {
    let depth = 0;
    for (let i = openIdx; i < str.length; i++) {
      if (str[i] === "(") depth++;
      else if (str[i] === ")") { depth--; if (depth === 0) return i; }
    }
    return -1;
  }
  // Faqat ENG TASHQI (top-level) vergullar bo'yicha bo'lish — ichkaridagi vergullarga tegmaydi
  function splitTopLevel(str) {
    const parts = []; let depth = 0, last = 0;
    for (let i = 0; i < str.length; i++) {
      const c = str[i];
      if (c === "(") depth++;
      else if (c === ")") depth--;
      else if (c === "," && depth === 0) { parts.push(str.slice(last, i)); last = i + 1; }
    }
    parts.push(str.slice(last));
    return parts;
  }

  // FRAC/ROOT/ABS/LOG_BASE/SUB/√(...) larni rekursiv ravishda LaTeX'ga aylantiradi —
  // ichidagi formulalar qanchalik chuqur ichma-ich bo'lishidan qat'iy nazar to'g'ri ishlaydi.
  // Bo'sh joylar (masalan hali yozilmagan maxraj) uchun nozik "□" belgisi qo'yiladi —
  // shunda foydalanuvchi qayerga yozish kerakligini ko'radi.
  const ph = a => a || "\\square";
  function convert(str) {
    let out = "";
    let i = 0;
    while (i < str.length) {
      // Aralash son: (N+FRAC(a,b)) yoki (N-FRAC(a,b)) → N\frac{a}{b}
      const mixed = /^\((-?\d+\.?\d*)([+-])FRAC\(/.exec(str.slice(i));
      if (mixed) {
        const fracOpen = i + mixed[0].length - 1;
        const fracClose = findMatchClose(str, fracOpen);
        if (fracClose !== -1 && str[fracClose + 1] === ")") {
          const args = splitTopLevel(str.slice(fracOpen + 1, fracClose)).map(convert);
          if (args.length === 2) {
            out += `${mixed[1]}${mixed[2] === "-" ? "-" : ""}\\frac{${ph(args[0])}}{${ph(args[1])}}`;
            i = fracClose + 2;
            continue;
          }
        }
      }
      const fn = /^(FRAC|ROOT|ABS|LOG_BASE|SUB|SUP)\(/.exec(str.slice(i));
      if (fn) {
        const openIdx = i + fn[0].length - 1;
        const closeIdx = findMatchClose(str, openIdx);
        if (closeIdx !== -1) {
          const args = splitTopLevel(str.slice(openIdx + 1, closeIdx)).map(convert);
          if (fn[1] === "FRAC" && args.length === 2) out += `\\frac{${ph(args[0])}}{${ph(args[1])}}`;
          else if (fn[1] === "ROOT" && args.length === 2) out += `\\sqrt[${ph(args[0])}]{${ph(args[1])}}`;
          else if (fn[1] === "ABS" && args.length === 1) out += `\\left|${ph(args[0])}\\right|`;
          else if (fn[1] === "LOG_BASE" && args.length === 2) out += `\\log_{${ph(args[0])}}(${ph(args[1])})`;
          else if (fn[1] === "SUB" && args.length === 2) out += `${ph(args[0])}_{${ph(args[1])}}`;
          else if (fn[1] === "SUP" && args.length === 2) out += `{${ph(args[0])}}^{${ph(args[1])}}`;
          else out += str.slice(i, closeIdx + 1);
          i = closeIdx + 1;
          continue;
        }
      }
      if (str[i] === "√" && str[i + 1] === "(") {
        const closeIdx = findMatchClose(str, i + 1);
        if (closeIdx !== -1) {
          out += `\\sqrt{${ph(convert(str.slice(i + 2, closeIdx)))}}`;
          i = closeIdx + 1;
          continue;
        }
      }
      out += str[i];
      i++;
    }
    return out;
  }

  let r = convert(s);
  r = r.replace(/√(\d+\.?\d*)/g, "\\sqrt{$1}");
  r = r.replace(/\^(-?\d+\.?\d*)/g, "^{$1}");
  const syms = {
    "π":"\\pi","α":"\\alpha","β":"\\beta","θ":"\\theta","γ":"\\gamma",
    "λ":"\\lambda","μ":"\\mu","σ":"\\sigma","φ":"\\phi","ω":"\\omega",
    "Δ":"\\Delta","∞":"\\infty","≤":"\\leq","≥":"\\geq","≠":"\\neq",
    "±":"\\pm","×":"\\times","÷":"\\div","∈":"\\in","∀":"\\forall",
    "∂":"\\partial","∑":"\\sum","∫":"\\int"
  };
  Object.entries(syms).forEach(([k,v]) => { r = r.split(k).join(v); });
  r = r.replace(/(?<![\\a-zA-Z])\*/g, "\\cdot ");
  r = r.replace(/log_\(/g, "\\log_{");
  ["sin","cos","tan","cot","arcsin","arccos","arctan","sinh","cosh","tanh","lim","ln","lg"].forEach(fn => {
    r = r.replace(new RegExp("(?<![\\\\a-zA-Z])" + fn + "\\(", "g"), "\\" + fn + "(");
  });
  return r;
}

// ===== MathInputField — real-time KaTeX preview =====
// Shows beautiful math as user types, LaTeX codes hidden
function MathInputField({ value, onFocus, style, placeholder, active }) {
  const [katexReady, setKatexReady] = useState(!!window.katex);
  const ref = useRef(null);

  useEffect(() => {
    if (!window.katex) loadKatex(() => setKatexReady(true));
    else setKatexReady(true);
  }, []);

  useEffect(() => {
    if (!katexReady || !ref.current) return;
    const latex = toLatex(value || "");
    if (!latex) { ref.current.innerHTML = ""; return; }
    try {
      window.katex.render(latex, ref.current, {
        throwOnError: false,
        displayMode: false,
        output: "html",
      });
    } catch { ref.current.textContent = value; }
  }, [value, katexReady]);

  return (
    <div
      onClick={e => { onFocus?.(); setTimeout(()=>e.currentTarget.scrollIntoView({behavior:"smooth",block:"center"}),350); }}
      style={{
        minHeight: 54, padding: "12px 18px",
        background: active ? "#F5F7FF" : "#FFFFFF",
        border: `2px solid ${active ? "#6366F1" : value ? "#22C55E" : "#E2E8F0"}`,
        borderRadius: 14, cursor: "pointer",
        display: "flex", alignItems: "center", flexWrap: "wrap",
        boxShadow: active ? "0 0 0 4px rgba(99,102,241,0.14), 0 2px 10px rgba(99,102,241,0.12)" : value ? "0 1px 4px rgba(34,197,94,0.08)" : "none",
        transition: "all 0.2s ease",
        ...style,
      }}
    >
      {value ? (
        <span ref={ref} style={{ fontSize: 23, fontFamily: "KaTeX_Main,serif", color: active ? "#4338CA" : "#15803D" }} />
      ) : (
        <span style={{ color: "#94A3B8", fontSize: 15 }}>{placeholder || "Javob yozish uchun bosing..."}</span>
      )}
      {active && !value && (
        <span style={{ display:"inline-block", width:2, height:24, background:"#6366F1", borderRadius:1, marginLeft:2, animation:"blink 1s step-end infinite" }} />
      )}
    </div>
  );
}

// ===== STORAGE =====
const ADMIN_LOGIN = "Dilshod_11";
const ADMIN_PW = "Dilshod_11";

// ===== 🔧 FIREBASE SOZLASH =====
// console.firebase.google.com da loyiha yarating → Project settings → Your apps → Web (</>)
// bo'limidan olingan konfiguratsiyani shu yerga qo'ying. Agar bo'sh qoldirilsa,
// sayt avvalgidek faqat shu qurilmaning localStorage'ida ishlayveradi (Firebase o'chiq).
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyCGpYIBW_TWHwyl-ddjbQ6TDjARKpgXn3k",
  authDomain: "math-platform-2dc2a.firebaseapp.com",
  projectId: "math-platform-2dc2a",
  storageBucket: "math-platform-2dc2a.firebasestorage.app",
  messagingSenderId: "109909968975",
  appId: "1:109909968975:web:7b5c4ba4a80a54085ac23a",
};
const FIREBASE_SYNC_COLLECTIONS = ["users", "tests", "results"];

let fbApp = null, fbFirestore = null, fbSdkLoading = false, fbListenersReady = false;
function isFirebaseConfigured() { return !!(FIREBASE_CONFIG.apiKey && FIREBASE_CONFIG.projectId); }

function loadFirebaseSdk(cb) {
  if (window.firebase && window.firebase.firestore) { cb(); return; }
  if (fbSdkLoading) { setTimeout(() => loadFirebaseSdk(cb), 300); return; }
  fbSdkLoading = true;
  const s1 = document.createElement("script");
  s1.src = "https://www.gstatic.com/firebasejs/9.23.0/firebase-app-compat.js";
  s1.onload = () => {
    const s2 = document.createElement("script");
    s2.src = "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore-compat.js";
    s2.onload = () => cb();
    s2.onerror = () => console.error("[Firebase] Firestore SDK yuklanmadi");
    document.head.appendChild(s2);
  };
  s1.onerror = () => console.error("[Firebase] App SDK yuklanmadi");
  document.head.appendChild(s1);
}

// Firestore'dagi o'zgarishlarni real-vaqtda localStorage'ga ko'chirib turadi —
// shu tufayli mavjud db.get() chaqiruvlari o'zgarishsiz ishlayveradi va barcha
// qurilmalar bir xil ma'lumotni ko'radi.
function initFirebaseSync() {
  if (!isFirebaseConfigured() || fbListenersReady) return;
  loadFirebaseSdk(() => {
    try {
      fbApp = window.firebase.apps.length ? window.firebase.app() : window.firebase.initializeApp(FIREBASE_CONFIG);
      fbFirestore = window.firebase.firestore();
      FIREBASE_SYNC_COLLECTIONS.forEach(col => {
        fbFirestore.collection(col).onSnapshot(snap => {
          const arr = snap.docs.map(d => d.data());
          try { localStorage.setItem(col, JSON.stringify(arr)); } catch {}
          window.dispatchEvent(new CustomEvent("firestore-sync", { detail: { collection: col } }));
        }, err => console.error("[Firebase] onSnapshot xatosi:", col, err));
      });
      fbListenersReady = true;
    } catch (e) { console.error("[Firebase] Ulanishda xato:", e); }
  });
}

// Butun kolleksiyani har safar Firestore'dan o'qib-qayta yozish (eskirgan yondashuv)
// 1000+ foydalanuvchi bir vaqtda ishlaganda juda sekin va qimmat bo'lardi (har bir kichik
// yozuvda BUTUN to'plam o'qilib-qayta yozilardi). Endi faqat O'ZGARGAN yozuvlarni
// yuboramiz — Firestore'dan o'qishga UMUMAN ehtiyoj yo'q, chunki eski holat allaqachon
// bizning qo'limizda (localStorage'da).
async function pushCollectionDiffToFirestore(col, prevArr, nextArr) {
  if (!fbFirestore) return;
  try {
    const coll = fbFirestore.collection(col);
    const prevMap = new Map((prevArr || []).map(i => [String(i.id), i]));
    const nextMap = new Map((nextArr || []).map(i => [String(i.id), i]));
    const batch = fbFirestore.batch();
    let ops = 0;
    for (const [id, item] of nextMap) {
      const old = prevMap.get(id);
      if (!old || JSON.stringify(old) !== JSON.stringify(item)) {
        batch.set(coll.doc(id), item);
        ops++;
      }
    }
    for (const id of prevMap.keys()) {
      if (!nextMap.has(id)) { batch.delete(coll.doc(id)); ops++; }
    }
    if (ops > 0) await batch.commit();
  } catch (e) { console.error("[Firebase] Yozishda xato:", col, e); }
}

const db = {
  get: (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set: (k, v) => {
    try {
      let prevForDiff = null;
      if (fbFirestore && FIREBASE_SYNC_COLLECTIONS.includes(k) && Array.isArray(v)) {
        prevForDiff = db.get(k); // eski holatni yozishdan OLDIN olib qolamiz (diff uchun)
      }
      localStorage.setItem(k, JSON.stringify(v));
      if (prevForDiff !== null || (fbFirestore && FIREBASE_SYNC_COLLECTIONS.includes(k) && Array.isArray(v))) {
        pushCollectionDiffToFirestore(k, prevForDiff, v);
      }
      return true;
    } catch (err) {
      console.error("[DB] Storage error:", err);
      if (err.name === "QuotaExceededError" || err.code === 22) {
        alert("Xotira to'ldi! PDF fayl juda katta bo'lishi mumkin. Iltimos kichikroq PDF yuklang yoki eski testlarni o'chiring.");
      }
      return false;
    }
  },
};
function initDB() {
  ["users","tests","results"].forEach(k => { if (!db.get(k)) db.set(k, []); });
  autoActivateScheduledTests();
}

// ===== SESSIYANI ESLAB QOLISH =====
// Login/parolni har safar qayta so'ramaslik uchun — bir marta kirgan
// foydalanuvchi (o'quvchi/admin/o'qituvchi) qurilmada eslab qolinadi.
// Telegram Mini App har ochilishida sahifa "yangi holatda" boshlangani
// uchun bu ayniqsa muhim (aks holda har safar qaytadan login so'raladi).
const SESSION_KEY = "app_session_v1";
function saveSession(s) { try { localStorage.setItem(SESSION_KEY, JSON.stringify(s)); } catch {} }
function loadSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY)); } catch { return null; } }
function clearSession() { try { localStorage.removeItem(SESSION_KEY); } catch {} }

// ===== SCHEDULED TEST AUTO-ACTIVATION =====
// Har bir testda ixtiyoriy `scheduledAt` (ms timestamp) bo'lishi mumkin — admin
// testni oldindan yuklab, qachon boshlanishini belgilab qo'yadi. Belgilangan vaqt
// kelganda test avtomatik "active" holatiga o'tadi.
function autoActivateScheduledTests() {
  const ts = db.get("tests") || [];
  const now = Date.now();
  let changed = false;
  const next = ts.map(t => {
    if (!t.active && t.scheduledAt && t.scheduledAt <= now) {
      changed = true;
      return { ...t, active: true, startedAt: t.scheduledAt };
    }
    return t;
  });
  if (changed) db.set("tests", next);
  return changed;
}

function tsToLocalInput(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function localInputToTs(s) {
  if (!s) return null;
  const t = new Date(s).getTime();
  return isNaN(t) ? null : t;
}
function formatScheduled(ts) {
  if (!ts) return "";
  return new Date(ts).toLocaleString("uz-UZ", { day:"2-digit", month:"2-digit", year:"numeric", hour:"2-digit", minute:"2-digit" });
}
const UZ_FLAG_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAABQCAIAAABd+SbeAAAFkUlEQVR42u3cS2yUVRQH8P+595uvM9OZPqYDLX1AQd7YaEADrUURFUKMuKAJxAcxxBhl49aNa12YGDVK0ISEqEsTXfiIQoIxRRJJQwoItPKo2BZqh3amncf3usfFlEKglGJwFs75Lybt7ZeZ5NfTc8/c+VLC/m8h+e+jhKA0seZ4HQGKiOjGimEYZhG8n9CKyDAHrg9jppaYYWkK6eKXkvsArRUFjgdLdzTVdTQkFscjAbg/nesevtYzMl68Yvr3wWAwiIiZrz8CuPEtio8CfasyUZB3O1vmvde+en19zfT6hOcPZHI/Xxn7sPfCH5kcEzFgPB9EIGI2ILBhKGIGmKEUmwDF1qOVQN9WywXvtYeW7NvYpgiGAfCnZy5/9vvA6bFJz5iF8WhlyNKKfIbFvHvNorZE1TvH+xrikTfbWr/sHzo5Mv7Wo8uHc86Bk5eebm3c3FT3/okLfeksNJVbWavZatnxnmqt3/94G4Ndw64xXT/2vHGop2dk3DHGEF1KZ0+Ppj2GIjJe0F5fu2dFs2NMMmzvXtbcEK3Iu96uJQs2LUjks4VVtbE9K5qNIjBT+VU0zThHEwBGVKtTO59ojUc8Y0JKvXKk9+CJ83Y84htmZgYIIIJhFOHqwnZIqeFcIaJVbdhOO17W8Rqror4xIxP5mmg4bltD2UIQBFBKWsdU0/Dz7vY1i1rjEdcYW6lfr44dPD1gxcJuYKYv4+sjR7ENjOYcGE7Gwk5ghtJZuyLUWBUdyuRAaKqODWbz47lCojJsYGUKHhNYWgczQLS1OckotmZ83j9Exsz+RJrwQFXkl+0bvnpmbVLrDzpW9+7oXFpd+diCupNdnXtXL6wLWd3Pt3+zZV1dSFOZNZCZK9qAodXy6koCLCIAvakMKzXLDmYAIhrIFn4aTA1nC6MF5+iVsRrbGi64HnBocPRcOjvueocHU1fzTsrxYKmy2g/v0KMJ7JtjOzrXz6/xDVuK2r8+emwwpW0rmJVHEwWeD0XQGq4HP0CkAszIu6gIwVLIuyBCRUjOOoqrhMBcnMgx4DMDWFkTIzZ0t7/2gFmFLFIqxLxpYf3r65YlLN0Qrdj7yPJVyapKxktti7tWtcSVoutbaFlDEwHMhwdThCmOXUsXMGguNIZZEbEfvLy8aV/ng1Hbao1HP+5cs6U5SYH5qGP122uXqal3jNI6ABiuDYfO7dqUqLANs6Vo2/e//dA3aMfCnpk6TSLcGDlu7zxLE/GWynD3lbGIpdYlq/szucvp7Lr62qwfnE1lSKuymjo0nntx5h9olcs5f+W9riUNPjMDz7bMOzIy/udoGkqRIlU8tTAMRTN262s559LYZKDI8c3FscmM58ctvXXR/JZYuO/apF9mPfqO0MzQltV7ZUxbanNjUhFFtN65rBFaX8jkJgou+0EsbDfGo5OeP+PcpxQprRkggrY0EUJEB558eHE88sXZywERS+u45Y34rpUt725YuSgWmV7vT2evZgtn0tlPTl3qTU2wmpOaRTQvYgfMf+fdcjvBo7t+lKWJAsetrgxva5m/sTHRGosS4Xwmd2Qo9d3ASN71YOl7eEFjAJq525Q59JS1MfB8MKAUwAgMiGBbSql7+pylOGyU4Ym0NZfzyuLEpipChKkjJLKJwQEzG3NvxXnTuFJe0Kz1HH3MzTUJBqgcwf41dLVTEIVS9OjD0VZRKEVFr3z1BVEoRUWz3C1Qos3QD0RBKvr/E7n3TqAFWiLQAi3QQiDQAi0RaIEWaIlAC7REoAVaoCUCLdASgRbo8o4l/5ygVNAkN9CVIjTQ2SUKpajo7tRxUShFRTce2igKpajoUEFuCSsJNMteKHO0QEsEWqAFWiLQAi0RaIEWaIlAC7REoAVaoCUCLdASgRZogZYItEBLBFqgyzj/ABi4aLDcDtFaAAAAAElFTkSuQmCC";
const RU_FLAG_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAABQCAIAAABd+SbeAAAAzElEQVR42u3asRFBURCG0bveGwGBSBNmtKoEM2I9KIYCZCLXVYPAH3C+Es7szG6wNcZo+n4LBKBBCzRo0AINWqBBgxZo0AINGrRAgxZo0KAFGrQ+a352fx2JygNNaKLPlyuFxES3/YlCYqKnzZJCArpbhs470AINGrRAgxZo0KAFGrRAgwYt0KAFGjRogf6V5qm9KCSge60oJKAPjyOFQHXbbSkkJvpeawqWofNOoEGDFmjQAg0atECDFmjQoAUatECDBi3QoAUa9N/2BrvhFxQCa2t4AAAAAElFTkSuQmCC";
const QQ_FLAG_IMG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAABQCAIAAABd+SbeAAABHGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGDiyUnOLWYSYGDIzSspCnJ3UoiIjFJgv8PAyCDJwMygyWCZmFxc4BgQ4MOAE3y7BlQNBJd1QWbhVocVcKWkFicD6T9AHJdcUFTCwMAYA2Rzl5cUgNgZQLZIUjaYXQNiFwEdCGRPALHTIewlYDUQ9g6wmpAgZyD7DJDtkI7ETkJiQ+0FAeZkIxJdTQQoSa0oAdFuTgwMoDCFiCLCCiHGLAbExgwMTEsQYvmLGBgsvgLFJyDEkmYyMGxvZWCQuIUQU1nAwMDfwsCw7XxyaVEZ1GopID7NeJI5mXUSRzb3NwF70UBpE8WPmhOMJKwnubEGlse+zS6oYu3cOKtmTeb+2suHXxr8/w8A3kFTfazGM+sAABQwSURBVHja7V1LjCXXWf6//5y6t7une8bzaGdiOxrPSMYYB40t2zJ2IotIVlbYAoOQkBALFrBjhXglQJRFtgiBYBHxEAsQkWAVAgQZTEIwJo6VxK/YjvFgz4wznvG8e/pxq/6PxTlVdep1+073dGeMXLbsvvdWnfOf//06p3BtddUpRl7FSJI0I2mkkEZj/FaSCyIioDS+7L0gQzexPd4H80IEngIoFAoACohgfWIARBEu7xQXVicvvncJQInVgGeR+ElIipgIAurC6AH7aGIOIgIIKSIUUoDq6xq5ZPlV/FADXBMAAFnOJQK0qZNSnyIBsBQeoE3ZCGQFCiMvIPmU/lGBEteLeijEKRAmDgADUKhA5rwev3XJK0ANY/iR8/9z5tyv/83X94zHRqsAgRARIghIho8QEVBFjBHjTX4sGR2ASXioWjKrhaV8zHpdcT6SaOAjrMMisWpCloNBKBQKkOC2ha0WkaqfQTFAgCbl4vNgDSXS+ahgQJKAFCEF8X5uFLx1KfuzX/zk4sibmaoC8KQRXJyfWxzNmVkLpARcRupF/DOZP4IcGb18EDWCKqaRElL0YIMBUAiFIErOR4P5BEQUN6kxy46+QnMFNbfWYwaykRW5Sw4Nt1XLhKDkqppR4rAseagkG8BJLguZNwvKWERMRHzQDoVZQTNaENjGlBZFxIwIq3dBL4AUgKVSoFjNqCQBimiJBna4OEF8VEIBLqvURRyVBkACd7GyDfF+pgJXD8iaAcqhwlellLHSfoEV4sKJFNdRSBMlWH0BgoJS81ULIkRMSDC3vCgEcEH3elKEQKJCG/KHGiPOgUTOgoUo1IGRkjXaiFQzIBJTpMt2TakpqVuyBmqtKyKiIkxluSXjiL/W2AEwMFGl34JuDiyl/WBF4S2NUVgbKxjZHjmAZhCCNCsKMwUCrumDmxEkHaIlW0mpi4UiChHi8tp6plyaG40Uq/nk2qTwmjmtxSoVwqa2tIGV10IatBAAIthfAVJWBa1cbDSMJW1QzQZJDCBN0LQgAc6EMWsNVtv2BE6UYEUBo6Y6vmUuK/us0cKjKMysEEQj6WvxiyhmLSEgBB5YLwpw48n7Pvb4Pbfftbx/YWHuzTPn/+W1t59+8fTVNVMFm95CLZWEoIe/0lVF7wIlIzdZsv5De9ABrUzYFEKyh8fRFq8U7OrvBhhsO6R9I9d/F1YUBQBQ1cx8zReJ6qAQFBNx4FphS2P77JOP/MSdh0Xk3LXVP3n6W8+/dfbyRCcFFS1ABYisVftDKQME8IHuIitm6YpyVKwNlVHKdWWmejRFlFRBvz5paPYSa7VA1xzQ9Z/B4KG16VF711bQCsKZmVUcXQ+dqFNCpDD13PjcU488eMdhM37vzPnP/t1/vn1hY88409JFTwEITk9u9K4eK6cp1FVGZ4jLhjyz1nqiehUaBRwaJPUrW+iQloYHU3o3biOa5h1tN79nzOjKFAXNTIHg4amQtGBWWQsUIAIHXVlfe/L+Iw/ecTjP85V88oV/eP4HV3hw78LIe+cUWmomRjW2OrGRxwN3HVrZ2JiY5IbLa5Pjdx7at5Bd3ihowW0TiAoR/03FGQSkG3OiFAKSFIt00qaaaj/DFAm9WA5mYPoNNRdqhDA4y0GWWlMz+JwUEVoxMZpZDAKVIkEKmPKOiIgYuDDyn/7xO0nx3j/zvbdfP3NlaT7L88ISwaE4QMU4drp/j//R5flffvTuw/PZoSW/vOT2z8nPP3DsE0eX989xcT5j4EPhoHXENE5HeXVXWF03JLYOBvm6I/IkMo0Oc0xr1DoaiMIYeUSBPLdb9vjDBxZBE+jLJ88DjomvHXU7CJqpTvLiqYeO/czxo/vG4z/+pcf+4tnXTpy9/Hs/9YkfOXzgvtsO3nfk0J8+/Uoh8GIWAbIZEx1kFRkhdWP6VRCRmruZdG6P5URX+3cfH9DgIgIzMzMozcw5p10HN7KMiJDOAarh97yAlvJSTxAWruLA0Uj/8plXz1xe9d69/O6FL3/7rZdPX/n7b745ct6N/J8/89L76/mcS8I5opdr+rlSKVpKLmR2zq2EoCsBAzdL5X31GeG2NptCs3QOHzQ06hxPHV5A9dp6vr66IYvzInL7/sXCqgRIiJPjjNc2uLq+8iuP3/u1V9957p2zX/zXb9/20UM/ec8dANad+8KXn3Mjf/DA4ic/fuBLz755dcK9cxnYSfAldok2TYFCtPTWtTcUYkP3DzJgkhdD4ptplSeI0V8pH/XjbARU0/DM+MHXAVUzf0TQQy9dXf3OO2c//WNHzPipe2//6+femBTiVQoyKloILL//2PKRWw4/dfzoLePs5VMX3r0y2bP32hPH79w3P/eP333rpVMX58f+/o8dfOqBu9yE71zdeOGt0ysT8aKAUboSjc1UClvxdhOJM2mkBGsxuxLD6ToHSDTm6nEuhieq4ItsrZWXkRr0ML4Jnfov/ffr60VB2pH9e3/t8Y+vra6uTkQVqoCImkxErly9/MS9R2/ds/DTx4/uHfkLa/nR/XsfufO2ew8feOToR1bW1ovCPnX37Yfmx7/w6N1qxdoGHbQdS2OLdiyV4unmcOA3Tvc9+tE4AG0EAK3sjmiXwqizVByPsldOrf7RP39LnYPwiePHPv+zD96xT1dWV8+vb0ykIGzssu+cuPj82++dvnjl9bMXn3757YvX+NWX3j65snr64pWvvXH6xLm1F06cffkHF0+ev/Tqu+effuUdwGkQSyLN45R2LyTnOKudbHtsmEKSINENhCIoQV7XRDXGuPmXJN1nfvu3Tl5c/bfvn8m8q/O/FMS4k6ORf/HU+98/c/7YrbfsX5g7trzv5x686+G7Dt+1vLgyyc9dWjWRvfPjw/sX/vCfXjh5YWW8Z/Hd81ePH1n++msn/+obL35kef/p91cO7J1fmh/9wVe+eWm9cKPR+ZU1Vbia6k3zjjJyQtttSLR5jCZiTSENz9IgbcAz6Sj9ymuUMsle/8s+nzNBaAP4wMrzHo8d2Tc38uqcqqoqVi5d+K8T53/3K9+dH40qzS0CgVW+h6qurG0sZbjvyP4jy/vHc6OVtfXXT77/xnvXCiqlALAxKZxzhRWZUyjMmOfmnBZFMR5lRq5P8sz7PC8yH7SONOxelYGcHjR23bimzmlza2Jgh/NzaYK/aS3FUr+wN0zvKpWJ8cAcPvPYkVsWxi4bZd57732lLlAnRWs5CjGp0faMfW74jzcv/Ptr54w0yNi58dg7sRAnjTNHwvuMQpopMM4cST/KAljzo0xIP/KpGZiaPp0Ny5v7FUyWWAZvmpAzKaMMZZp6BWIWF7N6zLfmqMJCJkl5FRipaovjTOZGYiHpTxqYsiNCOrZMaJR5n8RzLNPvnfQ/mjmKdpKsqRam133rVG2ZV029PUTYYuGgVfxqF4DIytQ2igKN/4VyFitlVyeNE+n0YfDCaEazUCAoXcjS3SkgQimIAlb6T4q0vhDFgCGzCohRonseWaiPNaq8t5m0mKsl4F2vKdWL1kmUNvN8ZaCQcC5Rj6JMxkFanRMIzVCXPOu/m7Wy0iemOaAwMauz44HxPUW8x+IejDKQSS4ODW+0WavWMkuJtGLUMrcUTYqEZa4itWY9yTodrve1OxSSH9EuiEvTq4C2qzJME3Wok6rNRzvVZ/S0SMRsHcoyDNS4MAdq41mcO3dukhcbuaMi4LbMJ1CGVNesrmZHmtEoa3V5OC2EsVN0raqrsViOKSo+lKTZXAAqHBObQU52qVtWhzuPp8QFhcpicuXy2fFoNBqNfLzm5q+98OyJ3/+deTfHKd0uacNMVRpgWqTuw2+NXCZeAMpaCqVZUmwncVpS0tHLpb8w1I7DpGgjSOSgL2Tvk5q6rtwI0pnYcSQSFqUkz/Plgwd/83OcG9fGsMg5Wt6451fXvdcy91lZdXS6MNDuRhn2BdCu3UAajR1VUW5ItzaLukBPh0v0RFGRDVU6NSFkg0JAxw53dJ40HcceOqLRxdNySbix7u1CtuGtllcPKjLIslI9e7IzXU3SU6Dru6f1APoG73Ed+hTu4BRoVu9SSnOwfatbWEDfl5usrvSmO6igiCgURBTdWJyNQOUUrd2R1EOcKRRGqgmajlrZHtYtgDJWqjtE7K+Y9PSMpUmlKcWRNns0ww2mgjzsLPbILtl1qEsfli1+9NMt6qzRg7BPP7Njn5mKXSnHaPbTsB8AdJBYBeiYwRj3fdML9nUsGpsk75hk7zy77soAg8wYC11HrWizAVOu7N7ZKp1smgO6gcCn+al+EigaaofUWvqmwrGFctysKcfhmyslM72c0eoqmX3e64Jw9pX0srnKzXENUTFlnFYPxvYZokshTE2xttO5vaSqffcGGL4sWIC7hbubYbSduKoWol5DqlVqFNuYYMsyuOmD29dXM8LW263QYuEp1quqrfQlc26E6phmE2Ygzy7w6XammKLQZho8dh5CRLQs/kyrAO2MiHEnsHzDx5xipYcEo5kobBZnf4gLm1HzbEdB/dC0dhKONsLuXme5t591O551117PMsjWtNP05UwLIzvf9/aEDLrzZRqnQq/fdDGN+HJqjHC96J7WCHoDW+hmINhQR1J37VPYK0F9nx9d8rZdl5meYp1n13pbaEvc1NvdMrVmvHmG2/o71nQW0b4hNmRWSz1DPHLz+9Q9HE3p7x3ZdDEpX+2Qkp2FltjGdFtjpk1dOmEPbFv3OnaZr4Y2wnBLz0pZJtoBv73xZ+VHe+mrIQ1Sqk4C11v5uNl86TdNO17fiM02l09tqRPZZIsd+zfSJj4Ce4txHG5rwCCWgFC/LuszIU0aohXSzPrrI2VRLq6m6gmXuqLVAyPJdu9A3SmS7JhMdwmRLVyx0V6Azs7zFiWaLNH9iVM8CplCoCHKV0wGarl7mkKh0XxBMCWfJw2K0cKi1zkmvkfds2LtuljZZIBOhgSxjBbZnj1kQ9jWm9QgaR0noVEEr7pU2hPVK26w65C/oWmNueUUJyPUpclkB2SyDbwX0QEiUKgGkUx0rKIGrQTeuwxXTxWvfP7U2O0h87qAirbIJ50PaXkUqPuE2YC33kfLnr35bSlg1Z0s/U1MrPaKp+0v3QMu2jXchHygUPrL4M0NgB3BaU/WGsKSrxX5Rn5o9cBvKOaTyc+eP4fL17L/PaPJRvm25ErV51CP16e82O87tqtNHNZ2HG71GjqrYHp32FSrhU0s0jSL1e4Baoy87uzs0pxbGI+zcWzrGDN7Cxe/Ov+Md1nvCRgNw1W32Up9kkP7AIFUIlG3AQ3YEUyhAJOOjCYA6ZwcGA6tboNEZaaSJ5yJnjN6ghAUYgs2/7B7eDHp/PI+y85ePfW33/jiwmipsUW5YeEaMrfZnuBG28umDaOY6qr0joCZmLI6M6EPm5wJc1tD9IT5oWz5oUfvV0Glgb0xh/P79hycH82zQ/E+IKXbpJNoB7bWwJlwXZ83kDoZ6fydzrtS46JSN4l1aOxt6Ach7exCehtSf4N9cKLdf4O0ggWzfGm0GFqyqhvLblJOChuT3IS7Uh1VySbYNqDsu21QAw45q8kSOOx5tU+1QeKZs9/+ttbYZhXp9TLacDbsbU0cFS1oOWnVvoQQsBBOhBS9PqNSbQsLPaVDDudAz0UbZA54ZtNimL6KEVuEwKxWMQGj1U3WZq7OpLE5ugIp7M+IW62N4hpp0i0nkLYah3/wskLXj4G4xaaqsBCCXV78UKpzF9InN7wcPOPj/ibjgg/21To/pR1TUP5/Lnv3aVxmeNj1PFU+vHbl+hDRHyJ699XrDR1vyBhym/DNqPu6VeTr7TvY8tQ7h+VZOi+263Vc7yJvYNvyjTJrs5xMIzM0GjR2h5chZrXVwe+CSN7kztyNazQI6+2P/hSCHRW6natG35z6uuzoZ5KuSk6g2R3Qd4i1d3Tw65UMpFt0Egi1TKNgR+VuJhCxdRi2D0Zv59s2RKE8pqo86UNl6OS9G62mt9tr/EFwEFvelIbDdsNJ6YqebW+7uSuiAZzg5tTaW1imOgdo1SzoBdB4GI0ONL2knzB8y1AKEC3LIJtvU+1p9kqPJeesNa8WA3Hg5t4HMTBuz3rrA3DjUQpQ1XDET41ohRaUa9dWZI60tBe4rA4kx7tDPrw2MRUA8mKyJnvEOadavbUCVy+vvLty+oWTz0FQWPynyIuiKPKiyPPcLBzubdwRVO8Y+RqnxW9nEJkOYXgfiDrnnDrnvc9UZX608NBHH5nLxupc5jPnHC5evKRAptlko8iLST7J8yKfbEwmk43JJJ9MJnk+yfPCrBjoKWzssEbneB6yqxp63imS8kU8+bi9pz6GA82dT+x9Q0n5EospCiSpVyF5GUS5xr73azSXk7xAwnvnnM+yLMsyn4Xjqlwu5jL1zjvnnHMeDpYXV/P1wMN5kRdFkTPPJc+RF64wMVGKxY36bHZUsS7/oXEuUl1Si5vC4psX2r1bNfoqPADN0TqHpySP9BzuHLrOpvoG6dtXGpakbHVrO8NIDoxovUgkHpevReFCPdjIomCWee/gtdTU3okWapl4lahRFBp0uVNXFEVhBcv3XAzbwvYB8GybG2zWgRGNZPIGFjZeojNd/bQYeMrrQvqAl9bJdamIbqI5IniqQXmo994577xz6tShRrRAVJ2JaSWlTkQEHgpXuMLMyjcNJZEFt1KXaRxhNTxCr3PSPICVQ0/tcrUI1VZCIGjp4G04DQrDBSyrqlfVcAg9SrKEVZmZqakpk9c6dXcmzb4vVQbOCus9aH963MjOm562lm5N550SN/WmdluTVggNV4rl+mCUCsvhNHpJ9uSo6gd328huxoQpxgKiteVHx0i8ieXwsTqf/kNEz45o51yF8eq/NUcHXFeHlQRRCuie5QSTDxFdaY+UtdOP/wcDaEZLCLxswwAAAABJRU5ErkJggg==";

// ===== Ko'p tillilik: test hujjatlari (PDF/LaTeX) 3 tilda: O'zbek, Qoraqalpoq, Rus =====
const DOC_LANGS = [
  { code:"uz", label:"UZ", full:"O'zbekcha", flag:null, flagImg:UZ_FLAG_IMG },
  { code:"qq", label:"QQ", full:"Qoraqalpoqcha", flag:null, flagImg:QQ_FLAG_IMG },
  { code:"ru", label:"РУ", full:"Русский", flag:null, flagImg:RU_FLAG_IMG },
];
// Til bayrog'ini chiqaradi: emoji mavjud bo'lsa emoji, aks holda (masalan Qoraqalpog'iston
// uchun) yuklangan rasm ishlatiladi (chunki 🇶🇶 unicode emoji sifatida mavjud emas).
function LangFlag({ lang, size }) {
  const s = size || 16;
  if (lang.flagImg) return <img src={lang.flagImg} alt={lang.label} style={{width:s*1.4,height:s,objectFit:"cover",borderRadius:3,verticalAlign:"middle",border:"1px solid rgba(0,0,0,0.08)"}}/>;
  return <span style={{fontSize:s}}>{lang.flag}</span>;
}
function emptyLangDoc() { return { docType:"pdf", pdfUrl:null, latexSource:"", latexFileName:"", latexImages:{} }; }
function emptyLangDocs() { return { uz: emptyLangDoc(), qq: emptyLangDoc(), ru: emptyLangDoc() }; }
// Testning berilgan tildagi hujjatini qaytaradi (mavjud bo'lsa). Eski (bir tilli) testlar uchun
// pdfUrl/latexSource maydonlari "uz" sifatida talqin qilinadi (orqaga moslik uchun).
function getLangDoc(test, lang) {
  if (!test) return null;
  const d = test.langDocs && test.langDocs[lang];
  if (d && (d.pdfUrl || d.latexSource)) return d;
  if (lang === "uz" && !test.langDocs && (test.pdfUrl || test.latexSource)) {
    return { docType: test.pdfUrl ? "pdf" : "latex", pdfUrl: test.pdfUrl||null, latexSource: test.latexSource||"", latexFileName: test.latexFileName||"", latexImages: test.latexImages||{} };
  }
  return null;
}
function testHasAnyDoc(test) {
  if (!test) return false;
  if (test.langDocs) return DOC_LANGS.some(l => { const d = getLangDoc(test, l.code); return d && (d.pdfUrl || d.latexSource); });
  return !!(test.pdfUrl || test.latexSource);
}
function availableDocLangs(test) {
  return DOC_LANGS.filter(l => { const d = getLangDoc(test, l.code); return d && (d.pdfUrl || d.latexSource); });
}
function formatCountdown(ms) {
  if (ms <= 0) return "Boshlanmoqda...";
  const totalMin = Math.ceil(ms / 60000); // daqiqagacha yaxlitlash — barqaror ko'rinish uchun
  const days = Math.floor(totalMin / 1440);
  const hours = Math.floor((totalMin % 1440) / 60);
  const mins = totalMin % 60;
  const parts = [];
  if (days > 0) parts.push(`${days} kun`);
  if (days > 0 || hours > 0) parts.push(`${hours} soat`);
  parts.push(`${mins} daqiqa`);
  return parts.join(" ") + " qoldi";
}

// ===== MATH CHECKER (math.js orqali) =====
// math.js CDN dan yuklanadi
let mathjs = null;
function loadMathJs(cb) {
  if (mathjs) { cb(mathjs); return; }
  if (window.math) { mathjs = window.math; cb(mathjs); return; }
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/mathjs/11.8.0/math.js";
  s.onload = () => { mathjs = window.math; cb(mathjs); };
  document.head.appendChild(s);
}

// ===== MATH CHECKER — PhotoMath-level accuracy =====

function matchParen(s, openPos) {
  let depth = 0;
  for (let k = openPos; k < s.length; k++) {
    if (s[k] === "(") depth++;
    else if (s[k] === ")") { depth--; if (depth === 0) return k; }
  }
  return -1;
}

function topLevelComma(s) {
  let depth = 0;
  for (let k = 0; k < s.length; k++) {
    if (s[k] === "(") depth++;
    else if (s[k] === ")") depth--;
    else if (s[k] === "," && depth === 0) return k;
  }
  return -1;
}

// Full recursive expand: FRAC, ROOT, SQRT, LOGBASE, trig, etc → math.js format
function toMathJs(s) {
  if (!s) return "0";
  let r = s.toString().trim().replace(/\s+/g, "");
  r = r.replace(/π/g, "pi");
  r = r.replace(/×/g, "*");
  r = r.replace(/÷/g, "/");
  r = r.replace(/−/g, "-");
  r = r.replace(/∞/g, "Infinity");
  r = r.replace(/;/g, ","); // interval separator → comma
  r = r.replace(/(\d+\.?\d*)°/g, "($1*pi/180)");
  r = r.replace(/°/g, "*pi/180");
  // Decimal comma handled inside processExpr to avoid corrupting arg separators
  return processExpr(r);
}

function processExpr(s) {
  if (!s) return "0";
  if (s.length > 50000) return "0"; // safety guard
  let result = "";
  let i = 0;
  let safety = 0;
  while (i < s.length && safety++ < 200000) {
    // FRAC(num,den) → (num)/(den)
    // MUHIM: agar FRAC'dan oldin to'g'ridan-to'g'ri raqam kelsa (masalan "3" dan keyin
    // darhol kasr, ular orasida ko'paytirish belgisisiz) — bu ARALASH SON deb hisoblanadi
    // (3 va 4/5 = 3.8), ko'paytirish emas (3×4/5 = 2.4). Agar oldin ")" (hisoblangan
    // ifoda) bo'lsa, bu haligacha ko'paytirish hisoblanadi.
    if (s.startsWith("FRAC(", i)) {
      const op = i + 4, cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const cm = topLevelComma(inner);
        const num = cm >= 0 ? inner.slice(0, cm) : inner;
        const den = cm >= 0 ? inner.slice(cm + 1) : "1";
        const lastCh0 = result.slice(-1);
        if (lastCh0 && /\d/.test(lastCh0)) result += "+";
        else if (lastCh0 === ")") result += "*";
        result += "(" + processExpr(num) + ")/(" + processExpr(den) + ")";
        i = cl + 1; continue;
      }
    }
    // SUP(base,exp) → (base)^(exp) — daraja (masalan x²)
    if (s.startsWith("SUP(", i)) {
      const op = i + 3, cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const cm = topLevelComma(inner);
        const base = cm >= 0 ? inner.slice(0, cm) : inner;
        const exp  = cm >= 0 ? inner.slice(cm + 1) : "";
        result += "(" + (processExpr(base) || "0") + ")^(" + (processExpr(exp) || "1") + ")";
        i = cl + 1; continue;
      }
    }
    // ROOT(n,x) → nthRoot(x,n)
    if (s.startsWith("ROOT(", i)) {
      const op = i + 4, cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const cm = topLevelComma(inner);
        const deg = cm >= 0 ? inner.slice(0, cm) : "2";
        const arg = cm >= 0 ? inner.slice(cm + 1) : inner;
        result += "nthRoot(" + processExpr(arg) + "," + deg + ")";
        i = cl + 1; continue;
      }
    }
    // LOG_BASE(base,arg) → log(arg,base)
    if (s.startsWith("LOG_BASE(", i)) {
      const op = i + 8, cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const cm = topLevelComma(inner);
        const base = cm >= 0 ? inner.slice(0, cm) : "10";
        const arg  = cm >= 0 ? inner.slice(cm + 1) : inner;
        const lastChL = result.slice(-1);
        if (lastChL && /[\d)]/.test(lastChL)) result += "*";
        result += "log(" + processExpr(arg) + "," + processExpr(base) + ")";
        i = cl + 1; continue;
      }
    }
    // INT(...) → 0 (integrals not evaluatable)
    if (s.startsWith("INT(", i)) {
      const op = i + 3, cl = matchParen(s, op);
      if (cl !== -1) { result += "0"; i = cl + 1; continue; }
    }
    // √(x) → sqrt(x)
    if (s.startsWith("√(", i)) {
      const op = i + 1, cl = matchParen(s, op);
      if (cl !== -1) {
        const lastChS = result.slice(-1);
        if (lastChS && /[\d)]/.test(lastChS)) result += "*";
        result += "sqrt(" + processExpr(s.slice(op + 1, cl)) + ")";
        i = cl + 1; continue;
      }
    }
    // |x| → abs(x)
    if (s[i] === "|") {
      const cl = s.indexOf("|", i + 1);
      if (cl !== -1) {
        result += "abs(" + processExpr(s.slice(i + 1, cl)) + ")";
        i = cl + 1; continue;
      }
    }
    // named functions with paren: sin(, cos(, tan(, ln(, lg(, log(, etc
    const fnMatch = s.slice(i).match(/^(arcsin|arccos|arctan|arccot|sinh|cosh|tanh|sin|cos|tan|cot|ln|lg|log|exp|sqrt|abs|nthRoot)\(/);
    if (fnMatch) {
      const fn = fnMatch[1];
      const op = i + fn.length;
      const cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const mjFn = fn === "lg" ? "log10" : fn === "ln" ? "log" : fn;
        // Implicit multiplication: 4sin( → 4*sin(
        const lastCh = result.slice(-1);
        if (lastCh && /[\d)a-zA-Z]/.test(lastCh)) result += "*";
        if (fn === "nthRoot") {
          const cm = topLevelComma(inner);
          const a1 = cm >= 0 ? inner.slice(0, cm) : inner;
          const a2 = cm >= 0 ? inner.slice(cm + 1) : "2";
          result += "nthRoot(" + processExpr(a1) + "," + processExpr(a2) + ")";
        } else if (fn === "log" && topLevelComma(inner) >= 0) {
          const cm = topLevelComma(inner);
          const a1 = inner.slice(0, cm);
          const a2 = inner.slice(cm + 1);
          result += "log(" + processExpr(a1) + "," + processExpr(a2) + ")";
        } else {
          result += mjFn + "(" + processExpr(inner) + ")";
        }
        i = cl + 1; continue;
      }
    }
    // Parenthesized group — also handle implicit mult: 2(3+1) → 2*(3+1)
    if (s[i] === "(") {
      const cl = matchParen(s, i);
      if (cl !== -1) {
        const lastCh2 = result.slice(-1);
        if (lastCh2 && /[\d)]/.test(lastCh2)) result += "*";
        result += "(" + processExpr(s.slice(i + 1, cl)) + ")";
        i = cl + 1; continue;
      }
    }
    // Default: copy character
    // Handle European decimal comma: digit,digit → digit.digit
    // BUT only when NOT a top-level argument separator (depth=0 is OK here since
    // we're already inside processExpr which handles structure)
    if (s[i] === ",") {
      const prev = s[i-1] || "";
      const next = s[i+1] || "";
      if (/\d/.test(prev) && /\d/.test(next)) {
        result += "."; i++; continue;
      }
      // Otherwise it's an argument separator — pass through
      result += ","; i++; continue;
    }
    const ch = s[i];
    const lastR = result.slice(-1);
    // Implicit multiplication (mathematical juxtaposition):
    // 2x→2*x, 2π→2*pi, 4sin→4*sin, 3pi→3*pi
    // )x→)*x, )2→)*2, )(→)*(
    // xpi→x*pi (letter+p where p starts "pi")
    const nextWord = s.slice(i);
    const startsSpecial = nextWord.startsWith("pi") || nextWord.startsWith("e");
    const needsMul = lastR !== "" && (
      // digit/letter/) followed by letter
      (/[\d)]/.test(lastR) && /[a-zA-Z√]/.test(ch)) ||
      // digit followed by π (already replaced to "p" of "pi")
      (/\d/.test(lastR) && startsSpecial) ||
      // ) followed by digit
      (lastR === ")" && /\d/.test(ch))
    );
    if (needsMul) result += "*";
    result += ch;
    i++;
  }
  return result;
}

// Smart numeric comparison using math.js
function mathEval(expr) {
  if (!window.math) return null;
  try {
    const r = window.math.evaluate(expr);
    if (typeof r === "number" && isFinite(r)) return r;
    // Complex number
    if (r && typeof r.re === "number") return r.re;
    return null;
  } catch { return null; }
}

function mathEvalScope(expr, scope) {
  if (!window.math) return null;
  try {
    const r = window.math.evaluate(expr, scope);
    if (typeof r === "number" && isFinite(r)) return r;
    return null;
  } catch { return null; }
}

// ── SET/INTERVAL EQUIVALENCE ──
// Normalizes various forms of the same mathematical set

function normalizeSet(s) {
  if (!s) return "";
  let r = s.toString().trim()
    .replace(/\s+/g, "")
    .replace(/π/g, "pi")
    .replace(/∞/g, "inf")
    .replace(/−/g, "-")
    .replace(/;/g, ",");

  // x∈R  →  (-inf,inf)  (only standalone R, not inside ROOT/FRAC etc)
  r = r.replace(/x∈R\b/gi, "(-inf,inf)");
  r = r.replace(/x∈ℝ/gi, "(-inf,inf)");
  // Standalone R as the whole answer (word boundary, not part of ROOT/etc)
  r = r.replace(/^R$/i, "(-inf,inf)");
  r = r.replace(/(?<![A-Za-z])R(?![A-Za-z])/g, "(-inf,inf)");

  // (-∞;+∞)  →  (-inf,inf)
  r = r.replace(/\(-inf,\+inf\)/g, "(-inf,inf)");
  r = r.replace(/\(-inf,inf\)/g, "(-inf,inf)");

  // Inequality → interval conversion
  // a < x < b  →  (a,b)
  // a < x ≤ b  →  (a,b]
  // a ≤ x < b  →  [a,b)
  // a ≤ x ≤ b  →  [a,b]
  r = r.replace(/([^<>≤≥]+)≤x≤([^<>≤≥]+)/g, (_, a, b) => "[" + a + "," + b + "]");
  r = r.replace(/([^<>≤≥]+)<x≤([^<>≤≥]+)/g,  (_, a, b) => "(" + a + "," + b + "]");
  r = r.replace(/([^<>≤≥]+)≤x<([^<>≤≥]+)/g,  (_, a, b) => "[" + a + "," + b + ")");
  r = r.replace(/([^<>≤≥]+)<x<([^<>≤≥]+)/g,  (_, a, b) => "(" + a + "," + b + ")");
  // With ≥: b ≥ x ≥ a  →  [a,b]
  r = r.replace(/([^<>≤≥]+)≥x≥([^<>≤≥]+)/g, (_, b, a) => "[" + a + "," + b + "]");
  r = r.replace(/([^<>≤≥]+)>x≥([^<>≤≥]+)/g,  (_, b, a) => "[" + a + "," + b + ")");
  r = r.replace(/([^<>≤≥]+)≥x>([^<>≤≥]+)/g,  (_, b, a) => "(" + a + "," + b + "]");
  r = r.replace(/([^<>≤≥]+)>x>([^<>≤≥]+)/g,  (_, b, a) => "(" + a + "," + b + ")");

  // x∈(a,b] form → normalize
  r = r.replace(/x∈/gi, "");
  r = r.replace(/∈/g, "IN");

  // Normalize spaces around brackets
  r = r.replace(/\s/g, "");
  return r.toLowerCase();
}

function setsAreEqual(c, s) {
  const nc = normalizeSet(c);
  const ns = normalizeSet(s);
  if (nc === ns) return true;

  // Also try toMathJs on normalized (for numeric bounds)
  // e.g. [sqrt(2), 3] vs [1.414..., 3]
  const evalBounds = (interval) => {
    // Extract bounds from (a,b] style
    const m = interval.match(/^([\[(])(.+),(.+)([\])])$/);
    if (!m) return null;
    const [, lb, a, b, rb] = m;
    const av = mathEval(toMathJs(a));
    const bv = mathEval(toMathJs(b));
    if (av === null || bv === null) return null;
    return lb + av.toPrecision(8) + "," + bv.toPrecision(8) + rb;
  };

  const ec = evalBounds(nc);
  const es = evalBounds(ns);
  if (ec && es && ec === es) return true;

  return false;
}

// ── AI-POWERED SEMANTIC CHECKER ──
// Used as fallback for interval notation, set theory, and other
// expressions that simple symbolic/numeric comparison can't handle.
// Calls Claude API (available in this artifact environment) to verify
// mathematical equivalence between two answer expressions.

const aiCheckCache = new Map();

async function aiCheckEquivalence(correct, student) {
  const cacheKey = correct + "|||" + student;
  if (aiCheckCache.has(cacheKey)) return aiCheckCache.get(cacheKey);

  console.log("[AI Check] Comparing:", correct, "vs", student);

  try {
    const prompt = `Siz matematik javoblarni tekshiruvchi yordamchisiz. Quyidagi ikkita matematik ifoda bir xil ma'noni anglatadimi tekshiring.

To'g'ri javob: ${correct}
O'quvchi javobi: ${student}

BELGILAR IZOHI (bular ichki yozuv formati, harfma-harf emas, MA'NOSIGA qarang):
- FRAC(a,b) — bu a/b kasr (surat/maxraj)
- √(x) — bu √x, x ning kvadrat ildizi
- Agar son va √(...) yonma-yon yozilgan bo'lsa (masalan "2√(3)"), bu KO'PAYTMA degani: 2·√3, YA'NI 2 koeffitsiyent, √3 esa alohida ildiz — ular orasida vergul yo'q, bu BITTA son (2√3 ≈ 3.464)
- x∈R — barcha haqiqiy sonlar to'plami, ya'ni (-∞;∞) bilan bir xil
- ; yoki , — interval ichidagi chegaralarni ajratuvchi belgi (interval format: (a;b], [a;b), va h.k.)
- ∞ — cheksizlik

MUHIM: Ifodalarni baholashdan oldin ularni to'g'ri parslang. Masalan "(3;2√(3)]" — bu interval bo'lib, pastki chegarasi 3, yuqori chegarasi 2√3 (≈3.464) bo'lgan, o'ngi yopiq (3 ta nuqta belgisi bilan emas, ] bilan yopiq) interval. "2√(3)" ni "2" va "√3" deb IKKITA alohida son sifatida emas, balki BITTA qiymat 2·√3 sifatida hisoblang.

Sonlarni taxminiy hisoblab solishtiring: √12 = 2√3 ≈ 3.4641. Agar ikkala ifoda xuddi shu sonlarni anglatsa — TRUE.

Bu ikkala ifoda MATEMATIK jihatdan bir xil narsani anglatadimi?

FAQAT "TRUE" yoki "FALSE" deb javob bering, boshqa hech narsa yozmang.`;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 10,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const errText = await response.text().catch(()=>"unknown");
      console.error("[AI Check] API error:", response.status, errText);
      aiCheckCache.set(cacheKey, null);
      return null;
    }

    const data = await response.json();
    const text = (data.content || []).map(b => b.text || "").join("").trim().toUpperCase();
    console.log("[AI Check] Response:", text);
    const result = text.includes("TRUE");
    aiCheckCache.set(cacheKey, result);
    return result;
  } catch (err) {
    console.error("[AI Check] Exception:", err);
    aiCheckCache.set(cacheKey, null);
    return null;
  }
}

function checkMath(correct, student) {
  if (!correct || correct.toString().trim() === "") return false;
  if (student === "" || student === undefined || student === null) return false;

  const rawC = correct.toString().trim();
  const rawS = student.toString().trim();

  // ── Normalize for comparison ──
  const norm = (x) => x
    .replace(/\s+/g, "")
    .replace(/π/g, "pi")
    .replace(/∞/g, "Infinity")
    .replace(/−/g, "-")
    .replace(/×/g, "*")
    .replace(/÷/g, "/")
    .replace(/;/g, ",")          // interval separator
    .replace(/(\d),(?=\d)/g, "$1.")  // decimal comma
    .replace(/≤/g, "<=")
    .replace(/≥/g, ">=")
    .replace(/≠/g, "!=")
    .replace(/∈/g, "in")
    .replace(/ℝ/g, "R")
    .replace(/ℤ/g, "Z")
    .replace(/ℕ/g, "N");

  const c = norm(rawC);
  const s = norm(rawS);

  // 1. Direct normalized match
  if (c === s) return true;

  // ── Interval / set equivalences ──
  // Normalize interval forms:
  // x∈R == (-∞;∞) == (-Inf,Inf) == R
  // 3<x<=5 == x∈(3;5] == (3,5]
  const normInterval = (x) => {
    let r = norm(x);
    // "x∈R" or "x in R" → "R"
    r = r.replace(/x\s*in\s*R/gi, "R");
    r = r.replace(/x\s*∈\s*R/g, "R");
    r = r.replace(/x\s*in\s*\(-Infinity,Infinity\)/gi, "R");
    r = r.replace(/\(-Infinity,Infinity\)/g, "R");
    r = r.replace(/\(-infinity,infinity\)/gi, "R");
    r = r.replace(/\(-∞[,;]∞\)/g, "R");
    r = r.replace(/\(-Inf,Inf\)/gi, "R");
    // "x∈(3;5]" → "(3,5]"
    r = r.replace(/x\s*in\s*([(\[].+?[\])])/gi, "$1");
    r = r.replace(/x\s*∈\s*([(\[].+?[\])])/g, "$1");
    // "3<x<=5" → "(3,5]"
    // Pattern: a < x <= b  or  a <= x < b etc.
    const inequalityToInterval = (str) => {
      // a < x <= b
      let m = str.match(/^(-?[\d.Infinity]+)<x<=(-?[\d.Infinity]+)$/i);
      if (m) return `(${m[1]},${m[2]}]`;
      // a <= x < b
      m = str.match(/^(-?[\d.Infinity]+)<=x<(-?[\d.Infinity]+)$/i);
      if (m) return `[${m[1]},${m[2]})`;
      // a < x < b
      m = str.match(/^(-?[\d.Infinity]+)<x<(-?[\d.Infinity]+)$/i);
      if (m) return `(${m[1]},${m[2]})`;
      // a <= x <= b
      m = str.match(/^(-?[\d.Infinity]+)<=x<=(-?[\d.Infinity]+)$/i);
      if (m) return `[${m[1]},${m[2]}]`;
      // x < b  (one-sided)
      m = str.match(/^x<(-?[\d.Infinity]+)$/i);
      if (m) return `(-Infinity,${m[1]})`;
      m = str.match(/^x<=(-?[\d.Infinity]+)$/i);
      if (m) return `(-Infinity,${m[1]}]`;
      m = str.match(/^x>(-?[\d.Infinity]+)$/i);
      if (m) return `(${m[1]},Infinity)`;
      m = str.match(/^x>=(-?[\d.Infinity]+)$/i);
      if (m) return `[${m[1]},Infinity)`;
      return str;
    };
    r = inequalityToInterval(r);
    // Normalize Infinity spellings
    r = r.replace(/\+?Infinity/gi, "Inf").replace(/-Infinity/gi, "-Inf");
    r = r.replace(/inf\b/gi, "Inf");
    return r.toLowerCase().trim();
  };

  const ic = normInterval(rawC);
  const is2 = normInterval(rawS);

  if (ic === is2) return true;

  // ── Numeric evaluation (existing logic) ──
  const EPS = 1e-6;
  const mathC = toMathJs(rawC);
  const mathS = toMathJs(rawS);

  if (mathC === mathS) return true;

  const cv = mathEval(mathC);
  const sv = mathEval(mathS);
  if (cv !== null && sv !== null) return Math.abs(cv - sv) < EPS;

  const diff = mathEval("(" + mathC + ")-(" + mathS + ")");
  if (diff !== null) return Math.abs(diff) < EPS;

  // Multi-value symbolic
  const testVals = [
    {x:1,n:1,t:1,a:1,b:1},{x:2,n:2,t:2,a:2,b:3},
    {x:0.5,n:3,t:0.5,a:0.5,b:2},{x:Math.PI/4,n:4,t:Math.PI/6,a:Math.PI,b:Math.E},
  ];
  let allMatch = true, tried = false;
  for (const scope of testVals) {
    const cv2 = mathEvalScope(mathC, scope);
    const sv2 = mathEvalScope(mathS, scope);
    if (cv2 === null || sv2 === null) { allMatch = false; break; }
    tried = true;
    if (Math.abs(cv2 - sv2) > EPS) { allMatch = false; break; }
  }
  if (tried && allMatch) return true;

  // DEG trig fallback
  if (window.math) {
    const hasTrig = /\b(sin|cos|tan|cot)\(/.test(mathC) || /\b(sin|cos|tan|cot)\(/.test(mathS);
    if (hasTrig) {
      try {
        const degConv = (e) => e
          .replace(/\bsin\(([^)]+)\)/g, "sin(($1)*pi/180)")
          .replace(/\bcos\(([^)]+)\)/g, "cos(($1)*pi/180)")
          .replace(/\btan\(([^)]+)\)/g, "tan(($1)*pi/180)");
        const cv3 = mathEval(degConv(mathC)), sv3 = mathEval(degConv(mathS));
        if (cv3 !== null && sv3 !== null && Math.abs(cv3 - sv3) < EPS) return true;
        const cv4 = mathEval(mathC), sv4 = mathEval(degConv(mathS));
        if (cv4 !== null && sv4 !== null && Math.abs(cv4 - sv4) < EPS) return true;
        const cv5 = mathEval(degConv(mathC)), sv5 = mathEval(mathS);
        if (cv5 !== null && sv5 !== null && Math.abs(cv5 - sv5) < EPS) return true;
      } catch {}
    }
  }

  // JS fallback
  try {
    const jsEv = (e) => {
      let r = e
        .replace(/\bsqrt\(/g, "Math.sqrt(")
        .replace(/\bnthRoot\(([^,]+),([^)]+)\)/g, "Math.pow($2,1/($1))")
        .replace(/\babs\(/g, "Math.abs(")
        .replace(/\blog10\(/g, "Math.log10(")
        .replace(/\blog\(([^,)]+)\)/g, "Math.log10($1)")
        .replace(/\blog\(([^,]+),([^)]+)\)/g, "(Math.log($1)/Math.log($2))")
        .replace(/\bsin\(/g, "Math.sin(")
        .replace(/\bcos\(/g, "Math.cos(")
        .replace(/\btan\(/g, "Math.tan(")
        .replace(/\bpi\b/g, "Math.PI")
        .replace(/\be\b/g, "Math.E")
        .replace(/\^/g, "**");
      return Function('"use strict";const x=1,n=1,t=1,a=1,b=1;return(' + r + ')')();
    };
    const jcv = jsEv(mathC), jsv = jsEv(mathS);
    if (!isNaN(jcv) && !isNaN(jsv) && isFinite(jcv) && isFinite(jsv)) {
      return Math.abs(jcv - jsv) < EPS;
    }
  } catch {}

  return false;
}

// Async version: tries fast symbolic/numeric check first,
// falls back to AI semantic check for set/interval expressions
async function checkMathAsync(correct, student) {
  if (checkMath(correct, student)) return true;

  const looksLikeSet = (s) => {
    const x = (s || "").toString();
    return /[\[\](){}]/.test(x) || /∈/.test(x) || /\bR\b/.test(x) || /;/.test(x) ||
           (/[<>≤≥]/.test(x) && (x.match(/[<>≤≥]/g) || []).length >= 2);
  };

  if (looksLikeSet(correct) || looksLikeSet(student)) {
    // Try symbolic set equality first (fast, no network needed)
    if (setsAreEqual(correct, student)) return true;
    // Fall back to AI for complex cases
    const aiResult = await aiCheckEquivalence(correct, student);
    if (aiResult === true) return true;
  }

  return false;
}

// Also used for real-time display in calculator
function normForMathJs(s) { return toMathJs(s); }

// ===== QAYTA BAHOLASH =====
// Test topshirilganda javoblar (o'quvchining tanlagan/yozgan xom javoblari) doim
// saqlanadi. Admin keyinroq to'g'ri javob kalitini tuzatsa, shu xom javoblarni
// YANGI kalit bilan qayta solishtirib, natijalarni to'g'rilash mumkin.
async function regradeTestResults(test) {
  const all = db.get("results") || [];
  let changed = 0;
  const updated = [];
  for (const r of all) {
    if (r.testId !== test.id) { updated.push(r); continue; }
    const scores = {}, subScores = {};
    let total = 0;
    for (let idx = 0; idx < test.questions.length; idx++) {
      const q = test.questions[idx];
      if (q.type === "closed") {
        const ok = r.answers?.[idx] !== undefined && r.answers[idx] === q.correctAnswer;
        scores[idx] = ok; if (ok) total++;
      } else if (q.subParts?.length > 0) {
        subScores[idx] = {};
        for (let si = 0; si < q.subParts.length; si++) {
          const sp = q.subParts[si];
          const ok = await checkMathAsync(sp.answer, r.subAnswers?.[idx]?.[si] || "");
          subScores[idx][si] = ok; if (ok) total++;
        }
      } else {
        const ok = await checkMathAsync(q.correctAnswer, r.openAnswers?.[idx] || "");
        scores[idx] = ok; if (ok) total++;
      }
    }
    changed++;
    updated.push({ ...r, scores, subScores, totalScore: total });
  }
  db.set("results", updated);
  return changed;
}

// ===== RASH (RASCH) MODELI =====
// Yuklangan "Rash_model_shablon.xlsx" jadvalidagi formulalarga aynan mos:
// theta (qobiliyat logiti) -> Z-ball (sinf o'rtachasiga nisbatan) -> BALL (markaz atrofida,
// lekin maksimal balldan hech qachon oshmaydigan asimptotik tanh formula) -> Daraja (harf baho).
const DEFAULT_RASCH_SETTINGS = {
  maxScore: 90.1,   // Maksimal ball (chegara)
  center: 50,       // Markaz (o'rtacha) ball
  ncMax: 45,        // NC chegarasi (bundan past)
  cMax: 50,         // C chegarasi
  cPlusMax: 55,     // C+ chegarasi
  bMax: 60,         // B chegarasi
  bPlusMax: 65,     // B+ chegarasi
  aMax: 70,         // A chegarasi (bundan yuqori = A+)
};
// Testdagi umumiy baholanadigan "punkt"lar soni (har bir sub-qism alohida punkt hisoblanadi —
// bu result.totalScore hisoblangan usul bilan bir xil bo'lishi kerak).
function testTotalItems(test) {
  let n = 0;
  for (const q of test.questions || []) n += q.subParts?.length > 0 ? q.subParts.length : 1;
  return n;
}
// Rash modeli logit (qobiliyat) qiymati: to'g'ri javoblar sonini 0.5..total-0.5 oralig'ida
// "clamp" qilib (0% yoki 100% cheksizlikka aylanib qolmasligi uchun), logit-ga aylantiradi.
function raschTheta(correct, total) {
  if (!total || total <= 0) return null;
  const c = Math.min(Math.max(correct, 0.5), total - 0.5);
  return Math.log(c / (total - c));
}
function raschGrade(ball, s) {
  if (ball < s.ncMax) return "NC";
  if (ball < s.cMax) return "C";
  if (ball < s.cPlusMax) return "C+";
  if (ball < s.bMax) return "B";
  if (ball < s.bPlusMax) return "B+";
  if (ball < s.aMax) return "A";
  return "A+";
}
// Bitta testning barcha topshirilgan natijalari uchun Rash modelini hisoblab, natijalarga
// "rasch" maydonini qo'shib db'ga saqlaydi. O'quvchi bu natijani keyinroq o'z profilida ko'radi.
function computeRaschForTest(test, settings = DEFAULT_RASCH_SETTINGS) {
  const all = db.get("results") || [];
  const total = testTotalItems(test);
  const idxList = [], thetas = [];
  all.forEach((r, i) => {
    if (r.testId !== test.id) return;
    const th = raschTheta(r.totalScore || 0, total);
    if (th === null || !isFinite(th)) return;
    idxList.push(i); thetas.push(th);
  });
  if (thetas.length === 0) return { count: 0 };
  const mean = thetas.reduce((a, b) => a + b, 0) / thetas.length;
  const variance = thetas.reduce((a, b) => a + (b - mean) ** 2, 0) / thetas.length; // STDEVP (populyatsiya)
  const std = Math.sqrt(variance);
  const range = settings.maxScore - settings.center;
  const updated = [...all];
  const calculatedAt = Date.now();
  idxList.forEach((i, k) => {
    const theta = thetas[k];
    const z = std !== 0 ? (theta - mean) / std : 0;
    const ball = range !== 0 ? settings.center + range * Math.tanh(10 * z / range) : settings.center;
    const daraja = raschGrade(ball, settings);
    updated[i] = { ...updated[i], rasch: { correct: updated[i].totalScore || 0, total, theta, zBall: z, ball, daraja, calculatedAt, settings } };
  });
  db.set("results", updated);
  return { count: thetas.length, mean, std, calculatedAt };
}

// ===== RASH NATIJALARINI EXCEL FAYLDAN YUKLASH (import) =====
// Admin/o'qituvchi natijalarni (masalan "Rash_model_shablon.xlsx" ga o'xshash, yoki o'zimiz
// eksport qilgan faylga BALL/Daraja ustunlari qo'shilgan holini) tashqarida to'ldirib qaytadan
// saytga yuklashi mumkin. Ustun nomlarini moslashuvchan (regex) tarzda topamiz, chunki fayl
// shablon nomlaridan biroz farq qilishi mumkin.
function parseUploadedResultsSheet(workbook) {
  const sheetName = workbook.SheetNames.includes("Natijalar") ? "Natijalar" : workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws) return { rows: [] };
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  let headerRowIdx = -1, cols = null;
  for (let r = 0; r < Math.min(10, aoa.length); r++) {
    const row = (aoa[r] || []).map(c => String(c ?? "").trim());
    const nameIdx = row.findIndex(c => /ism|f\.?\s*i\.?\s*o|f\.?\s*i\.?\s*sh/i.test(c));
    if (nameIdx >= 0) {
      headerRowIdx = r;
      cols = {
        name: nameIdx,
        group: row.findIndex(c => /guruh/i.test(c)),
        ball: row.findIndex(c => /^ball$/i.test(c) || /^jami\s*ball$/i.test(c)),
        daraja: row.findIndex(c => /daraja/i.test(c)),
        theta: row.findIndex(c => /theta/i.test(c)),
        zball: row.findIndex(c => /z[\s-]?ball/i.test(c)),
        correct: row.findIndex(c => /^to.?g.?ri$/i.test(c)),
        total: row.findIndex(c => /jami\s*savol/i.test(c)),
        sCols: [],
      };
      row.forEach((c, ci) => { if (/^savol\s*\d+/i.test(c) || /^\d+(\.0)?$/.test(c)) cols.sCols.push(ci); });
      break;
    }
  }
  if (headerRowIdx === -1) return { rows: [] };
  const rows = [];
  for (let r = headerRowIdx + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const name = String(row[cols.name] ?? "").trim();
    if (!name) continue;
    const rec = { name, group: cols.group >= 0 ? String(row[cols.group] ?? "").trim() : "" };
    const numOrNull = (v) => (v === "" || v === null || v === undefined || isNaN(Number(v))) ? null : Number(v);
    if (cols.ball >= 0) { const v = numOrNull(row[cols.ball]); if (v !== null) rec.ball = v; }
    if (cols.daraja >= 0 && row[cols.daraja]) rec.daraja = String(row[cols.daraja]).trim();
    if (cols.theta >= 0) { const v = numOrNull(row[cols.theta]); if (v !== null) rec.theta = v; }
    if (cols.zball >= 0) { const v = numOrNull(row[cols.zball]); if (v !== null) rec.zBall = v; }
    if (cols.correct >= 0) { const v = numOrNull(row[cols.correct]); if (v !== null) rec.correct = v; }
    if (cols.total >= 0) { const v = numOrNull(row[cols.total]); if (v !== null) rec.total = v; }
    if (cols.sCols.length) {
      let c = 0, t = 0;
      cols.sCols.forEach(ci => { const v = row[ci]; if (v === 0 || v === 1 || v === "0" || v === "1") { t++; if (Number(v) === 1) c++; } });
      if (t > 0) { if (rec.correct === undefined) rec.correct = c; if (rec.total === undefined) rec.total = t; }
    }
    rows.push(rec);
  }
  return { rows };
}
// Ismi bo'yicha "users" ro'yxati bilan solishtirib, tegishli natijaga rasch ma'lumotini yozadi.
// Fayl BALL ustunini o'z ichiga olsa — o'sha qiymatlar TO'G'RIDAN-TO'G'RI ishlatiladi.
// Aks holda (faqat To'g'ri/Jami savol yoki S1..S55 ustunlari bo'lsa) — Rash formulasi
// aynan shu fayldagi o'quvchilar populyatsiyasi asosida (Excel'dagidek) hisoblanadi.
// Fayldagi qatorlar uchun Rash modelini SOF hisoblaydi — hech qanday saytdagi
// ro'yxatdan o'tgan foydalanuvchi yoki testga bog'liq emas. Agar faylda tayyor BALL ustuni
// bo'lsa o'shani ishlatadi, aks holda To'g'ri/Jami (yoki S1..Sn) ustunlaridan Rash formulasi
// bilan (shu faylning o'z populyatsiyasi asosida) hisoblab chiqadi.
function computeRaschFromFileRows(rows, settings = DEFAULT_RASCH_SETTINGS) {
  const hasDirectBall = rows.some(r => typeof r.ball === "number");
  if (hasDirectBall) {
    return rows.filter(r => typeof r.ball === "number" && isFinite(r.ball)).map(r => ({ ...r, daraja: r.daraja || raschGrade(r.ball, settings) }));
  }
  const withCorrect = rows.filter(r => typeof r.correct === "number" && typeof r.total === "number" && r.total > 0);
  const thetas = withCorrect.map(r => raschTheta(r.correct, r.total));
  const mean = thetas.length ? thetas.reduce((a, b) => a + b, 0) / thetas.length : 0;
  const variance = thetas.length ? thetas.reduce((a, b) => a + (b - mean) ** 2, 0) / thetas.length : 0;
  const std = Math.sqrt(variance);
  const range = settings.maxScore - settings.center;
  return withCorrect.map((r, i) => {
    const theta = thetas[i];
    const z = std !== 0 ? (theta - mean) / std : 0;
    const ball = range !== 0 ? settings.center + range * Math.tanh(10 * z / range) : settings.center;
    return { ...r, theta, zBall: z, ball, daraja: raschGrade(ball, settings) };
  });
}
// Bitta testni saytda topshirganlar (results) VA yuklangan Excel fayldagi qatorlarni BITTA
// umumiy Rash populyatsiyasi sifatida birlashtirib hisoblaydi (o'rtacha/standart chetlanish
// ikkalasi bo'yicha ham hisoblanadi — ayri-ayri emas). Fayldagi ism saytdagi biror
// ro'yxatdan o'tgan o'quvchiga mos kelsa, o'sha bitta yozuv sifatida qo'shiladi (ikki marta
// hisoblanmaydi) va uning profiliga yoziladi. Saytdagi mos kelmagan (faylda yo'q) natijalar ham
// hisoblashga kiradi. Faylda bor-u, saytda profili topilmagan o'quvchilar esa hisobga kiradi,
// lekin natija hech kimning profiliga yozilmaydi — alohida ro'yxat (fileOnly) qilib qaytariladi.
function importRaschFromRows(test, rows, settings = DEFAULT_RASCH_SETTINGS) {
  const all = db.get("results") || [];
  const users = db.get("users") || [];
  // Fayldagi ism ustuniga ko'pincha qo'shimcha narsalar yopishtirilgan bo'ladi:
  // "Safarov Abbos (A. Sultanov)", "Omonboyeva Charos Math@32", "Baxshulloev Usmonbek  " kabi.
  // Shuning uchun avval qavs ichidagi va harf bo'lmagan belgilarni tozalaymiz,
  // so'ng FAQAT ismning birinchi ikkita so'ziga (F.I.Sh) qarab moslashtiramiz.
  const clean = (s) => (s || "")
    .replace(/\([^)]*\)/g, " ")        // (guruh/o'qituvchi nomi) — olib tashlanadi
    .replace(/[^\p{L}\s'’-]/gu, " ")   // raqam, @ va boshqa belgilar — olib tashlanadi
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  const nameToPhone = {};
  users.forEach(u => {
    nameToPhone[clean(`${u.firstName} ${u.lastName}`)] = u.phone;
    nameToPhone[clean(`${u.lastName} ${u.firstName}`)] = u.phone;
  });
  const findPhone = (rawName) => {
    const full = clean(rawName);
    if (nameToPhone[full]) return nameToPhone[full];
    const words = full.split(" ").filter(Boolean);
    if (words.length >= 2) {
      const a = words[0] + " " + words[1], b = words[1] + " " + words[0];
      if (nameToPhone[a]) return nameToPhone[a];
      if (nameToPhone[b]) return nameToPhone[b];
    }
    return null;
  };

  const total = testTotalItems(test);
  const combined = [];      // {name, correct, total, theta, siteIdx, directBall, directDaraja}
  const siteIdxUsed = new Set();
  const unusable = [];      // fayldan kelgan, hisoblab bo'lmaydigan (na ball, na to'g'ri/jami) qatorlar

  // 1) Fayldagi har bir qator — imkon bo'lsa saytdagi mos natijaga bog'lanadi
  rows.forEach(rec => {
    const correct = rec.correct, tot = rec.total;
    const hasRaw = typeof correct === "number" && typeof tot === "number" && tot > 0;
    if (!hasRaw && typeof rec.ball !== "number") { unusable.push(rec.name); return; }
    const phone = findPhone(rec.name);
    let siteIdx = -1;
    if (phone) siteIdx = all.findIndex(r => r.testId === test.id && r.userPhone === phone);
    if (siteIdx >= 0) siteIdxUsed.add(siteIdx);
    if (hasRaw) {
      const th = raschTheta(correct, tot);
      if (th !== null && isFinite(th)) { combined.push({ name: rec.name, correct, total: tot, theta: th, siteIdx }); return; }
    }
    // faqat tayyor BALL berilgan (xom to'g'ri/jami yo'q) — theta hisoblab bo'lmaydi, ball o'zi saqlanadi
    combined.push({ name: rec.name, correct: null, total: null, theta: null, siteIdx, directBall: rec.ball, directDaraja: rec.daraja });
  });

  // 2) Saytda shu testni topshirgan, lekin fayl orqali "yangilanmagan" (faylda yo'q) o'quvchilar
  all.forEach((r, idx) => {
    if (r.testId !== test.id || siteIdxUsed.has(idx)) return;
    const correct = r.totalScore || 0;
    const th = raschTheta(correct, total);
    if (th === null || !isFinite(th)) return;
    combined.push({ name: null, correct, total, theta: th, siteIdx: idx });
  });

  // 3) Umumiy o'rtacha/standart chetlanish — FAQAT theta hisoblangan yozuvlar bo'yicha (sayt + fayl birga)
  const withTheta = combined.filter(e => e.theta !== null && e.theta !== undefined);
  const mean = withTheta.length ? withTheta.reduce((a, b) => a + b.theta, 0) / withTheta.length : 0;
  const variance = withTheta.length ? withTheta.reduce((a, b) => a + (b.theta - mean) ** 2, 0) / withTheta.length : 0;
  const std = Math.sqrt(variance);
  const range = settings.maxScore - settings.center;
  const calculatedAt = Date.now();

  let matched = 0;
  const fileOnly = [];
  const updated = all.map(r => ({ ...r }));
  combined.forEach(e => {
    let ball, zBall = null, theta = e.theta ?? null;
    if (theta !== null) {
      zBall = std !== 0 ? (theta - mean) / std : 0;
      ball = range !== 0 ? settings.center + range * Math.tanh(10 * zBall / range) : settings.center;
    } else {
      ball = e.directBall;
    }
    const daraja = e.directDaraja || (typeof ball === "number" ? raschGrade(ball, settings) : null);
    if (e.siteIdx >= 0) {
      matched++;
      updated[e.siteIdx] = { ...updated[e.siteIdx], rasch: {
        correct: e.correct ?? updated[e.siteIdx].totalScore ?? null,
        total: e.total ?? total,
        theta, zBall, ball, daraja, calculatedAt, settings,
        source: e.name ? "site+upload" : "site",
      }};
    } else if (e.name) {
      fileOnly.push({ name: e.name, correct: e.correct, total: e.total, theta, zBall, ball, daraja });
    }
  });
  db.set("results", updated);
  return { matched, fileOnly, unusable, calculatedAt, combinedCount: combined.length };
}

// ===== EXCEL EXPORT (SheetJS .xlsx) =====

// Ma'lumotni tayyorlab, DATA URL qaytaradi (avtomatik yuklab olishga urinmaydi).
// Sabab: ko'plab embedded WebView muhitlar (masalan Telegram Mini App) dasturiy
// ravishda (.click()) boshlangan yuklab olishlarni bloklaydi, lekin foydalanuvchi
// O'ZI bosgan havolaga (haqiqiy user gesture) ruxsat beradi.
// Mustaqil "Excel -> Rash" hisob-kitobi natijasini yuklab olinadigan .xlsx qilib tayyorlaydi.
function buildRaschCalcExport(rows, settings) {
  const header = ["#","F.I.O","Guruh","To'g'ri","Jami","Theta (θ)","Z-ball","BALL","Daraja"];
  const sorted = [...rows].sort((a,b)=>(b.ball??-999)-(a.ball??-999));
  const data = sorted.map((r,i)=>[i+1, r.name||"", r.group||"", r.correct??"", r.total??"", r.theta!=null?+r.theta.toFixed(4):"", r.zBall!=null?+r.zBall.toFixed(4):"", r.ball!=null?+r.ball.toFixed(2):"", r.daraja||""]);
  if (typeof XLSX !== "undefined") {
    try {
      const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let C2 = range.s.c; C2 <= range.e.c; C2++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C2 });
        if (ws[addr]) ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: "6D28D9" } } };
      }
      ws["!cols"] = header.map((h,i)=>({ wch: i===1?26:12 }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Rash natijalari");
      const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
      return { dataUrl: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`, filename: "rash_natijalari.xlsx", isBinary: true };
    } catch {}
  }
  return null;
}

function buildExcelExport(test, results, users) {
  if (!test) return null;
  // Build question column list
  const allQ = [];
  test.questions.forEach((q, i) => {
    if (q.subParts?.length > 0) {
      q.subParts.forEach((_, si) => allQ.push({ label: `Savol${i+1}-${String.fromCharCode(97+si)}`, qIdx: i, sub: si }));
    } else {
      allQ.push({ label: `Savol${i+1}`, qIdx: i, sub: null });
    }
  });

  // Header row
  const header = ["T/r", "F.I.O", "Guruh", "Jami Ball", ...allQ.map(q => q.label)];

  // Data rows
  const rows = results.filter(r => r.testId === test.id)
    .sort((a, b) => b.totalScore - a.totalScore)
    .map((r, i) => {
      const u = users.find(u => u.phone === r.userPhone);
      let total = 0;
      const cols = allQ.map(({ qIdx, sub }) => {
        const v = sub !== null
          ? (r.subScores?.[qIdx]?.[sub] ? 1 : 0)
          : (r.scores?.[qIdx] ? 1 : 0);
        total += v;
        return v;
      });
      return [i + 1, u ? `${u.firstName} ${u.lastName}` : r.userPhone, u?.group || "-", total, ...cols];
    });

  // Try SheetJS if available, otherwise fall back to TSV
  if (typeof XLSX !== "undefined") {
    try {
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      // Style header row
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let C2 = range.s.c; C2 <= range.e.c; C2++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C2 });
        if (!ws[addr]) continue;
        ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: "4F6EF7" } } };
      }
      // Column widths
      ws["!cols"] = header.map((h, i) => ({ wch: i < 4 ? Math.max(h.length + 4, 12) : 10 }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Natijalar");
      const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
      return {
        dataUrl: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`,
        filename: `${test.name}_natijalari.xlsx`,
        isBinary: true,
      };
    } catch (e) { console.error("[Excel export]", e); }
  }
  // TSV fallback — opens cleanly in Excel (tab-separated), va nusxalab ham bo'ladi
  const tsv = [header, ...rows].map(row =>
    row.map(cell => String(cell).replace(/\t/g, " ")).join("\t")
  ).join("\n");
  const b64 = btoa(unescape(encodeURIComponent("\uFEFF" + tsv)));
  return {
    dataUrl: `data:text/tab-separated-values;charset=utf-8;base64,${b64}`,
    filename: `${test.name}_natijalari.tsv`,
    tsv,
    isBinary: false,
  };
}

// ===== COLORS & STYLES =====
const C = {
  bg:"#F0F4FF", card:"#FFFFFF", border:"#DDE3F0",
  primary:"#4F6EF7", primaryDark:"#3A56D4", primaryLight:"#EEF1FF",
  success:"#22C55E", successLight:"#F0FDF4", successDark:"#15803D",
  danger:"#EF4444", dangerLight:"#FEF2F2",
  warning:"#F59E0B", warningLight:"#FFFBEB",
  text:"#1E293B", textMid:"#64748B", textLight:"#94A3B8",
  purple:"#8B5CF6", purpleLight:"#F5F3FF",
};
const S = {
  page:{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:"'Segoe UI',system-ui,sans-serif", overflowX:"hidden", width:"100%", boxSizing:"border-box" },
  card:{ background:C.card, borderRadius:16, border:`1px solid ${C.border}`, boxShadow:"0 2px 12px rgba(79,110,247,0.07)" },
  input:{ width:"100%", padding:"11px 14px", background:C.card, border:`1.5px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:14, outline:"none", boxSizing:"border-box", marginBottom:10 },
  label:{ display:"block", color:C.textMid, fontSize:12, fontWeight:700, marginBottom:5, textTransform:"uppercase", letterSpacing:.5 },
  btnPrimary:{ width:"100%", padding:"12px", background:C.primary, border:"none", borderRadius:10, color:"white", fontSize:15, fontWeight:700, cursor:"pointer" },
  btnSuccess:{ padding:"10px 20px", background:C.success, border:"none", borderRadius:10, color:"white", fontSize:14, fontWeight:700, cursor:"pointer" },
  btnDanger:{ padding:"9px 16px", background:C.danger, border:"none", borderRadius:10, color:"white", fontSize:13, fontWeight:600, cursor:"pointer" },
  btnGhost:{ padding:"9px 16px", background:"transparent", border:`1.5px solid ${C.border}`, borderRadius:10, color:C.text, fontSize:13, fontWeight:600, cursor:"pointer" },
  btnSmall:{ padding:"7px 14px", border:"none", borderRadius:8, color:"white", fontSize:13, fontWeight:600, cursor:"pointer" },
  err:{ background:C.dangerLight, border:`1px solid #FECACA`, borderRadius:10, padding:"10px 14px", color:C.danger, fontSize:13, marginBottom:12 },
  badge:{ display:"inline-block", padding:"3px 10px", borderRadius:999, fontSize:12, fontWeight:700 },
  table:{ width:"100%", borderCollapse:"collapse", fontSize:14 },
  th:{ padding:"10px 14px", background:C.primaryLight, color:C.primary, textAlign:"left", fontWeight:700, borderBottom:`2px solid ${C.border}` },
  td:{ padding:"10px 14px", borderBottom:`1px solid ${C.border}`, color:C.text },
  empty:{ textAlign:"center", color:C.textLight, padding:"48px 0", fontSize:15 },
};

// ===== NODE TREE HELPERS =====
function uid() { return Math.random().toString(36).slice(2); }
function textNode(v="") { return { id: uid(), type: "text", value: v, nodes: null }; }
function slotNode(nodes) { const ns = nodes || null; return { id: uid(), type: "slot", value: "", nodes: ns !== null ? ns : [{ id: uid(), type: "text", value: "", nodes: null }] }; }
function createNode(type, value="", children=[]) {
  return { id: uid(), type, value, children };
}

function slotVal(slot) {
  if (!slot) return "";
  if (slot.nodes) return nodesToString(slot.nodes);
  return slot.value || "";
}
function nodeVal(n) {
  if (!n) return "";
  if (n.type === "text") return n.value || "";
  if (n.type === "slot") return slotVal(n);
  return nodesToString([n]);
}
function nodesToString(nodes) {
  if (!nodes) return "";
  return nodes.map(n => {
    if (n.type === "text") return n.value || "";
    if (n.type === "slot") return slotVal(n);
    if (n.type === "frac") { const num=slotVal(n.children[0]); const den=slotVal(n.children[1]); return `FRAC(${num},${den})`; }
    if (n.type === "sqrt") return `√(${slotVal(n.children[0])})`;
    if (n.type === "nthroot") { const deg=slotVal(n.children[0])||"n"; const arg=slotVal(n.children[1]); return `ROOT(${deg},${arg})`; }
    if (n.type === "sup") { const b=slotVal(n.children[0]); const e=slotVal(n.children[1]); return `SUP(${b},${e})`; }
    if (n.type === "abs") return `|${slotVal(n.children[0])}|`;
    if (n.type === "mixedfrac") {
      const w=slotVal(n.children[0]); const nm=slotVal(n.children[1]); const dn=slotVal(n.children[2]);
      // Mixed number: whole + fraction (e.g. 2¾ = 2 + 3/4), handle negative whole part too
      if (w && w.trim().startsWith("-")) {
        return "(" + w + "-FRAC(" + nm + "," + dn + "))";
      }
      return w ? "(" + w + "+FRAC(" + nm + "," + dn + "))" : "FRAC(" + nm + "," + dn + ")";
    }
    if (n.type === "defint") { const lo=slotVal(n.children[0]); const hi=slotVal(n.children[1]); const ex=slotVal(n.children[2]); const va=slotVal(n.children[3])||"x"; return "INT("+lo+","+hi+","+ex+","+va+")"; }
    if (n.type === "logbase") { const b=slotVal(n.children[0]); const a=slotVal(n.children[1]); return `LOG_BASE(${b},${a})`; }
    if (n.type === "func") return `${n.value||""}(${slotVal(n.children[0])})`;
    return n.value || "";
  }).join("");
}

// Find matching closing paren, respecting nested parens
function findClose(str, openPos) {
  let depth = 0;
  for (let k = openPos; k < str.length; k++) {
    if (str[k] === "(") depth++;
    else if (str[k] === ")") { depth--; if (depth === 0) return k; }
  }
  return -1;
}

// Find top-level comma (not inside nested parens)
function findTopComma(str) {
  let depth = 0;
  for (let k = 0; k < str.length; k++) {
    if (str[k] === "(") depth++;
    else if (str[k] === ")") depth--;
    else if (str[k] === "," && depth === 0) return k;
  }
  return -1;
}

function parseToNodes(str) {
  if (!str) return [{ id: uid(), type: "text", value: "", nodes: null }];
  const result = [];
  let i = 0;
  let parseNodesSafety = 0;
  while (i < str.length && parseNodesSafety++ < 100000) {
    // FRAC(num,den)
    if (str.startsWith("FRAC(", i)) {
      const openParen = i + 4;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const inner = str.slice(openParen + 1, close);
        const comma = findTopComma(inner);
        const num = comma >= 0 ? inner.slice(0, comma) : inner;
        const den = comma >= 0 ? inner.slice(comma + 1) : "";
        result.push(createNode("frac", "", [slotNode(parseToNodes(num)), slotNode(parseToNodes(den))]));
        i = close + 1; continue;
      }
    }
    // SUP(base,exp) — daraja (masalan x²)
    if (str.startsWith("SUP(", i)) {
      const openParen = i + 3;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const inner = str.slice(openParen + 1, close);
        const comma = findTopComma(inner);
        const base = comma >= 0 ? inner.slice(0, comma) : inner;
        const exp  = comma >= 0 ? inner.slice(comma + 1) : "";
        result.push(createNode("sup", "", [slotNode(parseToNodes(base)), slotNode(parseToNodes(exp))]));
        i = close + 1; continue;
      }
    }
    // LOG_BASE(base,arg)
    if (str.startsWith("LOG_BASE(", i)) {
      const openParen = i + 8;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const inner = str.slice(openParen + 1, close);
        const comma = findTopComma(inner);
        const base = comma >= 0 ? inner.slice(0, comma) : inner;
        const arg  = comma >= 0 ? inner.slice(comma + 1) : "";
        result.push(createNode("logbase", "", [slotNode(parseToNodes(base)), slotNode(parseToNodes(arg))]));
        i = close + 1; continue;
      }
    }
    // ROOT(deg,arg)
    if (str.startsWith("ROOT(", i)) {
      const openParen = i + 4;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const inner = str.slice(openParen + 1, close);
        const comma = findTopComma(inner);
        const deg = comma >= 0 ? inner.slice(0, comma) : "n";
        const arg = comma >= 0 ? inner.slice(comma + 1) : "";
        result.push(createNode("nthroot", "", [slotNode(parseToNodes(deg)), slotNode(parseToNodes(arg))]));
        i = close + 1; continue;
      }
    }
    // INT(lo,hi,expr,var)
    if (str.startsWith("INT(", i)) {
      const openParen = i + 3;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const inner = str.slice(openParen + 1, close);
        const parts = [];
        let rem = inner, off = 0;
        for (let p = 0; p < 4; p++) {
          const c2 = p < 3 ? findTopComma(rem) : -1;
          if (c2 >= 0) { parts.push(rem.slice(0, c2)); rem = rem.slice(c2 + 1); }
          else { parts.push(rem); rem = ""; }
        }
        const xSlot = slotNode(parseToNodes(parts[3] || "x"));
        result.push(createNode("defint", "", [slotNode(parseToNodes(parts[0]||"")), slotNode(parseToNodes(parts[1]||"")), slotNode(parseToNodes(parts[2]||"")), xSlot]));
        i = close + 1; continue;
      }
    }
    // √(arg)
    if (str.startsWith("√(", i)) {
      const openParen = i + 1;
      const close = findClose(str, openParen);
      if (close !== -1) {
        const arg = str.slice(openParen + 1, close);
        result.push(createNode("sqrt", "", [slotNode(parseToNodes(arg))]));
        i = close + 1; continue;
      }
    }
    // |expr| absolute value
    if (str[i] === "|") {
      const close = str.indexOf("|", i + 1);
      if (close !== -1) {
        const inner = str.slice(i + 1, close);
        result.push(createNode("abs", "", [slotNode(parseToNodes(inner))]));
        i = close + 1; continue;
      }
    }
    // Plain text — collect until next special token
    let j = i + 1;
    let jSafety2 = 0;
    while (j < str.length && jSafety2++ < 100000) {
      const ch = str[j];
      if (str.startsWith("FRAC(", j) || str.startsWith("LOG_BASE(", j) ||
          str.startsWith("ROOT(", j) || str.startsWith("INT(", j) ||
          str.startsWith("SUP(", j) ||
          str.startsWith("√(", j) || ch === "|") break;
      j++;
    }
    const chunk = str.slice(i, j);
    if (chunk) result.push(textNode(chunk));
    i = j;
  }
  return result.length ? result : [{ id: uid(), type: "text", value: "", nodes: null }];
}

// ===== CURSOR BOX & RENDER NODE =====
function CursorBox({ active, filled, children, onClick, small }) {
  const isEmpty = !children;
  const bg = active ? "rgba(99,102,241,0.22)" : isEmpty ? "rgba(251,191,36,0.25)" : "rgba(34,197,94,0.08)";
  const border = active ? "1.5px solid #6366F1" : isEmpty ? "1.5px solid #F59E0B" : "1.5px solid #86EFAC";
  return (
    <span onClick={e=>{e.stopPropagation(); onClick && onClick();}}
      style={{ display:"inline-flex", alignItems:"center", justifyContent:"center",
        minWidth:small?14:22, minHeight:small?18:26, padding:small?"1px 3px":"1px 5px",
        background:bg, border, borderRadius:5, cursor:"text", transition:"all 0.15s",
      }}>
      {children || (active
        ? <span style={{display:"inline-block",width:2,height:small?14:20,background:"#6366F1",borderRadius:1}}/>
        : <span style={{color:"#F59E0B",fontSize:small?11:13,fontWeight:700}}>□</span>
      )}
    </span>
  );
}

function renderSlot(slot, cursor, setCursor, style={}) {
  if (!slot) return null;
  const nodes = slot.nodes || [];
  const isEmpty = nodes.length===0||(nodes.length===1&&nodes[0].type==="text"&&!nodes[0].value);
  const isActive = nodes.some(n=>{
    if(n.id===cursor)return true;
    if(n.children)return n.children.some(c=>c&&c.nodes&&c.nodes.some(nn=>nn.id===cursor));
    return false;
  });
  const bg = isActive?"rgba(99,102,241,0.13)":isEmpty?"#E9EAEE":"transparent";
  const brd = isActive?"1.5px solid #6366F1":isEmpty?"1.5px solid #CBD5E1":"1px solid #E2E8F0";
  return (
    <span style={{ display:"inline-flex", alignItems:"center", flexWrap:"nowrap",
      minWidth:18, minHeight:22, padding:"0 3px", background:bg, border:brd, borderRadius:5, cursor:"text", verticalAlign:"middle",
      animation:isActive?"slotPulse 1.1s ease-in-out infinite":"none", ...style,
    }} onClick={e=>{e.stopPropagation(); if(isEmpty||!isActive){const first=nodes.find(n=>n.type==="text");if(first)setCursor(first.id);}}}>
      {isEmpty
        ? (isActive?<span style={{display:"inline-block",width:2,height:16,background:"#6366F1",borderRadius:1,animation:"blink 1s step-end infinite"}}/>:null)
        : nodes.map(n=>renderNode(n,cursor,setCursor))
      }
    </span>
  );
}

function renderNode(node, cursor, setCursor) {
  if (!node) return null;
  if (!setCursor) setCursor=()=>{};
  const active = cursor===node.id;
  const base = {display:"inline-flex",alignItems:"center",fontFamily:"KaTeX_Main,'Computer Modern',Georgia,serif"};

  if (node.type==="text") {
    const txtLatex=toLatex(node.value);
    const txtHtml=(node.value&&window.katex&&/[\^√πα-ω]/u.test(node.value))
      ?(()=>{try{return window.katex.renderToString(txtLatex,{throwOnError:false});}catch{return null;}})()
      :null;
    return (
      <span key={node.id} onClick={e=>{e.stopPropagation();setCursor(node.id);}}
        style={{...base,minWidth:6,minHeight:28,padding:"0 1px",
          background:active?"rgba(99,102,241,0.1)":"transparent",
          borderBottom:active?"2px solid #6366F1":"2px solid transparent",
          borderRadius:2,cursor:"text",fontSize:22,color:"#1E293B",verticalAlign:"middle"}}>
        {node.value
          ?(txtHtml?<span dangerouslySetInnerHTML={{__html:txtHtml}} style={{fontSize:20}}/>
            :<span style={{fontFamily:"KaTeX_Main,serif"}}>{node.value}</span>)
          :(active?<span style={{display:"inline-block",width:2,height:22,background:"#6366F1",borderRadius:1}}/>:"")}
      </span>
    );
  }
  if (node.type==="frac") {
    return (
      <span key={node.id} style={{...base,flexDirection:"column",verticalAlign:"middle",margin:"0 3px",gap:1,alignItems:"stretch"}}>
        {renderSlot(node.children[0],cursor,setCursor,{justifyContent:"center",minHeight:24})}
        <span style={{height:1.5,background:"#1E293B",display:"block",margin:"1px 0"}}/>
        {renderSlot(node.children[1],cursor,setCursor,{justifyContent:"center",minHeight:24})}
      </span>
    );
  }
  if (node.type==="sqrt") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 2px",alignItems:"center"}}>
      <span style={{fontSize:28,lineHeight:1,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>√</span>
      <span style={{borderTop:"1.8px solid #1E293B",paddingTop:2}}>{renderSlot(node.children[0],cursor,setCursor)}</span>
    </span>
  );
  if (node.type==="nthroot") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 4px",display:"inline-flex",alignItems:"flex-end"}}>
      <span style={{fontSize:12,lineHeight:1,marginBottom:10,marginRight:1}}>{renderSlot(node.children[0],cursor,setCursor,{fontSize:"11px",minHeight:14})}</span>
      <span style={{fontSize:30,lineHeight:1,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>√</span>
      <span style={{borderTop:"1.8px solid #1E293B",paddingTop:2}}>{renderSlot(node.children[1],cursor,setCursor)}</span>
    </span>
  );
  if (node.type==="sup") return (
    <span key={node.id} style={{...base,alignItems:"flex-start",verticalAlign:"middle",margin:"0 1px"}}>
      {renderSlot(node.children[0],cursor,setCursor,{minHeight:24})}
      <span style={{fontSize:"0.65em",lineHeight:1,marginTop:2}}>{renderSlot(node.children[1],cursor,setCursor,{fontSize:"12px",minHeight:14})}</span>
    </span>
  );
  if (node.type==="abs") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 2px"}}>
      <span style={{fontSize:24,color:"#1E293B"}}>|</span>{renderSlot(node.children[0],cursor,setCursor)}<span style={{fontSize:24,color:"#1E293B"}}>|</span>
    </span>
  );
  if (node.type==="mixedfrac") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 3px",alignItems:"center",gap:3}}>
      {renderSlot(node.children[0],cursor,setCursor,{minHeight:24})}
      <span style={{display:"inline-flex",flexDirection:"column",verticalAlign:"middle",gap:1,alignItems:"stretch"}}>
        {renderSlot(node.children[1],cursor,setCursor,{justifyContent:"center",minHeight:20})}
        <span style={{height:1.5,background:"#1E293B",display:"block"}}/>
        {renderSlot(node.children[2],cursor,setCursor,{justifyContent:"center",minHeight:20})}
      </span>
    </span>
  );
  if (node.type==="func") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 2px",alignItems:"center"}}>
      <span style={{fontSize:20,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>{node.value}</span>
      <span style={{fontSize:20,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>(</span>
      {renderSlot(node.children[0],cursor,setCursor)}
      <span style={{fontSize:20,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>)</span>
    </span>
  );
  if (node.type==="logbase") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 2px",alignItems:"flex-end"}}>
      <span style={{fontSize:18,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>log</span>
      <span style={{fontSize:"0.65em",lineHeight:1,marginBottom:2}}>{renderSlot(node.children[0],cursor,setCursor,{fontSize:"11px",minHeight:14})}</span>
      <span style={{fontSize:18,color:"#1E293B",fontFamily:"KaTeX_Main,serif",margin:"0 1px"}}>(</span>
      {renderSlot(node.children[1],cursor,setCursor,{minHeight:24})}
      <span style={{fontSize:18,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>)</span>
    </span>
  );
  if (node.type==="defint") return (
    <span key={node.id} style={{...base,verticalAlign:"middle",margin:"0 4px",alignItems:"center"}}>
      <span style={{display:"inline-flex",flexDirection:"column",alignItems:"center",marginRight:2}}>
        <span style={{fontSize:"0.65em"}}>{renderSlot(node.children[1],cursor,setCursor,{fontSize:"11px",minHeight:14})}</span>
        <span style={{fontSize:34,lineHeight:0.9,color:"#1E293B",fontFamily:"KaTeX_Main,serif"}}>∫</span>
        <span style={{fontSize:"0.65em"}}>{renderSlot(node.children[0],cursor,setCursor,{fontSize:"11px",minHeight:14})}</span>
      </span>
      {renderSlot(node.children[2],cursor,setCursor,{minHeight:24})}
      <span style={{fontSize:18,color:"#1E293B",fontFamily:"KaTeX_Main,serif",margin:"0 3px"}}>d</span>
      {renderSlot(node.children[3],cursor,setCursor,{fontSize:"14px",minHeight:14})}
    </span>
  );
  return <span key={node.id} style={{fontSize:20,fontFamily:"KaTeX_Main,serif"}}>{node.value}</span>;
}

// ===== DESMOS-STYLE SCIENTIFIC CALCULATOR KEYBOARD =====

// Real-time calculator result display
function calcResult(str) {
  if (!str) return null;
  try {
    const expr = toMathJs(str);
    if (!expr || expr.length < 1) return null;
    if (window.math) {
      const result = window.math.evaluate(expr);
      if (typeof result === "number" && !isNaN(result)) {
        // Format nicely
        if (Number.isInteger(result)) return result.toString();
        return parseFloat(result.toPrecision(10)).toString();
      }
    }
    return null;
  } catch { return null; }
}

function MathKeyboard({ initValue, onChange, onClose, isAdmin }) {
  const [nodes, setNodes] = useState(() => { if (initValue) return parseToNodes(initValue); return [{ id: Math.random().toString(36).slice(2), type: "text", value: "", nodes: null }]; });
  const [cursor, setCursor] = useState(null);
  const [tab, setTab] = useState("basic");
  const [popup, setPopup] = useState(null);
  const [calcVal, setCalcVal] = useState(null);
  const [shiftMode, setShiftMode] = useState(false); // lowercase/uppercase Latin
  const holdRef = useRef(null);

  const [katexReady, setKatexReady] = useState(!!window.katex);
  const [cursorBlink, setCursorBlink] = useState(true);

  // Vertikal kursor chizig'ini haqiqiy klaviaturalardagi kabi o'chib-yonib turishini
  // ta'minlaydi (500ms har birida ko'rinadi/yashiriladi).
  useEffect(() => {
    const iv = setInterval(() => setCursorBlink(b => !b), 500);
    return () => clearInterval(iv);
  }, []);

  useEffect(() => {
    const collectAll = (ns) => {
      const r = [];
      for (const n of ns) {
        if (n.type === "text") r.push(n);
        else if (n.type === "slot" && n.nodes) r.push(...collectAll(n.nodes));
        else if (n.children) n.children.forEach(s => { if (s && s.nodes) r.push(...collectAll(s.nodes)); });
      }
      return r;
    };
    const all = collectAll(nodes);
    if (all.length > 0) setCursor(all[all.length - 1].id);
    // Force KaTeX load immediately
    if (!window.katex) {
      loadKatex(() => setKatexReady(true));
    } else {
      setKatexReady(true);
    }
  }, []);

  useEffect(() => {
    const str = nodesToString(nodes);
    onChange(str);
    const res = calcResult(str);
    setCalcVal(res);
  }, [nodes]);

  const collectTextNodes = (nodesList) => {
    const result = [];
    for (const n of nodesList) {
      if (n.type === "text") result.push(n);
      else if (n.type === "slot" && n.nodes) result.push(...collectTextNodes(n.nodes));
      else if (n.children) for (const s of n.children) { if (s && s.nodes) result.push(...collectTextNodes(s.nodes)); }
    }
    return result;
  };

  const updateTextNode = (nodesList, cursorId, fn) => {
    return nodesList.map(n => {
      if (n.id === cursorId && n.type === "text") return { ...n, value: fn(n.value) };
      if (n.type === "slot" && n.nodes) { const u = updateTextNode(n.nodes, cursorId, fn); if (u !== n.nodes) return { ...n, nodes: u }; }
      if (n.children) {
        const uc = n.children.map(s => { if (!s || !s.nodes) return s; const u = updateTextNode(s.nodes, cursorId, fn); return u !== s.nodes ? { ...s, nodes: u } : s; });
        if (uc.some((c, i) => c !== n.children[i])) return { ...n, children: uc };
      }
      return n;
    });
  };

  const insertStructureInto = (nodesList, cursorId, struct, after) => {
    const topIdx = nodesList.findIndex(n => n.id === cursorId);
    if (topIdx >= 0) { const arr = [...nodesList]; arr.splice(topIdx + 1, 0, struct, after); return arr; }
    return nodesList.map(n => {
      if (n.type === "slot" && n.nodes) { const u = insertStructureInto(n.nodes, cursorId, struct, after); if (u !== n.nodes) return { ...n, nodes: u }; }
      if (n.children) {
        const uc = n.children.map(s => { if (!s || !s.nodes) return s; const u = insertStructureInto(s.nodes, cursorId, struct, after); return u !== s.nodes ? { ...s, nodes: u } : s; });
        if (uc.some((c, i) => c !== n.children[i])) return { ...n, children: uc };
      }
      return n;
    });
  };

  const insertChar = (ch) => setNodes(prev => updateTextNode(prev, cursor, v => v + ch));
  // Aqlli o'chirish: joriy joy bo'sh bo'lmasa — oxirgi belgini o'chiradi (avvalgidek).
  // Joriy joy BO'SH bo'lsa: agar shu tuzilma (masalan kasr)ning barcha qismlari bo'sh
  // bo'lsa — butun tuzilmani olib tashlab, atrofdagi matnni birlashtiradi (Desmos kabi).
  // Aks holda — kursorni oldingi joyga o'tkazadi.
  const deleteChar = () => {
    const curNode = collectTextNodes(nodes).find(t => t.id === cursor);
    if (curNode && curNode.value) {
      setNodes(prev => updateTextNode(prev, cursor, v => v.slice(0, -1)));
      return;
    }
    let newCursor = cursor;
    let handled = false;
    const process = (list) => {
      const out = [];
      for (let i = 0; i < list.length; i++) {
        const n = list[i];
        if (!handled && n.children && n.children.some(s => s?.nodes)) {
          const containsCursor = n.children.some(s => s?.nodes && collectTextNodes(s.nodes).some(t => t.id === cursor));
          if (containsCursor) {
            const allSlotsEmpty = n.children.every(s => !s?.nodes || collectTextNodes(s.nodes).every(t => !t.value));
            handled = true;
            if (allSlotsEmpty) {
              // Tuzilmani butunlay olib tashlab, chap va o'ng matnni birlashtiramiz
              const prevNode = out[out.length - 1];
              const nextNode = list[i + 1];
              if (prevNode && prevNode.type === "text" && nextNode && nextNode.type === "text") {
                out[out.length - 1] = { ...prevNode, value: prevNode.value + nextNode.value };
                newCursor = prevNode.id;
                i++; continue;
              } else if (prevNode && prevNode.type === "text") {
                newCursor = prevNode.id; continue;
              } else if (nextNode && nextNode.type === "text") {
                newCursor = nextNode.id; out.push(nextNode); i++; continue;
              } else {
                const nt = { id: uid(), type: "text", value: "", nodes: null };
                newCursor = nt.id; out.push(nt); continue;
              }
            } else {
              // Boshqa qismida matn bor — shunchaki kursorni oldingi joyga o'tkazamiz
              const all = collectTextNodes(nodes);
              const idx = all.findIndex(t => t.id === cursor);
              if (idx > 0) newCursor = all[idx - 1].id;
              out.push(n); continue;
            }
          }
        }
        if (!handled && n.type === "slot" && n.nodes) {
          const u = process(n.nodes);
          out.push(u !== n.nodes ? { ...n, nodes: u } : n);
          continue;
        }
        if (!handled && n.children) {
          const uc = n.children.map(s => { if (!s?.nodes) return s; const u = process(s.nodes); return u !== s.nodes ? { ...s, nodes: u } : s; });
          out.push(uc.some((c, i2) => c !== n.children[i2]) ? { ...n, children: uc } : n);
          continue;
        }
        out.push(n);
      }
      return out;
    };
    const result = process(nodes);
    if (handled) {
      setNodes(result);
      setCursor(newCursor);
      return;
    }
    // Tuzilma topilmadi (eng tashqi darajadagi bo'sh joy) — kursorni oldingi joyga o'tkazamiz
    const all = collectTextNodes(nodes);
    const idx = all.findIndex(t => t.id === cursor);
    if (idx > 0) setCursor(all[idx - 1].id);
  };
  const clearAll = () => { const r = { id: uid(), type: "text", value: "", nodes: null }; setNodes([r]); setCursor(r.id); };

  const movePrev = () => { const all = collectTextNodes(nodes); const i = all.findIndex(n => n.id === cursor); if (i > 0) setCursor(all[i - 1].id); };
  const moveNext = () => { const all = collectTextNodes(nodes); const i = all.findIndex(n => n.id === cursor); if (i >= 0 && i < all.length - 1) setCursor(all[i + 1].id); };

  const insertStructure = (type, fnName) => {
    const after = { id: uid(), type: "text", value: "", nodes: null };
    let struct = null;
    if (type === "frac") struct = createNode("frac", "", [slotNode(), slotNode()]);
    else if (type === "sqrt") struct = createNode("sqrt", "", [slotNode()]);
    else if (type === "nthroot") struct = createNode("nthroot", "", [slotNode(), slotNode()]);
    else if (type === "cbrt") { const ds = slotNode([textNode("3")]); const as2 = slotNode(); struct = createNode("nthroot", "", [ds, as2]); const a2 = textNode(""); setNodes(prev => insertStructureInto(prev, cursor, struct, a2)); setCursor(as2.nodes[0]?.id); return; }
    else if (type === "sup") {
      // Masalan foydalanuvchi "5" yozib, keyin daraja (x²) tugmasini bossa — "5" endi
      // yangi darajaning ASOSIGA (bazasiga) avtomatik ko'chiriladi, va kursor darhol
      // daraja qismiga o'tadi — foydalanuvchi to'g'ridan-to'g'ri "2" ni yoza oladi va
      // natijada to'g'ri "5²" hosil bo'ladi (avvalgidek bo'sh "^" chiqib qolmaydi).
      // MUHIM: agar joriy matnda oldin operator ham bo'lsa (masalan "3²+4" yozilgan
      // bo'lsa-yu, "+4" bitta tugunda tursa), FAQAT oxirgi son/harf qismi ("4") bazaga
      // olinadi, "+" belgisi esa formulada joyida (alohida matn sifatida) qoladi —
      // aks holda "+" yo'qolib, ikkita had bir-biriga ko'paytirilgandek hisoblanib
      // qolardi (masalan √(3²+4²) noto'g'ri √(3²·4²) bo'lib qolishi mumkin edi).
      const curNode = collectTextNodes(nodes).find(n => n.id === cursor);
      const fullVal = curNode ? curNode.value : "";
      const m = fullVal.match(/^(.*?)([0-9a-zA-Z.\u03c0]*)$/);
      const prefix = m ? m[1] : "";
      const baseVal = m ? m[2] : fullVal;
      const baseSlot = slotNode(baseVal ? [textNode(baseVal)] : []);
      const expSlot = slotNode();
      struct = createNode("sup", "", [baseSlot, expSlot]);
      setNodes(prev => {
        // Operator (prefix) qismi joriy tugunda qoladi, faqat baza qismi olib tashlanadi
        const cleared = fullVal ? updateTextNode(prev, cursor, () => prefix) : prev;
        return insertStructureInto(cleared, cursor, struct, after);
      });
      setCursor((expSlot.nodes && expSlot.nodes[0]?.id) || after.id);
      return;
    }
    else if (type === "abs") struct = createNode("abs", "", [slotNode()]);
    else if (type === "logbase") struct = createNode("logbase", "", [slotNode(), slotNode()]);
    else if (type === "mixedfrac") struct = createNode("mixedfrac", "", [slotNode(), slotNode(), slotNode()]);
    else if (type === "defint") { const xSlot = slotNode([textNode("x")]); struct = createNode("defint", "", [slotNode(), slotNode(), slotNode(), xSlot]); }
    else if (type === "func") struct = createNode("func", fnName || "", [slotNode()]);
    if (!struct) return;
    setNodes(prev => insertStructureInto(prev, cursor, struct, after));
    const fs = struct.children && struct.children[0];
    const fn2 = fs && fs.nodes && fs.nodes[0];
    setCursor(fn2?.id || after.id);
  };

  const handleKey = (btn) => {
    if (btn.c === "DEL") { deleteChar(); return; }
    if (btn.c === "CLR") { clearAll(); return; }
    if (btn.c === "PREV") { movePrev(); return; }
    if (btn.c === "NEXT") { moveNext(); return; }
    if (btn.c === "SHIFT") { setShiftMode(m => !m); return; }
    if (btn.struct) { insertStructure(btn.struct, btn.fn); return; }
    if (btn.c) {
      insertChar(btn.c);
      // After inserting uppercase letter, auto-switch back to lowercase
      if (shiftMode && btn.lat) setShiftMode(false);
    }
  };

  const startHold = (e, btn) => {
    if (!btn.sub?.length) return;
    e.preventDefault();
    const rect = e.currentTarget.getBoundingClientRect();
    holdRef.current = setTimeout(() => setPopup({ items: btn.sub, rect }), 380);
  };
  const endHold = (btn) => {
    if (holdRef.current) { clearTimeout(holdRef.current); holdRef.current = null; }
    if (!popup) handleKey(btn);
  };

  const ks = (lt) => {
    if (!lt || !katexReady || !window.katex) return null;
    try { return window.katex.renderToString(lt, { throwOnError: false }); }
    catch { return null; }
  };

  // ── KEY DEFINITIONS ──
  const KEYS = {
    basic: [
      [
        { lt:"x^{2}", struct:"sup", dot:true, sub:[{lt:"x^{3}",struct:"cbrt"},{lt:"x^{n}",struct:"sup"}] },
        { lt:"\\frac{\\square}{\\square}", struct:"frac", dot:true, sub:[{lt:"\\square\\frac{\\square}{\\square}",struct:"mixedfrac"}] },
        { lt:"\\sqrt{\\square}", struct:"sqrt", dot:true, sub:[{lt:"\\sqrt[3]{\\square}",struct:"cbrt"},{lt:"\\sqrt[n]{\\square}",struct:"nthroot"}] },
        { l:"(", c:"(", op:true, dot:true, sub:[{l:"[",c:"["},{l:"{",c:"{"}] },
        { l:")", c:")", op:true, dot:true, sub:[{l:"]",c:"]"},{l:"}",c:"}"}] },
        { l:"°", c:"°", lt:"^{\\circ}" },
      ],
      [
        { l:"7", c:"7", num:true },
        { l:"8", c:"8", num:true },
        { l:"9", c:"9", num:true },
        { l:"÷", c:"/", lt:"\\div", op:true },
        { l:"e", c:"e", lt:"e" },
        { l:"C", c:"CLR", clr:true },
      ],
      [
        { l:"4", c:"4", num:true },
        { l:"5", c:"5", num:true },
        { l:"6", c:"6", num:true },
        { l:"×", c:"*", lt:"\\times", op:true },
        { l:"π", c:"π", lt:"\\pi" },
        { l:"⌫", c:"DEL", del:true },
      ],
      [
        { l:"1", c:"1", num:true },
        { l:"2", c:"2", num:true },
        { l:"3", c:"3", num:true },
        { l:"−", c:"-", op:true },
        { l:"+", c:"+", op:true },
        { l:"±", c:"±", lt:"\\pm", op:true },
      ],
      [
        { l:"ln", struct:"func", fn:"ln", lt:"\\ln", dot:true, sub:[{l:"lg",struct:"func",fn:"lg",lt:"\\lg"},{l:"log",struct:"logbase",lt:"\\log_{\\square}"}] },
        { l:"0", c:"0", num:true },
        { l:",", c:",", num:true },
        { l:"x", c:"x", op:true, dot:true, sub:[{l:"y",c:"y"},{l:"z",c:"z"}] },
        { l:"=", c:"=", eq:true },
        { l:";", c:";", op:true },
      ],
    ],
    trig: [
      [
        { l:"sin", struct:"func", fn:"sin", lt:"\\sin" },
        { l:"cos", struct:"func", fn:"cos", lt:"\\cos" },
        { l:"tan", struct:"func", fn:"tan", lt:"\\tan" },
        { l:"asin", struct:"func", fn:"arcsin", lt:"\\arcsin" },
        { l:"acos", struct:"func", fn:"arccos", lt:"\\arccos" },
        { l:"atan", struct:"func", fn:"arctan", lt:"\\arctan" },
      ],
      [
        { l:"sinh", struct:"func", fn:"sinh", lt:"\\sinh" },
        { l:"cosh", struct:"func", fn:"cosh", lt:"\\cosh" },
        { l:"tanh", struct:"func", fn:"tanh", lt:"\\tanh" },
        { l:"ln", struct:"func", fn:"ln", lt:"\\ln" },
        { l:"lg", struct:"func", fn:"lg", lt:"\\lg" },
        { l:"log", struct:"logbase", lt:"\\log_{\\square}" },
      ],
      [
        { l:"°", c:"°", lt:"^{\\circ}" },
        { l:"π", c:"π", lt:"\\pi" },
        { l:"α", c:"α", lt:"\\alpha" },
        { l:"β", c:"β", lt:"\\beta" },
        { l:"γ", c:"γ", lt:"\\gamma" },
        { l:"θ", c:"θ", lt:"\\theta" },
      ],
      [
        { l:"λ", c:"λ", lt:"\\lambda" },
        { l:"μ", c:"μ", lt:"\\mu" },
        { l:"⌫", c:"DEL", del:true },
      ],
    ],
    extra: [
      [
        { l:"|x|", struct:"abs", lt:"|\\square|" },
        { l:"n!", c:"!", lt:"n!" },
        { l:"∞", c:"∞", lt:"\\infty" },
        { l:"≈", c:"≈", lt:"\\approx" },
        { l:"%", c:"%", op:true },
        { l:"<", c:"<", dot:true, sub:[{l:"≤",c:"≤",lt:"\\leq"}] },
      ],
      [
        { l:">", c:">", dot:true, sub:[{l:"≥",c:"≥",lt:"\\geq"}] },
        { l:"≠", c:"≠", lt:"\\neq" },
        { l:"∫", struct:"defint", lt:"\\int_{a}^{b}" },
        { l:"∑", c:"∑", lt:"\\sum" },
        { l:"∂", c:"∂", lt:"\\partial" },
        { l:"∈", c:"∈", lt:"\\in" },
      ],
      [
        { l:"≡", c:"≡", lt:"\\equiv" },
        { l:"[", c:"[" }, { l:"]", c:"]" },
        { l:"{", c:"{" }, { l:"}", c:"}" },
        { l:"⌫", c:"DEL", del:true },
      ],
    ],
    // latin: generated dynamically based on shiftMode
  };

    // Generate latin rows based on shiftMode
  const latinLetters = shiftMode
    ? ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    : ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"];

  const latinRows = [
    latinLetters.slice(0,6).map(l => ({ l, c:l, lat:true })),
    latinLetters.slice(6,12).map(l => ({ l, c:l, lat:true })),
    latinLetters.slice(12,18).map(l => ({ l, c:l, lat:true })),
    latinLetters.slice(18,24).map(l => ({ l, c:l, lat:true })),
    [
      ...latinLetters.slice(24,26).map(l => ({ l, c:l, lat:true })),
      { l: shiftMode ? "⇧" : "⇪", c:"SHIFT", shift:true },
      { l:"⌫", c:"DEL", del:true },
    ],
  ];

  const rows = tab === "latin" ? latinRows : (KEYS[tab] || KEYS.basic);

  const btnStyle = (btn) => {
    if (btn.num)   return { bg: "#FFFFFF", fg: "#1E293B", brd: "#DDE3F0", fs: 18, fw: 700, ff: "'SF Mono',monospace" };
    if (btn.del)   return { bg: "#FEF2F2", fg: "#EF4444", brd: "#FECACA", fs: 16, fw: 700, ff: "inherit" };
    if (btn.clr)   return { bg: "#FFF7ED", fg: "#EA580C", brd: "#FED7AA", fs: 13, fw: 700, ff: "inherit" };
    if (btn.op)    return { bg: "#F1F5FF", fg: "#4338CA", brd: "#C7D2FE", fs: 15, fw: 700, ff: "inherit" };
    if (btn.eq)    return { bg: "#6366F1", fg: "#FFFFFF", brd: "#4F46E5", fs: 15, fw: 800, ff: "inherit" };
    if (btn.nav)   return { bg: "#EEF1FF", fg: "#4F46E5", brd: "#C7D2FE", fs: 15, fw: 700, ff: "inherit" };
    if (btn.shift) return { bg: shiftMode ? "#6366F1" : "#E0E7FF", fg: shiftMode ? "#FFFFFF" : "#4338CA", brd: "#C7D2FE", fs: 14, fw: 700, ff: "inherit" };
    if (btn.lat)   return { bg: "#FFFBF0", fg: "#92400E", brd: "#FDE68A", fs: 17, fw: 700, ff: "'KaTeX_Math','Computer Modern',Georgia,serif" };
    return { bg: "#EEF1FF", fg: "#3730A3", brd: "#C7D2FE", fs: 12, fw: 600, ff: "inherit" };
  };

  const currentStr = nodesToString(nodes);
  const currentLatex = toLatex(currentStr);

  return (
    <div style={{
      background: "#F8FAFF",
      borderTop: "2px solid #6366F1",
      borderRadius: "20px 20px 0 0",
      boxShadow: "0 -6px 32px rgba(99,102,241,0.18)",
      fontFamily: "'SF Pro Display','Segoe UI',system-ui,sans-serif",
      overflow: "visible",
    }}>
      {/* ── LIVE FORMULA PREVIEW + OK ── Klaviatura ekranning pastki qismini butunlay
           egallab, javob maydonining o'zini yopib qo'yishi mumkin — shu sabab
           yozilayotgan formula shu yerda, klaviaturaning o'zida, haqiqiy vaqtda
           (real-time) KaTeX bilan ko'rsatiladi. Kursor — formulaning ICHIDA,
           aynan yozilayotgan joyda, chinakam yonib-o'chib turadigan VERTIKAL
           chiziq sifatida ko'rinadi (alohida oynachada emas). */}
      <div style={{
        background: "#FAFBFF", borderBottom: "1.5px solid #E0E7FF",
        borderRadius: "20px 20px 0 0",
        padding: "8px 10px", minHeight: 40, maxHeight: 60,
        display: "flex", alignItems: "center", gap: 8,
      }}>
        <div style={{ flex: 1, overflowX: "auto", overflowY: "hidden", display: "flex", alignItems: "center", gap: 3 }}>
          {(() => {
            const realStr = nodesToString(nodes);
            // Formula bo'sh bo'lsa — placeholder matn + (miltillasa) boshida kursor.
            if (!realStr) {
              let cursorHtml = null;
              try { cursorHtml = ks(cursorBlink ? "\\textcolor{#6366F1}{\\rule{0.1em}{0.95em}}" : "\\textcolor{#FAFBFF}{\\rule{0.1em}{0.95em}}"); } catch { cursorHtml = null; }
              return (
                <>
                  {cursorHtml && <span dangerouslySetInnerHTML={{ __html: cursorHtml }} style={{ color: "#6366F1" }}/>}
                  <span style={{ fontSize: 13, color: "#94A3B8" }}>Formula shu yerda ko'rinadi...</span>
                </>
              );
            }
            // Kursor turgan text-node'ning oxiriga ko'rinmas maxsus belgi (marker)
            // qo'shiladi. U toLatex orqali o'zgarishsiz o'tadi (chunki hech qanday
            // maxsus belgi/funksiya bilan mos kelmaydi), so'ng LaTeX matni tayyor
            // bo'lgach, o'sha belgi haqiqiy KaTeX \rule (vertikal chiziqcha) bilan
            // almashtiriladi — shu bilan kursor formulaning ICHIDA, aynan to'g'ri
            // joyda (hattoki kasr yoki ildiz ichida bo'lsa ham) chiqadi.
            const CURSOR_MARK = "\u0001";
            const withCursorMark = (list) => list.map(n => {
              if (n.id === cursor && n.type === "text") return { ...n, value: n.value + CURSOR_MARK };
              if (n.type === "slot" && n.nodes) return { ...n, nodes: withCursorMark(n.nodes) };
              if (n.children) return { ...n, children: n.children.map(s => s && s.nodes ? { ...s, nodes: withCursorMark(s.nodes) } : s) };
              return n;
            });
            const str = nodesToString(withCursorMark(nodes));
            let html = null;
            try {
              let latex = toLatex(str.startsWith("$") ? str.slice(1, -1) : str);
              // MUHIM: chiziqcha balandligini o'zgartirmaymiz (0 qilib yubormaymiz) —
              // aks holda KaTeX qator balandligini qayta hisoblab, atrofdagi
              // belgilar har miltillaganda "sakrab" katta-kichik bo'lib ko'rinardi.
              // Shuning uchun o'lcham DOIM bir xil, faqat rangi (ko'rinish/yo'qolish)
              // almashadi — shu bilan formula matni butunlay qimirlamaydi.
              const cursorLatex = cursorBlink ? "\\textcolor{#6366F1}{\\rule{0.1em}{0.95em}}" : "\\textcolor{#FAFBFF}{\\rule{0.1em}{0.95em}}";
              latex = latex.split(CURSOR_MARK).join(cursorLatex);
              html = latex ? ks(latex) : null;
            } catch { html = null; }
            return html
              ? <span dangerouslySetInnerHTML={{ __html: html }} style={{ fontSize: 17, whiteSpace: "nowrap", color: "#1E293B" }}/>
              : <span style={{ fontSize: 13, color: "#94A3B8" }}>Formula shu yerda ko'rinadi...</span>;
          })()}
        </div>
        {/* Result badge — faqat admin uchun */}
        {isAdmin && calcVal !== null && (
          <div style={{
            background: "linear-gradient(135deg,#4F46E5,#7C3AED)",
            color: "white", borderRadius: 8, padding: "4px 10px",
            fontSize: 13, fontWeight: 800, flexShrink: 0,
            boxShadow: "0 2px 8px rgba(99,102,241,0.3)",
            maxWidth: 100, overflow: "hidden", textOverflow: "ellipsis",
            whiteSpace: "nowrap",
          }}>
            = {calcVal}
          </div>
        )}
        <button onClick={onClose} style={{ padding: "6px 12px", background: "#6366F1", border: "none", borderRadius: 8, color: "white", fontWeight: 800, fontSize: 12, cursor: "pointer", flexShrink: 0 }}>OK ✓</button>
      </div>

      {/* ── NAV ROW: ← → — oynada yozilgan formulani ko'rib, kursorni bo'lak-bo'lak
           orasida siljitish uchun preview ostida joylashgan ── */}
      <div style={{
        background: "#FFFFFF", borderBottom: "1px solid #DDE3F0",
        padding: "5px 10px", display: "flex", justifyContent: "center", alignItems: "center", gap: 8,
      }}>
        <button onClick={movePrev} style={{ width: 68, height: 28, background: "#EEF1FF", border: "1px solid #C7D2FE", borderRadius: 8, color: "#4F46E5", fontSize: 14, cursor: "pointer", fontWeight: 700, flexShrink: 0 }}>←</button>
        <button onClick={moveNext} style={{ width: 68, height: 28, background: "#EEF1FF", border: "1px solid #C7D2FE", borderRadius: 8, color: "#4F46E5", fontSize: 14, cursor: "pointer", fontWeight: 700, flexShrink: 0 }}>→</button>
      </div>

      {/* ── TAB ROW ── */}
      <div style={{ display: "flex", background: "#F1F5FF", borderBottom: "1px solid #DDE3F0" }}>
        {[["basic", "Asosiy"], ["trig", "Trig/log"], ["extra", "Qo'shimcha"], ["latin", "Lotin"]].map(([t, l]) => (
          <button key={t} onClick={() => setTab(t)} style={{
            flex: 1, padding: "8px 3px",
            background: tab === t ? "#FFFFFF" : "transparent",
            border: "none", borderBottom: tab === t ? "2px solid #6366F1" : "2px solid transparent",
            color: tab === t ? "#4F46E5" : "#94A3B8",
            fontWeight: tab === t ? 800 : 500, fontSize: 12, cursor: "pointer",
          }}>{l}</button>
        ))}
      </div>

      {/* ── KEY GRID (ixcham, har doim 6 ustunli — bo'limlar orasida o'lcham sakramasligi uchun
           balandlik ham doimiy ushlab turiladi) ── */}
      <div style={{ padding: "6px 6px 10px", background: "#FFFFFF", minHeight: 236, boxSizing: "border-box" }}>
        {rows.map((row, ri) => {
          const GRID_COLS = 6;
          const padded = row.length < GRID_COLS ? [...row, ...Array(GRID_COLS - row.length).fill(null)] : row;
          return (
          <div key={ri} style={{ display: "grid", gridTemplateColumns: `repeat(${GRID_COLS}, 1fr)`, gap: 4, marginBottom: 4 }}>
            {padded.map((btn, ci) => {
              if (!btn) return <div key={ci} aria-hidden="true"/>;
              const st = btnStyle(btn);
              const khtml = ks(btn.lt);
              const hasSub = btn.sub?.length > 0;
              return (
                <button key={ci}
                  onPointerDown={e => { e.currentTarget.setPointerCapture(e.pointerId); startHold(e, btn); }}
                  onPointerUp={() => endHold(btn)}
                  onPointerCancel={() => { if (holdRef.current) { clearTimeout(holdRef.current); holdRef.current = null; } }}
                  style={{
                    minHeight: 38, background: st.bg,
                    border: `1.5px solid ${st.brd}`, borderRadius: 10,
                    color: st.fg, fontSize: st.fs, fontWeight: st.fw,
                    fontFamily: st.ff || "inherit",
                    fontStyle: btn.lat ? "italic" : "normal",
                    cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
                    position: "relative", WebkitTapHighlightColor: "transparent",
                    boxShadow: "0 1.5px 0 rgba(0,0,0,0.07)", userSelect: "none",
                    padding: "2px 2px", transition: "all 0.1s",
                  }}
                  onPointerEnter={e => { e.currentTarget.style.background = "#EEF1FF"; e.currentTarget.style.borderColor = "#818CF8"; }}
                  onPointerLeave={e => { e.currentTarget.style.background = st.bg; e.currentTarget.style.borderColor = st.brd; }}
                >
                  {khtml
                    ? <span dangerouslySetInnerHTML={{ __html: khtml }} style={{ color: st.fg, pointerEvents: "none", fontSize: 13, lineHeight: 1.2 }} />
                    : <span style={{ fontFamily: btn.num ? "'SF Mono',monospace" : "inherit", pointerEvents: "none" }}>{btn.l || "?"}</span>
                  }
                  {hasSub && <span style={{ position: "absolute", bottom: 2, right: 3, width: 4, height: 4, borderRadius: "50%", background: "#EF4444", boxShadow: "0 0 0 1px white" }} />}
                </button>
              );
            })}
          </div>
          );
        })}
      </div>

      {/* ── HOLD POPUP ── */}
      {popup && (
        <div style={{
          position: "fixed",
          left: Math.min(popup.rect.left - 10, window.innerWidth - 170),
          top: Math.max(popup.rect.top - popup.items.length * 52 - 16, 8),
          background: "white", borderRadius: 16, padding: 8,
          boxShadow: "0 8px 40px rgba(99,102,241,0.25)",
          zIndex: 99999, border: "2px solid #6366F1", minWidth: 150,
        }} onPointerDown={e => e.stopPropagation()}>
          <p style={{ color: "#94A3B8", fontSize: 11, margin: "2px 8px 8px", fontWeight: 700, textTransform: "uppercase" }}>Tanlang:</p>
          {popup.items.map((it, ii) => {
            const kh = ks(it.lt);
            return (
              <button key={ii}
                onClick={() => { if (it.struct) insertStructure(it.struct, it.fn); else if (it.c) insertChar(it.c); setPopup(null); }}
                style={{ display: "block", width: "100%", padding: "10px 14px", background: "transparent", border: "none", color: "#1E293B", fontSize: 16, cursor: "pointer", textAlign: "left", borderRadius: 10, fontWeight: 600 }}
                onMouseEnter={e => e.currentTarget.style.background = "#EEF1FF"}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}
              >
                {kh ? <span dangerouslySetInnerHTML={{ __html: kh }} /> : it.l}
              </button>
            );
          })}
          <button onClick={() => setPopup(null)} style={{ display: "block", width: "100%", padding: "8px", background: "#FEF2F2", border: "none", color: "#EF4444", fontSize: 12, cursor: "pointer", borderRadius: 10, marginTop: 4, fontWeight: 700 }}>✕ Yopish</button>
        </div>
      )}
    </div>
  );
}


// ===== AUTH =====
function AuthLayout({ children }) {
  return (
    <div style={{ minHeight:"100vh", background:`linear-gradient(135deg,#EEF1FF,#F0F4FF,#E8F4FF)`, display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ ...S.card, padding:36, width:"100%", maxWidth:420, boxShadow:"0 8px 40px rgba(79,110,247,0.15)" }}>{children}</div>
    </div>
  );
}
function LoginPage({ onLogin, onRegister, onAdmin }) {
  const [login, setLogin] = useState("");
  const [pwd, setPwd] = useState("");
  const [err, setErr] = useState("");
  const [showPwd, setShowPwd] = useState(false);

  const go = () => {
    if (!login.trim() || !pwd) { setErr("Login va parolni kiriting!"); return; }

    // 1. Check admin credentials first
    if (login.trim() === ADMIN_LOGIN && pwd === ADMIN_PW) {
      onAdmin(true); // true = full admin
      return;
    }

    // 2. Check teacher accounts
    const teachers = db.get("teachers") || [];
    const teacher = teachers.find(t => t.login === login.trim() && t.password === pwd);
    if (teacher) {
      onAdmin(false, teacher); // false = not full admin, pass teacher info
      return;
    }

    // 3. Check student database
    const u = (db.get("users")||[]).find(u => u.phone === login.trim());
    if (!u) { setErr("Login yoki parol noto'g'ri!"); return; }
    if (u.password !== pwd) { setErr("Login yoki parol noto'g'ri!"); return; }

    onLogin(u);
  };

  return (
    <AuthLayout>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{fontSize:52}}>📐</div>
        <h1 style={{margin:"8px 0 4px",fontSize:24,fontWeight:800}}>Matematika Testi</h1>
        <p style={{margin:0,color:C.textMid,fontSize:14}}>Platformaga xush kelibsiz</p>
      </div>
      {err&&<div style={S.err}>{err}</div>}
      <label style={S.label}>Login (telefon raqam)</label>
      <input value={login} onChange={e=>setLogin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} style={S.input} placeholder="+998901234567 yoki admin"/>
      <label style={S.label}>Parol</label>
      <div style={{position:"relative"}}>
        <input type={showPwd?"text":"password"} value={pwd} onChange={e=>setPwd(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}
          style={{...S.input,paddingRight:44}} placeholder="••••••"/>
        <button onClick={()=>setShowPwd(s=>!s)} type="button"
          style={{position:"absolute",right:10,top:10,background:"none",border:"none",cursor:"pointer",fontSize:18,color:C.textMid}}>
          {showPwd?"🙈":"👁"}
        </button>
      </div>
      <button onClick={go} style={S.btnPrimary}>Kirish</button>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,margin:"12px 0 0"}}>
        Hisobingiz yo'qmi? <span onClick={onRegister} style={{color:C.primary,cursor:"pointer",fontWeight:700}}>Ro'yxatdan o'ting</span>
      </p>
    </AuthLayout>
  );
}
function RegisterPage({ onDone, onLogin }) {
  const [f,setF]=useState({firstName:"",lastName:"",group:"",phone:"",password:"",password2:""});
  const [err,setErr]=useState("");
  const [showPwd,setShowPwd]=useState(false);

  const go=()=>{
    if(!f.firstName||!f.lastName||!f.group||!f.phone||!f.password){setErr("Barcha maydonlarni to'ldiring!");return;}
    if(f.password.length<4){setErr("Parol kamida 4 ta belgidan iborat bo'lsin!");return;}
    if(f.password!==f.password2){setErr("Parollar mos kelmadi!");return;}
    if(f.phone.trim()===ADMIN_LOGIN){setErr("Bu login band, boshqa raqam kiriting!");return;}
    const users=db.get("users")||[];
    if(users.find(u=>u.phone===f.phone)){setErr("Bu raqam allaqachon ro'yxatdan o'tgan!");return;}
    const u={firstName:f.firstName,lastName:f.lastName,group:f.group,phone:f.phone,password:f.password,id:Date.now()};
    users.push(u); db.set("users",users); onDone(u);
  };

  return (
    <AuthLayout>
      <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:52}}>📝</div><h1 style={{margin:"8px 0 0",fontSize:22,fontWeight:800}}>Ro'yxatdan O'tish</h1></div>
      {err&&<div style={S.err}>{err}</div>}
      {[["firstName","Ism"],["lastName","Familiya"],["group","Guruh"],["phone","Telefon raqam (login)"]].map(([k,ph])=>(
        <div key={k}><label style={S.label}>{ph}</label><input value={f[k]} onChange={e=>setF({...f,[k]:e.target.value})} style={S.input} placeholder={ph}/></div>
      ))}
      <label style={S.label}>Parol</label>
      <div style={{position:"relative"}}>
        <input type={showPwd?"text":"password"} value={f.password} onChange={e=>setF({...f,password:e.target.value})}
          style={{...S.input,paddingRight:44}} placeholder="Kamida 4 ta belgi"/>
        <button onClick={()=>setShowPwd(s=>!s)} type="button"
          style={{position:"absolute",right:10,top:10,background:"none",border:"none",cursor:"pointer",fontSize:18,color:C.textMid}}>
          {showPwd?"🙈":"👁"}
        </button>
      </div>
      <label style={S.label}>Parolni takrorlang</label>
      <input type={showPwd?"text":"password"} value={f.password2} onChange={e=>setF({...f,password2:e.target.value})}
        onKeyDown={e=>e.key==="Enter"&&go()} style={S.input} placeholder="Parolni qayta kiriting"/>
      <button onClick={go} style={S.btnPrimary}>Ro'yxatdan O'tish</button>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,marginTop:12}}>Hisobingiz bormi? <span onClick={onLogin} style={{color:C.primary,cursor:"pointer",fontWeight:700}}>Kirish</span></p>
    </AuthLayout>
  );
}


// ===== TEST CREATOR =====
function TestCreator({ existing, onSave, onCancel }) {
  const [name,setName]=useState(existing?.name||"");
  const [duration,setDuration]=useState(existing?.duration||60);
  const [scheduledAt,setScheduledAt]=useState(tsToLocalInput(existing?.scheduledAt));
  const [showAnswers,setShowAnswers]=useState(existing?.showAnswersAfter||"immediate");
  const [showStats,setShowStats]=useState(existing?.showStats!==false); // default true
  const [requireCode,setRequireCode]=useState(!!existing?.accessCode);
  const [accessCode,setAccessCode]=useState(existing?.accessCode||"");
  const [codePrice,setCodePrice]=useState(existing?.codePrice||"");
  const [pdfFile,setPdfFile]=useState(null);
  const [sections,setSections]=useState(existing?.sections||[{name:"Fan 1",count:10,type:"closed4"}]);
  const [questions,setQuestions]=useState(existing?.questions||[]);
  const [step,setStep]=useState(existing?2:1);
  const [err,setErr]=useState("");
  const [kbdOpen,setKbdOpen]=useState(null); // {qIdx, sub}
  const [restrictGroups,setRestrictGroups]=useState(!!(existing?.targetGroups && existing.targetGroups.length>0));
  const [targetGroups,setTargetGroups]=useState(existing?.targetGroups || []); // [] = barcha guruhlarga ko'rinadi
  const [customGroupInput,setCustomGroupInput]=useState("");
  const allGroups = useMemo(() => {
    const users = db.get("users") || [];
    const set = new Set(users.map(u=>u.group).filter(Boolean));
    (existing?.targetGroups||[]).forEach(g=>set.add(g));
    return [...set].sort();
  }, []);
  const toggleGroup = (g) => setTargetGroups(prev => prev.includes(g) ? prev.filter(x=>x!==g) : [...prev, g]);
  const addCustomGroup = () => {
    const g = customGroupInput.trim();
    if (!g) return;
    if (!targetGroups.includes(g)) setTargetGroups(prev=>[...prev, g]);
    setCustomGroupInput("");
  };

  const typeOpts=(t)=>({closed2:2,closed3:3,closed4:4,closed5:5,closed6:6,closed7:7,closed8:8,open:0})[t]||4;
  const typeLabel=(t)=>({closed2:"2 variant",closed3:"3 variant",closed4:"4 variant (A-D)",closed5:"5 variant (A-E)",closed6:"6 variant",closed7:"7 variant",closed8:"8 variant",open:"Ochiq (yozma)"})[t]||t;
  const totalQ=sections.reduce((s,sec)=>s+Number(sec.count),0);

  const [pdfLoading,setPdfLoading]=useState(false);
  const [showLatexPreview,setShowLatexPreview]=useState(true);
  const [latexMode,setLatexMode]=useState(existing?.latexFileName ? "upload" : "write"); // "upload" | "write"
  const latexTextareaRef = useRef(null);

  // ===== 3 tilli hujjatlar (UZ / QQ / RU): har bir test uchun PDF yoki LaTeX alohida-alohida yuklanadi =====
  const [docLang,setDocLang]=useState("uz"); // hozir tahrirlanayotgan til
  const [langDocs,setLangDocs]=useState(() => {
    if (existing?.langDocs) return existing.langDocs;
    // Eski (bir tilli) testlarni "uz" sifatida ko'chiramiz
    const base = emptyLangDocs();
    if (existing?.pdfUrl || existing?.latexSource) {
      base.uz = { docType: existing?.latexSource ? "latex" : "pdf", pdfUrl: existing?.pdfUrl||null, latexSource: existing?.latexSource||"", latexFileName: existing?.latexFileName||"", latexImages: existing?.latexImages||{} };
    }
    return base;
  });
  const cur = langDocs[docLang] || emptyLangDoc();
  const updateCur = (patch) => setLangDocs(p => ({ ...p, [docLang]: { ...(p[docLang]||emptyLangDoc()), ...patch } }));
  const setDocType = (t) => updateCur({ docType: t });
  const setPdfUrl = (v) => updateCur({ pdfUrl: v });
  const setLatexSource = (v) => updateCur({ latexSource: v });
  const setLatexFileName = (v) => updateCur({ latexFileName: v });
  const setLatexImages = (fn) => updateCur({ latexImages: typeof fn==="function" ? fn(cur.latexImages||{}) : fn });
  const docType = cur.docType || "pdf";
  const pdfUrl = cur.pdfUrl || null;
  const latexSource = cur.latexSource || "";
  const latexFileName = cur.latexFileName || "";
  const latexImages = cur.latexImages || {};

  const handlePdf=(e)=>{
    const file=e.target.files[0]; if(!file) return;
    if(file.size > 8*1024*1024){
      alert("PDF fayl hajmi 8MB dan oshmasligi kerak!");
      return;
    }
    setPdfLoading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      // Store as base64 data URL — persists in localStorage, works across sessions/devices
      setPdfUrl(ev.target.result);
      setPdfFile(file);
      setPdfLoading(false);
    };
    reader.onerror = () => {
      alert("PDF faylni o'qishda xatolik yuz berdi!");
      setPdfLoading(false);
    };
    reader.readAsDataURL(file);
  };

  const handleLatex=(e)=>{
    const file=e.target.files[0]; if(!file) return;
    if(file.size > 2*1024*1024){
      alert("LaTeX fayl hajmi 2MB dan oshmasligi kerak!");
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      setLatexSource(ev.target.result);
      setLatexFileName(file.name);
    };
    reader.onerror = () => alert("LaTeX faylni o'qishda xatolik yuz berdi!");
    reader.readAsText(file);
  };

  const handleLatexImage=(e)=>{
    const file=e.target.files[0]; if(!file) return;
    if(file.size > 3*1024*1024){
      alert("Rasm hajmi 3MB dan oshmasligi kerak!");
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      let n = Object.keys(latexImages).length + 1;
      let key = `rasm${n}`;
      while (latexImages[key]) { n++; key = `rasm${n}`; }
      const nextImages = {...latexImages,[key]:ev.target.result};
      setLatexImages(nextImages);
      const cmd = `\\includegraphics{${key}}`;
      const ta = latexTextareaRef.current;
      if (ta) {
        const start = ta.selectionStart ?? latexSource.length;
        const end = ta.selectionEnd ?? latexSource.length;
        const next = latexSource.slice(0,start) + cmd + latexSource.slice(end);
        setLatexSource(next);
        requestAnimationFrame(()=>{ ta.focus(); ta.selectionStart=ta.selectionEnd=start+cmd.length; });
      } else {
        setLatexSource((latexSource?latexSource+"\n":"") + cmd);
      }
    };
    reader.onerror = () => alert("Rasmni o'qishda xatolik yuz berdi!");
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const build=()=>{
    if(!name.trim()){setErr("Test nomini kiriting!");return;}
    setErr(""); setStep(2);
    // If questions already built and have answers, don't rebuild (preserve correctAnswers)
    const totalExpected = sections.reduce((s,sec)=>s+Number(sec.count),0);
    if(questions.length===totalExpected) return; // already built, keep answers
    const qs=[];
    sections.forEach(sec=>{
      for(let i=0;i<Number(sec.count);i++){
        const existingQ = existing?.questions?.[qs.length];
        qs.push({
          id:qs.length,
          type:sec.type==="open"?"open":"closed",
          optionsCount:typeOpts(sec.type),
          sectionName:sec.name,
          correctAnswer: existingQ?.correctAnswer || "",
          subParts: existingQ?.subParts || [],
        });
      }
    });
    setQuestions(qs);
  };

  const updQ=(i,f,v)=>setQuestions(p=>p.map((q,idx)=>idx===i?{...q,[f]:v}:q));
  const addSub=(i)=>setQuestions(p=>p.map((q,idx)=>idx!==i?q:{...q,subParts:[...(q.subParts||[]),{label:String.fromCharCode(97+(q.subParts||[]).length),answer:""}]}));
  const delSub=(qi,si)=>setQuestions(p=>p.map((q,i)=>i!==qi?q:{...q,subParts:q.subParts.filter((_,s)=>s!==si)}));
  const updSub=(qi,si,v)=>setQuestions(p=>p.map((q,i)=>i!==qi?q:{...q,subParts:q.subParts.map((s,idx)=>idx===si?{...s,answer:v}:s)}));

  const save=()=>{
    if(!name.trim()){setErr("Test nomini kiriting!");return;}
    if(requireCode && !accessCode.trim()){setErr("Kirish kodini kiriting yoki tasodifiy yarating!");return;}
    if(restrictGroups && targetGroups.length===0){setErr("Kamida bitta guruhni tanlang yoki \"Barcha guruhlarga\" ni belgilang!");return;}
    const schedTs = localInputToTs(scheduledAt);
    const wasActive = existing?.active||false;
    // Agar rejalashtirilgan vaqt kelajakda bo'lsa va test hali qo'lda faollashtirilmagan bo'lsa, uni nofaol holatda saqlaymiz — vaqt kelganda avtomatik faollashadi.
    const willAutoStart = schedTs && schedTs > Date.now() && !wasActive;
    // Har bir til uchun bo'sh bo'lmagan qismini saqlaymiz (docType bo'yicha tozalab)
    const cleanLangDocs = {};
    DOC_LANGS.forEach(l=>{
      const d = langDocs[l.code] || emptyLangDoc();
      cleanLangDocs[l.code] = {
        docType: d.docType||"pdf",
        pdfUrl: d.docType==="pdf" ? (d.pdfUrl||null) : null,
        latexSource: d.docType==="latex" ? (d.latexSource||null) : null,
        latexFileName: d.docType==="latex" ? (d.latexFileName||null) : null,
        latexImages: d.docType==="latex" ? (d.latexImages||null) : null,
      };
    });
    const uzDoc = cleanLangDocs.uz;
    onSave({id:existing?.id||Date.now(),name,duration,closedCount:questions.filter(q=>q.type==="closed").length,optionsCount:sections[0]?typeOpts(sections[0].type):4,sections,questions,active:willAutoStart?false:wasActive,scheduledAt:schedTs,showAnswersAfter:showAnswers,showStats,startedAt:existing?.startedAt||null,langDocs:cleanLangDocs,pdfUrl:uzDoc.pdfUrl,latexSource:uzDoc.latexSource,latexFileName:uzDoc.latexFileName,latexImages:uzDoc.latexImages,accessCode:requireCode?accessCode.trim().toUpperCase():null,codePrice:requireCode?codePrice:null,targetGroups:restrictGroups?targetGroups:[]});
  };

  const grouped=()=>{
    const g={};
    questions.forEach((q,i)=>{const k=q.sectionName||"Asosiy";if(!g[k])g[k]=[];g[k].push({q,i});});
    return g;
  };

  const kbdVal = kbdOpen ? (kbdOpen.sub!==null ? (questions[kbdOpen.qIdx]?.subParts?.[kbdOpen.sub]?.answer||"") : (questions[kbdOpen.qIdx]?.correctAnswer||"")) : "";

  return (
    <div style={{...S.card,padding:20,marginBottom:20,border:`2px solid ${C.primary}`,paddingBottom:kbdOpen?380:20}}>
      <h3 style={{margin:"0 0 16px",color:C.primary}}>{existing?"✏️ Tahrirlash":"➕ Yangi Test"}</h3>
      {err&&<div style={S.err}>{err}</div>}

      {step===1&&(
        <div>
          <label style={S.label}>Test nomi</label>
          <input value={name} onChange={e=>setName(e.target.value)} style={S.input} placeholder="Test nomi"/>

          {/* PDF / LaTeX upload — 3 tilda (UZ / QQ / RU) */}
          <label style={S.label}>📄 Test varianti (ixtiyoriy) — har til uchun alohida</label>
          <div style={{display:"flex",gap:8,marginBottom:10,background:"#F1F5F9",padding:5,borderRadius:11}}>
            {DOC_LANGS.map(l=>{
              const has = !!(cur && docLang===l.code ? (langDocs[l.code]?.pdfUrl||langDocs[l.code]?.latexSource) : (langDocs[l.code]?.pdfUrl||langDocs[l.code]?.latexSource));
              return (
                <button key={l.code} onClick={()=>setDocLang(l.code)} style={{
                  flex:1,padding:"8px 4px",borderRadius:8,border:"none",cursor:"pointer",
                  background:docLang===l.code?"white":"transparent",
                  boxShadow:docLang===l.code?"0 1px 4px rgba(0,0,0,0.12)":"none",
                  color:docLang===l.code?C.primary:C.textMid,fontWeight:800,fontSize:13,
                  display:"flex",alignItems:"center",justifyContent:"center",gap:5,position:"relative"
                }}>
                  <span><LangFlag lang={l} size={17}/></span><span>{l.label}</span>
                  {has && <span style={{position:"absolute",top:2,right:6,width:7,height:7,borderRadius:"50%",background:C.successDark}}/>}
                </button>
              );
            })}
          </div>
          <p style={{margin:"-6px 0 10px",fontSize:11.5,color:C.textLight}}>Hozir tahrirlanmoqda: <b>{DOC_LANGS.find(l=>l.code===docLang)?.full}</b>. Yashil nuqta — shu tilda hujjat yuklangan.</p>
          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <button onClick={()=>setDocType("pdf")} style={{
              flex:1,padding:"9px",borderRadius:9,border:`1.5px solid ${docType==="pdf"?C.primary:C.border}`,
              background:docType==="pdf"?C.primaryLight:C.card,color:docType==="pdf"?C.primary:C.textMid,
              fontWeight:700,fontSize:13,cursor:"pointer"
            }}>📄 PDF</button>
            <button onClick={()=>setDocType("latex")} style={{
              flex:1,padding:"9px",borderRadius:9,border:`1.5px solid ${docType==="latex"?C.primary:C.border}`,
              background:docType==="latex"?C.primaryLight:C.card,color:docType==="latex"?C.primary:C.textMid,
              fontWeight:700,fontSize:13,cursor:"pointer"
            }}>∑ LaTeX (.tex)</button>
          </div>

          {docType==="pdf" && (
            <div style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:14,marginBottom:10,textAlign:"center",background:"#FAFBFF"}}>
              {pdfLoading ? (
                <div style={{color:C.primary,fontWeight:700,fontSize:14}}>⏳ PDF yuklanmoqda...</div>
              ) : pdfUrl ? (
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <span style={{color:C.successDark,fontWeight:700,fontSize:14}}>✅ PDF yuklangan ({pdfFile?(pdfFile.size/1024/1024).toFixed(1):"?"} MB)</span>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>window.open(pdfUrl,"_blank")} style={{...S.btnSmall,background:C.primary,padding:"6px 12px",fontSize:12,border:"none",cursor:"pointer"}}>👁 Ko'rish</button>
                    <button onClick={()=>{setPdfUrl(null);setPdfFile(null);}} style={{...S.btnSmall,background:C.danger,padding:"6px 12px",fontSize:12}}>O'chirish</button>
                  </div>
                </div>
              ):(
                <label style={{cursor:"pointer",color:C.textMid,fontSize:14}}>
                  <span style={{fontSize:28,display:"block",marginBottom:4}}>📎</span>
                  PDF faylni tanlash uchun bosing (max 8MB)
                  <input type="file" accept=".pdf" onChange={handlePdf} style={{display:"none"}}/>
                </label>
              )}
            </div>
          )}

          {docType==="latex" && (
            <div style={{marginBottom:10}}>
              {/* Yuklash vs Yozish toggle */}
              <div style={{display:"flex",gap:6,marginBottom:10}}>
                <button onClick={()=>setLatexMode("upload")} style={{
                  flex:1,padding:"7px",borderRadius:8,border:`1.5px solid ${latexMode==="upload"?C.primary:C.border}`,
                  background:latexMode==="upload"?C.primaryLight:C.card,color:latexMode==="upload"?C.primary:C.textMid,
                  fontWeight:700,fontSize:12,cursor:"pointer"
                }}>📎 Fayl yuklash</button>
                <button onClick={()=>setLatexMode("write")} style={{
                  flex:1,padding:"7px",borderRadius:8,border:`1.5px solid ${latexMode==="write"?C.primary:C.border}`,
                  background:latexMode==="write"?C.primaryLight:C.card,color:latexMode==="write"?C.primary:C.textMid,
                  fontWeight:700,fontSize:12,cursor:"pointer"
                }}>✍️ Qo'lda yozish</button>
              </div>

              {latexMode==="upload" && (
                <div style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:14,textAlign:"center",background:"#FAFBFF"}}>
                  {latexSource ? (
                    <div>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                        <span style={{color:C.successDark,fontWeight:700,fontSize:14}}>✅ {latexFileName||"LaTeX hujjat"} yuklangan</span>
                        <button onClick={()=>{setLatexSource("");setLatexFileName("");}} style={{...S.btnSmall,background:C.danger,padding:"6px 12px",fontSize:12}}>O'chirish</button>
                      </div>
                      <p style={{margin:"0 0 8px",fontSize:12,color:C.textMid}}>O'quvchilarga chiroyli formula ko'rinishida (KaTeX) namoyish etiladi.</p>
                    </div>
                  ):(
                    <label style={{cursor:"pointer",color:C.textMid,fontSize:14}}>
                      <span style={{fontSize:28,display:"block",marginBottom:4}}>∑</span>
                      .tex faylni tanlash uchun bosing (max 2MB)
                      <input type="file" accept=".tex,.txt" onChange={handleLatex} style={{display:"none"}}/>
                    </label>
                  )}
                </div>
              )}

              {latexMode==="write" && (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap",gap:6}}>
                    <span style={{fontSize:12,color:C.textMid,fontWeight:600}}>LaTeX kodini kiriting:</span>
                    <div style={{display:"flex",gap:6}}>
                      <label style={{...S.btnSmall,background:"#F59E0B",padding:"4px 10px",fontSize:11,cursor:"pointer"}}>
                        🖼️ Rasm qo'shish
                        <input type="file" accept="image/*" onChange={handleLatexImage} style={{display:"none"}}/>
                      </label>
                      <button onClick={()=>setShowLatexPreview(p=>!p)} style={{...S.btnSmall,background:showLatexPreview?C.primary:"#E2E8F0",color:showLatexPreview?"white":C.textMid,padding:"4px 10px",fontSize:11}}>
                        {showLatexPreview?"👁 Preview ON":"👁 Preview OFF"}
                      </button>
                    </div>
                  </div>
                  <textarea
                    ref={latexTextareaRef}
                    value={latexSource}
                    onChange={e=>setLatexSource(e.target.value)}
                    placeholder={"\\section{Algebra masalalari}\n\n1-savol: $x^2 + 2x + 1 = 0$ tenglamani yeching.\n\n2-savol: Hisoblang: $$\\frac{a}{b} + \\sqrt{c}$$"}
                    style={{
                      width:"100%",minHeight:160,padding:"10px 12px",
                      border:`1.5px solid ${C.border}`,borderRadius:10,
                      fontFamily:"'SF Mono',monospace",fontSize:13,
                      color:C.text,background:C.card,resize:"vertical",
                      boxSizing:"border-box",
                    }}
                  />
                  {Object.keys(latexImages).length>0 && (
                    <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:8}}>
                      {Object.entries(latexImages).map(([key,src])=>(
                        <div key={key} style={{position:"relative",border:`1.5px solid ${C.border}`,borderRadius:8,padding:4,background:C.card}}>
                          <img src={src} alt={key} style={{width:56,height:56,objectFit:"cover",borderRadius:5,display:"block"}}/>
                          <p style={{margin:"3px 0 0",fontSize:9,color:C.textMid,textAlign:"center"}}>{key}</p>
                          <button onClick={()=>setLatexImages(p=>{const n={...p};delete n[key];return n;})} style={{position:"absolute",top:-6,right:-6,width:18,height:18,borderRadius:"50%",background:C.danger,color:"white",border:"none",fontSize:11,cursor:"pointer",lineHeight:1}}>✕</button>
                        </div>
                      ))}
                    </div>
                  )}
                  <p style={{margin:"6px 0 0",fontSize:11,color:C.textLight,lineHeight:1.6}}>
                    Formula: <code style={{background:"#F1F5FF",padding:"1px 5px",borderRadius:4}}>$x^2$</code> yoki <code style={{background:"#F1F5FF",padding:"1px 5px",borderRadius:4}}>$$x^2$$</code>. Sarlavha: <code style={{background:"#F1F5FF",padding:"1px 5px",borderRadius:4}}>\section{"{"}...{"}"}</code>.<br/>
                    Rasm: yuqoridagi <b>🖼️ Rasm qo'shish</b> tugmasi bilan yuklang — kursor turgan joyga <code style={{background:"#F1F5FF",padding:"1px 5px",borderRadius:4}}>\includegraphics{"{"}rasm1{"}"}</code> avtomatik qo'yiladi.<br/>
                    Chizma (TikZ/pgfplots): kodini to'g'ridan-to'g'ri yozing.
                  </p>

                  {showLatexPreview && latexSource && (
                    <div style={{marginTop:10,border:`1.5px solid ${C.border}`,borderRadius:10,overflow:"hidden"}}>
                      <div style={{background:C.primaryLight,padding:"6px 12px",fontSize:11,fontWeight:700,color:C.primary}}>👁 Jonli ko'rinish (o'quvchi shunday ko'radi)</div>
                      <div style={{maxHeight:280,overflowY:"auto",background:"white"}}>
                        <LatexDocViewer source={latexSource} images={latexImages}/>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:4}}>
            <div><label style={S.label}>⏱ Vaqt (daqiqa)</label><input type="number" value={duration} onChange={e=>{const v=e.target.value; setDuration(v===""?"":+v);}} onBlur={e=>{if(e.target.value===""||+e.target.value<1) setDuration(1);}} style={S.input} min={1}/></div>
            <div><label style={S.label}>👁 Natijalar</label>
              <select value={showAnswers} onChange={e=>setShowAnswers(e.target.value)} style={S.input}>
                <option value="immediate">Darhol</option><option value="manual">Qo'lda</option>
              </select>
            </div>
          </div>

          {/* Testni oldindan yuklab, qachon boshlanishini belgilash */}
          <label style={S.label}>🗓️ Test boshlanish vaqti (ixtiyoriy)</label>
          <input type="datetime-local" value={scheduledAt} onChange={e=>setScheduledAt(e.target.value)} style={S.input}/>
          <p style={{margin:"4px 0 12px",color:C.textMid,fontSize:12.5,lineHeight:1.5}}>
            {scheduledAt
              ? <>📅 Test <b>{formatScheduled(localInputToTs(scheduledAt))}</b> da avtomatik boshlanadi. O'quvchilar profilida bu vaqt oldindan ko'rinib turadi.</>
              : "Bo'sh qoldirsangiz, testni \"✅ Faollashtirish\" tugmasi orqali qo'lda boshlaysiz."}
          </p>
          {scheduledAt && (
            <button onClick={()=>setScheduledAt("")} style={{...S.btnGhost,marginBottom:12,padding:"6px 12px",fontSize:12,width:"auto"}}>✕ Vaqtni bekor qilish</button>
          )}

          {/* Statistika ko'rsatish */}
          <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",marginBottom:12,padding:"10px 12px",background:showStats?"#DCFCE7":"#F1F5F9",borderRadius:10,border:`1.5px solid ${showStats?"#22C55E":C.border}`}}>
            <input type="checkbox" checked={showStats} onChange={e=>setShowStats(e.target.checked)}
              style={{width:18,height:18,accentColor:"#22C55E",cursor:"pointer"}}/>
            <div>
              <span style={{fontWeight:700,fontSize:13,color:showStats?"#16A34A":C.textMid}}>📈 O'quvchilarga statistika ko'rsatilsin</span>
              <p style={{margin:0,fontSize:11,color:C.textLight}}>O'chirilsa, o'quvchi o'z natijalar statistikasini ko'ra olmaydi</p>
            </div>
          </label>

          {/* Maxfiy test / Kirish kodi */}
          <div style={{background:"#FFFBEB",border:"1.5px solid #FDE68A",borderRadius:12,padding:14,marginBottom:16}}>
            <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",marginBottom:requireCode?12:0}}>
              <input type="checkbox" checked={requireCode} onChange={e=>setRequireCode(e.target.checked)}
                style={{width:20,height:20,accentColor:"#F59E0B",cursor:"pointer"}}/>
              <span style={{fontWeight:700,color:"#92400E",fontSize:14}}>🔒 Maxfiy test (kirish kodi talab qilinsin)</span>
            </label>
            {requireCode && (
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10}}>
                <div>
                  <label style={{...S.label,color:"#92400E"}}>Kirish kodi</label>
                  <div style={{display:"flex",gap:8}}>
                    <input value={accessCode} onChange={e=>setAccessCode(e.target.value.toUpperCase())}
                      style={{...S.input,margin:0,fontFamily:"monospace",fontWeight:700,letterSpacing:1}}
                      placeholder="MATH2024"/>
                    <button onClick={()=>{
                      const rnd = Math.random().toString(36).slice(2,8).toUpperCase();
                      setAccessCode(rnd);
                    }} style={{...S.btnSmall,background:"#F59E0B",whiteSpace:"nowrap"}}>🎲 Yaratish</button>
                  </div>
                </div>
                <div>
                  <label style={{...S.label,color:"#92400E"}}>Narxi (so'm, ixtiyoriy)</label>
                  <input type="number" value={codePrice} onChange={e=>setCodePrice(e.target.value)}
                    style={{...S.input,margin:0}} placeholder="20000"/>
                </div>
              </div>
            )}
            {requireCode && (
              <p style={{margin:"8px 0 0",fontSize:12,color:"#92400E"}}>
                O'quvchi to'lov qilgach, ushbu kodni unga yuboring. Test ichiga kirishdan oldin shu kod so'raladi.
              </p>
            )}
          </div>

          {/* Qaysi guruhlarga ko'rinishi */}
          <div style={{background:"#EEF1FF",border:`1.5px solid ${C.border}`,borderRadius:12,padding:14,marginBottom:16}}>
            <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",marginBottom:restrictGroups?12:0}}>
              <input type="checkbox" checked={restrictGroups} onChange={e=>setRestrictGroups(e.target.checked)}
                style={{width:20,height:20,accentColor:C.primary,cursor:"pointer"}}/>
              <span style={{fontWeight:700,color:C.primary,fontSize:14}}>👥 Faqat tanlangan guruhlarga ko'rinsin</span>
            </label>
            {!restrictGroups && (
              <p style={{margin:0,fontSize:12,color:C.textMid}}>Hozircha barcha guruhlarga ko'rinadi.</p>
            )}
            {restrictGroups && (
              <div>
                {allGroups.length>0 && (
                  <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
                    {allGroups.map(g=>(
                      <button key={g} onClick={()=>toggleGroup(g)} style={{
                        padding:"6px 14px",borderRadius:20,border:`1.5px solid ${targetGroups.includes(g)?C.primary:C.border}`,
                        background:targetGroups.includes(g)?C.primary:"white",color:targetGroups.includes(g)?"white":C.text,
                        fontWeight:600,fontSize:13,cursor:"pointer"
                      }}>{targetGroups.includes(g)?"✓ ":""}{g}</button>
                    ))}
                  </div>
                )}
                <div style={{display:"flex",gap:8}}>
                  <input value={customGroupInput} onChange={e=>setCustomGroupInput(e.target.value)}
                    onKeyDown={e=>{if(e.key==="Enter"){e.preventDefault();addCustomGroup();}}}
                    style={{...S.input,margin:0}} placeholder="Yangi guruh nomi (masalan 10-A)"/>
                  <button onClick={addCustomGroup} style={{...S.btnSmall,background:C.primary,whiteSpace:"nowrap"}}>+ Qo'shish</button>
                </div>
                {targetGroups.length>0 && (
                  <p style={{margin:"8px 0 0",fontSize:12,color:C.textMid}}>Tanlangan: <b>{targetGroups.join(", ")}</b></p>
                )}
              </div>
            )}
          </div>

          {/* Sections */}
          <div style={{background:"#F8F9FF",borderRadius:12,padding:14,marginBottom:16,border:`1px solid ${C.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <span style={{fontWeight:800,fontSize:14}}>📚 Bo'limlar</span>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{color:C.textMid,fontSize:13}}>Jami: <b style={{color:C.primary}}>{totalQ}</b></span>
                <div style={{display:"flex",alignItems:"center",background:C.card,borderRadius:10,border:`1.5px solid ${C.border}`,overflow:"hidden"}}>
                  <button onClick={()=>setSections(p=>p.length>1?p.slice(0,-1):p)} style={{width:36,height:36,background:C.danger,border:"none",color:"white",fontSize:20,cursor:"pointer",fontWeight:900}}>−</button>
                  <span style={{minWidth:36,textAlign:"center",fontWeight:800,fontSize:16}}>{sections.length}</span>
                  <button onClick={()=>setSections(p=>[...p,{name:"Fan "+(p.length+1),count:10,type:"closed4"}])} style={{width:36,height:36,background:C.primary,border:"none",color:"white",fontSize:20,cursor:"pointer",fontWeight:900}}>+</button>
                </div>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"26px 1fr 110px 1fr 28px",gap:6,marginBottom:6}}>
              {["№","Bo'lim","Soni","Tur",""].map((h,i)=><span key={i} style={{color:C.textLight,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>{h}</span>)}
            </div>
            {sections.map((sec,i)=>(
              <div key={i} style={{display:"grid",gridTemplateColumns:"26px 1fr 110px 1fr 28px",gap:6,marginBottom:8,alignItems:"center"}}>
                <div style={{width:26,height:26,borderRadius:"50%",background:C.primaryLight,color:C.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:800}}>{i+1}</div>
                <input value={sec.name} onChange={e=>setSections(p=>p.map((s,idx)=>idx===i?{...s,name:e.target.value}:s))} style={{...S.input,margin:0,padding:"9px 10px",fontSize:13}} placeholder={"Fan "+(i+1)}/>
                <div style={{display:"flex",alignItems:"center",background:C.card,borderRadius:8,border:`1.5px solid ${C.border}`,overflow:"hidden",height:38}}>
                  <button onClick={()=>setSections(p=>p.map((s,idx)=>idx===i?{...s,count:Math.max(1,Number(s.count||0)-1)}:s))} style={{width:30,height:"100%",background:C.danger,border:"none",color:"white",fontSize:18,cursor:"pointer",fontWeight:900}}>−</button>
                  <input type="number" value={sec.count} onChange={e=>{const v=e.target.value; setSections(p=>p.map((s,idx)=>idx===i?{...s,count:v===""?"":+v}:s));}} onBlur={e=>{if(e.target.value===""||+e.target.value<1) setSections(p=>p.map((s,idx)=>idx===i?{...s,count:1}:s));}} style={{flex:1,background:"transparent",border:"none",color:C.text,fontWeight:800,fontSize:15,textAlign:"center",outline:"none",minWidth:0}} min={1}/>
                  <button onClick={()=>setSections(p=>p.map((s,idx)=>idx===i?{...s,count:Number(s.count||0)+1}:s))} style={{width:30,height:"100%",background:C.primary,border:"none",color:"white",fontSize:18,cursor:"pointer",fontWeight:900}}>+</button>
                </div>
                <select value={sec.type} onChange={e=>setSections(p=>p.map((s,idx)=>idx===i?{...s,type:e.target.value}:s))} style={{...S.input,margin:0,padding:"9px 8px",fontSize:12}}>
                  {["closed2","closed3","closed4","closed5","closed6","closed7","closed8","open"].map(t=><option key={t} value={t}>{typeLabel(t)}</option>)}
                </select>
                <button onClick={()=>sections.length>1&&setSections(p=>p.filter((_,idx)=>idx!==i))} style={{width:26,height:26,background:sections.length===1?"#E2E8F0":C.danger,border:"none",borderRadius:6,color:sections.length===1?C.textLight:"white",cursor:sections.length===1?"default":"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
              </div>
            ))}
            <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:8,paddingTop:8,borderTop:`1px solid ${C.border}`}}>
              {sections.map((sec,i)=><span key={i} style={{...S.badge,background:C.primaryLight,color:C.primary,fontSize:11}}>{sec.name}: <b>{sec.count}</b> × {typeLabel(sec.type)}</span>)}
            </div>
          </div>

          <div style={{display:"flex",gap:10}}>
            <button onClick={build} style={S.btnPrimary}>Davom etish →</button>
            <button onClick={onCancel} style={{...S.btnGhost,flex:1}}>Bekor</button>
          </div>
        </div>
      )}

      {step===2&&(
        <div>
          <div style={{display:"flex",gap:10,marginBottom:14}}>
            <button onClick={()=>setStep(1)} style={{...S.btnGhost,padding:"8px 14px"}}>← Orqaga</button>
            <button onClick={save} style={{...S.btnSuccess,flex:1}}>💾 Saqlash</button>
            <button onClick={onCancel} style={{...S.btnGhost,padding:"8px 14px"}}>Bekor</button>
          </div>
          <p style={{color:C.textMid,fontSize:13,marginBottom:12}}>To'g'ri javoblarni belgilang. Ochiq savollar uchun matematik formula yozish uchun <b>𝑓(𝑥)</b> tugmasini bosing.</p>
          <div style={{maxHeight:520,overflowY:"auto",paddingRight:6}}>
            {Object.entries(grouped()).map(([sec,items])=>(
              <div key={sec} style={{marginBottom:16}}>
                <div style={{background:C.primaryLight,borderRadius:8,padding:"8px 14px",marginBottom:8,display:"flex",justifyContent:"space-between"}}>
                  <span style={{color:C.primary,fontWeight:700,fontSize:14}}>📚 {sec}</span>
                  <span style={{color:C.textMid,fontSize:12}}>{items.length} savol</span>
                </div>
                {items.map(({q,i})=>(
                  <div key={i} style={{...S.card,marginBottom:7,padding:"11px 13px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:q.type==="open"?8:0,flexWrap:"wrap"}}>
                      <span style={{color:C.primary,fontWeight:700,minWidth:26,fontSize:14}}>{i+1}.</span>
                      {q.type==="open"&&<span style={{...S.badge,background:C.successLight,color:C.successDark,fontSize:11}}>📝 Ochiq</span>}
                      {q.type==="closed"&&(
                        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                          {Array.from({length:q.optionsCount},(_,oi)=>String.fromCharCode(65+oi)).map(opt=>(
                            <button key={opt} onClick={()=>updQ(i,"correctAnswer", q.correctAnswer===opt ? "" : opt)} style={{width:36,height:36,borderRadius:"50%",border:`2px solid`,borderColor:q.correctAnswer===opt?C.primary:C.border,background:q.correctAnswer===opt?C.primary:C.card,color:q.correctAnswer===opt?"white":C.text,cursor:"pointer",fontWeight:700,fontSize:13,transition:"all 0.15s"}}>{opt}</button>
                          ))}
                        </div>
                      )}
                    </div>
                    {q.type==="open"&&(
                      q.subParts?.length>0?(
                        <div>
                          {q.subParts.map((s,si)=>{
                            const isAct=kbdOpen?.qIdx===i&&kbdOpen?.sub===si;
                            return (
                              <div key={si} style={{marginBottom:8}}>
                                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:4}}>
                                  <span style={{color:C.warning,minWidth:38,fontSize:13,fontWeight:700}}>{i+1}{s.label})</span>
                                  <div onClick={()=>setKbdOpen(isAct?null:{qIdx:i,sub:si})} style={{flex:1,minHeight:44,padding:"10px 14px",background:isAct?"#EEF1FF":"#F8F9FF",border:`2px solid ${isAct?C.primary:s.answer?C.success:C.border}`,borderRadius:10,cursor:"pointer",fontSize:17,display:"flex",alignItems:"center",boxShadow:isAct?"0 0 0 3px rgba(79,110,247,0.2)":"none"}}>
                                    {s.answer?<span ref={el=>{if(el&&window.katex){try{window.katex.render(toLatex(s.answer),el,{throwOnError:false})}catch{}}}} style={{fontSize:18,fontFamily:"KaTeX_Main,serif",color:"#15803D"}}/>:<span style={{color:C.textLight,fontSize:13}}>Bosing...</span>}
                                  </div>
                                  <button onClick={()=>delSub(i,si)} style={{...S.btnSmall,background:C.danger,padding:"4px 10px"}}>✕</button>
                                </div>

                              </div>
                            );
                          })}
                          <button onClick={()=>addSub(i)} style={{...S.btnSmall,background:"#E2E8F0",color:C.textMid,fontSize:12,marginTop:4}}>+ Kichik band</button>
                        </div>
                      ):(()=>{
                        const isAct=kbdOpen?.qIdx===i&&kbdOpen?.sub===null;
                        return (
                          <div>
                            <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:4}}>
                              <div onClick={e=>{setKbdOpen(isAct?null:{qIdx:i,sub:null});if(!isAct)setTimeout(()=>e.currentTarget.scrollIntoView({behavior:"smooth",block:"center"}),350);}} style={{flex:1,minHeight:44,padding:"10px 14px",background:isAct?"#EEF1FF":"#F8F9FF",border:`2px solid ${isAct?C.primary:q.correctAnswer?C.success:C.border}`,borderRadius:10,cursor:"pointer",fontSize:17,display:"flex",alignItems:"center",boxShadow:isAct?"0 0 0 3px rgba(79,110,247,0.2)":"none"}}>
                                {q.correctAnswer?<span ref={el=>{if(el&&window.katex){try{window.katex.render(toLatex(q.correctAnswer),el,{throwOnError:false})}catch{}}}} style={{fontSize:18,fontFamily:"KaTeX_Main,serif",color:"#15803D"}}/>:<span style={{color:C.textLight,fontSize:13}}>Formulali javob uchun bosing...</span>}
                              </div>
                              <button onClick={()=>addSub(i)} style={{...S.btnSmall,background:"#E2E8F0",color:C.textMid,fontSize:12}}>+ Kichik band</button>
                            </div>

                          </div>
                        );
                      })()
                    )}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      )}


      {/* Fixed bottom keyboard for admin */}
      {kbdOpen && step===2 && (
        <div style={{
          position:"fixed", bottom:0, left:0, right:0, zIndex:9000,
          transition:"transform 0.3s cubic-bezier(0.32,0.72,0,1)",
        }}>
          <MathKeyboard
            key={kbdOpen.qIdx + "_" + (kbdOpen.sub ?? "x")}
            isAdmin={true}
            initValue={kbdOpen.sub!==null
              ? (questions[kbdOpen.qIdx]?.subParts?.[kbdOpen.sub]?.answer||"")
              : (questions[kbdOpen.qIdx]?.correctAnswer||"")}
            onChange={v=>{
              if(kbdOpen.sub!==null) updSub(kbdOpen.qIdx,kbdOpen.sub,v);
              else updQ(kbdOpen.qIdx,"correctAnswer",v);
            }}
            onClose={()=>setKbdOpen(null)}
          />
        </div>
      )}
    </div>
  );
}

// ===== ADMIN PANEL =====
// ── TeacherManager: proper top-level component (hooks work here) ──
function TeacherManager() {
  const [list, setList] = useState(()=>db.get("teachers")||[]);
  const [form, setForm] = useState({name:"",login:"",password:""});
  const [showPwd, setShowPwd] = useState(false);
  const [err, setErr] = useState("");
  const [confirmModal, setConfirmModal] = useState(null);

  const add = () => {
    if(!form.name.trim()||!form.login.trim()||!form.password){setErr("Barcha maydonlarni to'ldiring!");return;}
    if(form.password.length<4){setErr("Parol kamida 4 ta belgi!");return;}
    if(form.login.trim()===ADMIN_LOGIN){setErr("Bu login band!");return;}
    const cur=db.get("teachers")||[];
    if(cur.find(t=>t.login===form.login.trim())){setErr("Bu login band!");return;}
    const upd=[...cur,{id:Date.now(),name:form.name.trim(),login:form.login.trim(),password:form.password}];
    db.set("teachers",upd); setList(upd);
    setForm({name:"",login:"",password:""}); setErr("");
  };

  const del = (id) => {
    setConfirmModal({message:"O'qituvchi o'chirilsinmi?", onConfirm: () => {
      const upd=(db.get("teachers")||[]).filter(t=>t.id!==id);
      db.set("teachers",upd); setList(upd);
      setConfirmModal(null);
    }});
  };

  return (
    <div style={{...S.card,padding:16,marginBottom:18,border:`2px solid ${C.primary}`}}>
      {confirmModal && <ConfirmModal message={confirmModal.message} onConfirm={confirmModal.onConfirm} onCancel={()=>setConfirmModal(null)}/>}
      <h3 style={{margin:"0 0 14px",fontSize:15,color:C.primary}}>👩‍🏫 Yordamchi o'qituvchilar</h3>
      {err&&<div style={S.err}>{err}</div>}
      <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:12}}>
        {[["Ismi","name","text","To'liq ismi"],["Login","login","text","login (unikal)"],["Parol","password",showPwd?"text":"password","min 4 ta belgi"]].map(([lbl,k,tp,ph])=>(
          <div key={k} style={{display:"flex",gap:8,alignItems:"center"}}>
            <label style={{...S.label,margin:0,minWidth:56,fontSize:12}}>{lbl}</label>
            <div style={{flex:1,position:"relative"}}>
              <input type={tp} value={form[k]} onChange={e=>setForm(p=>({...p,[k]:e.target.value}))}
                style={{...S.input,margin:0,fontSize:13}} placeholder={ph}
                onKeyDown={e=>e.key==="Enter"&&add()}/>
              {k==="password"&&<button onClick={()=>setShowPwd(s=>!s)} type="button"
                style={{position:"absolute",right:8,top:10,background:"none",border:"none",cursor:"pointer",fontSize:16}}>{showPwd?"🙈":"👁"}</button>}
            </div>
          </div>
        ))}
        <button onClick={add} style={{...S.btnPrimary,margin:0,padding:"10px",fontSize:13}}>+ Qo'shish</button>
      </div>
      {list.length===0
        ? <p style={{color:C.textLight,fontSize:13,margin:0,textAlign:"center"}}>Hali yordamchi o'qituvchi qo'shilmagan</p>
        : <table style={{...S.table,fontSize:13}}>
            <thead><tr>{["Ismi","Login","Amal"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>{list.map(t=>(
              <tr key={t.id}>
                <td style={S.td}><b>{t.name}</b></td>
                <td style={S.td}><code style={{background:C.primaryLight,color:C.primary,padding:"2px 8px",borderRadius:4,fontSize:12}}>{t.login}</code></td>
                <td style={S.td}>
                  <button onClick={()=>del(t.id)} style={{...S.btnSmall,background:C.danger,padding:"4px 10px",fontSize:12}}>🗑 O'chirish</button>
                </td>
              </tr>
            ))}</tbody>
          </table>
      }
    </div>
  );
}


function AdminPanel({ onLogout, isFullAdmin=true, teacherInfo=null }) {
  const roleName = isFullAdmin ? "Admin" : (teacherInfo?.name || "O'qituvchi");
  const [tab,setTab]=useState("tests");
  const [tests,setTests]=useState([]); const [users,setUsers]=useState([]); const [results,setResults]=useState([]);
  const [creating,setCreating]=useState(false); const [editing,setEditing]=useState(null);
  const [adminDocPreview,setAdminDocPreview]=useState(null); // {type:"pdf"|"latex", url?, source?, name}
  const [confirmModal,setConfirmModal]=useState(null);
  const [exportModal,setExportModal]=useState(null); // {dataUrl, filename, tsv, isBinary}
  const [regradingId,setRegradingId]=useState(null); // hozir qayta baholanayotgan test id
  const [regradeDoneMsg,setRegradeDoneMsg]=useState(null);
  const [raschModal,setRaschModal]=useState(null); // {test, settings} — sozlamalarni tahrirlash oynasi
  const [raschBusyId,setRaschBusyId]=useState(null);
  const [raschDoneMsg,setRaschDoneMsg]=useState(null);
  const [raschUploadTarget,setRaschUploadTarget]=useState(null); // qaysi test uchun fayl yuklanmoqda
  const [raschUploading,setRaschUploading]=useState(false);
  const raschFileInputRef=useRef(null);

  // ===== Mustaqil "Excel -> Rash" kalkulyatori (saytda ro'yxatdan o'tgan bo'lish shart emas) =====
  const [raschCalcRawRows,setRaschCalcRawRows]=useState(null); // fayldan o'qilgan xom qatorlar (o'zgarmaydi)
  const [raschCalcRows,setRaschCalcRows]=useState(null); // hisoblangan (ko'rsatiladigan) qatorlar
  const [raschCalcFileName,setRaschCalcFileName]=useState(null);
  const [raschCalcBusy,setRaschCalcBusy]=useState(false);
  const [raschCalcError,setRaschCalcError]=useState(null);
  const [raschCalcSettings,setRaschCalcSettings]=useState({...DEFAULT_RASCH_SETTINGS});
  const raschCalcFileInputRef=useRef(null);
  const handleRaschCalcFile = (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setRaschCalcBusy(true); setRaschCalcError(null); setRaschCalcRows(null); setRaschCalcRawRows(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const { rows } = parseUploadedResultsSheet(wb);
        if (!rows.length) {
          setRaschCalcError("Faylda F.I.O / Ism ustuni topilmadi. Fayl tuzilishini tekshiring (birinchi ustun ism, keyingilari 1/0 javoblar bo'lishi kerak).");
        } else {
          const computed = computeRaschFromFileRows(rows, raschCalcSettings);
          if (!computed.length) {
            setRaschCalcError("Hech qanday qator hisoblanmadi — BALL ustuni yoki To'g'ri/Jami savol (yoki 1/0 javoblar) ustunlari topilmadi.");
          } else {
            setRaschCalcRawRows(rows);
            setRaschCalcRows(computed);
            setRaschCalcFileName(file.name);
          }
        }
      } catch (err) {
        setRaschCalcError("Faylni o'qib bo'lmadi. .xlsx formatida ekanligini tekshiring.");
      }
      setRaschCalcBusy(false);
    };
    reader.onerror = () => { setRaschCalcBusy(false); setRaschCalcError("Faylni o'qishda xatolik yuz berdi."); };
    reader.readAsArrayBuffer(file);
  };
  const recalcRaschCalc = () => {
    if (!raschCalcRawRows) return;
    setRaschCalcRows(computeRaschFromFileRows(raschCalcRawRows, raschCalcSettings));
  };
  // Yuklangan faylni (yuqoridagi kalkulyatorda hisoblangan) TANLANGAN bitta testga bog'lab,
  // shu testni saytda topshirganlar bilan birga (bitta Rash hisobida) profillarga saqlaydi.
  const [raschCalcAttachTestId,setRaschCalcAttachTestId]=useState("");
  const [raschCalcAttaching,setRaschCalcAttaching]=useState(false);
  const attachRaschCalcToTest = () => {
    if (!raschCalcRawRows || !raschCalcAttachTestId) return;
    const test = tests.find(t=>t.id===Number(raschCalcAttachTestId) || t.id===raschCalcAttachTestId);
    if (!test) return;
    setRaschCalcAttaching(true);
    setTimeout(() => {
      const settings = raschCalcSettings;
      const r = importRaschFromRows(test, raschCalcRawRows, settings);
      db.set("tests",(db.get("tests")||[]).map(t=>t.id===test.id?{...t,raschSettings:settings,raschCalculatedAt:r.calculatedAt}:t));
      reload();
      let msg = `💾 "${test.name}" — saytda topshirganlar va bu fayl birgalikda hisoblanib, ${r.matched} ta o'quvchi profiliga yozildi.`;
      if (r.fileOnly.length) msg += ` ${r.fileOnly.length} ta o'quvchi saytda topilmadi (yuqoridagi jadval endi ular bilan yangilandi).`;
      setRaschDoneMsg(msg);
      setRaschCalcAttaching(false);
      setTimeout(()=>setRaschDoneMsg(null), 8000);
    }, 30);
  };

  const triggerRaschUpload = (test) => { setRaschUploadTarget(test); requestAnimationFrame(()=>raschFileInputRef.current?.click()); };
  const handleRaschFile = (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file || !raschUploadTarget) return;
    const test = raschUploadTarget;
    setRaschUploading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const { rows } = parseUploadedResultsSheet(wb);
        if (!rows.length) {
          setRaschDoneMsg("⚠️ Faylda F.I.O / Ism ustuni topilmadi. Fayl tuzilishini tekshiring.");
        } else {
          const settings = test.raschSettings || DEFAULT_RASCH_SETTINGS;
          const r = importRaschFromRows(test, rows, settings);
          db.set("tests",(db.get("tests")||[]).map(t=>t.id===test.id?{...t,raschSettings:settings,raschCalculatedAt:r.calculatedAt}:t));
          reload();
          let msg = `📤 "${test.name}" — saytda topshirganlar VA fayldagilar birgalikda, bitta Rash hisobida qayta ishlandi. ${r.matched} ta o'quvchi profiliga yozildi.`;
          if (r.fileOnly.length) { msg += ` ${r.fileOnly.length} ta faylda bor, lekin saytda ro'yxatdan o'tmagan/topilmadi — natijalari "🎯 Rash (Excel)" bo'limida ko'rish/yuklab olish uchun tayyor.`; setRaschCalcRows(r.fileOnly); setRaschCalcRawRows(r.fileOnly); setRaschCalcFileName(`${test.name} — saytda topilmagan o'quvchilar`); }
          if (r.unusable.length) msg += ` ${r.unusable.length} ta qator hisoblab bo'lmadi (javob/ball yo'q).`;
          setRaschDoneMsg(msg);
        }
      } catch (err) {
        setRaschDoneMsg("❌ Faylni o'qib bo'lmadi. .xlsx formatida ekanligini tekshiring.");
      }
      setRaschUploading(false);
      setRaschUploadTarget(null);
      setTimeout(()=>setRaschDoneMsg(null), 7000);
    };
    reader.onerror = () => { setRaschUploading(false); setRaschUploadTarget(null); setRaschDoneMsg("❌ Faylni o'qishda xatolik yuz berdi."); setTimeout(()=>setRaschDoneMsg(null),5000); };
    reader.readAsArrayBuffer(file);
  };

  const openRaschModal = (test) => {
    setRaschModal({ test, settings: { ...(test.raschSettings || DEFAULT_RASCH_SETTINGS) } });
  };
  const runRasch = () => {
    const { test, settings } = raschModal;
    setRaschBusyId(test.id);
    setTimeout(() => { // UI qotib qolmasligi uchun keyingi tikda hisoblaymiz
      const r = computeRaschForTest(test, settings);
      if (r.count > 0) {
        db.set("tests", (db.get("tests")||[]).map(t=>t.id===test.id?{...t,raschSettings:settings,raschCalculatedAt:r.calculatedAt,raschMean:r.mean,raschStd:r.std}:t));
      }
      setRaschBusyId(null);
      setRaschModal(null);
      reload();
      setRaschDoneMsg(r.count>0?`🎯 "${test.name}" — ${r.count} ta o'quvchi natijasi Rash modeli bo'yicha hisoblandi.`:`"${test.name}" bo'yicha hali hech kim test topshirmagan.`);
      setTimeout(()=>setRaschDoneMsg(null), 4500);
    }, 30);
  };

  const handleRegrade = (test) => {
    setConfirmModal({
      message: `"${test.name}" testi bo'yicha barcha topshirilgan javoblar joriy (yangilangan) to'g'ri javob kaliti bilan qayta tekshiriladi. Davom etilsinmi?`,
      confirmLabel: "Qayta baholash",
      danger: false,
      onConfirm: async () => {
        setConfirmModal(null);
        setRegradingId(test.id);
        const count = await regradeTestResults(test);
        setRegradingId(null);
        reload();
        setRegradeDoneMsg(`✅ "${test.name}" — ${count} ta natija qayta baholandi.`);
        setTimeout(()=>setRegradeDoneMsg(null), 4000);
      },
    });
  };

  const reload=()=>{setTests(db.get("tests")||[]);setUsers(db.get("users")||[]);setResults(db.get("results")||[]);};
  useEffect(reload,[tab]);

  // Rejalashtirilgan testlarni har soniyada tekshirib, vaqti kelganlarini avtomatik faollashtiradi
  useEffect(()=>{
    const t=setInterval(()=>{ if(autoActivateScheduledTests()) reload(); },1000);
    return ()=>clearInterval(t);
  },[]);

  // Boshqa qurilmadan (masalan admin/o'qituvchi boshqa telefon/kompyuterda) kiritilgan
  // o'zgarish Firestore orqali kelganda darhol yangilaymiz
  useEffect(()=>{
    const h=()=>reload();
    window.addEventListener("firestore-sync",h);
    return ()=>window.removeEventListener("firestore-sync",h);
  },[]);

  const saveTest=(t)=>{let ts=db.get("tests")||[];ts=editing?ts.map(x=>x.id===t.id?t:x):[...ts,t];db.set("tests",ts);setCreating(false);setEditing(null);reload();};
  const deleteTest=(id)=>{
    setConfirmModal({message:"Test o'chirilsinmi? Bu amalni ortga qaytarib bo'lmaydi.", onConfirm: () => {
      db.set("tests",(db.get("tests")||[]).filter(t=>t.id!==id));reload();setConfirmModal(null);
    }});
  };
  const toggleActive=(id)=>{db.set("tests",(db.get("tests")||[]).map(t=>t.id===id?{...t,active:!t.active,startedAt:!t.active?Date.now():null,scheduledAt:null}:t));reload();};
  const deleteUser=(phone)=>{
    setConfirmModal({message:"Foydalanuvchi o'chirilsinmi?", onConfirm: () => {
      db.set("users",(db.get("users")||[]).filter(u=>u.phone!==phone));reload();setConfirmModal(null);
    }});
  };

  return (
    <div style={S.page}>
      {confirmModal && <ConfirmModal message={confirmModal.message} confirmLabel={confirmModal.confirmLabel} danger={confirmModal.danger} onConfirm={confirmModal.onConfirm} onCancel={()=>setConfirmModal(null)}/>}
      <input ref={raschFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleRaschFile} style={{display:"none"}}/>
      <input ref={raschCalcFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleRaschCalcFile} style={{display:"none"}}/>
      {regradeDoneMsg && (
        <div style={{position:"fixed",top:16,left:"50%",transform:"translateX(-50%)",zIndex:99999,background:C.successDark,color:"white",padding:"12px 20px",borderRadius:12,fontWeight:700,fontSize:13,boxShadow:"0 6px 20px rgba(0,0,0,0.2)",maxWidth:"90%",textAlign:"center"}}>{regradeDoneMsg}</div>
      )}
      {raschDoneMsg && (
        <div style={{position:"fixed",top:16,left:"50%",transform:"translateX(-50%)",zIndex:99999,background:"#6D28D9",color:"white",padding:"14px 20px",borderRadius:12,fontWeight:700,fontSize:12.5,boxShadow:"0 6px 20px rgba(0,0,0,0.2)",maxWidth:"92%",maxHeight:"70vh",overflowY:"auto",textAlign:"left",whiteSpace:"pre-line",lineHeight:1.6}}>{raschDoneMsg}</div>
      )}
      {raschModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 99999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={()=>setRaschModal(null)}>
          <div style={{ background: "white", borderRadius: 16, padding: 22, maxWidth: 420, width: "100%", maxHeight:"88vh", overflowY:"auto", boxShadow: "0 10px 40px rgba(0,0,0,0.25)" }} onClick={e=>e.stopPropagation()}>
            <p style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 800, color: C.text }}>🎯 Rash modeli bo'yicha hisoblash</p>
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: C.textMid, lineHeight: 1.5 }}>"<b>{raschModal.test.name}</b>" testini topshirgan barcha o'quvchilarning natijasi Rash (logistik) modeli asosida qayta baholanadi. Sozlamalarni xohlasangiz o'zgartiring (standart qiymatlar — shablon Excel bilan bir xil).</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
              {[
                ["maxScore","Maksimal ball (chegara)"],
                ["center","Markaz (o'rtacha) ball"],
                ["ncMax","NC chegarasi (bundan past)"],
                ["cMax","C chegarasi"],
                ["cPlusMax","C+ chegarasi"],
                ["bMax","B chegarasi"],
                ["bPlusMax","B+ chegarasi"],
                ["aMax","A chegarasi (yuqorisi A+)"],
              ].map(([key,label])=>(
                <label key={key} style={{fontSize:11.5,color:C.textMid,fontWeight:600}}>{label}
                  <input type="number" step="0.1" value={raschModal.settings[key]}
                    onChange={e=>setRaschModal(p=>({...p,settings:{...p.settings,[key]:e.target.value===""?"":Number(e.target.value)}}))}
                    style={{...S.input,marginTop:4,padding:"7px 9px",fontSize:13}}/>
                </label>
              ))}
            </div>
            <p style={{margin:"0 0 14px",fontSize:11,color:C.textLight}}>Topshirgan o'quvchilar soni: <b>{results.filter(r=>r.testId===raschModal.test.id).length}</b></p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setRaschModal(null)} style={{flex:1,padding:"11px",borderRadius:10,border:`1.5px solid ${C.border}`,background:"white",fontWeight:700,fontSize:13,cursor:"pointer",color:C.textMid}}>Bekor qilish</button>
              <button onClick={runRasch} style={{flex:1,padding:"11px",borderRadius:10,border:"none",background:"#6D28D9",color:"white",fontWeight:800,fontSize:13,cursor:"pointer"}}>🎯 Hisoblash</button>
            </div>
          </div>
        </div>
      )}
      {exportModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 99999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={()=>setExportModal(null)}>
          <div style={{ background: "white", borderRadius: 16, padding: 22, maxWidth: 360, width: "100%", boxShadow: "0 10px 40px rgba(0,0,0,0.25)" }} onClick={e=>e.stopPropagation()}>
            <p style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 800, color: C.text }}>📥 Excel fayl tayyor</p>
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: C.textMid, lineHeight: 1.5 }}>Yuklab olish uchun quyidagi tugmani bosing. Agar tugma ishlamasa (ba'zi ilovalarda cheklangan bo'lishi mumkin), saytni brauzerda (Chrome/Safari) oching.</p>
            <a href={exportModal.dataUrl} download={exportModal.filename} onClick={()=>setTimeout(()=>setExportModal(null),300)} style={{ display:"block", textAlign:"center", padding:"12px", borderRadius:10, background:C.successDark, color:"white", fontWeight:800, fontSize:14, textDecoration:"none", marginBottom:10 }}>⬇️ {exportModal.filename}</a>
            {!exportModal.isBinary && (
              <button onClick={()=>{ navigator.clipboard?.writeText(exportModal.tsv).then(()=>alert("Nusxalandi! Excel'ga joylashtirishingiz mumkin.")).catch(()=>{}); }} style={{ width:"100%", padding:"11px", borderRadius:10, border:`1.5px solid ${C.border}`, background:"white", fontWeight:700, fontSize:13, cursor:"pointer", color:C.textMid, marginBottom:10 }}>📋 Jadval sifatida nusxalash</button>
            )}
            <button onClick={()=>setExportModal(null)} style={{ width:"100%", padding:"10px", borderRadius:10, border:"none", background:"transparent", fontWeight:600, fontSize:13, cursor:"pointer", color:C.textLight }}>Yopish</button>
          </div>
        </div>
      )}
      <div style={{background:C.primary,padding:"14px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 2px 12px rgba(79,110,247,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:26}}>⚙️</span><div><p style={{margin:0,color:"white",fontWeight:800,fontSize:17}}>{isFullAdmin?"Admin Panel":`👩‍🏫 ${roleName}`}</p><p style={{margin:0,color:"rgba(255,255,255,0.7)",fontSize:12}}>{isFullAdmin?"Tizim boshqaruvi":"O'qituvchi paneli"}</p></div></div>
        <button onClick={onLogout} style={{...S.btnSmall,background:"rgba(255,255,255,0.2)",color:"white"}}>Chiqish</button>
      </div>

      {/* Admin document preview modal (PDF or LaTeX) */}
      {adminDocPreview&&(
        <div style={{position:"fixed",inset:0,background:"#1a1a1a",zIndex:9999,display:"flex",flexDirection:"column"}}>
          <div style={{background:C.card,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${C.border}`}}>
            <span style={{fontWeight:700,fontSize:16}}>{adminDocPreview.type==="pdf"?"📄":"∑"} {adminDocPreview.name}</span>
            <button onClick={()=>setAdminDocPreview(null)} style={{...S.btnDanger,padding:"8px 14px"}}>✕ Yopish</button>
          </div>
          {adminDocPreview.type==="pdf"
            ? <PdfViewer url={adminDocPreview.url} persistKey={"admin_pdf_"+adminDocPreview.url?.slice(-8)}/>
            : <ScrollPersistDiv persistKey={"admin_latex_scroll_"+(adminDocPreview.id||adminDocPreview.source?.length||"x")} style={{flex:1,overflowY:"auto",background:"white"}}><LatexDocViewer source={adminDocPreview.source} images={adminDocPreview.images}/></ScrollPersistDiv>
          }
        </div>
      )}

      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,display:"flex",padding:"0 16px"}}>
        {[["tests","📋 Testlar"],["users","👥 O'quvchilar"],["results","📊 Natijalar"],["raschcalc","🎯 Rash (Excel)"]].map(([t,l])=>(
          <button key={t} onClick={()=>{setTab(t);setCreating(false);setEditing(null);}} style={{padding:"14px 18px",background:"none",border:"none",cursor:"pointer",color:tab===t?C.primary:C.textMid,fontWeight:tab===t?800:500,borderBottom:tab===t?`3px solid ${C.primary}`:"3px solid transparent",fontSize:14}}>{l}</button>
        ))}
      </div>
      <div style={{padding:20,maxWidth:960,margin:"0 auto"}}>

        {tab==="tests"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <h3 style={{margin:0}}>Testlar ({tests.length})</h3>
              {!creating&&!editing&&<button onClick={()=>{setCreating(true);setEditing(null);}} style={{...S.btnSmall,background:C.primary}}>+ Yangi Test</button>}
            </div>
            {(creating||editing)&&<TestCreator existing={editing} onSave={saveTest} onCancel={()=>{setCreating(false);setEditing(null);}}/>}
            {tests.map(test=>{
              const tr=results.filter(r=>r.testId===test.id);
              const endAt=test.startedAt?new Date(test.startedAt+test.duration*60000).toLocaleTimeString("uz-UZ",{hour:"2-digit",minute:"2-digit"}):null;
              return (
                <div key={test.id} style={{...S.card,padding:18,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                    <div>
                      <h4 style={{margin:"0 0 5px",fontSize:16}}>{test.name}</h4>
                      <p style={{margin:"0 0 8px",color:C.textMid,fontSize:13}}>{test.questions?.length} savol • {test.duration} daqiqa</p>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                        <span style={{...S.badge,background:test.active?C.successLight:C.dangerLight,color:test.active?C.successDark:C.danger}}>{test.active?"✅ Faol":"⛔ Nofaol"}</span>
                        <span style={{...S.badge,background:C.primaryLight,color:C.primary}}>{tr.length} topshirdi</span>
                        {test.active&&endAt&&<span style={{...S.badge,background:C.warningLight,color:C.warning}}>⏱ {endAt} da tugaydi</span>}
                        {!test.active&&test.scheduledAt&&<span style={{...S.badge,background:"#EEF1FF",color:C.primary}}>📅 {formatScheduled(test.scheduledAt)} da boshlanadi</span>}
                        {availableDocLangs(test).map(l=>{
                          const d=getLangDoc(test,l.code);
                          return d.pdfUrl
                            ? <button key={l.code} onClick={()=>setAdminDocPreview({type:"pdf",url:d.pdfUrl,name:`${test.name} — ${l.full}`})} style={{...S.badge,background:"#FEF3C7",color:"#92400E",border:"none",cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4}}><LangFlag lang={l} size={13}/> 📄 {l.label}</button>
                            : <button key={l.code} onClick={()=>setAdminDocPreview({type:"latex",source:d.latexSource,name:`${test.name} — ${l.full}`,id:test.id+"_"+l.code,images:d.latexImages})} style={{...S.badge,background:"#EEF1FF",color:C.primary,border:"none",cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4}}><LangFlag lang={l} size={13}/> ∑ {l.label}</button>;
                        })}
                        {test.accessCode&&<span style={{...S.badge,background:"#FEF3C7",color:"#92400E"}}>🔒 Kod: {test.accessCode}{test.codePrice?` (${test.codePrice} so'm)`:""}</span>}
                        {test.targetGroups&&test.targetGroups.length>0&&<span style={{...S.badge,background:C.primaryLight,color:C.primary}}>👥 {test.targetGroups.join(", ")}</span>}
                        {test.raschCalculatedAt&&<span style={{...S.badge,background:"#EDE9FE",color:"#6D28D9"}}>🎯 Rash hisoblangan ({new Date(test.raschCalculatedAt).toLocaleDateString("uz-UZ")})</span>}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      <button onClick={()=>toggleActive(test.id)} style={{...S.btnSmall,background:test.active?C.danger:C.success}}>{test.active?"⛔ To'xtatish":"✅ Faollashtirish"}</button>
                      <button onClick={()=>{setEditing(test);setCreating(false);}} style={{...S.btnSmall,background:C.primary}}>✏️ Tahrirlash</button>
                      <button onClick={()=>handleRegrade(test)} disabled={regradingId===test.id} style={{...S.btnSmall,background:regradingId===test.id?"#94A3B8":"#7C3AED"}}>{regradingId===test.id?"⏳ Baholanmoqda...":"🔄 Qayta baholash"}</button>
                      <button onClick={()=>openRaschModal(test)} disabled={tr.length===0||raschBusyId===test.id} style={{...S.btnSmall,background:tr.length===0?"#CBD5E1":"#6D28D9",opacity:raschBusyId===test.id?0.6:1}}>{raschBusyId===test.id?"⏳ Hisoblanmoqda...":"🎯 Rash modeli"}</button>
                      <button onClick={()=>setExportModal(buildExcelExport(test,results,users))} style={{...S.btnSmall,background:C.successDark}}>📥 Excel</button>
                      <button onClick={()=>triggerRaschUpload(test)} disabled={raschUploading} style={{...S.btnSmall,background:"#0891B2",opacity:raschUploading?0.6:1}}>{raschUploading&&raschUploadTarget?.id===test.id?"⏳ Yuklanmoqda...":"📤 Natija yuklash"}</button>
                      <button onClick={()=>deleteTest(test.id)} style={{...S.btnSmall,background:C.danger}}>🗑️</button>
                    </div>
                  </div>
                </div>
              );
            })}
            {tests.length===0&&<div style={S.empty}>Hali test yaratilmagan</div>}
          </div>
        )}

        {tab==="users"&&(
          <div>
            {/* Teacher manager — only full admin sees this */}
            {isFullAdmin && <TeacherManager />}
            <h3 style={{marginBottom:16}}>O'quvchilar ({users.length})</h3>
            <div style={{overflowX:"auto"}}>
              <table style={S.table}>
                <thead><tr>{["#","Ism","Familiya","Guruh","Telefon","Testlar","Amal"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>{users.map((u,i)=>(
                  <tr key={u.phone} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                    <td style={S.td}>{i+1}</td><td style={S.td}>{u.firstName}</td><td style={S.td}>{u.lastName}</td>
                    <td style={S.td}>{u.group}</td>
                    <td style={S.td}><code style={{color:C.primary,background:C.primaryLight,padding:"2px 6px",borderRadius:4}}>{u.phone}</code></td>
                    <td style={S.td}>{results.filter(r=>r.userPhone===u.phone).length}</td>
                    <td style={S.td}>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={()=>{
                          const np = prompt("Yangi parol kiriting ("+u.firstName+" "+u.lastName+" uchun):");
                          if(np&&np.length>=4){
                            const us=(db.get("users")||[]).map(x=>x.phone===u.phone?{...x,password:np}:x);
                            db.set("users",us); reload();
                            alert("Parol yangilandi!");
                          } else if(np!==null){ alert("Parol kamida 4 ta belgidan iborat bo'lsin!"); }
                        }} style={{...S.btnSmall,background:C.primary,padding:"4px 10px"}}>🔑 Parol</button>
                        <button onClick={()=>deleteUser(u.phone)} style={{...S.btnSmall,background:C.danger,padding:"4px 10px"}}>O'chirish</button>
                      </div>
                    </td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
            {users.length===0&&<div style={S.empty}>O'quvchilar yo'q</div>}
          </div>
        )}

        {tab==="results"&&(
          <div>
            <h3 style={{margin:"0 0 16px"}}>Natijalar ({results.length})</h3>
            {(()=>{
              const paidTests = tests.filter(t=>t.accessCode && t.codePrice);
              if (!paidTests.length) return null;
              const totalRevenue = paidTests.reduce((sum,t) => {
                const count = results.filter(r=>r.testId===t.id).length;
                return sum + count * Number(t.codePrice||0);
              }, 0);
              return (
                <div style={{...S.card,padding:16,marginBottom:16,background:"linear-gradient(135deg,#FEF3C7,#FFFBEB)",border:"1.5px solid #FDE68A"}}>
                  <p style={{margin:"0 0 4px",color:"#92400E",fontSize:13,fontWeight:700}}>💰 Maxfiy testlardan tushum (taxminiy)</p>
                  <p style={{margin:0,color:"#92400E",fontSize:28,fontWeight:900}}>{totalRevenue.toLocaleString()} so'm</p>
                </div>
              );
            })()}
            {tests.map(test=>{
              const tr=results.filter(r=>r.testId===test.id);
              if(!tr.length) return null;
              return (
                <div key={test.id} style={{marginBottom:24}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                    <h4 style={{margin:0,color:C.primary}}>{test.name}</h4>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={()=>openRaschModal(test)} disabled={raschBusyId===test.id} style={{...S.btnSmall,background:"#6D28D9",opacity:raschBusyId===test.id?0.6:1}}>{raschBusyId===test.id?"⏳...":"🎯 Rash modeli"}</button>
                      <button onClick={()=>setExportModal(buildExcelExport(test,results,users))} style={{...S.btnSmall,background:C.successDark}}>📥 Excel</button>
                      <button onClick={()=>triggerRaschUpload(test)} disabled={raschUploading} style={{...S.btnSmall,background:"#0891B2",opacity:raschUploading?0.6:1}}>{raschUploading&&raschUploadTarget?.id===test.id?"⏳...":"📤 Natija yuklash"}</button>
                    </div>
                  </div>
                  <div style={{overflowX:"auto"}}>
                    <table style={S.table}>
                      <thead><tr>{["#","F.I.O","Guruh","Ball","Foiz","Rash ball","Daraja","Vaqt","Sana"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>{tr.sort((a,b)=>b.totalScore-a.totalScore).map((r,i)=>{
                        const u=users.find(u=>u.phone===r.userPhone);
                        const pct=Math.round((r.totalScore/test.questions.length)*100);
                        const dGrade=r.rasch?.daraja;
                        const dColor=dGrade==="NC"?C.danger:dGrade==="C"||dGrade==="C+"?C.warning:C.successDark;
                        return (
                          <tr key={r.id} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                            <td style={S.td}>{i+1}</td>
                            <td style={S.td}>{u?`${u.firstName} ${u.lastName}`:r.userPhone}</td>
                            <td style={S.td}>{u?.group||"-"}</td>
                            <td style={S.td}><b style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger}}>{r.totalScore}</b>/{test.questions.length}</td>
                            <td style={S.td}><span style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger,fontWeight:700}}>{pct}%</span></td>
                            <td style={S.td}>{r.rasch?<b style={{color:"#6D28D9"}}>{r.rasch.ball.toFixed(1)}</b>:<span style={{color:C.textLight}}>—</span>}</td>
                            <td style={S.td}>{dGrade?<span style={{...S.badge,background:dColor+"22",color:dColor,fontWeight:800}}>{dGrade}</span>:<span style={{color:C.textLight}}>—</span>}</td>
                            <td style={S.td}>{r.timeTaken?`${r.timeTaken} daq`:"-"}</td>
                            <td style={S.td}>{new Date(r.id).toLocaleDateString("uz-UZ")}</td>
                          </tr>
                        );
                      })}</tbody>
                    </table>
                  </div>
                </div>
              );
            })}
            {results.length===0&&<div style={S.empty}>Hali natijalar yo'q</div>}
          </div>
        )}

        {tab==="raschcalc"&&(
          <div>
            <h3 style={{margin:"0 0 6px"}}>🎯 Rash modeli — Excel fayldan hisoblash</h3>
            <p style={{margin:"0 0 18px",color:C.textMid,fontSize:13,lineHeight:1.5}}>Excel faylni yuklang — birinchi ustun F.I.O (yoki Ism), keyingi ustunlar har bir savol uchun 1/0 javoblar (yoki tayyor "BALL"/"Daraja" ustunlari) bo'lishi kifoya. Natija shu yerda ko'rinadi va Excel qilib yuklab olinadi. Xohlasangiz, pastda testlardan birini tanlab, natijani o'sha TEST bilan bog'lab (shu testni saytda topshirganlar bilan birgalikda hisoblab) o'quvchilar profiliga ham saqlashingiz mumkin — har bir test alohida-alohida, o'zining natijasi bilan hisoblanadi.</p>

            <div style={{...S.card,padding:16,marginBottom:16}}>
              <p style={{margin:"0 0 10px",fontWeight:800,fontSize:13}}>⚙️ Sozlamalar</p>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10,marginBottom:14}}>
                {[
                  ["maxScore","Maksimal ball"],["center","Markaz ball"],["ncMax","NC chegarasi"],["cMax","C chegarasi"],
                  ["cPlusMax","C+ chegarasi"],["bMax","B chegarasi"],["bPlusMax","B+ chegarasi"],["aMax","A chegarasi"],
                ].map(([key,label])=>(
                  <label key={key} style={{fontSize:11,color:C.textMid,fontWeight:600}}>{label}
                    <input type="number" step="0.1" value={raschCalcSettings[key]}
                      onChange={e=>setRaschCalcSettings(p=>({...p,[key]:e.target.value===""?"":Number(e.target.value)}))}
                      style={{...S.input,marginTop:4,padding:"7px 9px",fontSize:13,marginBottom:0}}/>
                  </label>
                ))}
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <button onClick={()=>raschCalcFileInputRef.current?.click()} disabled={raschCalcBusy} style={{...S.btnSmall,background:"#6D28D9",padding:"11px 20px",fontSize:13,opacity:raschCalcBusy?0.6:1}}>{raschCalcBusy?"⏳ Hisoblanmoqda...":"📤 Excel faylni yuklash"}</button>
                {raschCalcRawRows&&<button onClick={recalcRaschCalc} style={{...S.btnSmall,background:C.primary,padding:"11px 20px",fontSize:13}}>🔄 Sozlamalar bilan qayta hisoblash</button>}
                {raschCalcRows&&raschCalcRows.length>0&&<button onClick={()=>{const ex=buildRaschCalcExport(raschCalcRows,raschCalcSettings); if(ex) setExportModal(ex);}} style={{...S.btnSmall,background:C.successDark,padding:"11px 20px",fontSize:13}}>📥 Natijani Excel qilib olish</button>}
              </div>
              {raschCalcFileName&&<p style={{margin:"10px 0 0",fontSize:12,color:C.textLight}}>Fayl: <b>{raschCalcFileName}</b> • {raschCalcRawRows?.length||0} qator o'qildi</p>}
              {raschCalcError&&<p style={{margin:"10px 0 0",fontSize:12.5,color:C.danger,fontWeight:600}}>⚠️ {raschCalcError}</p>}
            </div>

            {raschCalcRawRows&&raschCalcRawRows.length>0&&(
              <div style={{...S.card,padding:16,marginBottom:16,background:"#F5F3FF",border:"1.5px solid #DDD6FE"}}>
                <p style={{margin:"0 0 10px",fontWeight:800,fontSize:13,color:"#6D28D9"}}>💾 Bu natijani bitta testga bog'lab, o'quvchilar profiliga saqlash</p>
                <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                  <select value={raschCalcAttachTestId} onChange={e=>setRaschCalcAttachTestId(e.target.value)} style={{...S.input,marginBottom:0,maxWidth:320,flex:"1 1 240px"}}>
                    <option value="">— Testni tanlang —</option>
                    {tests.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
                  </select>
                  <button onClick={attachRaschCalcToTest} disabled={!raschCalcAttachTestId||raschCalcAttaching} style={{...S.btnSmall,background:raschCalcAttachTestId?"#6D28D9":"#CBD5E1",padding:"11px 20px",fontSize:13,opacity:raschCalcAttaching?0.6:1}}>{raschCalcAttaching?"⏳ Saqlanmoqda...":"💾 Ushbu testga saqlash"}</button>
                </div>
                <p style={{margin:"8px 0 0",fontSize:11,color:C.textLight}}>Tanlangan test bo'yicha saytda topshirganlar + shu fayl birgalikda bitta Rash hisobida qayta hisoblanadi va mos o'quvchilarning profiliga yoziladi.</p>
              </div>
            )}

            {raschCalcRows&&raschCalcRows.length>0&&(
              <div style={{overflowX:"auto"}}>
                <table style={S.table}>
                  <thead><tr>{["#","F.I.O","Guruh","To'g'ri","Jami","Theta","Z-ball","BALL","Daraja"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>{[...raschCalcRows].sort((a,b)=>(b.ball??-999)-(a.ball??-999)).map((r,i)=>{
                    const dg=r.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                    return (
                      <tr key={i} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                        <td style={S.td}>{i+1}</td>
                        <td style={S.td}>{r.name}</td>
                        <td style={S.td}>{r.group||"-"}</td>
                        <td style={S.td}>{r.correct??"-"}</td>
                        <td style={S.td}>{r.total??"-"}</td>
                        <td style={S.td}>{r.theta!=null?r.theta.toFixed(3):"-"}</td>
                        <td style={S.td}>{r.zBall!=null?r.zBall.toFixed(3):"-"}</td>
                        <td style={S.td}><b style={{color:"#6D28D9"}}>{r.ball!=null?r.ball.toFixed(1):"-"}</b></td>
                        <td style={S.td}>{dg?<span style={{...S.badge,background:dc+"22",color:dc,fontWeight:800}}>{dg}</span>:"-"}</td>
                      </tr>
                    );
                  })}</tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ===== STUDENT DASHBOARD =====
function StudentDashboard({ user, onLogout }) {
  const [tab,setTab]=useState("tests");
  const [tests,setTests]=useState([]); const [results,setResults]=useState([]);
  const [activeTest,setActiveTest]=useState(null); const [viewResult,setViewResult]=useState(null);
  const [docModal,setDocModal]=useState(null); // {type:"pdf"|"latex", url?, source?, name}
  const [codeModal,setCodeModal]=useState(null); // test waiting for code entry
  const [codeInput,setCodeInput]=useState("");
  const [codeErr,setCodeErr]=useState("");
  // Test faqat tanlangan guruh(lar)ga mo'ljallangan bo'lsa, o'quvchi o'z guruhida bo'lsagina ko'rsin
  const visibleForMe = (t) => !t.targetGroups || t.targetGroups.length===0 || t.targetGroups.includes(user.group);
  const [unlockedTests,setUnlockedTests]=useState(()=>{
    try { return JSON.parse(localStorage.getItem("unlockedTests_"+user.phone)||"[]"); } catch { return []; }
  });
  const [now,setNow]=useState(Date.now());

  const reload=()=>{setTests(db.get("tests")||[]);setResults((db.get("results")||[]).filter(r=>r.userPhone===user.phone));};
  useEffect(reload,[user.phone,tab]);

  // Rejalashtirilgan testlar vaqti kelganda avtomatik faollashadi; sanoq va vaqt har soniya yangilanadi
  useEffect(()=>{
    const t=setInterval(()=>{
      if(autoActivateScheduledTests()) reload();
      setNow(Date.now());
    },1000);
    return ()=>clearInterval(t);
  },[]);

  // Admin boshqa qurilmadan test yaratsa/faollashtirsa, Firestore orqali darhol shu yerda ko'rinadi
  useEffect(()=>{
    const h=()=>reload();
    window.addEventListener("firestore-sync",h);
    return ()=>window.removeEventListener("firestore-sync",h);
  },[]);

  const startTest = (test) => {
    if (test.accessCode && !unlockedTests.includes(test.id)) {
      setCodeModal(test); setCodeInput(""); setCodeErr("");
      return;
    }
    setActiveTest(test);
  };

  const submitCode = () => {
    if (!codeModal) return;
    if (codeInput.trim().toUpperCase() === codeModal.accessCode) {
      const next = [...unlockedTests, codeModal.id];
      setUnlockedTests(next);
      localStorage.setItem("unlockedTests_"+user.phone, JSON.stringify(next));
      setActiveTest(codeModal);
      setCodeModal(null);
    } else {
      setCodeErr("Noto'g'ri kod! Qaytadan urinib ko'ring.");
    }
  };

  if(activeTest) return <TestTaking test={activeTest} user={user} onFinish={()=>{setActiveTest(null);setTab("monitoring");reload();}} onExit={()=>setActiveTest(null)}/>;
  if(viewResult){const t=(db.get("tests")||[]).find(t=>t.id===viewResult.testId);return <ResultDetail result={viewResult} test={t} onBack={()=>setViewResult(null)}/>;}

  return (
    <div style={S.page}>
      {/* Document Modal — PDF or LaTeX */}
      {docModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",zIndex:9999,display:"flex",flexDirection:"column"}}>
          <div style={{background:C.card,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${C.border}`}}>
            <span style={{fontWeight:700,fontSize:16}}>{docModal.type==="pdf"?"📄":"∑"} {docModal.name||"Test varianti"}</span>
            <button onClick={()=>setDocModal(null)} style={{...S.btnDanger,padding:"8px 14px"}}>✕ Yopish</button>
          </div>
          {docModal.type==="pdf"
            ? <PdfViewer url={docModal.url} persistKey={"student_pdf_"+docModal.url?.slice(-8)}/>
            : <ScrollPersistDiv persistKey={"student_latex_scroll_"+(docModal.id||"x")} style={{flex:1,overflowY:"auto",background:"white"}}><LatexDocViewer source={docModal.source} images={docModal.images}/></ScrollPersistDiv>
          }
        </div>
      )}

      {/* Access code modal */}
      {codeModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{...S.card,padding:28,maxWidth:380,width:"100%"}}>
            <div style={{textAlign:"center",marginBottom:16}}>
              <div style={{fontSize:44}}>🔒</div>
              <h3 style={{margin:"8px 0 4px",fontSize:18}}>Maxfiy test</h3>
              <p style={{margin:0,color:C.textMid,fontSize:14}}>{codeModal.name}</p>
              {codeModal.codePrice&&<p style={{margin:"6px 0 0",color:"#92400E",fontWeight:700,fontSize:16}}>{codeModal.codePrice} so'm</p>}
            </div>
            <p style={{color:C.textMid,fontSize:13,textAlign:"center",marginBottom:14}}>
              Testga kirish uchun to'lov qiling va sizga berilgan kodni kiriting.
            </p>
            {codeErr&&<div style={S.err}>{codeErr}</div>}
            <input value={codeInput} onChange={e=>setCodeInput(e.target.value.toUpperCase())}
              onKeyDown={e=>e.key==="Enter"&&submitCode()}
              style={{...S.input,textAlign:"center",fontFamily:"monospace",fontWeight:700,fontSize:18,letterSpacing:2}}
              placeholder="KIRISH KODI" autoFocus/>
            <button onClick={submitCode} style={S.btnPrimary}>Tasdiqlash ✓</button>
            <button onClick={()=>setCodeModal(null)} style={{...S.btnGhost,width:"100%",marginTop:8}}>Bekor qilish</button>
          </div>
        </div>
      )}
      <div style={{background:C.primary,padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 2px 12px rgba(79,110,247,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:40,height:40,borderRadius:"50%",background:"rgba(255,255,255,0.25)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:17,color:"white"}}>{user.firstName[0]}</div>
          <div><p style={{margin:0,color:"white",fontWeight:700}}>{user.firstName} {user.lastName}</p><p style={{margin:0,color:"rgba(255,255,255,0.7)",fontSize:12}}>{user.group}</p></div>
        </div>
        <button onClick={onLogout} style={{...S.btnSmall,background:"rgba(255,255,255,0.2)",color:"white"}}>Chiqish</button>
      </div>
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,display:"flex",padding:"0 16px",overflowX:"auto"}}>
        {[["tests","📋 Testlar"],["monitoring","📊 Monitoring"],["pdfs","📄 Test ko'rish"],["stats","📈 Statistika"]].map(([t,l])=>(
          <button key={t} onClick={()=>{setTab(t);reload();}} style={{padding:"14px 18px",background:"none",border:"none",cursor:"pointer",color:tab===t?C.primary:C.textMid,fontWeight:tab===t?800:500,borderBottom:tab===t?`3px solid ${C.primary}`:"3px solid transparent",fontSize:14,whiteSpace:"nowrap",flexShrink:0}}>{l}</button>
        ))}
      </div>
      <div style={{padding:20,maxWidth:800,margin:"0 auto"}}>
        {tab==="tests"&&(
          <div>
            {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now&&visibleForMe(t)).length>0&&(
              <div style={{marginBottom:24}}>
                <h3 style={{marginBottom:16}}>🗓️ Rejalashtirilgan Testlar</h3>
                {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now&&visibleForMe(t)).map(test=>(
                  <div key={test.id} style={{...S.card,padding:18,marginBottom:12,border:`1.5px dashed ${C.primary}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                      <div>
                        <h4 style={{margin:"0 0 5px",fontSize:16}}>{test.name}</h4>
                        <p style={{margin:"0 0 6px",color:C.textMid,fontSize:13}}>{test.questions?.length} savol • {test.duration} daqiqa</p>
                        <span style={{...S.badge,background:C.primaryLight,color:C.primary}}>📅 Boshlanadi: {formatScheduled(test.scheduledAt)}</span>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <p style={{margin:0,color:C.primary,fontWeight:800,fontSize:15}}>⏳ {formatCountdown(test.scheduledAt-now)}</p>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
            <h3 style={{marginBottom:16}}>Faol Testlar</h3>
            {tests.filter(t=>t.active&&visibleForMe(t)).map(test=>{
              const myRes=results.find(r=>r.testId===test.id);
              let timeInfo=null, expired=false;
              if(test.startedAt){
                const rem=Math.max(0,Math.floor((test.startedAt+test.duration*60000-now)/1000));
                expired=rem<=0;
                if(!expired){const m=Math.floor(rem/60),s=rem%60;timeInfo=`${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")} qoldi`;}
                else timeInfo="Vaqt tugadi";
              }
              return (
                <div key={test.id} style={{...S.card,padding:18,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                    <div>
                      <h4 style={{margin:"0 0 5px",fontSize:16}}>{test.name}</h4>
                      <p style={{margin:"0 0 6px",color:C.textMid,fontSize:13}}>{test.questions?.length} savol • {test.duration} daqiqa</p>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                        {timeInfo&&<span style={{...S.badge,background:expired?C.dangerLight:C.warningLight,color:expired?C.danger:C.warning,fontSize:12}}>⏱ {timeInfo}</span>}
                        <DocLangButtons test={test} onOpen={setDocModal} small/>
                      </div>
                    </div>
                    {myRes?(
                      <div style={{textAlign:"right"}}>
                        <p style={{margin:"0 0 4px",color:C.successDark,fontWeight:700}}>✅ Topshirildi</p>
                        <p style={{margin:0,color:C.textMid,fontSize:13}}>{myRes.totalScore}/{test.questions?.length} ball</p>
                      </div>
                    ):expired?(
                      <span style={{color:C.danger,fontWeight:700,fontSize:13}}>⛔ Vaqti o'tdi</span>
                    ):(
                      <button onClick={()=>startTest(test)} style={{...S.btnSmall,background:test.accessCode&&!unlockedTests.includes(test.id)?"#F59E0B":C.primary,padding:"10px 20px"}}>{test.accessCode&&!unlockedTests.includes(test.id)?"🔒 Kodni kiriting":"Boshlash →"}</button>
                    )}
                  </div>
                </div>
              );
            })}
            {tests.filter(t=>t.active&&visibleForMe(t)).length===0&&<div style={S.empty}>Hozircha faol testlar yo'q</div>}
          </div>
        )}
        {tab==="monitoring"&&(
          <div>
            <h3 style={{marginBottom:16}}>Mening Natijalarim</h3>
            {results.map(r=>{
              const test=(db.get("tests")||[]).find(t=>t.id===r.testId);
              if(!test) return null;
              const pct=Math.round((r.totalScore/test.questions.length)*100);
              const canView=test.showAnswersAfter==="immediate"||r.canViewAnswers;
              return (
                <div key={r.id} style={{...S.card,padding:18,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                    <div><h4 style={{margin:"0 0 4px",fontSize:15}}>{test.name}</h4><p style={{margin:0,color:C.textMid,fontSize:12}}>{new Date(r.id).toLocaleDateString("uz-UZ")}</p>
                      {r.rasch&&(()=>{ const dg=r.rasch.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                        return <p style={{margin:"5px 0 0",display:"inline-flex",alignItems:"center",gap:6}}><span style={{fontSize:11,color:C.textMid}}>🎯 Rash:</span><b style={{color:"#6D28D9",fontSize:13}}>{r.rasch.ball.toFixed(1)}</b><span style={{background:dc+"22",color:dc,padding:"2px 8px",borderRadius:6,fontSize:11,fontWeight:800}}>{dg}</span></p>; })()}
                    </div>
                    <div style={{textAlign:"right"}}>
                      <p style={{margin:"0 0 4px",fontWeight:800,fontSize:20,color:pct>=70?C.successDark:pct>=50?C.warning:C.danger}}>{r.totalScore}<span style={{color:C.textMid,fontWeight:400,fontSize:14}}>/{test.questions.length}</span></p>
                      {canView?<button onClick={()=>setViewResult(r)} style={{...S.btnSmall,background:C.primary,padding:"6px 14px",fontSize:12}}>Xatolarni Ko'rish</button>:<span style={{color:C.warning,fontSize:12}}>⏳ Keyinroq</span>}
                    </div>
                  </div>
                  <div style={{marginTop:10,background:"#E2E8F0",borderRadius:999,height:8,overflow:"hidden"}}>
                    <div style={{width:`${pct}%`,height:"100%",background:pct>=70?C.success:pct>=50?C.warning:C.danger,borderRadius:999,transition:"width 1s"}}/>
                  </div>
                </div>
              );
            })}
            {results.length===0&&<div style={S.empty}>Hali birorta test topshirilmagan</div>}
          </div>
        )}

        {tab==="pdfs"&&(
          <div>
            <h3 style={{marginBottom:16}}>📄 Test Ko'rish</h3>
            <p style={{color:C.textMid,fontSize:13,marginBottom:16}}>Faol testlarning variantlarini (PDF yoki LaTeX) bu yerda ko'rishingiz mumkin.</p>
            {tests.filter(t=>t.active && testHasAnyDoc(t)).map(test=>(
              <div key={test.id} style={{...S.card,padding:18,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <div style={{display:"flex",alignItems:"center",gap:12}}>
                  <div style={{width:44,height:44,borderRadius:10,background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>📄</div>
                  <div>
                    <h4 style={{margin:"0 0 2px",fontSize:15}}>{test.name}</h4>
                    <p style={{margin:0,color:C.textMid,fontSize:12}}>{test.questions?.length} savol • {availableDocLangs(test).map(l=>l.label).join(" / ")}</p>
                  </div>
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><DocLangButtons test={test} onOpen={setDocModal}/></div>
              </div>
            ))}
            {tests.filter(t=>t.active && testHasAnyDoc(t)).length===0&&<div style={S.empty}>Hozircha test variantlari mavjud emas</div>}
          </div>
        )}

        {tab==="stats"&&(()=>{
          // Only include tests where admin enabled statistics
          const allTests = tests.filter(t => t.showStats !== false);
          if (allTests.length === 0) return (
            <div style={{textAlign:"center",padding:"48px 20px"}}>
              <p style={{fontSize:40,margin:"0 0 12px"}}>📊</p>
              <h3 style={{margin:"0 0 8px",color:C.text}}>Statistika mavjud emas</h3>
              <p style={{color:C.textMid,fontSize:14}}>Hozircha statistika ko'rishga ruxsat berilgan test yo'q.</p>
            </div>
          );
          const myResults = results;

          // Score color based on percentage
          const scoreColor = (score, total) => {
            if (!total) return "#94A3B8";
            const pct = score / total * 100;
            if (pct >= 85) return "#16A34A";
            if (pct >= 70) return "#2563EB";
            if (pct >= 50) return "#D97706";
            return "#DC2626";
          };
          const scoreBg = (score, total) => {
            if (!total) return "#F1F5F9";
            const pct = score / total * 100;
            if (pct >= 85) return "#DCFCE7";
            if (pct >= 70) return "#DBEAFE";
            if (pct >= 50) return "#FEF3C7";
            return "#FEE2E2";
          };

          const attempted = allTests.filter(t => myResults.find(r => r.testId === t.id));
          const missed = allTests.filter(t => !myResults.find(r => r.testId === t.id));
          const totalScore = myResults.reduce((s,r) => s + (r.totalScore||0), 0);
          const totalMax = myResults.reduce((s,r) => {
            const t = allTests.find(x => x.id === r.testId);
            return s + (t?.questions?.length||0);
          }, 0);
          const avgPct = totalMax > 0 ? Math.round(totalScore / totalMax * 100) : 0;

          return (
            <div>
              {/* Summary cards */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:18}}>
                {[
                  {label:"Jami testlar",value:allTests.length,icon:"📋",bg:"#EEF1FF",fg:"#4338CA"},
                  {label:"Topshirilgan",value:attempted.length,icon:"✅",bg:"#DCFCE7",fg:"#16A34A"},
                  {label:"O'tkazilgan",value:missed.length,icon:"❌",bg:"#FEE2E2",fg:"#DC2626"},
                  {label:"O'rtacha ball",value:avgPct+"%",icon:"📊",bg:"#FEF3C7",fg:"#D97706"},
                ].map((c,i)=>(
                  <div key={i} style={{background:c.bg,borderRadius:14,padding:"14px 12px",textAlign:"center"}}>
                    <div style={{fontSize:26,marginBottom:4}}>{c.icon}</div>
                    <div style={{fontSize:22,fontWeight:900,color:c.fg}}>{c.value}</div>
                    <div style={{fontSize:12,color:c.fg,fontWeight:600,opacity:0.8}}>{c.label}</div>
                  </div>
                ))}
              </div>

              {/* Bar chart — score per test */}
              {attempted.length > 0 && (
                <div style={{...S.card,padding:16,marginBottom:14}}>
                  <h4 style={{margin:"0 0 12px",fontSize:14,fontWeight:700}}>📊 Test natijalari (ball foizi)</h4>
                  {attempted.map(t => {
                    const r = myResults.find(x => x.testId === t.id);
                    const total = t.questions?.length || 1;
                    const score = r?.totalScore || 0;
                    const pct = Math.round(score / total * 100);
                    const clr = scoreColor(score, total);
                    return (
                      <div key={t.id} style={{marginBottom:10}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                          <span style={{fontSize:12,color:C.text,fontWeight:600,maxWidth:"75%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.name}</span>
                          <span style={{fontSize:12,fontWeight:800,color:clr}}>{score}/{total} ({pct}%)</span>
                        </div>
                        <div style={{height:10,background:"#F1F5F9",borderRadius:6,overflow:"hidden"}}>
                          <div style={{height:"100%",width:pct+"%",background:clr,borderRadius:6,transition:"width 0.6s ease"}}/>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Full results table */}
              <div style={{...S.card,padding:0,marginBottom:14,overflow:"hidden"}}>
                <div style={{padding:"12px 16px",borderBottom:`1px solid ${C.border}`}}>
                  <h4 style={{margin:0,fontSize:14,fontWeight:700}}>📋 Barcha testlar jadvali</h4>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                    <thead>
                      <tr style={{background:"#F8FAFF"}}>
                        {["#","Test nomi","Sana","Ball","Natija","Rash","Holat"].map(h=>(
                          <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,color:C.textMid,fontSize:12,whiteSpace:"nowrap",borderBottom:`1px solid ${C.border}`}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {allTests.map((t,i) => {
                        const r = myResults.find(x => x.testId === t.id);
                        const total = t.questions?.length || 0;
                        const score = r?.totalScore ?? null;
                        const pct = score !== null && total ? Math.round(score/total*100) : null;
                        const date = r ? new Date(r.id).toLocaleDateString("uz-UZ") : "—";
                        const rowBg = i%2===0 ? "#FFFFFF" : "#FAFBFF";
                        const missed2 = score === null;
                        return (
                          <tr key={t.id} style={{background: missed2 ? "#F8F9FA" : rowBg, opacity: missed2 ? 0.65 : 1}}>
                            <td style={{padding:"10px 12px",color:C.textMid,fontWeight:700}}>{i+1}</td>
                            <td style={{padding:"10px 12px",fontWeight:600,color: missed2 ? C.textMid : C.text}}>{t.name}</td>
                            <td style={{padding:"10px 12px",color:C.textMid,whiteSpace:"nowrap"}}>{date}</td>
                            <td style={{padding:"10px 12px",fontWeight:700,color: missed2 ? C.textLight : scoreColor(score,total)}}>
                              {missed2 ? "—" : `${score}/${total}`}
                            </td>
                            <td style={{padding:"10px 12px"}}>
                              {r?.rasch
                                ? (()=>{ const dg=r.rasch.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                                  return <span style={{display:"inline-flex",alignItems:"center",gap:5}}><b style={{color:"#6D28D9",fontSize:13}}>{r.rasch.ball.toFixed(1)}</b><span style={{background:dc+"22",color:dc,padding:"2px 7px",borderRadius:6,fontSize:11,fontWeight:800}}>{dg}</span></span>; })()
                                : <span style={{color:C.textLight,fontSize:12}}>—</span>
                              }
                            </td>
                            <td style={{padding:"10px 12px"}}>
                              {missed2
                                ? <span style={{background:"#F1F5F9",color:C.textMid,padding:"3px 8px",borderRadius:6,fontSize:12,fontWeight:600}}>Topshirilmagan</span>
                                : <span style={{background:scoreBg(score,total),color:scoreColor(score,total),padding:"3px 8px",borderRadius:6,fontSize:12,fontWeight:700}}>{pct}%</span>
                              }
                            </td>
                            <td style={{padding:"10px 12px"}}>
                              {missed2
                                ? <span style={{fontSize:16}}>⚪</span>
                                : pct>=85 ? <span title="A'lo">🏆</span>
                                : pct>=70 ? <span title="Yaxshi">🟢</span>
                                : pct>=50 ? <span title="Qoniqarli">🟡</span>
                                : <span title="Qoniqarsiz">🔴</span>
                              }
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {allTests.length===0&&<div style={{...S.empty,padding:24}}>Hali testlar mavjud emas</div>}
                </div>
              </div>

              {/* Pie chart — donut style (CSS) */}
              {attempted.length > 0 && (
                <div style={{...S.card,padding:16,marginBottom:14}}>
                  <h4 style={{margin:"0 0 16px",fontSize:14,fontWeight:700}}>🎯 Natijalar taqsimoti</h4>
                  <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
                    {/* Donut chart via conic-gradient */}
                    {(()=>{
                      const alo = attempted.filter(t=>{const r=myResults.find(x=>x.testId===t.id);const tot=t.questions?.length||1;return r&&(r.totalScore/tot*100)>=85;}).length;
                      const yax = attempted.filter(t=>{const r=myResults.find(x=>x.testId===t.id);const tot=t.questions?.length||1;const p=r?.totalScore/tot*100||0;return p>=70&&p<85;}).length;
                      const qon = attempted.filter(t=>{const r=myResults.find(x=>x.testId===t.id);const tot=t.questions?.length||1;const p=r?.totalScore/tot*100||0;return p>=50&&p<70;}).length;
                      const yom = attempted.filter(t=>{const r=myResults.find(x=>x.testId===t.id);const tot=t.questions?.length||1;const p=r?.totalScore/tot*100||0;return p<50;}).length;
                      const tot = attempted.length || 1;
                      const aloD = alo/tot*360, yaxD = yax/tot*360, qonD = qon/tot*360, yomD = yom/tot*360;
                      const grad = `conic-gradient(#16A34A 0deg ${aloD}deg, #2563EB ${aloD}deg ${aloD+yaxD}deg, #D97706 ${aloD+yaxD}deg ${aloD+yaxD+qonD}deg, #DC2626 ${aloD+yaxD+qonD}deg 360deg)`;
                      return (
                        <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap",width:"100%"}}>
                          <div style={{width:100,height:100,borderRadius:"50%",background:grad,flexShrink:0,boxShadow:"0 2px 12px rgba(0,0,0,0.1)",position:"relative"}}>
                            <div style={{position:"absolute",inset:"20%",borderRadius:"50%",background:"white",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column"}}>
                              <span style={{fontSize:16,fontWeight:900,color:C.text}}>{attempted.length}</span>
                              <span style={{fontSize:9,color:C.textMid}}>test</span>
                            </div>
                          </div>
                          <div style={{flex:1,minWidth:150}}>
                            {[["🏆 A'lo (85%+)",alo,"#16A34A"],["🟢 Yaxshi (70-85%)",yax,"#2563EB"],["🟡 Qoniqarli (50-70%)",qon,"#D97706"],["🔴 Yomon (<50%)",yom,"#DC2626"]].map(([lbl,cnt,clr])=>(
                              <div key={lbl} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                                <span style={{fontSize:12,color:C.text}}>{lbl}</span>
                                <span style={{fontSize:13,fontWeight:800,color:clr,minWidth:24,textAlign:"right"}}>{cnt}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                </div>
              )}

              {/* Score trend — bar per test in chronological order */}
              {attempted.length > 1 && (
                <div style={{...S.card,padding:16,marginBottom:14}}>
                  <h4 style={{margin:"0 0 12px",fontSize:14,fontWeight:700}}>📈 Natijalar dinamikasi</h4>
                  <div style={{display:"flex",alignItems:"flex-end",gap:6,height:80,padding:"0 4px"}}>
                    {[...myResults].sort((a,b)=>a.id-b.id).map((r,i)=>{
                      const t=allTests.find(x=>x.id===r.testId);
                      const total=t?.questions?.length||1;
                      const pct=Math.round((r.totalScore||0)/total*100);
                      const clr=scoreColor(r.totalScore,total);
                      const h=Math.max(8,pct*0.7);
                      return (
                        <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                          <span style={{fontSize:9,color:clr,fontWeight:700}}>{pct}%</span>
                          <div style={{width:"100%",height:h,background:clr,borderRadius:"4px 4px 0 0",transition:"height 0.5s ease",minHeight:4}}/>
                          <span style={{fontSize:8,color:C.textLight,maxWidth:30,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textAlign:"center"}}>{t?.name||""}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          );
        })()}
      </div>
    </div>
  );
}

// ===== TEST TAKING =====
function TestTaking({ test, user, onFinish, onExit }) {
  const calcRem = useCallback(() => {
    if(test.startedAt) return Math.max(0,Math.floor((test.startedAt+test.duration*60000-Date.now())/1000));
    return test.duration*60;
  },[test]);

  const progressKey = "test_progress_"+test.id+"_"+user.phone;
  const loadProgress = () => {
    try { return JSON.parse(localStorage.getItem(progressKey)) || {}; } catch { return {}; }
  };

  const [timeLeft,setTimeLeft]=useState(calcRem);
  const [answers,setAnswers]=useState(()=>loadProgress().answers||{});
  const [openAns,setOpenAns]=useState(()=>loadProgress().openAns||{});
  const [subAns,setSubAns]=useState(()=>loadProgress().subAns||{});
  const [kbd,setKbd]=useState(null);
  const [showConfirm,setShowConfirm]=useState(false);
  const [pdfViewOpen,setPdfViewOpen]=useState(false);
  const testDocLangs = useMemo(()=>availableDocLangs(test),[test]);
  const [pdfViewLang,setPdfViewLang]=useState(testDocLangs[0]?.code||"uz");
  const activeLangDoc = getLangDoc(test, pdfViewLang) || getLangDoc(test, testDocLangs[0]?.code);
  const startedAt=useRef(Date.now());
  const didSubmit=useRef(false);

  const vibrated10 = useRef(false);
  const [confirmModal,setConfirmModal]=useState(null);

  // O'quvchi javoblarini har o'zgarishda saqlab boradi — sahifa yopilib qayta
  // ochilsa (ilova yopilsa, brauzer yangilansa va h.k.) javoblar yo'qolmaydi.
  useEffect(()=>{
    try { localStorage.setItem(progressKey, JSON.stringify({answers,openAns,subAns})); } catch {}
  },[answers,openAns,subAns]);

  // Savollar ro'yxatida qayerda to'xtagan bo'lsa, qayta ochilganda o'sha yerdan davom etadi
  const windowScrollKey = "test_scroll_"+test.id+"_"+user.phone;
  useEffect(()=>{
    const saved = sessionStorage.getItem(windowScrollKey);
    if (saved) requestAnimationFrame(()=>window.scrollTo(0,+saved));
    let ticking=null;
    const onScroll=()=>{
      if(ticking) return;
      ticking = requestAnimationFrame(()=>{ sessionStorage.setItem(windowScrollKey, window.scrollY); ticking=null; });
    };
    window.addEventListener("scroll",onScroll,{passive:true});
    return ()=>{ window.removeEventListener("scroll",onScroll); if(ticking) cancelAnimationFrame(ticking); };
  },[]);

  useEffect(()=>{
    const t=setInterval(()=>{
      const rem=calcRem();
      setTimeLeft(rem);
      // 10-minute warning vibration (5x)
      if(rem<=600 && rem>595 && !vibrated10.current){
        vibrated10.current = true;
        if(navigator.vibrate){
          navigator.vibrate([200,150,200,150,200,150,200,150,200]);
        }
      }
      if(rem<=0&&!didSubmit.current){didSubmit.current=true;submit({});}
    },1000);
    return ()=>clearInterval(t);
  },[]);

  // Submit uses current state via ref pattern
  const answersRef=useRef(answers); answersRef.current=answers;
  const openRef=useRef(openAns); openRef.current=openAns;
  const subRef=useRef(subAns); subRef.current=subAns;

  const [grading, setGrading] = useState(false);

  const submit=useCallback(async (overrides={})=>{
    setGrading(true);
    const ans={...answersRef.current,...(overrides.answers||{})};
    const oa={...openRef.current,...(overrides.openAns||{})};
    const sa={...subRef.current,...(overrides.subAns||{})};
    const scores={},subScores={};
    let total=0;
    for (let idx=0; idx<test.questions.length; idx++) {
      const q = test.questions[idx];
      if(q.type==="closed"){
        const ok=ans[idx]!==undefined&&ans[idx]===q.correctAnswer;
        scores[idx]=ok; if(ok)total++;
      } else if(q.subParts?.length>0){
        subScores[idx]={};
        for (let si=0; si<q.subParts.length; si++) {
          const sp = q.subParts[si];
          const ok = await checkMathAsync(sp.answer, sa[idx]?.[si]||"");
          subScores[idx][si]=ok; if(ok)total++;
        }
      } else {
        const ok = await checkMathAsync(q.correctAnswer, oa[idx]||"");
        scores[idx]=ok; if(ok)total++;
      }
    }
    const res={id:Date.now(),testId:test.id,userPhone:user.phone,answers:ans,openAnswers:oa,subAnswers:sa,scores,subScores,totalScore:total,canViewAnswers:test.showAnswersAfter==="immediate",timeTaken:Math.round((Date.now()-startedAt.current)/60000)};
    const all=db.get("results")||[]; all.push(res); db.set("results",all);
    try { localStorage.removeItem(progressKey); sessionStorage.removeItem(windowScrollKey); } catch {}
    setGrading(false);
    onFinish();
  },[test,user,onFinish]);

  const fmt=(s)=>`${String(Math.floor(s/60)).padStart(2,"0")}:${String(s%60).padStart(2,"0")}`;
  const tColor=timeLeft<60?C.danger:timeLeft<300?C.warning:C.successDark;
  const closedQs=test.questions.filter(q=>q.type==="closed");
  const openQs=test.questions.filter(q=>q.type==="open");
  const answered=Object.keys(answers).length;
  const total=test.questions.length;

  const kbdVal=kbd?(kbd.sub!==null?(subAns[kbd.qIdx]?.[kbd.sub]||""):(openAns[kbd.qIdx]||"")):"";
  const setKbdVal=(v)=>{
    if(!kbd) return;
    if(kbd.sub!==null) setSubAns(p=>({...p,[kbd.qIdx]:{...(p[kbd.qIdx]||{}),[kbd.sub]:v}}));
    else setOpenAns(p=>({...p,[kbd.qIdx]:v}));
  };

  return (
    <div style={{...S.page, paddingBottom: kbd ? 340 : 0, transition:'padding 0.3s'}}>
      {confirmModal && <ConfirmModal message={confirmModal.message} confirmLabel={confirmModal.confirmLabel} onConfirm={confirmModal.onConfirm} onCancel={()=>setConfirmModal(null)}/>}
      {/* Confirm modal */}
      {showConfirm&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{...S.card,padding:28,maxWidth:380,width:"100%"}}>
            <h3 style={{margin:"0 0 12px",fontSize:18}}>✅ Testni yakunlash</h3>
            <p style={{color:C.textMid,margin:"0 0 6px",fontSize:14}}>Javob berilgan: <b style={{color:C.primary}}>{answered}</b> / {closedQs.length} (yopiq)</p>
            <p style={{color:C.textMid,margin:"0 0 16px",fontSize:14}}>Belgilanmagan savollar uchun avtomatik <b style={{color:C.danger}}>0 ball</b> beriladi.</p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{didSubmit.current=true;submit();}} style={{...S.btnSuccess,flex:1}}>Tasdiqlash ✓</button>
              <button onClick={()=>setShowConfirm(false)} style={{...S.btnGhost,flex:1}}>Bekor qilish</button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{position:"sticky",top:0,zIndex:100,background:C.card,borderBottom:`1px solid ${C.border}`,padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 2px 8px rgba(0,0,0,0.08)"}}>
        <div>
          <p style={{margin:0,fontWeight:700,fontSize:15}}>{test.name}</p>
          <p style={{margin:0,color:C.textMid,fontSize:12}}>{answered}/{closedQs.length} belgilangan</p>
        </div>
        <div style={{textAlign:"center"}}>
          <p style={{margin:0,fontSize:28,fontWeight:900,color:tColor,fontFamily:"monospace",lineHeight:1}}>{fmt(timeLeft)}</p>
          <p style={{margin:0,fontSize:10,color:C.textLight}}>Qolgan vaqt</p>
        </div>
        <div style={{display:"flex",gap:8}}>
          {testDocLangs.length>0&&<button onClick={()=>setPdfViewOpen(true)} style={{...S.btnSmall,background:"#F59E0B",padding:"9px 12px",fontSize:12,border:"none",cursor:"pointer"}}>{activeLangDoc?.pdfUrl?"📄":"∑"} Test ko'rish</button>}
          <button onClick={()=>setShowConfirm(true)} disabled={grading} style={{...S.btnSuccess,padding:"9px 16px",fontSize:13,opacity:grading?0.6:1}}>{grading?"⏳ Tekshirilmoqda...":"✅ Yakunlash"}</button>
          <button onClick={()=>setConfirmModal({message:"Testdan chiqasizmi? Belgilagan javoblaringiz saqlanadi, keyinroq davom ettirishingiz mumkin.", confirmLabel:"Chiqish", onConfirm:()=>{setConfirmModal(null);onExit();}})} style={{...S.btnDanger,padding:"9px 10px",fontSize:13}}>✕</button>
        </div>
      </div>

      {/* In-test PDF viewer — stays within TestTaking, timer keeps running,
          answers preserved. User can switch back and forth freely, and can
          switch between the languages the test was uploaded in (UZ/QQ/RU). */}
      {pdfViewOpen&&testDocLangs.length>0&&activeLangDoc&&(
        <div style={{position:"fixed",inset:0,background:activeLangDoc.pdfUrl?"#1a1a1a":"white",zIndex:9000,display:"flex",flexDirection:"column"}}>
          <div style={{background:"#F59E0B",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
            <div>
              <span style={{fontWeight:800,fontSize:15,color:"white"}}>{activeLangDoc.pdfUrl?"📄":"∑"} {test.name}</span>
              <p style={{margin:0,fontSize:11,color:"rgba(255,255,255,0.85)"}}>Vaqt davom etmoqda: {fmt(timeLeft)} • Javoblaringiz saqlanadi</p>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {testDocLangs.length>1&&(
                <div style={{display:"flex",gap:4,background:"rgba(0,0,0,0.15)",padding:4,borderRadius:9}}>
                  {testDocLangs.map(l=>(
                    <button key={l.code} onClick={()=>setPdfViewLang(l.code)} style={{
                      padding:"6px 10px",borderRadius:6,border:"none",cursor:"pointer",fontSize:12,fontWeight:800,
                      background:pdfViewLang===l.code?"white":"transparent",
                      color:pdfViewLang===l.code?"#92400E":"white"
                    }}><LangFlag lang={l} size={13}/> {l.label}</button>
                  ))}
                </div>
              )}
              <button onClick={()=>setPdfViewOpen(false)} style={{...S.btnSuccess,padding:"9px 16px",fontSize:13,background:"white",color:"#92400E",fontWeight:800}}>
                ✏️ Javob berishga qaytish
              </button>
            </div>
          </div>
          {activeLangDoc.pdfUrl
            ? <PdfViewer url={activeLangDoc.pdfUrl} persistKey={"pdf_scroll_"+test.id+"_"+pdfViewLang}/>
            : <ScrollPersistDiv persistKey={"taking_latex_scroll_"+test.id+"_"+pdfViewLang} style={{flex:1,overflowY:"auto"}}><LatexDocViewer source={activeLangDoc.latexSource} images={activeLangDoc.latexImages}/></ScrollPersistDiv>
          }
        </div>
      )}

      <div style={{padding:"16px",maxWidth:760,margin:"0 auto"}}>
        {closedQs.length>0&&(
          <div style={{...S.card,padding:18,marginBottom:16}}>
            <h3 style={{margin:"0 0 14px",color:C.primary,fontSize:15}}>🔵 I-Qism: Test Savollar (1–{closedQs.length})</h3>
            <table style={{width:"100%",borderCollapse:"separate",borderSpacing:"0 2px"}}>
              <tbody>
                {closedQs.map((q,ri)=>{
                    // Use actual question index in test.questions for answers key
                    const actualIdx = test.questions.indexOf(q);
                    const key = actualIdx >= 0 ? actualIdx : ri;
                    return (
                  <tr key={ri} style={{background:ri%2===0?C.card:C.bg}}>
                    <td style={{width:32,color:C.textMid,fontWeight:700,fontSize:13,paddingRight:8,paddingLeft:4,textAlign:"right",whiteSpace:"nowrap"}}>{ri+1}</td>
                    {Array.from({length:q.optionsCount},(_,oi)=>String.fromCharCode(65+oi)).map(opt=>{
                      const sel=answers[key]===opt;
                      return (
                        <td key={opt} style={{padding:"3px 3px",textAlign:"center"}}>
                          <button onClick={()=>setAnswers(p=>{
                            const cur=p[key];
                            // Toggle: click same = deselect, click different = select
                            return {...p,[key]: cur===opt ? undefined : opt};
                          })} style={{width:36,height:36,borderRadius:"50%",border:`2px solid ${sel?C.success:C.border}`,background:sel?C.success:C.card,color:sel?"white":C.textMid,cursor:"pointer",fontWeight:700,fontSize:13,transition:"all 0.15s",outline:"none"}}>{opt}</button>
                        </td>
                      );
                    })}
                  </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        )}

        {openQs.length>0&&(
          <div style={{...S.card,padding:18}}>
            <h3 style={{margin:"0 0 14px",color:C.successDark,fontSize:15}}>📝 II-Qism: Ochiq Savollar ({closedQs.length+1}–{total})</h3>
            {openQs.map((q,ri)=>{
              const idx=closedQs.length+ri;
              return (
                <div key={idx} style={{marginBottom:14,padding:14,background:"#F8F9FF",borderRadius:12,border:`1px solid ${C.border}`}}>
                  <p style={{margin:"0 0 10px",fontWeight:700,color:C.warning,fontSize:14}}>Savol {idx+1}</p>
                  {q.subParts?.length>0?q.subParts.map((sp,si)=>{
                    const val=subAns[idx]?.[si]||"";
                    const isA=kbd?.qIdx===idx&&kbd?.sub===si;
                    return (
                      <div key={si} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                        <span style={{color:C.warning,minWidth:44,fontSize:14,fontWeight:700}}>{idx+1}{sp.label})</span>
                        <MathInputField
                          value={val}
                          active={isA}
                          onFocus={()=>setKbd({qIdx:idx,sub:si})}
                          placeholder="Javob yozish uchun bosing..."
                          style={{flex:1}}
                        />
                      </div>
                    );
                  }):(()=>{
                    const val=openAns[idx]||"";
                    const isA=kbd?.qIdx===idx&&kbd?.sub===null;
                    return (
                      <MathInputField
                        value={val}
                        active={isA}
                        onFocus={()=>setKbd({qIdx:idx,sub:null})}
                        placeholder="Javob yozish uchun bosing..."
                      />
                    );
                  })()}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Fixed bottom keyboard */}
      {kbd && (
        <div style={{
          position:"fixed", bottom:0, left:0, right:0, zIndex:9000,
          transition: "transform 0.3s cubic-bezier(0.32,0.72,0,1)",
        }}>
          <MathKeyboard
            key={kbd.qIdx + "_" + (kbd.sub ?? "x")}
            initValue={kbdVal}
            onChange={setKbdVal}
            onClose={()=>setKbd(null)}
          />
        </div>
      )}
    </div>
  );
}

// ===== RESULT DETAIL =====
function ResultDetail({ result, test, onBack }) {
  if(!test) return <div style={{padding:24}}><button onClick={onBack}>← Orqaga</button> Test topilmadi.</div>;
  const pct=Math.round((result.totalScore/test.questions.length)*100);
  const closedQs=test.questions.filter(q=>q.type==="closed");
  return (
    <div style={S.page}>
      <div style={{background:C.primary,padding:"14px 20px",display:"flex",alignItems:"center",gap:14,boxShadow:"0 2px 12px rgba(79,110,247,0.3)"}}>
        <button onClick={onBack} style={{...S.btnSmall,background:"rgba(255,255,255,0.2)",color:"white"}}>← Orqaga</button>
        <h2 style={{margin:0,color:"white",fontSize:17}}>Xatolar Tahlili: {test.name}</h2>
      </div>
      <div style={{padding:20,maxWidth:800,margin:"0 auto"}}>
        <div style={{...S.card,padding:24,textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:72,fontWeight:900,color:pct>=70?C.successDark:pct>=50?C.warning:C.danger,lineHeight:1}}>{result.totalScore}</div>
          <div style={{color:C.textMid,fontSize:18,margin:"4px 0 12px"}}>/ {test.questions.length} ball ({pct}%)</div>
          <div style={{background:C.bg,borderRadius:999,height:14,overflow:"hidden",maxWidth:400,margin:"0 auto"}}>
            <div style={{width:`${pct}%`,height:"100%",background:pct>=70?C.success:pct>=50?C.warning:C.danger,borderRadius:999,transition:"width 1.5s"}}/>
          </div>
        </div>
        {result.rasch&&(()=>{ const dg=result.rasch.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
          return (
            <div style={{...S.card,padding:20,textAlign:"center",marginBottom:20,background:"linear-gradient(135deg,#F5F3FF,#FFFFFF)",border:"1.5px solid #DDD6FE"}}>
              <p style={{margin:"0 0 8px",color:"#6D28D9",fontSize:13,fontWeight:800}}>🎯 RASH MODELI BO'YICHA BAHOLASH</p>
              <div style={{display:"flex",justifyContent:"center",alignItems:"baseline",gap:14,flexWrap:"wrap"}}>
                <div><div style={{fontSize:40,fontWeight:900,color:"#6D28D9",lineHeight:1}}>{result.rasch.ball.toFixed(1)}</div><div style={{fontSize:11,color:C.textMid}}>Rash ball</div></div>
                <div><span style={{background:dc+"22",color:dc,padding:"6px 16px",borderRadius:10,fontSize:20,fontWeight:900}}>{dg}</span><div style={{fontSize:11,color:C.textMid,marginTop:4}}>Daraja</div></div>
              </div>
              <p style={{margin:"10px 0 0",fontSize:11,color:C.textLight}}>Sinf/guruh o'rtachasiga nisbatan hisoblangan (logistik model). Hisoblangan sana: {new Date(result.rasch.calculatedAt).toLocaleDateString("uz-UZ")}</p>
            </div>
          );
        })()}
        {closedQs.length>0&&(
          <div style={{...S.card,padding:18,marginBottom:16}}>
            <h3 style={{margin:"0 0 14px",color:C.primary}}>Yopiq Savollar</h3>
            <div style={{overflowX:"auto"}}>
              <table style={S.table}>
                <thead><tr>{["#","Sizning javob","To'g'ri javob","Natija"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>{closedQs.map((q,i)=>{
                  const ok=result.scores?.[i];
                  return (
                    <tr key={i} style={{background:ok?C.successLight:i%2===0?C.card:C.dangerLight}}>
                      <td style={S.td}>{i+1}</td>
                      <td style={{...S.td,fontWeight:700,color:result.answers?.[i]?(ok?C.successDark:C.danger):C.textLight}}>{result.answers?.[i]||"—"}</td>
                      <td style={{...S.td,fontWeight:700,color:C.successDark}}>{q.correctAnswer}</td>
                      <td style={S.td}>{ok?"✅":"❌"}</td>
                    </tr>
                  );
                })}</tbody>
              </table>
            </div>
          </div>
        )}
        {test.questions.filter(q=>q.type==="open").map((q,ri)=>{
          const idx=closedQs.length+ri;
          return (
            <div key={idx} style={{...S.card,padding:16,marginBottom:12}}>
              <h4 style={{color:C.warning,margin:"0 0 10px"}}>Savol {idx+1}</h4>
              {q.subParts?.length>0?q.subParts.map((sp,si)=>{
                const ok=result.subScores?.[idx]?.[si];
                const stu=result.subAnswers?.[idx]?.[si];
                return (
                  <div key={si} style={{display:"flex",gap:8,alignItems:"flex-start",flexWrap:"wrap",padding:"10px",borderRadius:8,background:ok?C.successLight:C.dangerLight,marginBottom:6}}>
                    <span style={{color:C.warning,fontWeight:700,minWidth:40}}>{idx+1}{sp.label})</span>
                    <span style={{color:ok?C.successDark:C.danger,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}>{stu?<KatexSpan latex={toLatex(stu)} fontSize={18}/>:<span style={{color:C.textLight}}>—</span>}</span>
                    <span style={{color:C.textMid}}>→</span>
                    <span style={{display:"inline-flex",flexWrap:"wrap",gap:2,alignItems:"center",color:C.successDark,fontWeight:600,fontSize:18}}><KatexSpan latex={toLatex(sp.answer||"")} fontSize={18}/></span>
                    <span style={{marginLeft:"auto"}}>{ok?"✅":"❌"}</span>
                  </div>
                );
              }):(()=>{
                const ok=result.scores?.[idx];
                const stu=result.openAnswers?.[idx];
                return (
                  <div style={{display:"flex",gap:8,alignItems:"flex-start",flexWrap:"wrap",padding:"10px",borderRadius:8,background:ok?C.successLight:C.dangerLight}}>
                    <span style={{color:ok?C.successDark:C.danger,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}>{stu?<KatexSpan latex={toLatex(stu)} fontSize={18}/>:<span style={{color:C.textLight,fontSize:14}}>Javob berilmadi</span>}</span>
                    <span style={{color:C.textMid}}>→</span>
                    <span style={{color:C.successDark,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}><KatexSpan latex={toLatex(q.correctAnswer||"")} fontSize={18}/></span>
                    <span style={{marginLeft:"auto"}}>{ok?"✅":"❌"}</span>
                  </div>
                );
              })()}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ===== MAIN =====
// Splash screen component
function SplashScreen({ onDone }) {
  const [phase, setPhase] = useState(0); // 0=show, 1=fade

  useEffect(() => {
    const t1 = setTimeout(() => setPhase(1), 1800);
    const t2 = setTimeout(() => onDone(), 2400);
    return () => { clearTimeout(t1); clearTimeout(t2); };
  }, []);

  return (
    <div style={{
      position:"fixed", inset:0, zIndex:99999,
      background:"linear-gradient(135deg,#4F46E5 0%,#7C3AED 50%,#2563EB 100%)",
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      opacity: phase===1 ? 0 : 1,
      transition:"opacity 0.6s ease",
      pointerEvents: phase===1 ? "none" : "all",
    }}>
      {/* Animated circles background */}
      <div style={{position:"absolute",inset:0,overflow:"hidden"}}>
        {[...Array(6)].map((_,i) => (
          <div key={i} style={{
            position:"absolute",
            width: 80+i*60, height: 80+i*60,
            borderRadius:"50%",
            border:"1px solid rgba(255,255,255,0.15)",
            top:"50%", left:"50%",
            transform:"translate(-50%,-50%)",
            animation:`pulse-ring ${1.5+i*0.3}s ease-out infinite`,
          }}/>
        ))}
      </div>
      {/* Logo */}
      <div style={{
        width:90, height:90, borderRadius:24,
        background:"rgba(255,255,255,0.15)",
        backdropFilter:"blur(10px)",
        display:"flex", alignItems:"center", justifyContent:"center",
        fontSize:44, marginBottom:20,
        boxShadow:"0 8px 32px rgba(0,0,0,0.2)",
        animation:"logo-pop 0.6s cubic-bezier(0.34,1.56,0.64,1) 0.3s both",
      }}>📐</div>
      <h1 style={{
        color:"white", fontSize:26, fontWeight:900, margin:"0 0 8px",
        letterSpacing:"-0.5px",
        animation:"slide-up 0.5s ease 0.5s both",
      }}>Matematika Testi</h1>
      <p style={{
        color:"rgba(255,255,255,0.75)", fontSize:15, margin:0,
        animation:"slide-up 0.5s ease 0.7s both",
      }}>Bilim platformasi</p>
      {/* Loading dots */}
      <div style={{display:"flex",gap:8,marginTop:32,animation:"slide-up 0.5s ease 0.9s both"}}>
        {[0,1,2].map(i=>(
          <div key={i} style={{
            width:8,height:8,borderRadius:"50%",
            background:"rgba(255,255,255,0.6)",
            animation:`dot-bounce 1.2s ease ${i*0.2}s infinite`,
          }}/>
        ))}
      </div>
      <style>{`
        @keyframes pulse-ring { 0%{transform:translate(-50%,-50%) scale(0.8);opacity:0.6} 100%{transform:translate(-50%,-50%) scale(1.4);opacity:0} }
        @keyframes logo-pop { from{transform:scale(0);opacity:0} to{transform:scale(1);opacity:1} }
        @keyframes slide-up { from{transform:translateY(20px);opacity:0} to{transform:translateY(0);opacity:1} }
        @keyframes dot-bounce { 0%,80%,100%{transform:translateY(0)} 40%{transform:translateY(-10px)} }
      `}</style>
    </div>
  );
}

export default function App() {
  // Saqlangan sessiya bo'lsa (avval login qilingan bo'lsa), avtomatik tiklaymiz —
  // login/parol qayta so'ralmaydi. Foydalanuvchi ma'lumoti bazadan yangilanib olinadi
  // (masalan boshqa qurilmada tahrirlangan bo'lishi mumkin).
  const initFromSession = () => {
    const s = loadSession();
    if (!s) return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null };
    if (s.role === "admin") return { page: "admin", user: null, isAdmin: true, isTeacher: false, teacherInfo: null };
    if (s.role === "teacher") {
      const t = (db.get("teachers")||[]).find(x => x.id === s.teacherId);
      if (!t) { clearSession(); return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null }; }
      return { page: "teacher", user: null, isAdmin: false, isTeacher: true, teacherInfo: t };
    }
    if (s.role === "student") {
      const u = (db.get("users")||[]).find(x => x.phone === s.phone);
      if (!u) { clearSession(); return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null }; }
      return { page: "student", user: u, isAdmin: false, isTeacher: false, teacherInfo: null };
    }
    return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null };
  };
  const initState = initFromSession();
  const [page,setPage]=useState(initState.page);
  const [user,setUser]=useState(initState.user);
  const [isAdmin,setIsAdmin]=useState(initState.isAdmin);
  const [isTeacher,setIsTeacher]=useState(initState.isTeacher);
  const [teacherInfo,setTeacherInfo]=useState(initState.teacherInfo);
  const [showSplash,setShowSplash]=useState(true);
  useEffect(()=>{
    initDB();
    initFirebaseSync();
    // Load math.js for answer checking
    loadMathJs(() => {});
    // Gorizontal siljish/tirnash (masalan A4 hujjat ko'rinishida) tasodifan butun
    // sahifani yon tomonga surib, bo'sh joy ko'rsatib qo'ymasligi uchun bloklaymiz
    document.documentElement.style.overflowX = "hidden";
    document.body.style.overflowX = "hidden";
    // Disable zoom on mobile
    let meta = document.querySelector("meta[name=viewport]");
    if (!meta) { meta = document.createElement("meta"); meta.name = "viewport"; document.head.appendChild(meta); }
    meta.content = "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no";
    // Blink cursor style
    if (!document.getElementById("math-blink-style")) {
      const st = document.createElement("style");
      st.id = "math-blink-style";
      st.textContent = "@keyframes blink{0%,100%{opacity:1}50%{opacity:0}}@keyframes slotPulse{0%,100%{border-color:#6366F1;box-shadow:0 0 0 0 rgba(99,102,241,0.35)}50%{border-color:#A5B4FC;box-shadow:0 0 0 3px rgba(99,102,241,0.12)}}";
      document.head.appendChild(st);
    }
    // Prevent double-tap zoom
    document.documentElement.style.touchAction = "manipulation";
  },[]);

  if(showSplash) return <SplashScreen onDone={()=>setShowSplash(false)}/>;
  if(page==="admin"&&isAdmin) return <AdminPanel isFullAdmin={true} onLogout={()=>{setIsAdmin(false);setPage("login");clearSession();}}/>;
  if(page==="teacher"&&isTeacher) return <AdminPanel isFullAdmin={false} teacherInfo={teacherInfo} onLogout={()=>{setIsTeacher(false);setTeacherInfo(null);setPage("login");clearSession();}}/>;
  if(page==="student"&&user) return <StudentDashboard user={user} onLogout={()=>{setUser(null);setPage("login");clearSession();}}/>;
  if(page==="register") return <RegisterPage onDone={u=>{setUser(u);setPage("student");saveSession({role:"student",phone:u.phone});}} onLogin={()=>setPage("login")}/>;
  return <LoginPage
    onLogin={u=>{setUser(u);setPage("student");saveSession({role:"student",phone:u.phone});}}
    onRegister={()=>setPage("register")}
    onAdmin={(fullAdmin, teacher)=>{
      if(fullAdmin){ setIsAdmin(true); setPage("admin"); saveSession({role:"admin"}); }
      else { setIsTeacher(true); setTeacherInfo(teacher); setPage("teacher"); saveSession({role:"teacher",teacherId:teacher.id}); }
    }}
  />;
}
