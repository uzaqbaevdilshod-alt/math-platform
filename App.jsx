import { useState, useEffect, useRef, useCallback, useMemo, memo } from "react";
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

function KatexSpan({ latex, block }) {
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
  return <span ref={ref} style={{ fontFamily: "KaTeX_Main, serif", fontSize: block ? 20 : 16 }} />;
}

// Convert display string to LaTeX for render
function toLatex(s) {
  if (!s) return "";
  let r = s;

  // Mixed number pattern: (N+FRAC(a,b)) → N\frac{a}{b}
  let mfPrev = "", mfIter = 0;
  while (mfPrev !== r && mfIter++ < 20) {
    mfPrev = r;
    r = r.replace(/\((-?\d+\.?\d*)\+FRAC\(([^()]+),([^()]+)\)\)/g, "$1\\frac{$2}{$3}");
    r = r.replace(/\((-\d+\.?\d*)-FRAC\(([^()]+),([^()]+)\)\)/g, "$1\\frac{$2}{$3}");
  }

  // FRAC(a,b) -> \frac{a}{b}  (max 30 iterations safety)
  let prev = "", prevIter = 0;
  while (prev !== r && prevIter++ < 30) {
    prev = r;
    r = r.replace(/FRAC\(([^()]+),([^()]+)\)/g, "\\frac{$1}{$2}");
  }
  r = r.replace(/SUB\(([^,]+),([^)]+)\)/g, "$1_{$2}");
  r = r.replace(/ROOT\(([^,]+),([^)]+)\)/g, "\\sqrt[$1]{$2}");
  r = r.replace(/ABS\(([^)]+)\)/g, "\\left|$1\\right|");
  r = r.replace(/√\(([^)]+)\)/g, "\\sqrt{$1}");
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
  // LOG_BASE(base,arg) → \log_{base}(arg)
  r = r.replace(/LOG_BASE\(([^,]+),([^)]+)\)/g, "\\log_{$1}($2)");
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
      onClick={onFocus}
      style={{
        minHeight: 48, padding: "10px 16px",
        background: active ? "#F8FAFF" : "#FFFFFF",
        border: `2px solid ${active ? "#6366F1" : value ? "#22C55E" : "#E2E8F0"}`,
        borderRadius: 12, cursor: "pointer",
        display: "flex", alignItems: "center", flexWrap: "wrap",
        boxShadow: active ? "0 0 0 3px rgba(99,102,241,0.12)" : "none",
        transition: "all 0.2s",
        ...style,
      }}
    >
      {value ? (
        <span ref={ref} style={{ fontSize: 20, fontFamily: "KaTeX_Main,serif", color: active ? "#4F46E5" : "#15803D" }} />
      ) : (
        <span style={{ color: "#94A3B8", fontSize: 14 }}>{placeholder || "Javob yozish uchun bosing..."}</span>
      )}
      {active && !value && (
        <span style={{ display:"inline-block", width:2, height:22, background:"#6366F1", borderRadius:1, marginLeft:2, animation:"blink 1s step-end infinite" }} />
      )}
    </div>
  );
}

// ===== STORAGE =====
const ADMIN_LOGIN = "admin";
const ADMIN_PW = "admin123";

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

// Butun kolleksiyani Firestore'ga yozadi (id bo'yicha upsert + o'chirilganlarni tozalash).
async function pushCollectionToFirestore(col, arr) {
  if (!fbFirestore) return;
  try {
    const coll = fbFirestore.collection(col);
    const snap = await coll.get();
    const existingIds = new Set(snap.docs.map(d => d.id));
    const newIds = new Set(arr.map(item => String(item.id)));
    const batch = fbFirestore.batch();
    arr.forEach(item => { batch.set(coll.doc(String(item.id)), item); });
    existingIds.forEach(id => { if (!newIds.has(id)) batch.delete(coll.doc(id)); });
    await batch.commit();
  } catch (e) { console.error("[Firebase] Yozishda xato:", col, e); }
}

const db = {
  get: (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set: (k, v) => {
    try {
      localStorage.setItem(k, JSON.stringify(v));
      if (fbFirestore && FIREBASE_SYNC_COLLECTIONS.includes(k) && Array.isArray(v)) {
        pushCollectionToFirestore(k, v);
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
    if (s.startsWith("FRAC(", i)) {
      const op = i + 4, cl = matchParen(s, op);
      if (cl !== -1) {
        const inner = s.slice(op + 1, cl);
        const cm = topLevelComma(inner);
        const num = cm >= 0 ? inner.slice(0, cm) : inner;
        const den = cm >= 0 ? inner.slice(cm + 1) : "1";
        const lastCh0 = result.slice(-1);
        if (lastCh0 && /[\d)]/.test(lastCh0)) result += "*";
        result += "(" + processExpr(num) + ")/(" + processExpr(den) + ")";
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

// ===== EXCEL EXPORT (SheetJS .xlsx) =====
// Ma'lumotni tayyorlab, DATA URL qaytaradi (avtomatik yuklab olishga urinmaydi).
// Sabab: ko'plab embedded WebView muhitlar (masalan Telegram Mini App) dasturiy
// ravishda (.click()) boshlangan yuklab olishlarni bloklaydi, lekin foydalanuvchi
// O'ZI bosgan havolaga (haqiqiy user gesture) ruxsat beradi.
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
    if (n.type === "sup") { const b=slotVal(n.children[0]); const e=slotVal(n.children[1]); return b+"^"+e; }
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
  const bg = isActive?"rgba(99,102,241,0.13)":isEmpty?"rgba(251,191,36,0.18)":"transparent";
  const brd = isActive?"1.5px solid #6366F1":isEmpty?"1.5px solid #F59E0B":"1px solid #E2E8F0";
  return (
    <span style={{ display:"inline-flex", alignItems:"center", flexWrap:"nowrap",
      minWidth:18, minHeight:22, padding:"0 3px", background:bg, border:brd, borderRadius:5, cursor:"text", verticalAlign:"middle", ...style,
    }} onClick={e=>{e.stopPropagation(); if(isEmpty||!isActive){const first=nodes.find(n=>n.type==="text");if(first)setCursor(first.id);}}}>
      {isEmpty
        ? <span style={{color:isActive?"#6366F1":"#F59E0B",fontSize:12,fontWeight:700}}>□</span>
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
    const norm = (s) => {
      let r = s;
      let prev = "";
      let _ci=0; while (prev !== r && _ci++<30) { prev = r; r = r.replace(/FRAC\(([^()]+),([^()]+)\)/g, "($1)/($2)"); }
      r = r.replace(/ROOT\(([^,]+),([^)]+)\)/g, "nthRoot($2,$1)");
      r = r.replace(/LOG_BASE\(([^,]+),([^)]+)\)/g, "log($2,$1)");
      r = r.replace(/√\(([^)]+)\)/g, "sqrt($1)");
      r = r.replace(/√(\w+)/g, "sqrt($1)");
      r = r.replace(/π/g, "pi");
      r = r.replace(/×/g, "*");
      r = r.replace(/÷/g, "/");
      r = r.replace(/INT\([^)]+\)/g, "0");
      return r;
    };
    const expr = norm(str);
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
  const [angleMode, setAngleMode] = useState("deg"); // deg | rad
  const [shiftMode, setShiftMode] = useState(false); // lowercase/uppercase Latin
  const holdRef = useRef(null);

  const [katexReady, setKatexReady] = useState(!!window.katex);

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
  const deleteChar = () => setNodes(prev => updateTextNode(prev, cursor, v => v.slice(0, -1)));
  const clearAll = () => { const r = { id: uid(), type: "text", value: "", nodes: null }; setNodes([r]); setCursor(r.id); };

  const movePrev = () => { const all = collectTextNodes(nodes); const i = all.findIndex(n => n.id === cursor); if (i > 0) setCursor(all[i - 1].id); };
  const moveNext = () => { const all = collectTextNodes(nodes); const i = all.findIndex(n => n.id === cursor); if (i >= 0 && i < all.length - 1) setCursor(all[i + 1].id); };

  const insertStructure = (type) => {
    const after = { id: uid(), type: "text", value: "", nodes: null };
    let struct = null;
    if (type === "frac") struct = createNode("frac", "", [slotNode(), slotNode()]);
    else if (type === "sqrt") struct = createNode("sqrt", "", [slotNode()]);
    else if (type === "nthroot") struct = createNode("nthroot", "", [slotNode(), slotNode()]);
    else if (type === "cbrt") { const ds = slotNode([textNode("3")]); const as2 = slotNode(); struct = createNode("nthroot", "", [ds, as2]); const a2 = textNode(""); setNodes(prev => insertStructureInto(prev, cursor, struct, a2)); setCursor(as2.nodes[0]?.id); return; }
    else if (type === "sup") struct = createNode("sup", "", [slotNode(), slotNode()]);
    else if (type === "abs") struct = createNode("abs", "", [slotNode()]);
    else if (type === "logbase") struct = createNode("logbase", "", [slotNode(), slotNode()]);
    else if (type === "mixedfrac") struct = createNode("mixedfrac", "", [slotNode(), slotNode(), slotNode()]);
    else if (type === "defint") { const xSlot = slotNode([textNode("x")]); struct = createNode("defint", "", [slotNode(), slotNode(), slotNode(), xSlot]); }
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
    if (btn.struct) { insertStructure(btn.struct); return; }
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
        { l:"ln", c:"ln(", lt:"\\ln", dot:true, sub:[{l:"lg",c:"lg(",lt:"\\lg"},{lt:"\\log_{\\square}",struct:"logbase"},{l:"log",c:"log(",lt:"\\log"}] },
        { l:"(", c:"(", op:true, dot:true, sub:[{l:"[",c:"["},{l:"{",c:"{"}] },
        { l:")", c:")", op:true, dot:true, sub:[{l:"]",c:"]"},{l:"}",c:"}"}] },
      ],
      [
        { l:"sin", c:"sin(", lt:"\\sin" },
        { l:"cos", c:"cos(", lt:"\\cos" },
        { l:"tan", c:"tan(", lt:"\\tan" },
        { l:"e", c:"e", lt:"e" },
        { l:"π", c:"π", lt:"\\pi" },
        { l:"⌫", c:"DEL", del:true },
      ],
      [
        { l:"7", c:"7", num:true },
        { l:"8", c:"8", num:true },
        { l:"9", c:"9", num:true },
        { l:"÷", c:"/", lt:"\\div", op:true },
        { l:"×", c:"*", lt:"\\times", op:true },
        { l:"C", c:"CLR", clr:true },
      ],
      [
        { l:"4", c:"4", num:true },
        { l:"5", c:"5", num:true },
        { l:"6", c:"6", num:true },
        { l:"−", c:"-", op:true },
        { l:"+", c:"+", op:true },
        { l:"=", c:"=", eq:true },
      ],
      [
        { l:"1", c:"1", num:true },
        { l:"2", c:"2", num:true },
        { l:"3", c:"3", num:true },
        { l:",", c:",", num:true },
        { l:";", c:";", op:true },
        { l:"←", c:"PREV", nav:true },
      ],
      [
        { l:"0", c:"0", num:true },
        { l:"x", c:"x", op:true },
        { l:"±", c:"±", lt:"\\pm", op:true },
        { l:"π", c:"π", lt:"\\pi" },
        { l:"%", c:"%", op:true },
        { l:"→", c:"NEXT", nav:true },
      ],
    ],
    extra: [
      [
        { l:"asin", c:"arcsin(", lt:"\\arcsin" },
        { l:"acos", c:"arccos(", lt:"\\arccos" },
        { l:"atan", c:"arctan(", lt:"\\arctan" },
        { l:"|x|", struct:"abs", lt:"|\\square|" },
        { l:"n!", c:"!", lt:"n!" },
      ],
      [
        { l:"sinh", c:"sinh(", lt:"\\sinh" },
        { l:"cosh", c:"cosh(", lt:"\\cosh" },
        { l:"tanh", c:"tanh(", lt:"\\tanh" },
        { l:"∞", c:"∞", lt:"\\infty" },
        { l:"%", c:"%", op:true },
      ],
      [
        { l:"α", c:"α", lt:"\\alpha" },
        { l:"β", c:"β", lt:"\\beta" },
        { l:"θ", c:"θ", lt:"\\theta" },
        { l:"λ", c:"λ", lt:"\\lambda" },
        { l:"μ", c:"μ", lt:"\\mu" },
      ],
      [
        { l:"<", c:"<", dot:true, sub:[{l:"≤",c:"≤",lt:"\\leq"}] },
        { l:">", c:">", dot:true, sub:[{l:"≥",c:"≥",lt:"\\geq"}] },
        { l:"≠", c:"≠", lt:"\\neq" },
        { l:"±", c:"±", lt:"\\pm" },
        { l:"°", c:"°", lt:"^{\\circ}" },
      ],
      [
        { l:"∫", struct:"defint", lt:"\\int_{a}^{b}" },
        { l:"∑", c:"∑", lt:"\\sum" },
        { l:"∂", c:"∂", lt:"\\partial" },
        { l:"∈", c:"∈", lt:"\\in" },
        { l:"≡", c:"≡", lt:"\\equiv" },
      ],
      [
        { l:"[", c:"[" }, { l:"]", c:"]" },
        { l:"{", c:"{" }, { l:"}", c:"}" },
        { l:"→", c:"NEXT", nav:true },
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
      { l:"←", c:"PREV", nav:true },
      { l:"→", c:"NEXT", nav:true },
    ],
  ];

  const rows = tab === "latin" ? latinRows : (KEYS[tab] || KEYS.basic);

  const btnStyle = (btn) => {
    if (btn.num)   return { bg: "#FFFFFF", fg: "#1E293B", brd: "#DDE3F0", fs: 22, fw: 700, ff: "'SF Mono',monospace" };
    if (btn.del)   return { bg: "#FEF2F2", fg: "#EF4444", brd: "#FECACA", fs: 18, fw: 700, ff: "inherit" };
    if (btn.clr)   return { bg: "#FFF7ED", fg: "#EA580C", brd: "#FED7AA", fs: 15, fw: 700, ff: "inherit" };
    if (btn.op)    return { bg: "#F1F5FF", fg: "#4338CA", brd: "#C7D2FE", fs: 17, fw: 700, ff: "inherit" };
    if (btn.eq)    return { bg: "#6366F1", fg: "#FFFFFF", brd: "#4F46E5", fs: 17, fw: 800, ff: "inherit" };
    if (btn.nav)   return { bg: "#EEF1FF", fg: "#4F46E5", brd: "#C7D2FE", fs: 18, fw: 700, ff: "inherit" };
    if (btn.shift) return { bg: shiftMode ? "#6366F1" : "#E0E7FF", fg: shiftMode ? "#FFFFFF" : "#4338CA", brd: "#C7D2FE", fs: 16, fw: 700, ff: "inherit" };
    if (btn.lat)   return { bg: "#FFFBF0", fg: "#92400E", brd: "#FDE68A", fs: 20, fw: 700, ff: "'KaTeX_Math','Computer Modern',Georgia,serif" };
    return { bg: "#EEF1FF", fg: "#3730A3", brd: "#C7D2FE", fs: 13, fw: 600, ff: "inherit" };
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
      {/* ── TOP BAR: input display + calculator result ── */}
      <div style={{
        background: "#FFFFFF",
        borderBottom: "1.5px solid #E0E7FF",
        borderRadius: "20px 20px 0 0",
        padding: "10px 12px",
        display: "flex", alignItems: "center", gap: 8,
      }}>
        <button onClick={movePrev} style={{ width: 34, height: 34, background: "#EEF1FF", border: "1px solid #C7D2FE", borderRadius: 9, color: "#4F46E5", fontSize: 16, cursor: "pointer", fontWeight: 700, flexShrink: 0 }}>←</button>
        <button onClick={moveNext} style={{ width: 34, height: 34, background: "#EEF1FF", border: "1px solid #C7D2FE", borderRadius: 9, color: "#4F46E5", fontSize: 16, cursor: "pointer", fontWeight: 700, flexShrink: 0 }}>→</button>

        {/* Formula display */}
        <div
          onClick={e => { e.stopPropagation(); const all = collectTextNodes(nodes); const last = all[all.length - 1]; if (last) setCursor(last.id); }}
          style={{
            flex: 1, minHeight: 44, background: "#F8FAFF",
            border: "2px solid #6366F1", borderRadius: 10, padding: "6px 12px",
            display: "flex", alignItems: "center", flexWrap: "wrap", gap: 3,
            cursor: "text", overflowX: "auto",
            boxShadow: "0 0 0 3px rgba(99,102,241,0.1)",
          }}>
          {nodes.length === 0 || (nodes.length === 1 && nodes[0].type === "text" && !nodes[0].value)
            ? <span style={{ color: "#94A3B8", fontSize: 15 }}>Javob yozing...</span>
            : nodes.map(n => renderNode(n, cursor, id => setCursor(id)))
          }
        </div>

        {/* Result badge — faqat admin uchun */}
        {isAdmin && calcVal !== null && (
          <div style={{
            background: "linear-gradient(135deg,#4F46E5,#7C3AED)",
            color: "white", borderRadius: 10, padding: "6px 12px",
            fontSize: 15, fontWeight: 800, flexShrink: 0,
            boxShadow: "0 2px 8px rgba(99,102,241,0.3)",
            maxWidth: 110, overflow: "hidden", textOverflow: "ellipsis",
            whiteSpace: "nowrap",
          }}>
            = {calcVal}
          </div>
        )}

        <button onClick={deleteChar} style={{ width: 34, height: 34, background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: 9, color: "#EF4444", fontSize: 16, cursor: "pointer", flexShrink: 0 }}>⌫</button>
        <button onClick={onClose} style={{ padding: "8px 14px", background: "#6366F1", border: "none", borderRadius: 9, color: "white", fontWeight: 800, fontSize: 13, cursor: "pointer", flexShrink: 0 }}>OK ✓</button>
      </div>

      {/* ── TAB ROW ── */}
      <div style={{ display: "flex", background: "#F1F5FF", borderBottom: "1px solid #DDE3F0" }}>
        {[["basic", "Asosiy"], ["extra", "Qo'shimcha"], ["latin", "Lotin"]].map(([t, l]) => (
          <button key={t} onClick={() => setTab(t)} style={{
            flex: 1, padding: "9px 4px",
            background: tab === t ? "#FFFFFF" : "transparent",
            border: "none", borderBottom: tab === t ? "2.5px solid #6366F1" : "2.5px solid transparent",
            color: tab === t ? "#4F46E5" : "#94A3B8",
            fontWeight: tab === t ? 800 : 500, fontSize: 13, cursor: "pointer",
          }}>{l}</button>
        ))}
        {/* Angle mode toggle */}
        <button onClick={() => setAngleMode(m => m === "deg" ? "rad" : "deg")} style={{
          padding: "9px 14px", background: "transparent", border: "none",
          borderBottom: "2.5px solid transparent",
          color: "#64748B", fontSize: 12, cursor: "pointer", fontWeight: 600,
          flexShrink: 0,
        }}>
          <span style={{ background: "#E0E7FF", borderRadius: 6, padding: "3px 8px", color: "#4338CA" }}>{angleMode.toUpperCase()}</span>
        </button>
      </div>

      {/* ── KEY GRID (Desmos layout: 5 columns × 6 rows) ── */}
      <div style={{ padding: "8px 8px 16px", background: "#FFFFFF" }}>
        {rows.map((row, ri) => (
          <div key={ri} style={{ display: "grid", gridTemplateColumns: `repeat(${row.length}, 1fr)`, gap: 5, marginBottom: 5 }}>
            {row.map((btn, ci) => {
              const st = btnStyle(btn);
              const khtml = ks(btn.lt);
              const hasSub = btn.sub?.length > 0;
              return (
                <button key={ci}
                  onPointerDown={e => { e.currentTarget.setPointerCapture(e.pointerId); startHold(e, btn); }}
                  onPointerUp={() => endHold(btn)}
                  onPointerCancel={() => { if (holdRef.current) { clearTimeout(holdRef.current); holdRef.current = null; } }}
                  style={{
                    minHeight: 50, background: st.bg,
                    border: `1.5px solid ${st.brd}`, borderRadius: 12,
                    color: st.fg, fontSize: st.fs, fontWeight: st.fw,
                    fontFamily: st.ff || "inherit",
                    fontStyle: btn.lat ? "italic" : "normal",
                    cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
                    position: "relative", WebkitTapHighlightColor: "transparent",
                    boxShadow: "0 2px 0 rgba(0,0,0,0.07)", userSelect: "none",
                    padding: "3px 2px", transition: "all 0.1s",
                  }}
                  onPointerEnter={e => { e.currentTarget.style.background = "#EEF1FF"; e.currentTarget.style.borderColor = "#818CF8"; }}
                  onPointerLeave={e => { e.currentTarget.style.background = st.bg; e.currentTarget.style.borderColor = st.brd; }}
                >
                  {khtml
                    ? <span dangerouslySetInnerHTML={{ __html: khtml }} style={{ color: st.fg, pointerEvents: "none", fontSize: 14, lineHeight: 1.3 }} />
                    : <span style={{ fontFamily: btn.num ? "'SF Mono',monospace" : "inherit", pointerEvents: "none" }}>{btn.l || "?"}</span>
                  }
                  {hasSub && <span style={{ position: "absolute", bottom: 3, right: 4, width: 5, height: 5, borderRadius: "50%", background: "#EF4444", boxShadow: "0 0 0 1px white" }} />}
                </button>
              );
            })}
          </div>
        ))}
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
                onClick={() => { if (it.struct) insertStructure(it.struct); else if (it.c) insertChar(it.c); setPopup(null); }}
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
      {[["firstName","Ism"],["lastName","Familiya"],["group","Guruh/Sinf (10-A)"],["phone","Telefon raqam (login)"]].map(([k,ph])=>(
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
  const [pdfUrl,setPdfUrl]=useState(existing?.pdfUrl||null);
  const [sections,setSections]=useState(existing?.sections||[{name:"Fan 1",count:10,type:"closed4"}]);
  const [questions,setQuestions]=useState(existing?.questions||[]);
  const [step,setStep]=useState(existing?2:1);
  const [err,setErr]=useState("");
  const [kbdOpen,setKbdOpen]=useState(null); // {qIdx, sub}

  const typeOpts=(t)=>({closed2:2,closed3:3,closed4:4,closed5:5,closed6:6,closed7:7,closed8:8,open:0})[t]||4;
  const typeLabel=(t)=>({closed2:"2 variant",closed3:"3 variant",closed4:"4 variant (A-D)",closed5:"5 variant (A-E)",closed6:"6 variant",closed7:"7 variant",closed8:"8 variant",open:"Ochiq (yozma)"})[t]||t;
  const totalQ=sections.reduce((s,sec)=>s+Number(sec.count),0);

  const [pdfLoading,setPdfLoading]=useState(false);
  const [docType,setDocType]=useState(existing?.latexSource ? "latex" : "pdf"); // "pdf" | "latex"
  const [latexSource,setLatexSource]=useState(existing?.latexSource || "");
  const [latexFileName,setLatexFileName]=useState(existing?.latexFileName || "");
  const [latexMode,setLatexMode]=useState(existing?.latexFileName ? "upload" : "write"); // "upload" | "write"
  const [showLatexPreview,setShowLatexPreview]=useState(true);
  const [latexImages,setLatexImages]=useState(existing?.latexImages || {}); // {rasm1: dataURL, ...}
  const latexTextareaRef = useRef(null);


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
      setLatexImages(p=>({...p,[key]:ev.target.result}));
      const cmd = `\\includegraphics{${key}}`;
      const ta = latexTextareaRef.current;
      if (ta) {
        const start = ta.selectionStart ?? latexSource.length;
        const end = ta.selectionEnd ?? latexSource.length;
        const next = latexSource.slice(0,start) + cmd + latexSource.slice(end);
        setLatexSource(next);
        requestAnimationFrame(()=>{ ta.focus(); ta.selectionStart=ta.selectionEnd=start+cmd.length; });
      } else {
        setLatexSource(p=>p + (p?"\n":"") + cmd);
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
    const schedTs = localInputToTs(scheduledAt);
    const wasActive = existing?.active||false;
    // Agar rejalashtirilgan vaqt kelajakda bo'lsa va test hali qo'lda faollashtirilmagan bo'lsa, uni nofaol holatda saqlaymiz — vaqt kelganda avtomatik faollashadi.
    const willAutoStart = schedTs && schedTs > Date.now() && !wasActive;
    onSave({id:existing?.id||Date.now(),name,duration,closedCount:questions.filter(q=>q.type==="closed").length,optionsCount:sections[0]?typeOpts(sections[0].type):4,sections,questions,active:willAutoStart?false:wasActive,scheduledAt:schedTs,showAnswersAfter:showAnswers,showStats,startedAt:existing?.startedAt||null,pdfUrl:docType==="pdf"?pdfUrl:null,latexSource:docType==="latex"?latexSource:null,latexFileName:docType==="latex"?latexFileName:null,latexImages:docType==="latex"?latexImages:null,accessCode:requireCode?accessCode.trim().toUpperCase():null,codePrice:requireCode?codePrice:null});
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
          <input value={name} onChange={e=>setName(e.target.value)} style={S.input} placeholder="Algebra imtihoni"/>

          {/* PDF / LaTeX upload */}
          <label style={S.label}>📄 Test varianti (ixtiyoriy)</label>
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
            <div><label style={S.label}>⏱ Vaqt (daqiqa)</label><input type="number" value={duration} onChange={e=>setDuration(+e.target.value)} style={S.input} min={1}/></div>
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

          {/* Pullik test / Kirish kodi */}
          <div style={{background:"#FFFBEB",border:"1.5px solid #FDE68A",borderRadius:12,padding:14,marginBottom:16}}>
            <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",marginBottom:requireCode?12:0}}>
              <input type="checkbox" checked={requireCode} onChange={e=>setRequireCode(e.target.checked)}
                style={{width:20,height:20,accentColor:"#F59E0B",cursor:"pointer"}}/>
              <span style={{fontWeight:700,color:"#92400E",fontSize:14}}>🔒 Pullik test (kirish kodi talab qilinsin)</span>
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
                  <button onClick={()=>setSections(p=>p.map((s,idx)=>idx===i?{...s,count:Math.max(1,s.count-1)}:s))} style={{width:30,height:"100%",background:C.danger,border:"none",color:"white",fontSize:18,cursor:"pointer",fontWeight:900}}>−</button>
                  <input type="number" value={sec.count} onChange={e=>setSections(p=>p.map((s,idx)=>idx===i?{...s,count:Math.max(1,+e.target.value)}:s))} style={{flex:1,background:"transparent",border:"none",color:C.text,fontWeight:800,fontSize:15,textAlign:"center",outline:"none",minWidth:0}} min={1}/>
                  <button onClick={()=>setSections(p=>p.map((s,idx)=>idx===i?{...s,count:s.count+1}:s))} style={{width:30,height:"100%",background:C.primary,border:"none",color:"white",fontSize:18,cursor:"pointer",fontWeight:900}}>+</button>
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
                              <div onClick={()=>setKbdOpen(isAct?null:{qIdx:i,sub:null})} style={{flex:1,minHeight:44,padding:"10px 14px",background:isAct?"#EEF1FF":"#F8F9FF",border:`2px solid ${isAct?C.primary:q.correctAnswer?C.success:C.border}`,borderRadius:10,cursor:"pointer",fontSize:17,display:"flex",alignItems:"center",boxShadow:isAct?"0 0 0 3px rgba(79,110,247,0.2)":"none"}}>
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
      {confirmModal && <ConfirmModal message={confirmModal.message} onConfirm={confirmModal.onConfirm} onCancel={()=>setConfirmModal(null)}/>}
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
        {[["tests","📋 Testlar"],["users","👥 O'quvchilar"],["results","📊 Natijalar"]].map(([t,l])=>(
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
                        {test.pdfUrl&&<button onClick={()=>setAdminDocPreview({type:"pdf",url:test.pdfUrl,name:test.name})} style={{...S.badge,background:"#FEF3C7",color:"#92400E",border:"none",cursor:"pointer"}}>📄 PDF</button>}
                        {test.latexSource&&<button onClick={()=>setAdminDocPreview({type:"latex",source:test.latexSource,name:test.name,id:test.id,images:test.latexImages})} style={{...S.badge,background:"#EEF1FF",color:C.primary,border:"none",cursor:"pointer"}}>∑ LaTeX</button>}
                        {test.accessCode&&<span style={{...S.badge,background:"#FEF3C7",color:"#92400E"}}>🔒 Kod: {test.accessCode}{test.codePrice?` (${test.codePrice} so'm)`:""}</span>}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      <button onClick={()=>toggleActive(test.id)} style={{...S.btnSmall,background:test.active?C.danger:C.success}}>{test.active?"⛔ To'xtatish":"✅ Faollashtirish"}</button>
                      <button onClick={()=>{setEditing(test);setCreating(false);}} style={{...S.btnSmall,background:C.primary}}>✏️ Tahrirlash</button>
                      <button onClick={()=>setExportModal(buildExcelExport(test,results,users))} style={{...S.btnSmall,background:C.successDark}}>📥 Excel</button>
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
                  <p style={{margin:"0 0 4px",color:"#92400E",fontSize:13,fontWeight:700}}>💰 Pullik testlardan tushum (taxminiy)</p>
                  <p style={{margin:0,color:"#92400E",fontSize:28,fontWeight:900}}>{totalRevenue.toLocaleString()} so'm</p>
                </div>
              );
            })()}
            {tests.map(test=>{
              const tr=results.filter(r=>r.testId===test.id);
              if(!tr.length) return null;
              return (
                <div key={test.id} style={{marginBottom:24}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <h4 style={{margin:0,color:C.primary}}>{test.name}</h4>
                    <button onClick={()=>setExportModal(buildExcelExport(test,results,users))} style={{...S.btnSmall,background:C.successDark}}>📥 Excel</button>
                  </div>
                  <div style={{overflowX:"auto"}}>
                    <table style={S.table}>
                      <thead><tr>{["#","F.I.O","Guruh","Ball","Foiz","Vaqt","Sana"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>{tr.sort((a,b)=>b.totalScore-a.totalScore).map((r,i)=>{
                        const u=users.find(u=>u.phone===r.userPhone);
                        const pct=Math.round((r.totalScore/test.questions.length)*100);
                        return (
                          <tr key={r.id} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                            <td style={S.td}>{i+1}</td>
                            <td style={S.td}>{u?`${u.firstName} ${u.lastName}`:r.userPhone}</td>
                            <td style={S.td}>{u?.group||"-"}</td>
                            <td style={S.td}><b style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger}}>{r.totalScore}</b>/{test.questions.length}</td>
                            <td style={S.td}><span style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger,fontWeight:700}}>{pct}%</span></td>
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
            <span style={{fontWeight:700,fontSize:16}}>{docModal.type==="pdf"?"📄":"∑"} Test varianti</span>
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
              <h3 style={{margin:"8px 0 4px",fontSize:18}}>Pullik test</h3>
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
            {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now).length>0&&(
              <div style={{marginBottom:24}}>
                <h3 style={{marginBottom:16}}>🗓️ Rejalashtirilgan Testlar</h3>
                {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now).map(test=>(
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
            {tests.filter(t=>t.active).map(test=>{
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
                        {test.pdfUrl&&<button onClick={()=>setDocModal({type:"pdf",url:test.pdfUrl})} style={{...S.badge,background:"#FEF3C7",color:"#92400E",border:"none",cursor:"pointer",fontSize:12}}>📄 Variantni ko'rish</button>}
                        {test.latexSource&&<button onClick={()=>setDocModal({type:"latex",source:test.latexSource,id:test.id,images:test.latexImages})} style={{...S.badge,background:C.primaryLight,color:C.primary,border:"none",cursor:"pointer",fontSize:12}}>∑ Variantni ko'rish</button>}
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
            {tests.filter(t=>t.active).length===0&&<div style={S.empty}>Hozircha faol testlar yo'q</div>}
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
                    <div><h4 style={{margin:"0 0 4px",fontSize:15}}>{test.name}</h4><p style={{margin:0,color:C.textMid,fontSize:12}}>{new Date(r.id).toLocaleDateString("uz-UZ")}</p></div>
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
            {tests.filter(t=>t.active && (t.pdfUrl || t.latexSource)).map(test=>(
              <div key={test.id} style={{...S.card,padding:18,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <div style={{display:"flex",alignItems:"center",gap:12}}>
                  <div style={{width:44,height:44,borderRadius:10,background:test.pdfUrl?"#FEF3C7":C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{test.pdfUrl?"📄":"∑"}</div>
                  <div>
                    <h4 style={{margin:"0 0 2px",fontSize:15}}>{test.name}</h4>
                    <p style={{margin:0,color:C.textMid,fontSize:12}}>{test.questions?.length} savol • {test.pdfUrl?"PDF":"LaTeX"}</p>
                  </div>
                </div>
                <button onClick={()=>setDocModal(test.pdfUrl?{type:"pdf",url:test.pdfUrl,id:test.id}:{type:"latex",source:test.latexSource,id:test.id,images:test.latexImages})} style={{...S.btnSmall,background:C.primary,padding:"9px 18px"}}>👁 Ko'rish</button>
              </div>
            ))}
            {tests.filter(t=>t.active && (t.pdfUrl || t.latexSource)).length===0&&<div style={S.empty}>Hozircha test variantlari mavjud emas</div>}
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
                        {["#","Test nomi","Sana","Ball","Natija","Holat"].map(h=>(
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
          {(test.pdfUrl||test.latexSource)&&<button onClick={()=>setPdfViewOpen(true)} style={{...S.btnSmall,background:"#F59E0B",padding:"9px 12px",fontSize:12,border:"none",cursor:"pointer"}}>{test.pdfUrl?"📄":"∑"} Test ko'rish</button>}
          <button onClick={()=>setShowConfirm(true)} disabled={grading} style={{...S.btnSuccess,padding:"9px 16px",fontSize:13,opacity:grading?0.6:1}}>{grading?"⏳ Tekshirilmoqda...":"✅ Yakunlash"}</button>
          <button onClick={()=>setConfirmModal({message:"Testdan chiqasizmi? Belgilagan javoblaringiz saqlanadi, keyinroq davom ettirishingiz mumkin.", confirmLabel:"Chiqish", onConfirm:()=>{setConfirmModal(null);onExit();}})} style={{...S.btnDanger,padding:"9px 10px",fontSize:13}}>✕</button>
        </div>
      </div>

      {/* In-test PDF viewer — stays within TestTaking, timer keeps running,
          answers preserved. User can switch back and forth freely. */}
      {pdfViewOpen&&(test.pdfUrl||test.latexSource)&&(
        <div style={{position:"fixed",inset:0,background:test.pdfUrl?"#1a1a1a":"white",zIndex:9000,display:"flex",flexDirection:"column"}}>
          <div style={{background:"#F59E0B",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <span style={{fontWeight:800,fontSize:15,color:"white"}}>{test.pdfUrl?"📄":"∑"} {test.name}</span>
              <p style={{margin:0,fontSize:11,color:"rgba(255,255,255,0.85)"}}>Vaqt davom etmoqda: {fmt(timeLeft)} • Javoblaringiz saqlanadi</p>
            </div>
            <button onClick={()=>setPdfViewOpen(false)} style={{...S.btnSuccess,padding:"9px 16px",fontSize:13,background:"white",color:"#92400E",fontWeight:800}}>
              ✏️ Javob berishga qaytish
            </button>
          </div>
          {test.pdfUrl
            ? <PdfViewer url={test.pdfUrl} persistKey={"pdf_scroll_"+test.id}/>
            : <ScrollPersistDiv persistKey={"taking_latex_scroll_"+test.id} style={{flex:1,overflowY:"auto"}}><LatexDocViewer source={test.latexSource} images={test.latexImages}/></ScrollPersistDiv>
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
                    <span style={{color:ok?C.successDark:C.danger,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}>{stu?parseToNodes(stu).map(n=>renderNode(n,null,()=>{})):<span style={{color:C.textLight}}>—</span>}</span>
                    <span style={{color:C.textMid}}>→</span>
                    <span style={{display:"inline-flex",flexWrap:"wrap",gap:2,alignItems:"center",color:C.successDark,fontWeight:600,fontSize:18}}>{parseToNodes(sp.answer||"").map(n=>renderNode(n,null,()=>{}))}</span>
                    <span style={{marginLeft:"auto"}}>{ok?"✅":"❌"}</span>
                  </div>
                );
              }):(()=>{
                const ok=result.scores?.[idx];
                const stu=result.openAnswers?.[idx];
                return (
                  <div style={{display:"flex",gap:8,alignItems:"flex-start",flexWrap:"wrap",padding:"10px",borderRadius:8,background:ok?C.successLight:C.dangerLight}}>
                    <span style={{color:ok?C.successDark:C.danger,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}>{stu?parseToNodes(stu).map(n=>renderNode(n,null,()=>{})):<span style={{color:C.textLight,fontSize:14}}>Javob berilmadi</span>}</span>
                    <span style={{color:C.textMid}}>→</span>
                    <span style={{color:C.successDark,fontWeight:600,fontSize:18,display:"inline-flex",alignItems:"center",flexWrap:"wrap",gap:2}}>{parseToNodes(q.correctAnswer||"").map(n=>renderNode(n,null,()=>{}))}</span>
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
  const [page,setPage]=useState("login");
  const [user,setUser]=useState(null);
  const [isAdmin,setIsAdmin]=useState(false);
  const [isTeacher,setIsTeacher]=useState(false);
  const [teacherInfo,setTeacherInfo]=useState(null);
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
      st.textContent = "@keyframes blink{0%,100%{opacity:1}50%{opacity:0}}";
      document.head.appendChild(st);
    }
    // Prevent double-tap zoom
    document.documentElement.style.touchAction = "manipulation";
  },[]);

  if(showSplash) return <SplashScreen onDone={()=>setShowSplash(false)}/>;
  if(page==="admin"&&isAdmin) return <AdminPanel isFullAdmin={true} onLogout={()=>{setIsAdmin(false);setPage("login");}}/>;
  if(page==="teacher"&&isTeacher) return <AdminPanel isFullAdmin={false} teacherInfo={teacherInfo} onLogout={()=>{setIsTeacher(false);setTeacherInfo(null);setPage("login");}}/>;
  if(page==="student"&&user) return <StudentDashboard user={user} onLogout={()=>{setUser(null);setPage("login");}}/>;
  if(page==="register") return <RegisterPage onDone={u=>{setUser(u);setPage("student");}} onLogin={()=>setPage("login")}/>;
  return <LoginPage
    onLogin={u=>{setUser(u);setPage("student");}}
    onRegister={()=>setPage("register")}
    onAdmin={(fullAdmin, teacher)=>{
      if(fullAdmin){ setIsAdmin(true); setPage("admin"); }
      else { setIsTeacher(true); setTeacherInfo(teacher); setPage("teacher"); }
    }}
  />;
}
