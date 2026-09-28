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
  // Belgidan keyin bo'sh joy qo'yamiz: aks holda "πe" -> "\\pie" bo'lib, KaTeX uni noma'lum buyruq deb o'qiydi
  Object.entries(syms).forEach(([k,v]) => { r = r.split(k).join(v + " "); });
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
// Admin paroli kodda OCHIQ holda saqlanmaydi — faqat uning xeshi (PBKDF2-SHA256). Admin panelidagi
// "🔐 Xavfsizlik" bo'limida parol almashtirilsa, yangi xesh "appSettings" ichida saqlanadi va shu
// standart qiymat o'rniga ishlatiladi. Standart parolni albatta almashtiring!
const ADMIN_PW_HASH_DEFAULT = "pbkdf2$100000$39CP2S4n4KJlCe+YIgYf5w==$O/i0b3rLdFWAotHCTQC3alAYRoJrRGXrsIZ6q3dPDAo=";

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
const FIREBASE_SYNC_COLLECTIONS = ["users", "tests", "results", "teachers", "partners", "partnerUploads", "raschResults", "appSettings", "groups"];

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


// ===== XAVFSIZLIK: PAROLLAR =====
// Parollar ochiq matn holida saqlanmaydi: PBKDF2-SHA256 (100 000 marta, har bir parolga alohida
// tasodifiy "tuz") bilan xeshlanadi. Saqlash formati: pbkdf2$<takror>$<tuz>$<xesh>.
// Eski (ochiq) parollar kirishda tekshiriladi va darhol xeshlanadi; admin paneli ochilganda qolgan
// barcha ochiq parollar ham fon rejimida xeshlab chiqiladi.
const PWD_ITER = 100000;
const bufToB64 = (buf) => { const a = new Uint8Array(buf); let s = ""; for (let i = 0; i < a.length; i++) s += String.fromCharCode(a[i]); return btoa(s); };
const b64ToBytes = (s) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
function cryptoReady() { return !!(window.crypto && window.crypto.subtle); }
async function pbkdf2Bits(password, salt, iterations) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(String(password)), "PBKDF2", false, ["deriveBits"]);
  return crypto.subtle.deriveBits({ name: "PBKDF2", salt, iterations, hash: "SHA-256" }, key, 256);
}
function isHashedPassword(s) { return typeof s === "string" && s.startsWith("pbkdf2$"); }
async function hashPassword(password) {
  if (!cryptoReady()) throw new Error("Brauzer xavfsiz shifrlashni qo'llamaydi. Saytni https orqali oching.");
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const bits = await pbkdf2Bits(password, salt, PWD_ITER);
  return `pbkdf2$${PWD_ITER}$${bufToB64(salt)}$${bufToB64(bits)}`;
}
async function verifyPassword(stored, password) {
  if (!stored || typeof password !== "string") return false;
  if (!isHashedPassword(stored)) return stored === password; // eski, hali xeshlanmagan parol
  if (!cryptoReady()) return false;
  const [, it, saltB64, hashB64] = stored.split("$");
  const bits = bufToB64(await pbkdf2Bits(password, b64ToBytes(saltB64), +it));
  // vaqt bo'yicha teng solishtirish
  let diff = bits.length ^ hashB64.length;
  for (let i = 0; i < Math.min(bits.length, hashB64.length); i++) diff |= bits.charCodeAt(i) ^ hashB64.charCodeAt(i);
  return diff === 0;
}
// Kirish muvaffaqiyatli bo'lganda eski ochiq parolni xeshlab saqlaydi
async function upgradePasswordIfPlain(collection, match, plain) {
  const rec = (db.get(collection) || []).find(match);
  if (!rec || isHashedPassword(rec.password) || !cryptoReady()) return rec;
  const h = await hashPassword(plain);
  const next = (db.get(collection) || []).map(x => match(x) ? { ...x, password: h } : x);
  db.set(collection, next);
  return next.find(match);
}
function getAdminPasswordHash() {
  const s = (db.get("appSettings") || []).find(x => x.id === "admin");
  return (s && s.passwordHash) || ADMIN_PW_HASH_DEFAULT;
}
function adminUsesDefaultPassword() { return getAdminPasswordHash() === ADMIN_PW_HASH_DEFAULT; }
// Barcha to'plamlardagi ochiq parollarni fon rejimida (bo'lib-bo'lib) xeshlaydi.
// Saqlashda eng oxirgi ma'lumot qayta o'qiladi — shu vaqt ichida ro'yxatdan o'tganlar yo'qolmaydi.
let passwordMigrationRunning = false;
async function migratePlainPasswords(onProgress) {
  if (passwordMigrationRunning || !cryptoReady()) return 0;
  passwordMigrationRunning = true;
  let done = 0;
  try {
    for (const col of ["users", "teachers", "partners"]) {
      const plain = (db.get(col) || []).filter(x => x.password && !isHashedPassword(x.password));
      for (let i = 0; i < plain.length; i += 25) {
        const chunk = plain.slice(i, i + 25);
        const hashed = new Map();
        for (const x of chunk) hashed.set(x.password + "\u0000" + (x.phone || x.id), await hashPassword(x.password));
        const next = (db.get(col) || []).map(x => {
          const k = x.password + "\u0000" + (x.phone || x.id);
          return (!isHashedPassword(x.password) && hashed.has(k)) ? { ...x, password: hashed.get(k) } : x;
        });
        db.set(col, next);
        done += chunk.length;
        onProgress && onProgress(done);
        await new Promise(r => setTimeout(r, 0));
      }
    }
  } finally { passwordMigrationRunning = false; }
  return done;
}
function countPlainPasswords() {
  return ["users", "teachers", "partners"].reduce((a, col) => a + (db.get(col) || []).filter(x => x.password && !isHashedPassword(x.password)).length, 0);
}

// ===== XAVFSIZLIK: KIRISH URINISHLARINI CHEKLASH =====
// Ketma-ket 5 marta noto'g'ri parol kiritilsa, shu qurilmada kirish vaqtincha bloklanadi
// (har safar blok vaqti ikki barobar uzayadi: 1, 2, 4 ... daqiqa, ko'pi bilan 30 daqiqa).
const LOGIN_GUARD_KEY = "login_guard_v1";
function loginGuardState() { try { return JSON.parse(localStorage.getItem(LOGIN_GUARD_KEY)) || { fails: 0, until: 0, level: 0 }; } catch { return { fails: 0, until: 0, level: 0 }; } }
function loginGuardFail() {
  const g = loginGuardState();
  g.fails = (g.fails || 0) + 1;
  if (g.fails >= 5) { g.level = (g.level || 0) + 1; g.until = Date.now() + Math.min(30, 2 ** (g.level - 1)) * 60000; g.fails = 0; }
  try { localStorage.setItem(LOGIN_GUARD_KEY, JSON.stringify(g)); } catch {}
  return g;
}
function loginGuardReset() { try { localStorage.removeItem(LOGIN_GUARD_KEY); } catch {} }

// ===== SESSIYANI ESLAB QOLISH =====
// Login/parolni har safar qayta so'ramaslik uchun — bir marta kirgan
// foydalanuvchi (o'quvchi/admin/o'qituvchi) qurilmada eslab qolinadi.
// Telegram Mini App har ochilishida sahifa "yangi holatda" boshlangani
// uchun bu ayniqsa muhim (aks holda har safar qaytadan login so'raladi).
const SESSION_KEY = "app_session_v1";
// Sessiyaga hisobning joriy parol xeshidan olingan "tasdiq" (v) va muddat (30 kun) yoziladi.
// Parol almashtirilsa yoki muddat o'tsa — boshqa qurilmalardagi eski sessiyalar bekor bo'ladi.
// Faqat {role:"admin"} kabi qo'lda yozilgan sessiya endi admin panelini ochmaydi.
const SESSION_TTL = 30 * 24 * 3600 * 1000;
function sessionVerifier(passwordStored) { return isHashedPassword(passwordStored) ? passwordStored.slice(-16) : ""; }
function saveSession(s) { try { localStorage.setItem(SESSION_KEY, JSON.stringify({ ...s, exp: Date.now() + SESSION_TTL })); } catch {} }
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
      return { ...t, active: true, startedAt: t.scheduledAt, everActivated: true };
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
const QQ_FLAG_IMG = "data:image/png;base64,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";

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
  r = r.replace(/π/g, "(pi)"); // qavsda: "πe" -> "(pi)*e", "eπ" -> "e*(pi)" bo'lishi uchun
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
        const lastChSup = result.slice(-1);
        if (lastChSup && /[\d)]/.test(lastChSup)) result += "*";
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
        const lastChRoot = result.slice(-1);
        if (lastChRoot && /[\d)]/.test(lastChRoot)) result += "*";
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
        const lastChAbs = result.slice(-1);
        if (lastChAbs && /[\d)]/.test(lastChAbs)) result += "*";
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
        if (lastCh2 && /[\d)a-zA-Z]/.test(lastCh2)) result += "*";
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
// "RASH_30000" Excel shablonidagi (Malumotlar -> Natijalar -> Reyting) formulalarga AYNAN mos.
// Qadamlar:
//  1) Har bir band (55 ta) uchun: P = to'g'ri javoblar soni / o'quvchilar soni (N).
//     Hamma to'g'ri (P=1) yoki hamma xato (P=0) bo'lgan bandning og'irligi 0.
//  2) Qiyinlik b = -ln(P/(1-P)); og'irlik w = b - min(b, shu fan bandlari) + 1.
//  3) Algebra va Geometriya ALOHIDA: xom ball = Σ(javob × w), maks = Σw,
//     logit = ln(x/(maks-x)), bu yerda x = maks·0.3% .. maks·99.7% oralig'iga siqilgan xom ball.
//  4) Z = (logit - o'rtacha) / standart chetlanish (STDEV.S, n-1).
//  5) Algebra bali = (50+10Z)·2/3, Geometriya bali = (50+10Z)/3, Umumiy = yig'indi;
//     90 dan oshsa: 90 + (ortig'i)/100. Umuman to'g'ri javobi yo'q o'quvchi = 0.
//  6) Foiz: <46 bo'lsa yo'q, >=65 bo'lsa 100, aks holda ball/65·100.  BMBA = 93·foiz/100 + 11.
//  7) Baho: <46 NC, <50 C, <55 C+, <60 B, <65 B+, <70 A, qolgani A+.
const DEFAULT_RASCH_SETTINGS = {
  ncMax: 46,     // NC chegarasi (bundan past)
  cMax: 50,      // C chegarasi
  cPlusMax: 55,  // C+ chegarasi
  bMax: 60,      // B chegarasi
  bPlusMax: 65,  // B+ chegarasi (foiz 100% bo'ladigan ball ham shu)
  aMax: 70,      // A chegarasi (bundan yuqori = A+)
};
// Rash modeli faqat aynan 55 ta savol/banddan iborat testlar uchun ishlaydi (shablon shu
// formatga mo'ljallangan): 1–35 oddiy savollar, 36–45 har biri (a) va (b) bandli.
const RASCH_REQUIRED_ITEMS = 55;
// Band indeksi (0 dan) Geometriyaga tegishlimi: 23–30, 33–35 va 41(a)–45(b) savollar.
// Qolganlari (1–22, 31–32, 36(a)–40(b)) — Algebra.
function raschIsGeometry(i) {
  return (i >= 22 && i <= 29) || (i >= 32 && i <= 34) || (i >= 45 && i <= 54);
}
// Testdagi umumiy baholanadigan "punkt"lar soni (har bir sub-qism alohida punkt hisoblanadi —
// bu result.totalScore hisoblangan usul bilan bir xil bo'lishi kerak).
function testTotalItems(test) {
  let n = 0;
  for (const q of test.questions || []) n += q.subParts?.length > 0 ? q.subParts.length : 1;
  return n;
}
function testEligibleForRasch(test) {
  return testTotalItems(test) === RASCH_REQUIRED_ITEMS;
}
function raschGrade(ball, s = DEFAULT_RASCH_SETTINGS) {
  if (ball === null || ball === undefined || !isFinite(ball)) return null;
  if (ball < s.ncMax) return "NC";
  if (ball < s.cMax) return "C";
  if (ball < s.cPlusMax) return "C+";
  if (ball < s.bMax) return "B";
  if (ball < s.bPlusMax) return "B+";
  if (ball < s.aMax) return "A";
  return "A+";
}
// Saytdagi bitta natijani (result) Excel'dagi kabi 55 ta 0/1 javobdan iborat qatorga aylantiradi.
// Tartib: savollar ketma-ket, kichik bandli savol esa (a), (b) ... bo'lib yoyiladi.
function resultItemVector(test, r) {
  const v = [];
  (test.questions || []).forEach((q, i) => {
    if (q.subParts?.length > 0) q.subParts.forEach((_, si) => v.push(r.subScores?.[i]?.[si] ? 1 : 0));
    else v.push(r.scores?.[i] ? 1 : 0);
  });
  return v;
}
// Bir guruh (populyatsiya) o'quvchilar uchun Rash hisobining yadrosi.
// vectors: har biri 55 ta 0/1 dan iborat massivlar. Natija — shu tartibdagi obyektlar.
function raschComputeCohort(vectors, settings = DEFAULT_RASCH_SETTINGS) {
  const N = vectors.length;
  const K = RASCH_REQUIRED_ITEMS;
  if (N === 0) return [];
  // 1-2) Bandlar bo'yicha qiyinlik va og'irlik
  const S = new Array(K).fill(0);
  vectors.forEach(v => { for (let i = 0; i < K; i++) S[i] += v[i] ? 1 : 0; });
  const b = S.map(s => (s === 0 || s === N) ? null : -Math.log((s / N) / (1 - s / N)));
  const w = new Array(K).fill(0);
  [false, true].forEach(g => {
    const idx = [];
    for (let i = 0; i < K; i++) if (raschIsGeometry(i) === g && b[i] !== null) idx.push(i);
    if (!idx.length) return;
    const mn = Math.min(...idx.map(i => b[i]));
    idx.forEach(i => { w[i] = b[i] - mn + 1; });
  });
  // 3-4) Fan bo'yicha xom ball, logit va Z
  const subject = (g) => {
    let mx = 0;
    for (let i = 0; i < K; i++) if (raschIsGeometry(i) === g) mx += w[i];
    const cor = mx * 0.003;
    const raw = vectors.map(v => { let s = 0; for (let i = 0; i < K; i++) if (raschIsGeometry(i) === g && v[i]) s += w[i]; return s; });
    if (!(mx > 0)) return { raw, logit: raw.map(() => null), z: raw.map(() => null) };
    const logit = raw.map(x => { const c = Math.min(Math.max(x, cor), mx - cor); return Math.log(c / (mx - c)); });
    const mean = logit.reduce((a, c) => a + c, 0) / N;
    const sd = N > 1 ? Math.sqrt(logit.reduce((a, c) => a + (c - mean) ** 2, 0) / (N - 1)) : 0;
    const z = logit.map(l => (sd > 0 && isFinite(sd)) ? (l - mean) / sd : null);
    return { raw, logit, z };
  };
  const A = subject(false), G = subject(true);
  // 5-7) Yakuniy ballar
  const out = vectors.map((v, k) => {
    const correct = v.reduce((a, c) => a + (c ? 1 : 0), 0);
    let alg, geo, ball;
    if (correct === 0) { alg = 0; geo = 0; ball = 0; }
    else if (A.z[k] === null || G.z[k] === null) { alg = null; geo = null; ball = null; }
    else {
      alg = (50 + 10 * A.z[k]) * 2 / 3;
      geo = (50 + 10 * G.z[k]) / 3;
      ball = alg + geo;
      if (ball > 90) ball = (ball - 90) / 100 + 90;
    }
    const foiz = ball === null ? null : ball < 46 ? null : ball >= 65 ? 100 : ball / 65 * 100;
    const bmba = foiz === null ? null : 93 * foiz / 100 + 11;
    return {
      correct, total: K,
      algRaw: A.raw[k], geoRaw: G.raw[k], algLogit: A.logit[k], geoLogit: G.logit[k], algZ: A.z[k], geoZ: G.z[k],
      alg, geo, ball, foiz, bmba, daraja: raschGrade(ball, settings), rank: null,
    };
  });
  // Reyting o'rni: ball bo'yicha kamayish tartibida, teng ballarda — ro'yxatdagi tartib bo'yicha
  out.map((o, i) => i).filter(i => out[i].ball !== null)
    .sort((a, c) => (out[c].ball - out[a].ball) || (a - c))
    .forEach((i, r) => { out[i].rank = r + 1; });
  return out;
}
// Bitta testning saytda topshirilgan barcha natijalari uchun Rash hisoblab, "rasch" maydoniga yozadi.
function computeRaschForTest(test, settings = DEFAULT_RASCH_SETTINGS) {
  const all = db.get("results") || [];
  const idxList = [];
  all.forEach((r, i) => { if (r.testId === test.id) idxList.push(i); });
  if (!idxList.length) return { count: 0 };
  const calc = raschComputeCohort(idxList.map(i => resultItemVector(test, all[i])), settings);
  const calculatedAt = Date.now();
  const updated = [...all];
  idxList.forEach((i, k) => { updated[i] = { ...updated[i], rasch: { ...calc[k], calculatedAt, settings, source: "site" } }; });
  db.set("results", updated);
  return { count: idxList.length, calculatedAt };
}

// ===== RASH NATIJALARINI EXCEL FAYLDAN YUKLASH (import) =====
// Fayl "RASH_30000" shablonidagi kabi bo'lishi mumkin: F.I.Sh ustuni, keyin 55 ta band
// (1..35, "36 (a)", "36(b)" ... "45(b)") uchun 0/1 javoblar va "O'quv markazi" ustuni.
// Barcha varaqlar ko'rib chiqiladi va eng ko'p javob ustuni bor varaq tanlanadi
// (masalan "Malumotlar"). Javob ustunlari topilmasa — tayyor BALL/Daraja ustunlari o'qiladi.
function parseUploadedResultsSheet(workbook) {
  const isItemHeader = (c) => /^\d+(\.0)?$/.test(c) || /^\d+\s*\(?\s*[a-zа-я]\s*\)?$/i.test(c) || /^savol\s*\d+/i.test(c) || /^s\s*\d+$/i.test(c);
  const isNameHeader = (c) => c.length <= 30 && /^(ism|f\.?\s*i\.?\s*o\.?|f\.?\s*i\.?\s*sh\.?|familiya|o.?quvchi)/i.test(c);
  let best = null;
  workbook.SheetNames.forEach((sheetName, si) => {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws["!ref"]) return;
    const full = XLSX.utils.decode_range(ws["!ref"]);
    const headRange = { s: full.s, e: { r: Math.min(full.e.r, full.s.r + 9), c: full.e.c } };
    const head = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", range: headRange });
    head.forEach((rowRaw, r) => {
      const row = (rowRaw || []).map(c => String(c ?? "").trim());
      const nameIdx = row.findIndex(isNameHeader);
      if (nameIdx < 0) return;
      const itemCols = [];
      row.forEach((c, ci) => { if (ci > nameIdx && isItemHeader(c)) itemCols.push(ci); });
      const score = itemCols.length * 10 + (sheetName === "Natijalar" ? 1 : 0) - si * 0.001;
      if (!best || score > best.score) best = { score, sheetName, headerRowIdx: r, row, nameIdx, itemCols };
    });
  });
  if (!best) return { rows: [] };
  const ws = workbook.Sheets[best.sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  const row = best.row;
  const cols = {
    name: best.nameIdx,
    group: row.findIndex(c => /guruh|o.?quv\s*markaz|markaz/i.test(c)),
    ball: row.findIndex(c => /^ball$/i.test(c) || /^(jami|umumiy)\s*ball$/i.test(c)),
    daraja: row.findIndex(c => /daraja|^baho$/i.test(c)),
    correct: row.findIndex(c => /^(jami\s*)?to.?g.?ri$/i.test(c)),
    total: row.findIndex(c => /jami\s*savol/i.test(c)),
    items: best.itemCols,
  };
  const numOrNull = (v) => (v === "" || v === null || v === undefined || isNaN(Number(v))) ? null : Number(v);
  const rows = [];
  for (let r = best.headerRowIdx + 1; r < aoa.length; r++) {
    const rr = aoa[r] || [];
    const name = String(rr[cols.name] ?? "").trim();
    if (!name) continue;
    const rec = { name, group: cols.group >= 0 ? String(rr[cols.group] ?? "").trim() : "" };
    if (cols.ball >= 0) { const v = numOrNull(rr[cols.ball]); if (v !== null) rec.ball = v; }
    if (cols.daraja >= 0 && rr[cols.daraja]) rec.daraja = String(rr[cols.daraja]).trim();
    if (cols.correct >= 0) { const v = numOrNull(rr[cols.correct]); if (v !== null) rec.correct = v; }
    if (cols.total >= 0) { const v = numOrNull(rr[cols.total]); if (v !== null) rec.total = v; }
    if (cols.items.length) {
      // Bo'sh katak — Excel'dagi kabi 0 (xato) hisoblanadi
      const items = cols.items.map(ci => (rr[ci] === 1 || rr[ci] === "1" || rr[ci] === true) ? 1 : 0);
      rec.correct = items.reduce((a, c) => a + c, 0);
      rec.total = items.length;
      if (items.length === RASCH_REQUIRED_ITEMS) rec.items = items;
    }
    rows.push(rec);
  }
  return { rows, sheetName: best.sheetName, itemCount: cols.items.length };
}
// Fayldagi qatorlar uchun Rash modelini SOF hisoblaydi (saytdagi foydalanuvchi yoki testga
// bog'liq emas). 55 ta javob ustuni bo'lsa — Excel shablonidagi algoritm bilan hisoblanadi;
// bo'lmasa, faylda tayyor BALL ustuni bo'lsa o'sha ishlatiladi.
function computeRaschFromFileRows(rows, settings = DEFAULT_RASCH_SETTINGS) {
  const withItems = rows.filter(r => Array.isArray(r.items) && r.items.length === RASCH_REQUIRED_ITEMS);
  if (withItems.length) {
    const calc = raschComputeCohort(withItems.map(r => r.items), settings);
    return withItems.map((r, i) => ({ ...r, ...calc[i], items: r.items }));
  }
  return rows.filter(r => typeof r.ball === "number" && isFinite(r.ball))
    .map(r => ({ ...r, daraja: r.daraja || raschGrade(r.ball, settings) }));
}
// Bitta testni saytda topshirganlar (results) VA yuklangan Excel fayldagi qatorlarni BITTA
// umumiy Rash populyatsiyasi sifatida birlashtirib hisoblaydi (bandlar qiyinligi, o'rtacha va
// standart chetlanish hammasi bo'yicha birga). Fayldagi ism saytdagi biror ro'yxatdan o'tgan
// o'quvchiga mos kelsa, o'sha bitta yozuv sifatida qo'shiladi (ikki marta hisoblanmaydi) va
// uning profiliga yoziladi. Faylda bor-u, saytda profili topilmagan o'quvchilar hisobga kiradi,
// lekin natija hech kimning profiliga yozilmaydi — alohida ro'yxat (fileOnly) qilib qaytariladi.
function importRaschFromRows(test, rows, settings = DEFAULT_RASCH_SETTINGS) {
  const all = db.get("results") || [];
  const users = db.get("users") || [];
  // Fayldagi ism ustuniga ko'pincha qo'shimcha narsalar yopishtirilgan bo'ladi:
  // "Safarov Abbos (A. Sultanov)", "Omonboyeva Charos Math@32" kabi. Shuning uchun avval
  // qavs ichidagi va harf bo'lmagan belgilarni tozalaymiz, so'ng birinchi ikkita so'zga qaraymiz.
  const clean = (s) => (s || "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^\p{L}\s'’-]/gu, " ")
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

  const combined = [];   // {name, group, vector|null, siteIdx, directBall, directDaraja, batchId}
  const siteIdxUsed = new Set();
  const unusable = [];   // na 55 ta javob, na tayyor ball bor qatorlar

  // 1) Fayldagi har bir qator — imkon bo'lsa saytdagi mos natijaga bog'lanadi
  rows.forEach(rec0 => {
    const rec = rec0.items || !rec0.itemsStr ? rec0 : { ...rec0, items: [...rec0.itemsStr].map(Number) };
    const hasItems = Array.isArray(rec.items) && rec.items.length === RASCH_REQUIRED_ITEMS;
    if (!hasItems && typeof rec.ball !== "number") { unusable.push(rec.name); return; }
    const phone = findPhone(rec.name);
    let siteIdx = -1;
    if (phone) siteIdx = all.findIndex((r, i) => r.testId === test.id && r.userPhone === phone && !siteIdxUsed.has(i));
    if (siteIdx >= 0) siteIdxUsed.add(siteIdx);
    combined.push({ name: rec.name, group: rec.group || "", source: rec.source || "Excel fayl", vector: hasItems ? rec.items : null, siteIdx, directBall: rec.ball, directDaraja: rec.daraja, batchId: rec.batchId ?? null });
  });

  // 2) Saytda shu testni topshirgan, lekin faylda yo'q o'quvchilar
  all.forEach((r, idx) => {
    if (r.testId !== test.id || siteIdxUsed.has(idx)) return;
    const u = users.find(x => x.phone === r.userPhone);
    combined.push({ name: null, siteName: u ? `${u.firstName} ${u.lastName}` : r.userPhone, group: u?.group || "", source: "Sayt", vector: resultItemVector(test, r), siteIdx: idx });
  });

  // 3) Javob qatori bor hammasi BITTA populyatsiya sifatida hisoblanadi
  const withVec = combined.filter(e => e.vector);
  const calc = raschComputeCohort(withVec.map(e => e.vector), settings);
  withVec.forEach((e, i) => { e.calc = calc[i]; });
  const calculatedAt = Date.now();

  let matched = 0;
  const fileOnly = [];
  const fileRows = [];
  const allRows = []; // umumiy hisobdagi HAR BIR o'quvchi (sayt + barcha fayllar) — admin uchun
  const updated = all.map(r => ({ ...r }));
  combined.forEach(e => {
    const c = e.calc || { correct: null, total: RASCH_REQUIRED_ITEMS, alg: null, geo: null, ball: e.directBall, foiz: null, bmba: null, daraja: e.directDaraja || raschGrade(e.directBall, settings), rank: null };
    if (e.siteIdx >= 0) {
      matched++;
      updated[e.siteIdx] = { ...updated[e.siteIdx], rasch: { ...c, calculatedAt, settings, source: e.name ? "site+upload" : "site" } };
    }
    allRows.push(slimRaschRow({ name: e.name || e.siteName, group: e.group, source: e.name ? (e.siteIdx >= 0 ? e.source + " (saytda ham bor)" : e.source) : "Sayt", ...c }));
    if (e.name) {
      const row = slimRaschRow({ name: e.name, group: e.group, ...c, matchedSite: e.siteIdx >= 0, batchId: e.batchId ?? null });
      fileRows.push(row);
      if (e.siteIdx < 0) fileOnly.push(row);
    }
  });
  db.set("results", updated);
  // Barcha ishtirokchilar bo'yicha har bir savolning to'g'ri javob ulushi (PDF grafigi uchun)
  const SH = new Array(RASCH_REQUIRED_ITEMS).fill(0);
  withVec.forEach(e => e.vector.forEach((v, i) => { SH[i] += v ? 1 : 0; }));
  const overallShare = withVec.length ? SH.map(x => x / withVec.length) : null;
  return { matched, fileOnly, fileRows, allRows, overallShare, vecCount: withVec.length, unusable, calculatedAt, combinedCount: combined.length };
}

// Saqlash uchun ixcham qator: faqat ko'rsatiladigan maydonlar, sonlar 4 xonagacha yaxlitlanadi
// (Firestore'da bitta hujjat 1MB dan oshmasligi uchun — minglab o'quvchi bo'lsa ham sig'adi).
function slimRaschRow(r) {
  const n = (v) => (v === null || v === undefined || !isFinite(v)) ? null : Math.round(v * 10000) / 10000;
  const o = { name: r.name, group: r.group || "", correct: r.correct ?? null, total: r.total ?? null,
    alg: n(r.alg), geo: n(r.geo), ball: n(r.ball), foiz: n(r.foiz), bmba: n(r.bmba), daraja: r.daraja || null, rank: r.rank ?? null };
  if (r.source) o.source = r.source;
  if (r.matchedSite !== undefined) o.matchedSite = r.matchedSite;
  if (r.batchId !== undefined && r.batchId !== null) o.batchId = r.batchId;
  return o;
}
// Fayldan o'qilgan qatorlarni saqlashga tayyorlaydi: 55 ta javob "1101..." qatori sifatida.
function packUploadRows(rows) {
  return rows.map(r => {
    const { items, ...rest } = r;
    return items ? { ...rest, itemsStr: items.join("") } : rest;
  });
}
// Yangi fayl yuklamasini saqlaydi. Bir hamkor markaz shu test uchun qayta yuklasa — eski fayli
// almashtiriladi (bir o'quvchi ikki marta hisoblanmasligi uchun). Admin yuklagan fayllar esa
// fayl nomi bo'yicha almashtiriladi.
function saveRaschUpload(batch) {
  const all = db.get("partnerUploads") || [];
  const same = (u) => u.testId === batch.testId && (batch.partnerId === "admin"
    ? (u.partnerId === "admin" && u.fileName === batch.fileName)
    : u.partnerId === batch.partnerId);
  const replaced = all.some(same);
  db.set("partnerUploads", [...all.filter(u => !same(u)), batch]);
  return replaced;
}
// Bitta test uchun UMUMIY Rash hisobi: saytda topshirganlar + shu testga yuklangan BARCHA
// Excel fayllar (hamkor markazlar va admin yuklagani) — hammasi bitta populyatsiya sifatida.
// Natija: saytdagi o'quvchilar profiliga, har bir markazning "Natijalarim" bo'limiga va
// admin uchun "raschResults" to'plamiga (umumiy reyting jadvali) yoziladi.
function calculateRaschCombined(test, settings = DEFAULT_RASCH_SETTINGS) {
  const allUploads = db.get("partnerUploads") || [];
  const partners = db.get("partners") || [];
  const batches = allUploads.filter(u => u.testId === test.id);
  const taggedRows = [];
  batches.forEach(b => {
    const src = b.partnerId === "admin" ? `Admin: ${b.fileName || "fayl"}` : (partners.find(p => p.id === b.partnerId)?.name || b.partnerName || "Hamkor markaz");
    (b.rawRows || b.rows || []).forEach(r => taggedRows.push({ ...r, batchId: b.id, source: src, group: r.group || (b.partnerId === "admin" ? "" : src) }));
  });
  const r = importRaschFromRows(test, taggedRows, settings);
  const ballsAll = r.allRows.map(x => x.ball).filter(v => typeof v === "number" && isFinite(v));
  const overallMean = ballsAll.length ? ballsAll.reduce((a, b) => a + b, 0) / ballsAll.length : null;
  const byBatch = {};
  r.fileRows.forEach(fr => { if (fr.batchId != null) (byBatch[fr.batchId] = byBatch[fr.batchId] || []).push(fr); });
  if (batches.length) {
    db.set("partnerUploads", allUploads.map(u => u.testId === test.id
      ? { ...u, status: "processed", processedAt: r.calculatedAt, rows: byBatch[u.id] || [], stats: { itemShare: batchItemShare(u.rawRows), overallShare: r.overallShare, overallMean, overallCount: r.allRows.length } }
      : u));
  }
  const summary = {
    id: test.id, testId: test.id, testName: test.name, calculatedAt: r.calculatedAt,
    count: r.allRows.length, siteCount: r.allRows.filter(x => x.source === "Sayt").length,
    files: batches.map(b => ({ id: b.id, partnerId: b.partnerId, source: b.partnerId === "admin" ? `Admin: ${b.fileName || "fayl"}` : (partners.find(p => p.id === b.partnerId)?.name || "Hamkor markaz"), rows: (b.rawRows || []).length, uploadedAt: b.uploadedAt })),
    rows: r.allRows,
    itemShare: r.overallShare, itemN: r.vecCount,
  };
  db.set("raschResults", [...(db.get("raschResults") || []).filter(x => x.id !== test.id), summary]);
  db.set("tests", (db.get("tests")||[]).map(t=>t.id===test.id?{...t,raschSettings:settings,raschCalculatedAt:r.calculatedAt}:t));
  return { ...r, partnerBatchesProcessed: batches.filter(b => b.partnerId !== "admin").length, partnerCentersCount: new Set(batches.filter(b => b.partnerId !== "admin").map(b=>b.partnerId)).size, filesCount: batches.length };
}

// Admin buyrug'i: bitta test bo'yicha umumiy hisobni yangilab, natijalarni BARCHA o'quv
// markazlariga BIR VAQTDA ochadi. Har bir markaz fayliga shu paytdagi natijalar nusxasi
// (publishedRows) yoziladi — keyin yangi fayl kelib qayta hisoblansa ham, admin yana e'lon
// qilmaguncha markazlar faqat e'lon qilingan natijani ko'radi.
function publishRaschToPartners(test, settings = DEFAULT_RASCH_SETTINGS) {
  const r = calculateRaschCombined(test, settings);
  const now = Date.now();
  let count = 0;
  const centers = new Set();
  db.set("partnerUploads", (db.get("partnerUploads") || []).map(u => {
    if (u.testId !== test.id || u.partnerId === "admin") return u;
    count++; centers.add(u.partnerId);
    return { ...u, publishedRows: u.rows || [], publishedStats: u.stats || null, publishedAt: now };
  }));
  // Saytda topshirgan o'quvchilar ham o'z Rash natijasini aynan shu buyruq bilan ko'radi
  let siteCount = 0;
  db.set("results", (db.get("results") || []).map(x => {
    if (x.testId !== test.id || !x.rasch) return x;
    siteCount++;
    return { ...x, raschPublished: { ...x.rasch, publishedAt: now } };
  }));
  db.set("tests", (db.get("tests") || []).map(t => t.id === test.id ? { ...t, raschPublishedAt: now } : t));
  return { ...r, publishedSite: siteCount, publishedFiles: count, publishedCenters: centers.size, publishedAt: now };
}


// ===== O'QUV MARKAZI UCHUN PDF HISOBOT (jsPDF) =====
// "Piramida o'quv markazi Jondor" namunasidagi shakl: 1) markaz o'quvchilari jadvali
// (Algebra, Geometriya, Ball, rangli Baho), 2) oxirgi sahifada markaz statistikasi, savollar
// bo'yicha to'g'ri javoblar ulushi grafigi, eng qiyin / ortda qolgan / ustun savollar va
// minnatdorchilik xati. PDF to'g'ridan-to'g'ri fayl qilib yaratiladi (chop etish oynasisiz).
let jsPdfLoading = null;
function loadJsPdf() {
  if (window.jspdf && window.jspdf.jsPDF) return Promise.resolve(window.jspdf.jsPDF);
  if (jsPdfLoading) return jsPdfLoading;
  jsPdfLoading = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    s.onload = () => resolve(window.jspdf.jsPDF);
    s.onerror = () => { jsPdfLoading = null; reject(new Error("PDF kutubxonasi yuklanmadi. Internet aloqasini tekshiring.")); };
    document.head.appendChild(s);
  });
  return jsPdfLoading;
}
// PDF uchun shrift: Liberation Sans (Helvetica bilan bir xil o'lchamli, SIL OFL litsenziyali),
// lotin + kirill harflari va №, «», ʻ belgilari bilan qisqartirilgan nusxa. Standart PDF shrifti
// kirill ismlarni va "№" belgisini chiqara olmaydi, shuning uchun shu shrift PDF ichiga joylanadi.
const PDF_FONT = "LiberationSans";
const PDF_FONT_DATA = {
  normal: "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AjkAuwI5ALgErABlBKwAZASsAGUEcwBUCB8AoQVWAAQFVgCoBccAaAXHAKgFVgCoBOMAqAY5AGcFxwCoAjkAvQQAACAFVgCoBHMAqAaqAKgFxwCoBjkAYQVWAKgGOQBhBccAqAVWAF0E4wAuBccAngVWAAkHjQAJBVYALgVWAC0E4wBBAjkAkgI5AAACOQAQA8EACgRz/+ECqgBqBHMAVwRzAIQEAABXBHMAVgRzAFcCOQAdBHMAVgRzAI4BxwCJAcf/zgQAAIoBxwCKBqoAiARzAIgEcwBWBHMAhARzAFYCqgCIBAAAOQI5AB8EcwCFBAAABwXH//0EAAAXBAAABQQAAFMCrAAiAhQAtwKsACIErABcAjkAAAKqAPIEcwCHBHMAOgRzAHEEc//+AhQAtwRzAHMCqgAtBeUAHwL2ABoEcwBTBKwAZAKqAFsF5QAfBGv/7wMzAHoEZABBAqoAKQKqABsCqgBIBJwAjARMAFACqgDzAqoAdwKqAFAC7AAbBHMAUwasADgGrAA4BqwASQTjAIMFVgAEBVYABAVWAAQFVgAEBVYABAVWAAQIAAAYBccAaAVWAKgFVgCoBVYAqAVWAKgCOQAJAjkAjgI5/9ICOQAHBccADgXHAKgGOQBhBjkAYQY5AGEGOQBhBjkAYQSsAI4GOQBHBccAngXHAJ4FxwCeBccAngVWAC0FVgCoBOMAjgRzAFcEcwBXBHMAVwRzAFcEcwBXBHMAVwcdAEIEAABXBHMAVwRzAFcEcwBXBHMAVwI5AAoCOQCHAjn/0wI5AAgEcwBWBHMAjARzAFYEcwBWBHMAVgRzAFYEcwBWBGQAQQTjACwEcwCLBHMAiwRzAIsEcwCLBAAABQRzAIoEAAAFBVYABARzAFcFVgAEBHMAVwVWAAQEcwBXBccAaAQAAFcFxwBoBAAAVwXHAGgEAABXBccAaAQAAFcFxwCoBOsAVgXHAA4EcwBWBVYAqARzAFcFVgCoBHMAVwVWAKgEcwBXBVYAqARzAFcFVgCoBHMAVwY5AGcEcwBWBjkAZwRzAFYGOQBnBHMAVgY5AGcEcwBWBccAqARzAI4FxwAOBHMACgI5/7gCOf+4AjkADAI5AAwCOf/SAjn/0gI5AFwBxwAfAjkAvQI5AMIF4QCqA40AiQQAACABx/+ZBVYAqAQAAIoEAACKBHMAqAHHAFsEcwCoAccAfgRzAKgCVQCKBHMAqAKsAIoEcwAUAccAEAXHAKgEcwCMBccAqARzAIwFxwCoBHMAjATV//4FyQClBHMAjAY5AGEEcwBWBjkAYQRzAFYGOQBhBHMAVggAAGEHjQBWBccAqAKqAIgFxwCoAqoAgQXHAKgCqgA4BVYAXQQAADkFVgBdBAAAOQVWAF0EAAA5BVYAXQQAADkE4wAuAjkAHwTjAC4DAAAfBOMALgI5AB8FxwCeBHMAhQXHAJ4EcwCFBccAngRzAIUFxwCeBHMAhQXHAJ4EcwCFBccAngRzAIUHjQAJBcf//QVWAC0EAAAFBVYALQTjAEEEAABTBOMAQQQAAFME4wBBBAAAUwHHAIoGBQBeA04AswVWAF0EAAA5BOMALgI5AB8EcwBXBFwARAHHAH8BxwB2AqoAAAKqAAACqgAzAqr/3QKqAJwCqgAzAqoAUAKq/+kCqgAWAAD+fwRoAKgFxwCoBVYAqAVXAKgG6wAuBFUAqAXAAGgFVgBdAjkAvQI5AAcEAAAgCHUAEggVAKgG1QAuBKkAqAXAAKgFFQA3BcAAqAVWAAQFQACoBVYAqARVAKgFawAPBVYAqAdjABwE1QBDBcAAqAXAAKgEqQCoBUAAEgaqAKgFxwCoBjkAYQXAAKYFVgCoBccAaATjAC4FFQA3BhUAdgVWAC4F6wCoBVUAoAdVAKgHgACoBlUALgcVAKgFQACoBcAAaQgVAKgFxwBgBHMAVwSVAHgEQACOAusAjgSrABQEcwBXBVoABwOrADEEeACOBHgAjgOAAIoEqwALBYAAjgRrAI4EcwBWBFUAjgRzAIQEAABXA6oAIwQAAAUGlQBWBAAAFwSVAI4EKwB6BmsAjgaVAI4FAAAtBcAAjgQrAI4EFQA3BgAAjgRVABEEcwBXBHMAVwRzAAoC6wCOBBUAVwQAADkBxwCJAjn/+AHH/84HQAALBoAAjgRzAAoDgACKBHgAjgQAAAUEawCOCrQAcgT+AA8GOQAVBOcAFQeZAL0FtQCIBVgABAQAAAYHLgC/BZAAhgahAHkFewB6CG0AvQbwAIgE1QBgA6sAIAZfAEMFggBEBjkAYQRzAFYGbQAJBQwABwZtAAkFDAAHCJgAYQcsAFUGqgA2BOYAKwmHAGwG0ABXCrQAcgT+AA8FxwBoBAAAUAQHAAsAAP6iAAD+vAAA/58AAP+f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",
  bold: "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BHMAPwRzAEsEcwBYBHMAQQRzAEcCqgDFAqoAwwSsAFYErABVBKwAVgTjAF4HzQB1BccAMwXHAIkFxwBUBccAiQVWAIkE4wCJBjkAVAXHAIkCOQCJBHMAHwXHAIkE4wCJBqoAiQXHAIkGOQBUBVYAiQY5AFQFxwCJBVYAOwTjABcFxwB7BVYADgeNAAIFVgASBVYAIwTjAD0CqgBzAjkAFQKqABkErAAtBHP/7AKqAEIEcwA8BOMAhwRzAFAE4wBUBHMAUAKqACME4wBUBOMAjwI5AI8COf/gBHMAjwI5AI8HHQCHBOMAhwTjAFAE4wCHBOMAVAMdAIcEcwBIAqoAGQTjAH8EcwAIBjn/+gRzAA4EcwAQBAAAUgMdACECPQCcAx0AKwSsAFECOQAAAqoAwgRzADMEcwAVBHMAOQRzAAgCPQCcBHMANQKqABAF5QAgAvYALQRzAFwErABUAqoAUAXlACAEa//vAzMAWgRkADECqgAzAqoALAKqAFcEnACGBHMARwKqAI0CqgBgAqoAUgLsAC0EcwBdBqwAXgasAF4GrABnBOMAcgXHADMFxwAzBccAMwXHADMFxwAzBccAMwgAAAQFxwBUBVYAiQVWAIkFVgCJBVYAiQI5/9gCOQBoAjn/qwI5/9cFxwAIBccAiQY5AFQGOQBUBjkAVAY5AFQGOQBUBKwAVgY5AFQFxwB7BccAewXHAHsFxwB7BVYAIwVWAIkE4wCPBHMANARzADQEcwA0BHMANARzADQEcwA0Bx0AQgRzAFAEcwBQBHMAUARzAFAEcwBQAjn/wAI5AHICOf+tAjn/2ATjAFAE4wCHBOMAUATjAFAE4wBQBOMAUATjAFAEZAAxBOMAAQTjAH8E4wB/BOMAfwTjAH8EcwAQBOMAjwRzABAFxwAzBHMAPAXHADMEcwA8BccAMwRzADwFxwBUBHMAUAXHAFQEcwBQBccAVARzAFAFxwBUBHMAUAXHAIkFwABUBccACATjAFQFVgCJBHMAUAVWAIkEcwBQBVYAiQRzAFAFVgCJBHMAUAVWAIkEcwBQBjkAVATjAFQGOQBUBOMAVAY5AFQE4wBUBjkAVATjAFQFxwCJBOMAjwXHAA4E4wAKAjn/mQI5/5kCOf/sAjn/7AI5/7wCOf+9AjkAWAI5AEUCOQCJAjkAkQZHAIkEcwCPBHMAHwI5/6YFxwCJ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AYkAGQGLADMBkgAZAZcAGQAJAWsAGQFy/+UBdf/NAXYAMwF3ADMBewAZAX4ATAGEABkBkv/lAAkBav/NAW//zQFy/+UBdf/NAXb/sgF3/7IBeP/NAXv/sgGD/80ACgFrAEwBdf/lAXYAMwF3ADMBeP/lAXsAMwGBABkBhAAZAYsAGQGV/+UAAQGEADMAAwGEABkBlf/lAZv/5QAJAWT/zQFo/+UBav/lAW//zQF3/80Bef+yAYP/5QGI/80Bj//lABYADf7lAA/+zQAb/+UAHP/lAWT/TAFo/38Bav/lAWv/5QFv/38BcP/lAXL/5QF1/80Bdv/lAXf/5QF4/+UBef+yAYP/zQGE/+UBiP9/AYn/zQGS/80Bo//lABIBZP/NAWj/5QFr/+UBb//NAXD/5QFy/80Bdv/NAXf/5QF5/7IBe//NAX7/5QGB/+UBhAAZAYX/5QGKABkBlf/lAZj/5QGb/+UAHgAN/zMAD/8ZABv/5QAc/+UAa//NAWT/fwFo/80BagAzAW//zQFy/80BeP+yAYP/5QGE/80Bhv+yAYn/mgGM/7IBjv+yAY//fwGQ/5gBkv9/AZP/5QGU/7IBlf9/AZf/sgGZ/7IBnf/lAZ//5QGg/+UBov/lAaP/zQAkAA3/GQAP/wAAG//NABz/zQBr/7IAe/+yAWT/ZgFo/5gBb/+yAXL/zQF4/7IBgf/lAYP/zQGF/80Bhv+yAYf/sgGI/38Bif9/AYr/5QGL/5oBjP+yAY3/sgGO/7IBj/9/AZD/mgGR/7IBkv9/AZP/sgGU/7IBlf9/AZn/zQGa/7IBnP+yAZ3/sgGi/7IBo/+aAAkBZP/NAWj/mAFv/7IBdv+yAXf/sgF4/+UBe//lAYP/ywGP/80ABwFr/+UBcv/LAXX/sgF4/7IBgf/NAZL/5QGX/+UAAwGEAEwBiQAZAZIAGQABAYQAGQACAYP/sgJc/zMADgFk/80BaP/lAWr/sgFr/80Bb/+yAXD/zQFy/+UBdf/NAXb/MQF5/5gBe/9mAYH/sgGD/7ICXP9mAAoBaP/NAWr/zQFr/+UBb/+yAXj/5QF5/80Bg//NAYj/5QGKABkBj//NAAsBZP/NAWj/zQFq/80Bb/+yAXX/5QF2/80Bef+yAXv/5QGI/80BigAZAY//zQACAYsAGQGb/80ADAGI/80Biv/lAYv/5QGP/80BkP/lAZX/5QGX/+UBmf/NAZv/zQGe/+UBof/lAaP/5QAOAYT/5QGF/+UBif/lAYr/5QGP/+UBkP/lAZL/5QGV/80Blv/lAZf/zQGY/+UBm/+yAZ7/zQGj/+UABgAN/0wAD/8zAYj/sgGP/80Bkv/lAZX/5QAEAYsAGQGV/+UBlwAZAZ7/5QAIAYX/5QGK/+UBkP/lAZX/5QGW/+UBl//lAZn/5QGb/80ABgGEABkBif/lAYsAGQGS/80Blf/lAZv/zQANAYT/5QGF/80Bif/lAYr/5QGL/+UBj//lAZD/zQGS/80Blf/NAZf/zQGY/+UBm/+yAZ7/zQAGAY8AGQGS/+UBlf/lAZj/5QGb/+UBof/lAAYBhf/lAYn/5QGS/+UBlf/lAZf/5QGb/80ACAGE/+UBhf/lAYv/5QGS/+UBlf/lAZf/5QGY/+UBof/lAAwBiP/lAYr/5QGL/+UBj//NAZD/5QGV/+UBlv/NAZf/5QGZ/80Bm//NAaH/5QGj/+UACwGI/80Biv/lAYv/5QGP/80BkP/lAZb/5QGX/+UBmf/lAZv/zQGh/+UBo//lAAUBhAAZAYX/5QGZ/+UBm//LAZ7/5QAIAA3/ZgAP/0wBiP/lAYoAGQGP/+UBkv/lAZX/5QGXADMAFAAN/2YAD/9MABv/5QAc/+UAewAzAYT/5QGI/80Bif/lAYoAGQGL/+UBj//lAZD/5QGS/80BlP/lAZX/zQGWABkBmP/lAZkAGQGh/+UBo//lAAoBhf/lAYj/5QGJ/+UBj//NAZD/5QGS/+UBlv/lAZf/5QGb/80Bo//lAAkBhf/lAYn/zQGL/+UBkv/NAZX/zQGXABkBmP/lAZv/zQGh/+UAAwGEABkBiwAZAZX/5QACAYQAGQGS/+UAAgGW/0wBm/9MAAoBhf/lAYj/5QGK/+UBi//lAY//zQGQ/+UBlf/lAZb/5QGZ/+UBo//lAAoBhf/lAYj/5QGK/+UBj//NAZD/5QGV/+UBlv/NAZn/zQGb/7IBof/lAAUADf9mAA//TAAb/+UAHP/lAGv/zQAGAA3/fwAP/2YAG//lABz/5QBr/80Ae//NAAECW/+0AAMAAf+PAFT/tAJc/7QABQFW/zMBX/8zAXb/TAF7/zEBfv9MAAQAAAABAAgAAQLyAAwAAQL8ADQAAgAGACIAOwAAAEIAWwAaAIAAlgA0AJgAtgBLALgBPwBqAUQBRQDyAPQCbgJiAnoCYgJiAoYB6gJuApgCJgJiAjICCAJuAnoChgJ6AoYCYgKeAm4ChgJ0AoYChgKMAoAChgJ6AmgCaAHwAlwCjAKYAlYCaAKYAfYCngJEAg4CgAKYAmgCpAKeAoACegKAAoACkgJuAm4CbgJuAm4CbgH8AnoCbgJiAmICbgKYApgCmAKYAmICbgJ6AnoCegJ6AnoCegJuAm4CbgJuAoYCMgICAoACgAKAAoACgAKAAggCIAJoAmgCaAJoApgCmAKYApgCngKeAkQCRAJEAkQCRAJEAp4CngKeAp4CFAIOAhQCbgKAAm4CgAJuAoACegJ6AnoCegJ6AnoCegJ6AmICgAJiAoACbgJoAm4CaAIaAmgCJgJcAhoCaAJ6AiACegIgAnoCIAJ6AiACbgKMAm4CjAKYApgCmAKYApgCmAKYAlYCmAKYAm4CgAImAlYCYgIsAmgCMgKYAjICVgIyApgCMgKYAp4CmAJuAp4CbgIyAm4CngI4Aj4CngJ6AkQCegJEAnoCRAJKAlAChgKYAoYCVgKGApgCYgJoAmICaAJiAlwCYgJoAp4CpAKeAqQCngKkAm4CngJuAp4CbgKeAm4CngJuAp4CbgKeAnQCegKGAoAChgKMApICjAKSAowCkgKYAp4CpAABA0gGBAABAcwGLAABA5gEsAABBRQGBAABAlgGLAABA1wGBAABAqgEsAABAjAGLAABAsYHbAABAlgEsAABAtAGBAABAkQGLAABAooGBAABAz4EsAABAvgGBAABAnYEsAABBEwGBAABA/wEsAABARgGLAABAkQEsAABArwGBAABAkQGBAABAtoGBAABA8oGBAABAyAGBAABAjAEsAABAqgGBAABAoAGBAABAfQEsAABARgGBAABAmwGBAABAVQFyAAGAgAAAQAIAAEADAAMAAEAFgA6AAIAAQHXAdoAAAAEAAAAEgAAABgAAAAeAAAAHgABAAAETAABAAAEsAABAAAEOAAEAAoACgAKAAoAAQAABpAAAQAAAAoARABGAAdERkxUACxib3BvADZjb3B0ADZjeXJsADZncmVrADZoZWJyADZsYXRuADYABAAAAAD//wAAAAAAAAAAAAA=",
  italic: "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A6oFzBAmAFUAABAHAncBpABLAAEAuAAABVwFgQAPAAABAyEHIQMjEyE3IRMhNyEHA1pYARke/ud9vn3+6R4BF1j+HB4Ehh4E5f48mv15AoeaAcScnAAAAQAZ/+wCgAUsACAAACUUMzI3BwYjIiY1ND8BIzczEyM3MzczBzMHIwMzByMHBgEVWio6E2NKVWEPKn0ZfTt9Gn9peC/IGsg7yBnJKAzTWA6FGGZUN0vcgwEwg/Lyg/7Qg9I8//8Amf/sBdEHBhImADYAABAHAn4BlAAA//8AVv/tBEoFvRImAFYAABAHAVEBDwAA//8Amf/sBdEGnBImADYAABAHAUwB+AFJ//8AVv/tBEoFUxImAFYAABAHAUwBCgAA//8Amf/sBdEG6hImADYAABAHAoAB5wAA//8AVv/tBEoF5hImAFYAABAHAU0BAAAA//8Amf/sBdEHwRImADYAABAHAU8B3QFO//8AVv/tBEoGcxImAFYAABAHAU8A9gAA//8Amf/sBdEG8RImADYAABAHAn8B9wAA//8AVv/tBKEF5BImAFYAABAHAVIBYQAA//8Amf5VBdEFgRAmADYAABAHAVABsgAA//8AVv5VBEoEOhImAFYAABAHAVABwAAA//8AsQAACC4G/hImADgAABAHAnsCXQAA//8AZgAABjUF0xImAFgAABAHAUoBpwAA//8A1QAABdEG/hImADoAABAHAnsBUAAA////jP5XBGcF0xImAFoAABAHAUoA2gAA//8A1QAABdEGshImADoAABAHAn0BQAAA////2AAABRwG8BImADsAABAHAnoBoAAA////9AAAA+QF5BImAFsAABAHAHQBNQAA////2AAABRwGphImADsAABAHAU4BYwDa////9AAAA+QFzBImAFsAABAHAU4A0QAA////2AAABRwG/hImADsAABAHAnwBCwAA////6QAAA+oF0xImAFv1ABAHAUsAyAAAAAEAIQAAArIFzAAOAAAzEz4BMzIXByYjIg4BBwMh6h2QhEYwGCspKDMgEOYEtJx8DIkIHTpN+2UAAgB6/+wF3QWWAAgAIQAAJTIANyEGFRQWATY1NCYjIgYHJzYkMyAAERQCBCMgABE0NwK/4gEuPfwpCdEDJwLPuJrpW6h3AUXUAQkBK8T+kvT+7/7UJX4BEf05N8nVAqwsDr7YeoM8s6r+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BHMAQwRzACIEcwBJBHMASQRzAEkEcwBJBHMASQRkAFIE4wAsBHMAVgRzAFYEcwBWBHMAVgQA/4wEc//NBAD/jAVW/5sEcwAuBVb/mwRzAC4FVv+bBHMALgXHAHEEAABDBccAcQQAAEMFxwBxBAAAQwXHAHEEAABDBccAPwUAAEUFxwAiBHMARQVWAD8EcwBFBVYAPwRzAEUFVgA/BHMARQVWAD8EcwBFBVYAPwRzAEUGOQBlBHMABAY5AGUEcwAEBjkAZQRzAAQGOQBlBHMABAXHAD8EcwAiBccAPgRzACICOQBRAcf/8gI5AFEBxwAiAjkAUQHHACICOf/AAcf/gwI5AFECOQBZBd4AUQONACEEAP/7Acf/GwVWAD8EAAAiBAAAIgRzAD8BxwAhBHMAPwHH/7wEcwA/AkAAIQRzAD8DMwAhBHMAAgHHAAIFxwA/BHMAJgXHAD8EcwAmBccAPwRzACYE6wAzBckAPwRzACIGOQBvBHMAQwY5AG8EcwBDBjkAbwRzAEMIAABlB40ARQXHAD8CqgAiBccAPwKq/8sFxwA/AqoAIgVWADoEAAAFBVYAOgQAAAUFVgA6BAAABQVWADoEAAAFBOMAuAI5AA4E4wC4AtUAXQTjALgCOQAZBccAmQRzAFYFxwCZBHMAVgXHAJkEcwBWBccAmQRzAFYFxwCZBHMAVgXHAJkEcwBWB40AsQXHAGYFVgDVBAD/jAVWANUE4//YBAD/9ATj/9gEAP/0BOP/2AQA/+kBxwAhBjkAegNOAMcFVgA6BAAABQTjALgCOQA7BHMAQQQA/5gBxwDgAccA1wKqAE0CqgCIAqoAjgKqAIICqgFWAqoAyQKqAEcCqgA1AqoAFwAA/vwEkAA/BcwAPwVWAD8FVgA/Br4AuARaAD8FqQBlBVYAOgI5AFECOQBRBAD/+wiH/5MH3AA/BsEAuAS3AD8FuQA/BR3//QXKAD8FVv+bBTUAPgVWAD8EWgA/BaL/QwVWAD8HVv+tBOkAEwW5AD8FuQA/BLcAPwV8/5MGqgA/BccAPwY5AG8FzAA/BVYAPwXHAHEE4wC4BR3//QZcAHoFVv/ZBdEAPwVjAM4HXQA/B2IAPwZxALgHFwA+BTUAPgWNABQILgA/BXT/pQRzAC4EgABpBC0AJwPyACwEbABDBHMARQWB/7YDuP/zBHMAVgRzAFYDxwAiBIP/ngV8ACIEZwAiBHMAQwRnACIEc//NBAAAQwaqACIEAP+MBq8ARQQA/64EkwBZBCQAiQakAFYGzwBWBPgAQgXjACYENQAmA+8ACQYEACIERv/QBHMARQRzAEUEcwAiAwYAIQP3AEMEAAAFAccAIQHHACsBx/8bB0j/ngafACIEcwAiA8cAIgRzAFYEAP+MBHMAVgpqAG8EwgBCBhQAZwY7ACIHOQA/BY4AIgVY/50EAP+dBwYAPwVyACIGoQAUBXsADQgxAD8G2QAjBOn/7QO4/5wGpQDwBdIApgY5AG8EcwBCBhQAsQS4AHAGFACxBLgAcAjeAG8HWQBCBp8AbwTMAEIKQgByBwIAZgpqAG8EwgBCBccAcQQAAEMEcwA6AAD/LQAA/zgAAAAOAAAAHwAA/mwAAPwlAAD8ZgW5AD8EcwBWBTUAPgQ1ACYFVgA/BHP/zQPwAD8CtwAiBFoAJgMGAA0FKAA/BDQAIgdW/60Fgf+oBOkAOwO4//MEtwA/A8cAIgS3AD8DxwAiBLcAPwPHACIF8QCrBIoAdQXHAD8EZwAiByEAPwWbACIIswA/BrcAIgXHAG8EAABCBccAcQQAAEME4wC4BqoAIgRzALwEAABwBHMAQwQA/90FVv/ZBAD/rgcuALgFKQBYBWMAzgQkAIkFYwDOBCQAiQVjAD8EcwAiBs0AHAUU//UGzQAcBRT/9QI5AFEHVv+tBYH/tgUEAEAEKAAdBXz/kwSD/54FxwA/BGcAGQXHAD8EZwAiBWMAzgQkAIkGqgA/BXwAIgHHACEFVv+bBHMALgVW/5sEcwAuCAD/rwcdACAFVgA/BHMARQY5AHoEcwBBBjkAegRzAEEHVv+tBYH/tgTpABMDuP/zBOkAEwQA/5gFuQA/BHMAVgW5AD8EcwBWBjkAaARzAEMGOQBvBHMAQwY5AG8EcwBDBY0AFAPvAAkFHf/9BAD/jAUd//0EAP+MBR3//QQA/4wFYwDOBCQAiQRaAD8DBgAjBxcAPgXjACYEWv+pAwb/WwVW/9kEAP+uBVb/2QQA/64CqgBpBHP/8wRz//MIAP/zCAD/8wNOAMcEa/9gAccAxwHHAMABx//6AccA/QKqAJQCqgCJAqr/uQKqALkEcwDYBHMARALNAGEIAAC6CAAAMwGAALYC1QC2AtUAYwKqAFYCqgAQAVb90wN3ACcIqgBU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",
};
function registerPdfFonts(doc) {
  Object.entries(PDF_FONT_DATA).forEach(([style, b64]) => {
    const file = `${PDF_FONT}-${style}.ttf`;
    doc.addFileToVFS(file, b64);
    doc.addFont(file, PDF_FONT, style);
  });
  doc.setFont(PDF_FONT, "normal");
}
// Hisobotdagi standart rasmlar (namunadagi logotip, kitoblar reklamasi va ijtimoiy tarmoq belgilari).
// Admin "PDF hisobot sozlamalari"da o'z rasmini yuklasa, o'shasi ishlatiladi.
const PDF_DEFAULT_IMAGES = {
  logo: "data:image/jpeg;base64,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",
  ad: "data:image/jpeg;base64,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",
  tg: "data:image/jpeg;base64,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",
  ig: "data:image/jpeg;base64,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",
  yt: "data:image/jpeg;base64,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",
};

// Hisobotdagi tashkilotchi ma'lumotlari (admin "🏆 Umumiy natijalar" → "PDF hisobot sozlamalari"da o'zgartiradi).
// logo / adImage: undefined — standart rasm, "" — rasm qo'yilmaydi, data URL — admin yuklagan rasm.
const DEFAULT_PDF_SETTINGS = {
  id: "pdf",
  orgName: "Rustambek oqiw orayi",
  telegram: "@Rustambek_oqiw_orayi",
  phone: "+99894 123 51 51",
  subtitle: "Milliy sertifikat formatida",
  adTitle: "Bizning kitoblarimiz",
  socialTelegram: "@rustambek_oqiw_orayi",
  socialInstagram: "@rustambek_oqiw_orayi",
  socialYoutube: "@rustambek_oqiw_orayi",
};
function getPdfSettings() {
  const s = (db.get("appSettings") || []).find(x => x.id === "pdf");
  const cfg = { ...DEFAULT_PDF_SETTINGS, ...(s || {}) };
  // Eski sozlamada sarlavha ostiga "· o'quv markazi hisoboti" qo'shilgan bo'lsa — endi u avtomatik qo'shiladi
  cfg.subtitle = String(cfg.subtitle || "").replace(/\s*·\s*(o'quv markazi hisoboti|barcha ishtirokchilar)\s*$/i, "");
  cfg.logoSrc = cfg.logo === undefined ? PDF_DEFAULT_IMAGES.logo : (cfg.logo || null);
  cfg.adSrc = cfg.adImage === undefined ? PDF_DEFAULT_IMAGES.ad : (cfg.adImage || null);
  return cfg;
}
// Band nomlari: 1..35, 36(a), 36(b) ... 45(b)
function raschItemLabel(i) {
  if (i < 35) return String(i + 1);
  const k = i - 35;
  return `${36 + Math.floor(k / 2)}(${k % 2 === 0 ? "a" : "b"})`;
}
// Markaz fayli bo'yicha savollar statistikasi (hisob vaqtida saqlanadi)
function batchItemShare(rawRows) {
  const vecs = (rawRows || []).map(r => r.itemsStr ? [...r.itemsStr].map(Number) : r.items).filter(v => v && v.length === RASCH_REQUIRED_ITEMS);
  if (!vecs.length) return null;
  const S = new Array(RASCH_REQUIRED_ITEMS).fill(0);
  vecs.forEach(v => v.forEach((x, i) => { S[i] += x ? 1 : 0; }));
  return S.map(x => x / vecs.length);
}
const PDF_GRADE_COLORS = { "A+": [30, 110, 50], "A": [76, 165, 80], "B+": [46, 117, 182], "B": [111, 168, 220], "C+": [237, 125, 49], "C": [214, 160, 30], "NC": [192, 57, 43] };
const pdfNum = (v) => (v === null || v === undefined || !isFinite(v)) ? "-" : String(+v.toFixed(2));
const PDF_C = { BLUE: [31, 78, 121], GREY: [120, 128, 140], LIGHT: [241, 244, 248], LINE: [205, 212, 222], TEXT: [30, 41, 59] };
const PDF_W = 210, PDF_M = 14, PDF_CW = PDF_W - 2 * PDF_M;
// Rasmning haqiqiy o'lchami (reklama rasmini buzmasdan joylash uchun)
function pdfImageSize(src) {
  return new Promise(resolve => {
    if (!src || typeof Image === "undefined") { resolve(null); return; }
    const im = new Image();
    im.onload = () => resolve({ w: im.naturalWidth, h: im.naturalHeight });
    im.onerror = () => resolve(null);
    im.src = src;
  });
}
function pdfImgFormat(src) { return /^data:image\/png/i.test(src || "") ? "PNG" : "JPEG"; }

// ---- Umumiy qismlar
function pdfFirstPageHeader(doc, cfg, { title, subtitle, heading, subheading }) {
  const { BLUE, TEXT } = PDF_C, M = PDF_M, W = PDF_W;
  doc.setFont(PDF_FONT, "bold"); doc.setFontSize(9);
  doc.setTextColor(192, 0, 0); doc.text(`Telegram kanalimiz: ${cfg.telegram}`, M, 11);
  doc.setTextColor(46, 125, 50); doc.text(cfg.phone, W - M, 11, { align: "right" });
  let bx = M;
  if (cfg.logoSrc) {
    try { doc.addImage(cfg.logoSrc, pdfImgFormat(cfg.logoSrc), M, 15, 18, 18); bx = M + 20; } catch { bx = M; }
  }
  const bw = W - M - bx;
  doc.setFillColor(...BLUE); doc.rect(bx, 15, bw, 18, "F");
  doc.setTextColor(255, 255, 255); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(14); doc.text(title, bx + bw / 2, 23, { align: "center" });
  doc.setFont(PDF_FONT, "italic"); doc.setFontSize(8); doc.text(subtitle, bx + bw / 2, 29.5, { align: "center" });
  doc.setTextColor(...BLUE); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(16); doc.text(heading, M, 44);
  doc.setFontSize(10); doc.setTextColor(...TEXT); doc.text(subheading, M, 49.5);
  doc.setDrawColor(...BLUE); doc.setLineWidth(0.5); doc.line(M, 51.5, W - M, 51.5);
}
function pdfResultsTable(doc, rows, startY) {
  const { BLUE, LINE } = PDF_C, M = PDF_M, CW = PDF_CW;
  // Ustunlar Excel shablonidagi "Hisobot" varag'i tartibida: ... Ball, Baho, Foiz, BMBA
  const cols = [
    { k: "n", h: "№", w: 10, a: "center" }, { k: "name", h: "Ism Familiya", w: 64, a: "left" },
    { k: "alg", h: "Algebra", w: 19, a: "center" }, { k: "geo", h: "Geometriya", w: 21, a: "center" },
    { k: "ball", h: "Ball", w: 18, a: "center" }, { k: "baho", h: "Baho", w: 14, a: "center" },
    { k: "foiz", h: "Foiz", w: 18, a: "center" }, { k: "bmba", h: "BMBA", w: 18, a: "center" },
  ];
  const pct = (v) => (v === null || v === undefined || !isFinite(v)) ? "-" : `${+v.toFixed(1)}%`;
  const RH = 7.2, HH = 8, BOTTOM = 278;
  const drawHeader = (y) => {
    doc.setFillColor(...BLUE); doc.rect(M, y, CW, HH, "F");
    doc.setTextColor(255, 255, 255); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(8.5);
    let x = M;
    cols.forEach(c => { doc.text(c.h, x + c.w / 2, y + 5.3, { align: "center" }); x += c.w; });
    return y + HH;
  };
  let y = drawHeader(startY);
  rows.forEach((r, i) => {
    if (y + RH > BOTTOM) { doc.addPage(); y = drawHeader(16); }
    if (i % 2 === 1) { doc.setFillColor(248, 250, 252); doc.rect(M, y, CW, RH, "F"); }
    doc.setDrawColor(...LINE); doc.setLineWidth(0.2); doc.rect(M, y, CW, RH);
    let x = M;
    const vals = { n: String(i + 1), name: r.name || "", alg: pdfNum(r.alg), geo: pdfNum(r.geo), ball: pdfNum(r.ball), baho: r.daraja || "", foiz: pct(r.foiz), bmba: r.bmba != null && isFinite(r.bmba) ? String(+r.bmba.toFixed(1)) : "-" };
    cols.forEach((c, ci) => {
      if (ci > 0) doc.line(x, y, x, y + RH);
      if (c.k === "baho" && r.daraja && PDF_GRADE_COLORS[r.daraja]) {
        doc.setFillColor(...PDF_GRADE_COLORS[r.daraja]); doc.rect(x + 0.2, y + 0.2, c.w - 0.4, RH - 0.4, "F");
        doc.setTextColor(255, 255, 255); doc.setFont(PDF_FONT, "bold");
      } else {
        if (c.k === "n") doc.setTextColor(100, 116, 139); else doc.setTextColor(...PDF_C.TEXT);
        doc.setFont(PDF_FONT, c.k === "ball" ? "bold" : "normal");
      }
      doc.setFontSize(8.5);
      let t = vals[c.k];
      if (c.k === "name") t = doc.splitTextToSize(t, c.w - 4)[0] || "";
      if (c.a === "left") doc.text(t, x + 2.5, y + 4.8);
      else doc.text(t, x + c.w / 2, y + 4.8, { align: "center" });
      x += c.w;
    });
    y += RH;
  });
}
function pdfSectionTitle(doc, t, yy) {
  const { BLUE } = PDF_C;
  doc.setTextColor(...BLUE); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(12); doc.text(t, PDF_M, yy);
  doc.setDrawColor(...BLUE); doc.setLineWidth(0.5); doc.line(PDF_M, yy + 2, PDF_W - PDF_M, yy + 2);
}
function pdfStatCards(doc, cards, y) {
  const { LIGHT, BLUE, GREY, TEXT } = PDF_C;
  const cw = (PDF_CW - 3 * 4) / 4;
  cards.forEach(([l, v, sub, subColor], k) => {
    const x = PDF_M + k * (cw + 4);
    doc.setFillColor(...LIGHT); doc.rect(x, y, cw, 21, "F");
    doc.setFillColor(...BLUE); doc.rect(x, y, 1.1, 21, "F");
    doc.setTextColor(...GREY); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(6.5); doc.text(l, x + 4, y + 5.5);
    doc.setTextColor(...TEXT); doc.setFontSize(15); doc.text(v, x + 4, y + 13.5);
    doc.setTextColor(...(subColor || GREY)); doc.setFont(PDF_FONT, "italic"); doc.setFontSize(6.3);
    doc.text(doc.splitTextToSize(sub || "", cw - 6)[0] || "", x + 4, y + 18.3);
  });
}
// Savollar bo'yicha to'g'ri javoblar ulushi grafigi; overall berilsa — qora chiziq bilan
function pdfItemChart(doc, share, overall, yy) {
  const { GREY } = PDF_C, M = PDF_M, W = PDF_W;
  pdfSectionTitle(doc, "Savollar bo'yicha to'g'ri javoblar ulushi", yy);
  const cx0 = M + 8, cx1 = W - M, cy0 = yy + 8, cy1 = yy + 54;
  doc.setFont(PDF_FONT, "normal"); doc.setFontSize(6); doc.setTextColor(...GREY);
  [0, 25, 50, 75, 100].forEach(p => {
    const gy = cy1 - (cy1 - cy0) * p / 100;
    doc.setDrawColor(228, 232, 238); doc.setLineWidth(0.15); doc.line(cx0, gy, cx1, gy);
    doc.text(`${p}%`, cx0 - 1.5, gy + 1, { align: "right" });
  });
  const slot = (cx1 - cx0) / RASCH_REQUIRED_ITEMS, bw = slot * 0.72;
  share.forEach((p, i) => {
    const color = p >= 0.7 ? [46, 125, 50] : p >= 0.4 ? [47, 116, 181] : p >= 0.15 ? [237, 125, 49] : [192, 0, 0];
    const bx = cx0 + i * slot + (slot - bw) / 2, h = Math.max(0.3, (cy1 - cy0) * p);
    doc.setFillColor(...color); doc.rect(bx, cy1 - h, bw, h, "F");
    if (overall && overall[i] != null) {
      const oy = cy1 - (cy1 - cy0) * overall[i];
      doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.5); doc.line(bx - 0.4, oy, bx + bw + 0.4, oy);
    }
    doc.setTextColor(...GREY); doc.setFontSize(4.3);
    const lb = raschItemLabel(i);
    if (i < 35) doc.text(lb, bx + bw / 2, cy1 + 3, { align: "center" });
    else { doc.text(lb.slice(0, 2), bx + bw / 2, cy1 + 3, { align: "center" }); doc.text(lb.slice(2), bx + bw / 2, cy1 + 5.2, { align: "center" }); }
  });
  let lx = M; const ly = cy1 + 10;
  [[[46, 125, 50], "70% va yuqori"], [[47, 116, 181], "40–70%"], [[237, 125, 49], "15–40%"], [[192, 0, 0], "15% dan past"]].forEach(([c, t]) => {
    doc.setFillColor(...c); doc.rect(lx, ly - 2.4, 3, 3, "F");
    doc.setTextColor(...GREY); doc.setFont(PDF_FONT, "normal"); doc.setFontSize(6.5); doc.text(t, lx + 4.2, ly);
    lx += doc.getTextWidth(t) + 10;
  });
  if (overall) {
    doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.5); doc.line(lx, ly - 1, lx + 5, ly - 1);
    doc.text("qora chiziq — barcha ishtirokchilar", lx + 6.5, ly);
  }
  return ly;
}
function pdfItemList(doc, x, y, colW, head, color, list, rightText, pctPos = 0.34, gap = 5.6) {
  const { LINE, GREY, TEXT } = PDF_C;
  doc.setTextColor(...color); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(8.5); doc.text(head, x, y);
  doc.setDrawColor(...LINE); doc.setLineWidth(0.3); doc.line(x, y + 1.8, x + colW, y + 1.8);
  list.forEach((it, k) => {
    const ry = y + 7 + k * gap;
    doc.setTextColor(...TEXT); doc.setFont(PDF_FONT, "normal"); doc.setFontSize(7.5); doc.text(`${raschItemLabel(it.i)}-savol`, x, ry);
    doc.setFont(PDF_FONT, "bold"); doc.text(`${Math.round(it.p * 100)}%`, x + colW * pctPos, ry);
    doc.setTextColor(...GREY); doc.setFont(PDF_FONT, "normal"); doc.setFontSize(6.5); doc.text(rightText(it), x + colW, ry, { align: "right" });
  });
  if (!list.length) { doc.setTextColor(...GREY); doc.setFont(PDF_FONT, "italic"); doc.setFontSize(7); doc.text("—", x, y + 7); }
  return y + 7 + Math.max(1, list.length) * gap;
}
function pdfFooters(doc, footerLeft) {
  const { LINE, GREY } = PDF_C;
  const n = doc.getNumberOfPages();
  for (let p = 1; p <= n; p++) {
    doc.setPage(p);
    doc.setDrawColor(...LINE); doc.setLineWidth(0.2); doc.line(PDF_M, 285, PDF_W - PDF_M, 285);
    doc.setFont(PDF_FONT, "italic"); doc.setFontSize(7); doc.setTextColor(...GREY);
    doc.text(footerLeft, PDF_M, 289.5);
    doc.setFont(PDF_FONT, "normal"); doc.text(`${p} / ${n}`, PDF_W - PDF_M, 289.5, { align: "right" });
  }
}
function pdfTitle(testName) { return `${String(testName || "").toUpperCase()}${/test/i.test(testName || "") ? "" : " TEST"} NATIJALARI`; }
function pdfFinish(doc, name) {
  const blob = doc.output("blob");
  return { url: URL.createObjectURL(blob), filename: name, blob };
}
const pdfSafeName = (s) => String(s || "").replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, "_");
const pdfSortRows = (rows) => [...rows].sort((a, b) => (b.ball ?? -999) - (a.ball ?? -999));
const pdfMedian = (arr) => { const s = [...arr].sort((a, b) => a - b); return s.length ? (s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2) : null; };

// ===== 1) O'quv markazi hisoboti =====
async function buildCenterReportPdf({ centerName, testName, rows, stats }) {
  const JsPDF = await loadJsPdf();
  const cfg = getPdfSettings();
  const doc = new JsPDF({ unit: "mm", format: "a4" });
  registerPdfFonts(doc);
  const { BLUE, LIGHT, TEXT } = PDF_C, M = PDF_M, W = PDF_W, CW = PDF_CW;
  const sorted = pdfSortRows(rows);
  pdfFirstPageHeader(doc, cfg, { title: pdfTitle(testName), subtitle: `${cfg.subtitle} · o'quv markazi hisoboti`, heading: centerName, subheading: "O'quvchilar natijalari" });
  pdfResultsTable(doc, sorted, 55);

  doc.addPage();
  pdfSectionTitle(doc, "Markaz statistikasi", 18);
  const balls = sorted.map(r => r.ball).filter(v => typeof v === "number" && isFinite(v));
  const mean = balls.length ? balls.reduce((a, b) => a + b, 0) / balls.length : null;
  const top = sorted.find(r => typeof r.ball === "number");
  const diff = (mean != null && stats?.overallMean != null) ? mean - stats.overallMean : null;
  pdfStatCards(doc, [
    ["O'QUVCHILAR SONI", String(rows.length), stats?.overallCount ? `barcha ishtirokchilar: ${stats.overallCount}` : "", null],
    ["O'RTACHA BALL", pdfNum(mean), diff != null ? `umumiy o'rtachadan ${diff >= 0 ? "+" : ""}${diff.toFixed(1)}` : "", diff != null ? (diff >= 0 ? [46, 125, 50] : [192, 0, 0]) : null],
    ["ENG YUQORI BALL", pdfNum(top?.ball), top?.name || "", null],
    ["MEDIANA", pdfNum(pdfMedian(balls)), "o'quvchilarning yarmi yuqorida", null],
  ], 25);

  let yy = 58;
  const share = stats?.itemShare, overall = stats?.overallShare;
  if (share && share.length === RASCH_REQUIRED_ITEMS) {
    const ly = pdfItemChart(doc, share, overall, yy);
    yy = ly + 12;
    const items = share.map((p, i) => ({ i, p, o: overall ? overall[i] : null }));
    const pct = (v) => `${Math.round(v * 100)}%`;
    const withO = items.filter(x => x.o != null).map(x => ({ ...x, d: x.p - x.o }));
    const colW = (CW - 8) / 3;
    pdfItemList(doc, M, yy, colW, "Eng qiyin savollar", TEXT, [...items].sort((a, b) => a.p - b.p).slice(0, 6), it => it.o != null ? `umumiy ${pct(it.o)}` : "");
    pdfItemList(doc, M + colW + 4, yy, colW, "Umumiydan ortda qolgan", [192, 0, 0], withO.filter(x => x.d < 0).sort((a, b) => a.d - b.d).slice(0, 6), it => `${Math.round(it.d * 100)} (umumiy ${pct(it.o)})`);
    pdfItemList(doc, M + 2 * (colW + 4), yy, colW, "Umumiydan ustun", [46, 125, 50], withO.filter(x => x.d > 0).sort((a, b) => b.d - a.d).slice(0, 6), it => `+${Math.round(it.d * 100)} (umumiy ${pct(it.o)})`);
    yy = yy + 7 + 6 * 5.6 + 10;
  }

  // Minnatdorchilik xati
  const body = "Mock testimizda faol ishtirok etganingiz va biz bilan hamkorlik qilganingiz uchun samimiy minnatdorchilik bildiramiz. O'quvchilaringizning natijalari ustozlarning mashaqqatli mehnati, sabr-toqati va fidoyiligining samarasidir. Ushbu tahlil keyingi tayyorgarlikda kuchli tomonlarni mustahkamlash va bo'shliqlarni to'ldirishda yordam beradi, degan umiddamiz. Barcha o'quvchilaringizga Milliy sertifikat imtihonida yuqori natijalar, jamoangizga esa yangi muvaffaqiyatlar tilaymiz!";
  doc.setFont(PDF_FONT, "normal"); doc.setFontSize(9);
  const lines = doc.splitTextToSize(body, CW - 16);
  const boxH = 24 + lines.length * 4.6 + 12;
  const by = Math.min(yy, 280 - boxH);
  doc.setFillColor(...LIGHT); doc.rect(M, by, CW, boxH, "F");
  doc.setFillColor(...BLUE); doc.rect(M, by, 1.5, boxH, "F");
  doc.setTextColor(...BLUE); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(11); doc.text(`Hurmatli «${centerName}» jamoasi!`, M + 8, by + 10);
  doc.setTextColor(51, 65, 85); doc.setFont(PDF_FONT, "normal"); doc.setFontSize(9); doc.text(lines, M + 8, by + 18, { lineHeightFactor: 1.45 });
  const sy = by + 18 + lines.length * 4.6 + 5;
  doc.setFont(PDF_FONT, "italic"); doc.setTextColor(...TEXT); doc.text(`Hurmat bilan, ${cfg.orgName} jamoasi`, W - M - 8, sy, { align: "right" });
  doc.setFont(PDF_FONT, "normal"); doc.setFontSize(7.5); doc.setTextColor(...PDF_C.GREY);
  doc.text(`Telegram kanalimiz: ${cfg.telegram}   ·   ${cfg.phone}`, W - M - 8, sy + 6, { align: "right" });

  pdfFooters(doc, `${centerName} · ${testName}`);
  return pdfFinish(doc, `${pdfSafeName(centerName)}_${pdfSafeName(testName)}.pdf`);
}

// ===== 2) Admin: barcha ishtirokchilar bo'yicha umumiy hisobot =====
async function buildOverallReportPdf({ testName, rows, itemShare, itemN }) {
  const JsPDF = await loadJsPdf();
  const cfg = getPdfSettings();
  const adSize = cfg.adSrc ? await pdfImageSize(cfg.adSrc) : null;
  const doc = new JsPDF({ unit: "mm", format: "a4" });
  registerPdfFonts(doc);
  const { TEXT, GREY } = PDF_C, M = PDF_M, W = PDF_W, CW = PDF_CW;
  const sorted = pdfSortRows(rows);
  pdfFirstPageHeader(doc, cfg, { title: pdfTitle(testName), subtitle: `${cfg.subtitle} · barcha ishtirokchilar`, heading: "Umumiy natijalar", subheading: "Ishtirokchilar natijalari" });
  pdfResultsTable(doc, sorted, 55);

  doc.addPage();
  pdfSectionTitle(doc, "Umumiy statistika", 18);
  const withBall = sorted.filter(r => typeof r.ball === "number" && isFinite(r.ball));
  const avg = (k) => { const v = withBall.map(r => r[k]).filter(x => typeof x === "number" && isFinite(x)); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  const top = withBall[0];
  const aCount = withBall.filter(r => r.daraja === "A+" || r.daraja === "A").length;
  pdfStatCards(doc, [
    ["ISHTIROKCHILAR", String(rows.length), "", null],
    ["O'RTACHA BALL", pdfNum(avg("ball")), (avg("alg") != null && avg("geo") != null) ? `algebra ${avg("alg").toFixed(1)} · geometriya ${avg("geo").toFixed(1)}` : "", null],
    ["ENG YUQORI BALL", pdfNum(top?.ball), top?.name || "", null],
    ["A+ VA A", `${aCount} ta`, rows.length ? `ishtirokchilarning ${(aCount / rows.length * 100).toFixed(1)}%` : "", null],
  ], 25);

  let yy = 58;
  if (itemShare && itemShare.length === RASCH_REQUIRED_ITEMS) {
    const ly = pdfItemChart(doc, itemShare, null, yy);
    yy = ly + 11;
    const items = itemShare.map((p, i) => ({ i, p }));
    const colW = (CW - 8) / 2;
    const solved = (it) => itemN ? `${Math.round(it.p * itemN)} ta ishtirokchi to'g'ri yechgan` : "";
    pdfItemList(doc, M, yy, colW, "Eng qiyin savollar", [192, 0, 0], [...items].sort((a, b) => a.p - b.p).slice(0, 8), solved, 0.3, 5.2);
    pdfItemList(doc, M + colW + 8, yy, colW, "Eng oson savollar", [46, 125, 50], [...items].sort((a, b) => b.p - a.p).slice(0, 8), solved, 0.3, 5.2);
    yy = yy + 7 + 8 * 5.2 + 8;
  }

  // Reklama bloki (kitoblar) va ijtimoiy tarmoqlar
  const socials = [["tg", "Telegram", cfg.socialTelegram], ["ig", "Instagram", cfg.socialInstagram], ["yt", "YouTube", cfg.socialYoutube]].filter(x => (x[2] || "").trim());
  const socialH = socials.length * 7;
  if (cfg.adSrc || socials.length) {
    let imgW = 0, imgH = 0;
    if (cfg.adSrc) {
      const ratio = adSize ? adSize.h / adSize.w : 0.63;
      imgW = 110; imgH = imgW * ratio;
    }
    const needed = 8 + imgH + 4 + socialH;
    if (yy + needed > 281) {
      // joy yetmasa — rasm shu sahifaga sig'adigan qilib kichraytiriladi, juda kichik bo'lib qolsa — yangi sahifa
      const room = 281 - yy - 8 - 4 - socialH;
      if (cfg.adSrc && room >= 35) { imgH = room; imgW = imgH / (adSize ? adSize.h / adSize.w : 0.63); }
      else { doc.addPage(); yy = 18; }
    }
    pdfSectionTitle(doc, cfg.adTitle || "Bizning kitoblarimiz", yy);
    let cy = yy + 5;
    if (cfg.adSrc) {
      try { doc.addImage(cfg.adSrc, pdfImgFormat(cfg.adSrc), (W - imgW) / 2, cy, imgW, imgH); cy += imgH + 7; } catch { /* rasm buzilgan bo'lsa tashlab ketiladi */ }
    }
    socials.forEach(([k, label, handle]) => {
      const x = W / 2 - 34;
      try { doc.addImage(PDF_DEFAULT_IMAGES[k], "JPEG", x, cy - 3.6, 4.6, 4.6); } catch {}
      doc.setTextColor(...TEXT); doc.setFont(PDF_FONT, "bold"); doc.setFontSize(9); doc.text(`${label}:`, x + 7, cy);
      doc.setTextColor(...GREY); doc.setFont(PDF_FONT, "normal"); doc.text(handle, x + 7 + doc.getTextWidth(`${label}: `) + 1, cy);
      cy += 7;
    });
  }

  pdfFooters(doc, `Barcha ishtirokchilar · ${testName}`);
  return pdfFinish(doc, `${pdfSafeName(testName)}_${pdfSafeName(cfg.orgName)}.pdf`);
}

// ===== BARCHA MARKAZLAR HISOBOTLARI — BITTA ZIP =====
let jsZipLoading = null;
function loadJsZip() {
  if (window.JSZip) return Promise.resolve(window.JSZip);
  if (jsZipLoading) return jsZipLoading;
  jsZipLoading = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = () => resolve(window.JSZip);
    s.onerror = () => { jsZipLoading = null; reject(new Error("ZIP kutubxonasi yuklanmadi. Internet aloqasini tekshiring.")); };
    document.head.appendChild(s);
  });
  return jsZipLoading;
}
// Har bir markaz uchun alohida PDF + umumiy natijalar PDF'i bitta ZIP papkaga joylanadi.
// onProgress(tayyor, jami) — jarayonni ko'rsatish uchun.
async function buildAllCentersZip({ summary, uploads, onProgress }) {
  const JSZip = await loadJsZip();
  const partners = db.get("partners") || [];
  const files = uploads.filter(u => String(u.testId) === String(summary.id) && u.partnerId !== "admin" && (u.rows || []).length);
  const total = files.length + 1;
  const zip = new JSZip();
  const folderName = pdfSafeName(`${summary.testName}_markazlar_natijalari`);
  const folder = zip.folder(folderName);
  const used = new Set();
  const uniq = (name) => { let n = name, k = 2; while (used.has(n.toLowerCase())) n = name.replace(/\.pdf$/i, `_${k++}.pdf`); used.add(n.toLowerCase()); return n; };
  let done = 0;
  onProgress && onProgress(done, total);
  const overall = await buildOverallReportPdf({ testName: summary.testName, rows: summary.rows || [], itemShare: summary.itemShare, itemN: summary.itemN });
  folder.file(uniq(`00_Umumiy_natijalar_${pdfSafeName(summary.testName)}.pdf`), await overall.blob.arrayBuffer());
  URL.revokeObjectURL(overall.url);
  onProgress && onProgress(++done, total);
  for (const u of files) {
    const centerName = partners.find(p => p.id === u.partnerId)?.name || u.partnerName || "O'quv markazi";
    const pdf = await buildCenterReportPdf({ centerName, testName: summary.testName, rows: u.rows || [], stats: u.stats });
    folder.file(uniq(pdf.filename), await pdf.blob.arrayBuffer());
    URL.revokeObjectURL(pdf.url);
    onProgress && onProgress(++done, total);
    await new Promise(r => setTimeout(r, 0)); // sahifa qotib qolmasligi uchun
  }
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const blob = new Blob([bytes], { type: "application/zip" });
  return { url: URL.createObjectURL(blob), filename: `${folderName}.zip`, blob, count: files.length };
}

// PDF tayyor bo'lgach ko'rsatiladigan oyna — yuklab olish havolasi foydalanuvchining o'zi
// bosadigan tugma (ba'zi ilovalar, masalan Telegram, avtomatik yuklab olishni bloklaydi).
function PdfReadyModal({ pdf, onClose, title }) {
  if (!pdf) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 99999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div style={{ background: "white", borderRadius: 16, padding: 22, maxWidth: 360, width: "100%", boxShadow: "0 10px 40px rgba(0,0,0,0.25)" }} onClick={e => e.stopPropagation()}>
        <p style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 800, color: C.text }}>{pdf.title || title || "📄 PDF hisobot tayyor"}</p>
        <p style={{ margin: "0 0 16px", fontSize: 12.5, color: C.textMid, lineHeight: 1.5 }}>Yuklab olish uchun tugmani bosing. Ishlamasa, "Yangi oynada ochish" orqali oching va u yerdan saqlang.</p>
        <a href={pdf.url} download={pdf.filename} style={{ display: "block", textAlign: "center", padding: "12px", borderRadius: 10, background: "#0891B2", color: "white", fontWeight: 800, fontSize: 14, textDecoration: "none", marginBottom: 10, wordBreak: "break-all" }}>⬇️ {pdf.filename}</a>
        <a href={pdf.url} target="_blank" rel="noopener noreferrer" style={{ display: "block", textAlign: "center", padding: "11px", borderRadius: 10, border: `1.5px solid ${C.border}`, color: C.textMid, fontWeight: 700, fontSize: 13, textDecoration: "none", marginBottom: 10 }}>↗ Yangi oynada ochish</a>
        <button onClick={onClose} style={{ width: "100%", padding: "10px", borderRadius: 10, border: "none", background: "transparent", fontWeight: 600, fontSize: 13, cursor: "pointer", color: C.textLight }}>Yopish</button>
      </div>
    </div>
  );
}

// ===== EXCEL EXPORT (SheetJS .xlsx) =====

// Ma'lumotni tayyorlab, DATA URL qaytaradi (avtomatik yuklab olishga urinmaydi).
// Sabab: ko'plab embedded WebView muhitlar (masalan Telegram Mini App) dasturiy
// ravishda (.click()) boshlangan yuklab olishlarni bloklaydi, lekin foydalanuvchi
// O'ZI bosgan havolaga (haqiqiy user gesture) ruxsat beradi.
// Mustaqil "Excel -> Rash" hisob-kitobi natijasini yuklab olinadigan .xlsx qilib tayyorlaydi.
function buildRaschCalcExport(rows, settings) {
  // Excel shablonidagi "Hisobot" varag'i bilan bir xil ustunlar
  const withSource = rows.some(r => r.source);
  const header = ["O'rin","F.I.Sh","Algebra","Geometriya","Umumiy ball","Baho","Foiz (%)","BMBA ball","O'quv markazi","To'g'ri javoblar", ...(withSource?["Manba"]:[])];
  const sorted = [...rows].sort((a,b)=>(b.ball??-999)-(a.ball??-999));
  const n2 = (v) => v!=null && isFinite(v) ? +v.toFixed(2) : "";
  const data = sorted.map((r,i)=>[r.ball!=null?i+1:"", r.name||"", n2(r.alg), n2(r.geo), n2(r.ball), r.daraja||"", n2(r.foiz), n2(r.bmba), r.group||"", r.correct??"", ...(withSource?[r.source||""]:[])]);
  if (typeof XLSX !== "undefined") {
    try {
      const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let C2 = range.s.c; C2 <= range.e.c; C2++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C2 });
        if (ws[addr]) ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: "6D28D9" } } };
      }
      ws["!cols"] = header.map((h,i)=>({ wch: i===1?28:i===8?20:12 }));
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

function slotValRaw(slot) {
  if (!slot) return "";
  if (slot.nodes) return nodesToString(slot.nodes);
  return slot.value || "";
}
function slotVal(slot) { return balanceParens(slotValRaw(slot)); }
function nodeVal(n) {
  if (!n) return "";
  if (n.type === "text") return n.value || "";
  if (n.type === "slot") return slotVal(n);
  return nodesToString([n]);
}
// Kasr/daraja/ildiz kataklari ichidagi qavslarni muvozanatlaydi. Formulalar ichki ko'rinishda
// FRAC(a,b), SUP(a,b) kabi qavslar bilan saqlanadi — katak ichida yopilmagan "(" yoki ortiqcha ")"
// qolsa (masalan o'chirish paytida), butun tuzilma buzilib "SUP((x+3,)" kabi xom matn chiqib qolardi.
function balanceParens(str) {
  let depth = 0, pre = 0;
  for (const ch of String(str || "")) {
    if (ch === "(") depth++;
    else if (ch === ")") { if (depth === 0) pre++; else depth--; }
  }
  return "(".repeat(pre) + (str || "") + ")".repeat(depth);
}
function nodesToString(nodes) {
  const slotVal = (slot) => balanceParens(slotValRaw(slot));
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
  // Kursor turgan eng ichki daraja (sup) tuzilmasini topadi
  const findSupAtCursor = (list) => {
    for (const n of list) {
      if (n.children) {
        for (const sl of n.children) { if (sl?.nodes) { const deeper = findSupAtCursor(sl.nodes); if (deeper) return deeper; } }
        if (n.type === "sup" && n.children.some(sl => sl?.nodes && collectTextNodes(sl.nodes).some(t => t.id === cursor))) return n;
      } else if (n.type === "slot" && n.nodes) { const d = findSupAtCursor(n.nodes); if (d) return d; }
    }
    return null;
  };
  // Daraja tuzilmasini "yechadi": uning o'rniga faqat asosdagi elementlar qoladi
  const unwrapNode = (list, id, replacement) => {
    const out = [];
    for (const n of list) {
      if (n.id === id) { out.push(...replacement); continue; }
      if (n.type === "slot" && n.nodes) { out.push({ ...n, nodes: unwrapNode(n.nodes, id, replacement) }); continue; }
      if (n.children) { out.push({ ...n, children: n.children.map(sl => sl?.nodes ? { ...sl, nodes: unwrapNode(sl.nodes, id, replacement) } : sl) }); continue; }
      out.push(n);
    }
    return out;
  };
  const deleteChar = () => {
    // Daraja bo'sh bo'lsa — o'chirish tugmasi daraja belgisini olib tashlaydi (asos oddiy matn bo'lib qoladi).
    // Kursor asosda bo'lsa, shu bilan birga oxirgi belgi ham o'chiriladi.
    const sup = findSupAtCursor(nodes);
    if (sup) {
      const expNodes = sup.children[1]?.nodes || [];
      const expEmpty = expNodes.every(n => n.type === "text" && !n.value);
      if (expEmpty) {
        const inExp = collectTextNodes(expNodes).some(t => t.id === cursor);
        const baseNodes = (sup.children[0]?.nodes && sup.children[0].nodes.length) ? sup.children[0].nodes : [textNode("")];
        const next = unwrapNode(nodes, sup.id, baseNodes);
        const baseTexts = collectTextNodes(baseNodes);
        if (inExp) { setNodes(next); setCursor(baseTexts[baseTexts.length - 1].id); return; }
        const cur = baseTexts.find(t => t.id === cursor);
        setNodes(cur && cur.value ? updateTextNode(next, cursor, v => v.slice(0, -1)) : next);
        if (!cur) setCursor(baseTexts[baseTexts.length - 1].id);
        return;
      }
    }
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
      // YANA MUHIM: agar matn "13π" kabi RAQAM+O'ZGARUVCHI/KONSTANTA ketma-ketligi
      // bo'lsa, daraja FAQAT oxirgi bitta belgiga (π) tegishli bo'lishi kerak — "13"
      // koeffitsiyent bo'lib qolaveradi (standart matematik qoida: 13π² = 13·π²,
      // (13π)² emas). Shuning uchun: agar matn RAQAM bilan tugasa — butun sonni
      // (masalan "12.5") bazaga olamiz; aks holda (harf/π/e bilan tugasa) — faqat
      // O'SHA BITTA oxirgi belgini bazaga olamiz, oldingi son/harflar prefiks bo'lib qoladi.
      const curNode = collectTextNodes(nodes).find(n => n.id === cursor);
      const fullVal = curNode ? curNode.value : "";
      let prefix = "", baseVal = "";
      if (fullVal) {
        const numMatch = fullVal.match(/[0-9]*\.?[0-9]+$/);
        if (fullVal.endsWith(")")) {
          // Qavs bilan tugagan bo'lsa — butun qavsli ifoda (masalan "(x+2)") darajaning asosi bo'ladi.
          // Mos ochiluvchi qavsni orqaga qarab qidiramiz (ichma-ich qavslarni hisobga olib).
          let depth = 0, open = -1;
          for (let k = fullVal.length - 1; k >= 0; k--) {
            if (fullVal[k] === ")") depth++;
            else if (fullVal[k] === "(") { depth--; if (depth === 0) { open = k; break; } }
          }
          if (open >= 0) {
            // Qavsdan oldin funksiya nomi bo'lsa (sin(x) kabi) — u ham asosga kiradi
            const fnm = fullVal.slice(0, open).match(/(arcsin|arccos|arctan|sinh|cosh|tanh|sin|cos|tan|cot|ln|lg|log)$/);
            const start = fnm ? open - fnm[1].length : open;
            baseVal = fullVal.slice(start);
            prefix = fullVal.slice(0, start);
          } else {
            prefix = fullVal;
          }
        } else if (numMatch && /[0-9]$/.test(fullVal)) {
          baseVal = numMatch[0];
          prefix = fullVal.slice(0, fullVal.length - baseVal.length);
        } else if (/[a-zA-Z\u03c0]$/.test(fullVal)) {
          baseVal = fullVal.slice(-1);
          prefix = fullVal.slice(0, -1);
        } else {
          prefix = fullVal;
        }
      }
      const baseSlot = slotNode(baseVal ? [textNode(baseVal)] : undefined);
      const expSlot = slotNode();
      struct = createNode("sup", "", [baseSlot, expSlot]);
      setNodes(prev => {
        // Operator/koeffitsiyent (prefix) qismi joriy tugunda qoladi, faqat baza qismi olib tashlanadi
        const cleared = fullVal ? updateTextNode(prev, cursor, () => prefix) : prev;
        return insertStructureInto(cleared, cursor, struct, after);
      });
      // Asos bo'sh qolsa — kursor avval asosga qo'yiladi (bo'sh "□" qolib ketmasligi uchun)
      setCursor(baseVal ? ((expSlot.nodes && expSlot.nodes[0]?.id) || after.id) : ((baseSlot.nodes && baseSlot.nodes[0]?.id) || after.id));
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

  // ── KOMPYUTER KLAVIATURASI ── Klaviatura ochiq turganda, foydalanuvchi
  // ekrandagi tugmalarni bosmasdan, kompyuterning o'z klaviaturasi bilan ham
  // yoza olishi uchun (raqamlar, harflar, +−×÷, qavslar, orqaga qaytarish,
  // strelkalar, Enter — OK tugmasi bilan bir xil ishlaydi).
  useEffect(() => {
    const onKeyDown = (e) => {
      if (e.ctrlKey || e.metaKey || e.altKey) return; // Ctrl/Cmd/Alt bilan birga — brauzerning o'z ishiga aralashmaymiz
      const k = e.key;
      if (/^[0-9]$/.test(k)) { e.preventDefault(); insertChar(k); return; }
      if (/^[a-zA-Z]$/.test(k)) { e.preventDefault(); insertChar(k); return; }
      if (k === "+") { e.preventDefault(); insertChar("+"); return; }
      if (k === "-") { e.preventDefault(); insertChar("-"); return; }
      if (k === "*") { e.preventDefault(); insertChar("*"); return; }
      if (k === "/") { e.preventDefault(); insertChar("/"); return; }
      if (k === "(") { e.preventDefault(); insertChar("("); return; }
      if (k === ")") { e.preventDefault(); insertChar(")"); return; }
      if (k === "," || k === ".") { e.preventDefault(); insertChar(","); return; }
      if (k === ";") { e.preventDefault(); insertChar(";"); return; }
      if (k === "=") { e.preventDefault(); insertChar("="); return; }
      if (k === "^") { e.preventDefault(); insertStructure("sup"); return; }
      if (k === "Backspace") { e.preventDefault(); deleteChar(); return; }
      if (k === "ArrowLeft") { e.preventDefault(); movePrev(); return; }
      if (k === "ArrowRight") { e.preventDefault(); moveNext(); return; }
      if (k === "Enter") { e.preventDefault(); onClose(); return; }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [cursor, nodes]);

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
// O'quvchi ism-familiyasini solishtirish uchun: katta-kichik harf, ortiqcha bo'sh joy va
// tutuq belgisining turli ko'rinishlari (ʻ ’ ` ') farq qilmaydi.
function normPersonName(s) {
  return String(s || "").toLowerCase().replace(/[ʻʼ‘’`´]/g, "'").replace(/\s+/g, " ").trim();
}
function sameStudentName(u, text) {
  const n = normPersonName(text);
  return !!n && (normPersonName(`${u.firstName} ${u.lastName}`) === n || normPersonName(`${u.lastName} ${u.firstName}`) === n);
}
async function findStudentByLogin(login, pwd) {
  const users = db.get("users") || [];
  const cands = users.filter(u => sameStudentName(u, login) || u.phone === String(login || "").trim());
  for (const u of cands) {
    if (await verifyPassword(u.password, pwd)) return upgradePasswordIfPlain("users", x => x.phone === u.phone, pwd);
  }
  return null;
}
// O'quv markazi o'z NOMI bilan kiradi (alohida login yo'q). Nom band emasligini tekshiradi:
// boshqa markaz nomi/logini, o'qituvchi logini yoki admin login bilan bir xil bo'lmasin.
function partnerNameTaken(name, exceptId) {
  const n = normPersonName(name);
  if (!n) return false;
  if (n === normPersonName(ADMIN_LOGIN)) return true;
  if ((db.get("teachers") || []).some(t => normPersonName(t.login) === n)) return true;
  return (db.get("partners") || []).some(p => p.id !== exceptId && (normPersonName(p.name) === n || normPersonName(p.login) === n));
}
// ===== GURUHLAR =====
// Guruhni admin/o'qituvchi ochadi ("groups" to'plami) yoki testni biror guruhga mo'ljallaganda
// o'sha guruh ham ochilgan hisoblanadi. O'quvchi guruhlar ro'yxatini ko'rmaydi — nomini o'zi
// yozadi va shunday guruh ochilgan bo'lsagina qo'shiladi (katta-kichik harf, bo'sh joy va
// tutuq belgisining turi farq qilmaydi).
function normGroupName(s) {
  return String(s || "").toLowerCase().replace(/[ʻʼ‘’`´]/g, "'").replace(/\s+/g, " ").trim();
}
function openGroups() {
  const map = new Map();
  (db.get("groups") || []).forEach(g => { const n = normGroupName(g.name); if (n && !map.has(n)) map.set(n, g.name.trim()); });
  (db.get("tests") || []).forEach(t => (t.targetGroups || []).forEach(g => { const n = normGroupName(g); if (n && !map.has(n)) map.set(n, String(g).trim()); }));
  return [...map.values()].sort((a, b) => a.localeCompare(b));
}
function findOpenGroup(text) {
  const n = normGroupName(text);
  return n ? (openGroups().find(g => normGroupName(g) === n) || null) : null;
}
// Guruh(lar)ni "groups" to'plamiga qo'shadi (allaqachon bo'lsa — qayta qo'shmaydi)
function ensureGroups(names) {
  const cur = db.get("groups") || [];
  const have = new Set(cur.map(g => normGroupName(g.name)));
  const add = [];
  names.forEach(nm => { const n = normGroupName(nm); if (n && !have.has(n)) { have.add(n); add.push({ id: `g${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, name: String(nm).trim().replace(/\s+/g, " "), createdAt: Date.now() }); } });
  if (add.length) db.set("groups", [...cur, ...add]);
  return add.length;
}
function allKnownGroups() { return openGroups(); }

function LoginPage({ onLogin, onRegister, onAdmin, onPartner, onPartnerRegister }) {
  const [login, setLogin] = useState("");
  const [pwd, setPwd] = useState("");
  const [err, setErr] = useState("");
  const [showPwd, setShowPwd] = useState(false);

  const [busy, setBusy] = useState(false);
  const [lockLeft, setLockLeft] = useState(() => Math.max(0, loginGuardState().until - Date.now()));
  useEffect(() => {
    if (lockLeft <= 0) return;
    const t = setInterval(() => setLockLeft(Math.max(0, loginGuardState().until - Date.now())), 1000);
    return () => clearInterval(t);
  }, [lockLeft > 0]);

  const go = async () => {
    if (busy) return;
    if (!login.trim() || !pwd) { setErr("Login va parolni kiriting!"); return; }
    const guard = loginGuardState();
    if (guard.until > Date.now()) { setLockLeft(guard.until - Date.now()); return; }
    setBusy(true); setErr("");
    try {
      const L = login.trim();
      // 1. Admin — parol faqat xesh bilan solishtiriladi
      if (L === ADMIN_LOGIN && await verifyPassword(getAdminPasswordHash(), pwd)) {
        loginGuardReset(); onAdmin(true); return;
      }
      // 2. Yordamchi o'qituvchi
      for (const t of (db.get("teachers") || []).filter(t => t.login === L)) {
        if (await verifyPassword(t.password, pwd)) {
          const fresh = await upgradePasswordIfPlain("teachers", x => x.id === t.id, pwd);
          loginGuardReset(); onAdmin(false, fresh || t); return;
        }
      }
      // 3. Hamkor markaz — markaz nomi (katta-kichik harf farqsiz) yoki eski login bilan
      for (const p of (db.get("partners") || []).filter(p => p.login === L || normPersonName(p.name) === normPersonName(L))) {
        if (await verifyPassword(p.password, pwd)) {
          const fresh = await upgradePasswordIfPlain("partners", x => x.id === p.id, pwd);
          loginGuardReset();
          if (!p.approved) { setErr("Hisobingiz hali admin tomonidan tasdiqlanmagan. Iltimos, kuting yoki admin bilan bog'laning."); return; }
          onPartner(fresh || p); return;
        }
      }
      // 4. O'quvchi: ism + familiya (istalgan tartibda) va parol. Eski hisoblar — telefon bilan ham.
      const u = await findStudentByLogin(L, pwd);
      if (u) { loginGuardReset(); onLogin(u); return; }

      const g = loginGuardFail();
      if (g.until > Date.now()) { setLockLeft(g.until - Date.now()); setErr(""); }
      else setErr(`Login yoki parol noto'g'ri! (${5 - g.fails} ta urinish qoldi)`);
    } catch (e) {
      setErr(e.message || "Kirishda xatolik yuz berdi.");
    } finally { setBusy(false); }
  };

  return (
    <AuthLayout>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{fontSize:52}}>📐</div>
        <h1 style={{margin:"8px 0 4px",fontSize:24,fontWeight:800}}>Matematika Testi</h1>
        <p style={{margin:0,color:C.textMid,fontSize:14}}>Platformaga xush kelibsiz</p>
      </div>
      {err&&<div style={S.err}>{err}</div>}
      <label style={S.label}>Ism Familiya / Markaz nomi</label>
      <input value={login} onChange={e=>setLogin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} style={S.input} placeholder="Masalan: Ali Valiyev"/>
      <p style={{margin:"-4px 0 10px",fontSize:11.5,color:C.textLight}}>O'quv markazlari — markaz nomini, admin va o'qituvchilar — o'z loginini kiritadi.</p>
      <label style={S.label}>Parol</label>
      <div style={{position:"relative"}}>
        <input type={showPwd?"text":"password"} value={pwd} onChange={e=>setPwd(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}
          style={{...S.input,paddingRight:44}} placeholder="••••••"/>
        <button onClick={()=>setShowPwd(s=>!s)} type="button"
          style={{position:"absolute",right:10,top:10,background:"none",border:"none",cursor:"pointer",fontSize:18,color:C.textMid}}>
          {showPwd?"🙈":"👁"}
        </button>
      </div>
      {lockLeft > 0 && <div style={S.err}>🔒 Ko'p marta noto'g'ri parol kiritildi. {Math.ceil(lockLeft/1000)} soniyadan keyin qayta urinib ko'ring.</div>}
      <button onClick={go} disabled={busy || lockLeft > 0} style={{...S.btnPrimary,opacity:(busy||lockLeft>0)?0.6:1}}>{busy ? "⏳ Tekshirilmoqda..." : "Kirish"}</button>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,margin:"12px 0 0"}}>
        Hisobingiz yo'qmi? <span onClick={onRegister} style={{color:C.primary,cursor:"pointer",fontWeight:700}}>Ro'yxatdan o'ting</span>
      </p>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,margin:"6px 0 0"}}>
        Hamkor o'quv markazmisiz? <span onClick={onPartnerRegister} style={{color:"#0891B2",cursor:"pointer",fontWeight:700}}>Shu yerda ro'yxatdan o'ting</span>
      </p>
    </AuthLayout>
  );
}
function RegisterPage({ onDone, onLogin }) {
  const [f,setF]=useState({firstName:"",lastName:"",password:"",password2:""});
  const [err,setErr]=useState("");
  const [showPwd,setShowPwd]=useState(false);

  const [busy,setBusy]=useState(false);
  const go=async()=>{
    if(busy) return;
    const firstName=f.firstName.trim().replace(/\s+/g," "), lastName=f.lastName.trim().replace(/\s+/g," ");
    if(!firstName||!lastName||!f.password){setErr("Barcha maydonlarni to'ldiring!");return;}
    if(f.password.length<4){setErr("Parol kamida 4 ta belgidan iborat bo'lsin!");return;}
    if(f.password!==f.password2){setErr("Parollar mos kelmadi!");return;}
    const users=db.get("users")||[];
    const full=`${firstName} ${lastName}`;
    // Bir xil ism-familiyali o'quvchilar bo'lishi mumkin — ular parol bilan ajratiladi.
    // Shuning uchun aynan shu ism-familiya VA shu parol bilan hisob bo'lsa, boshqa parol so'raladi.
    setBusy(true);
    try {
    for(const x of users.filter(u=>sameStudentName(u,full))){ if(await verifyPassword(x.password,f.password)){ setBusy(false); setErr("Bu ism-familiya bilan hisob allaqachon bor. Agar bu siz bo'lsangiz — \"Kirish\"ni bosing, aks holda boshqa parol tanlang.");return;} }
    // Ichki identifikator (natijalar shu bo'yicha bog'lanadi); o'quvchi uni ko'rmaydi va kiritmaydi
    let key; do { key="id"+Date.now().toString(36)+Math.random().toString(36).slice(2,6); } while(users.some(u=>u.phone===key));
    const u={firstName,lastName,group:"",phone:key,password:await hashPassword(f.password),id:Date.now()};
    db.set("users",[...(db.get("users")||[]),u]); onDone(u);
    } catch(e){ setErr(e.message||"Xatolik yuz berdi."); }
    setBusy(false);
  };

  return (
    <AuthLayout>
      <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:52}}>📝</div><h1 style={{margin:"8px 0 0",fontSize:22,fontWeight:800}}>Ro'yxatdan O'tish</h1></div>
      {err&&<div style={S.err}>{err}</div>}
      {[["firstName","Ism"],["lastName","Familiya"]].map(([k,ph])=>(
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
      <button onClick={go} disabled={busy} style={{...S.btnPrimary,opacity:busy?0.6:1}}>{busy?"⏳ Saqlanmoqda...":"Ro'yxatdan O'tish"}</button>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,marginTop:12}}>Hisobingiz bormi? <span onClick={onLogin} style={{color:C.primary,cursor:"pointer",fontWeight:700}}>Kirish</span></p>
    </AuthLayout>
  );
}


// ===== HAMKOR MARKAZ — O'ZI RO'YXATDAN O'TISHI =====
// Ro'yxatdan o'tgach hisob DARHOL faollashmaydi — admin PartnerManager orqali
// tasdiqlaguncha "kutilmoqda" holatida turadi va kirish rad etiladi.
function PartnerRegisterPage({ onDone, onLogin }) {
  const [f,setF]=useState({name:"",password:"",password2:""});
  const [err,setErr]=useState("");
  const [showPwd,setShowPwd]=useState(false);
  const [done,setDone]=useState(false);

  const go=async()=>{
    const name=f.name.trim().replace(/\s+/g," ");
    if(!name||!f.password){setErr("Markaz nomi va parolni kiriting!");return;}
    if(f.password.length<4){setErr("Parol kamida 4 ta belgidan iborat bo'lsin!");return;}
    if(f.password!==f.password2){setErr("Parollar mos kelmadi!");return;}
    if(partnerNameTaken(name)){setErr("Bu nomdagi markaz allaqachon ro'yxatdan o'tgan. Nomga shahar yoki tumanni qo'shing (masalan: \"Piramida o'quv markazi Jondor\").");return;}
    let hashed; try { hashed=await hashPassword(f.password); } catch(e){ setErr(e.message); return; }
    const p={id:Date.now(),name,login:name,password:hashed,approved:false,createdAt:Date.now()};
    db.set("partners",[...(db.get("partners")||[]),p]);
    setDone(true);
  };

  if (done) {
    return (
      <AuthLayout>
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:52}}>⏳</div>
          <h1 style={{margin:"8px 0 0",fontSize:20,fontWeight:800}}>So'rov yuborildi</h1>
        </div>
        <p style={{textAlign:"center",color:C.textMid,fontSize:14,lineHeight:1.6,margin:"0 0 20px"}}>
          Arizangiz qabul qilindi. Admin tasdiqlagandan so'ng, markaz nomi va parolingiz bilan kira olasiz. Iltimos, admin bilan bog'lanib xabar bering.
        </p>
        <button onClick={onLogin} style={S.btnPrimary}>Kirish sahifasiga qaytish</button>
      </AuthLayout>
    );
  }

  return (
    <AuthLayout>
      <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:52}}>🤝</div><h1 style={{margin:"8px 0 0",fontSize:20,fontWeight:800}}>Hamkor Markaz — Ro'yxatdan O'tish</h1></div>
      <p style={{textAlign:"center",color:C.textLight,fontSize:12.5,margin:"0 0 16px",lineHeight:1.5}}>Ro'yxatdan o'tgach, admin tasdiqlashini kutasiz — shundan keyingina kira olasiz.</p>
      {err&&<div style={S.err}>{err}</div>}
      <label style={S.label}>Markaz nomi</label>
      <input value={f.name} onChange={e=>setF({...f,name:e.target.value})} style={S.input} placeholder="Masalan: 'Iqtidor' o'quv markazi"/>
      <p style={{margin:"-4px 0 10px",fontSize:11.5,color:C.textLight}}>Kirishda shu nom va parolni yozasiz.</p>
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
      <button onClick={go} style={{...S.btnPrimary,background:"#0891B2"}}>Ro'yxatdan O'tish</button>
      <p style={{textAlign:"center",color:C.textMid,fontSize:13,marginTop:12}}>Hisobingiz bormi? <span onClick={onLogin} style={{color:"#0891B2",cursor:"pointer",fontWeight:700}}>Kirish</span></p>
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
    const set = new Set(openGroups());
    (existing?.targetGroups||[]).forEach(g=>set.add(g));
    return [...set].sort((a,b)=>a.localeCompare(b));
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
    if(restrictGroups && targetGroups.length) ensureGroups(targetGroups);
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
    onSave({id:existing?.id||Date.now(),name,duration,closedCount:questions.filter(q=>q.type==="closed").length,optionsCount:sections[0]?typeOpts(sections[0].type):4,sections,questions,active:willAutoStart?false:wasActive,everActivated:existing?.everActivated||(willAutoStart?false:wasActive),scheduledAt:schedTs,showAnswersAfter:showAnswers,showStats,startedAt:existing?.startedAt||null,langDocs:cleanLangDocs,pdfUrl:uzDoc.pdfUrl,latexSource:uzDoc.latexSource,latexFileName:uzDoc.latexFileName,latexImages:uzDoc.latexImages,accessCode:requireCode?accessCode.trim().toUpperCase():null,codePrice:requireCode?codePrice:null,targetGroups:restrictGroups?targetGroups:[]});
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
                  <span style={{color:C.textMid,fontSize:12}}>{items.reduce((s,{q})=>s+(q.subParts?.length>0?q.subParts.length:1),0)} savol</span>
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
  const [open, setOpen] = useState(false);
  const [list, setList] = useState(()=>db.get("teachers")||[]);
  const [form, setForm] = useState({name:"",login:"",password:""});
  const [showPwd, setShowPwd] = useState(false);
  const [err, setErr] = useState("");
  const [confirmModal, setConfirmModal] = useState(null);

  const add = async () => {
    if(!form.name.trim()||!form.login.trim()||!form.password){setErr("Barcha maydonlarni to'ldiring!");return;}
    if(form.password.length<4){setErr("Parol kamida 4 ta belgi!");return;}
    if(form.login.trim()===ADMIN_LOGIN){setErr("Bu login band!");return;}
    const cur=db.get("teachers")||[];
    if(cur.find(t=>t.login===form.login.trim())){setErr("Bu login band!");return;}
    let hashed; try { hashed=await hashPassword(form.password); } catch(e){ setErr(e.message); return; }
    const upd=[...(db.get("teachers")||[]),{id:Date.now(),name:form.name.trim(),login:form.login.trim(),password:hashed}];
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
      <button onClick={()=>setOpen(o=>!o)} aria-expanded={open} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,width:"100%",background:"none",border:"none",padding:0,cursor:"pointer",textAlign:"left",marginBottom:open?14:0}}>
        <span style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><b style={{fontSize:15,color:C.primary}}>👩‍🏫 Yordamchi o'qituvchilar</b><span style={{...S.badge,background:C.primaryLight,color:C.primary}}>{list.length} ta</span></span>
        <span style={{fontSize:13,color:C.textMid,fontWeight:700,transform:open?"rotate(180deg)":"none",transition:"transform 0.2s"}}>▼</span>
      </button>
      {open&&(<>
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
      </>)}
    </div>
  );
}

// Hamkor o'quv markazlar — faqat o'z Excel natijalarini yuklash huquqiga ega,
// boshqa hech narsani (o'quvchilar, boshqa testlar, natijalar ro'yxati) ko'rmaydi.
// Admin ular uchun shu yerda login/parol yaratib beradi.
function PartnerManager() {
  const [open, setOpen] = useState(false);
  const [list, setList] = useState(()=>db.get("partners")||[]);
  const [form, setForm] = useState({name:"",login:"",password:""});
  const [showPwd, setShowPwd] = useState(false);
  const [err, setErr] = useState("");
  const [confirmModal, setConfirmModal] = useState(null);

  const add = async () => {
    const name=form.name.trim().replace(/\s+/g," ");
    if(!name||!form.password){setErr("Markaz nomi va parolni kiriting!");return;}
    if(form.password.length<4){setErr("Parol kamida 4 ta belgi!");return;}
    if(partnerNameTaken(name)){setErr("Bu nomdagi markaz allaqachon bor!");return;}
    let hashed; try { hashed=await hashPassword(form.password); } catch(e){ setErr(e.message); return; }
    // Admin o'zi to'g'ridan-to'g'ri qo'shgani uchun darhol tasdiqlangan hisoblanadi
    const upd=[...(db.get("partners")||[]),{id:Date.now(),name,login:name,password:hashed,approved:true,createdAt:Date.now()}];
    db.set("partners",upd); setList(upd);
    setForm({name:"",login:"",password:""}); setErr("");
  };

  const del = (id) => {
    setConfirmModal({message:"Hamkor markaz o'chirilsinmi?", onConfirm: () => {
      const upd=(db.get("partners")||[]).filter(t=>t.id!==id);
      db.set("partners",upd); setList(upd);
      setConfirmModal(null);
    }});
  };

  const approve = (id) => {
    const upd=(db.get("partners")||[]).map(p=>p.id===id?{...p,approved:true}:p);
    db.set("partners",upd); setList(upd);
  };

  const pending = list.filter(p=>!p.approved);
  const approved = list.filter(p=>p.approved);

  return (
    <div style={{...S.card,padding:16,marginBottom:18,border:`2px solid #0891B2`}}>
      {confirmModal && <ConfirmModal message={confirmModal.message} onConfirm={confirmModal.onConfirm} onCancel={()=>setConfirmModal(null)}/>}
      <button onClick={()=>setOpen(o=>!o)} aria-expanded={open} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,width:"100%",background:"none",border:"none",padding:0,cursor:"pointer",textAlign:"left",marginBottom:open?14:0}}>
        <span style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><b style={{fontSize:15,color:"#0891B2"}}>🤝 Hamkor o'quv markazlar</b><span style={{...S.badge,background:"#CFFAFE",color:"#0891B2"}}>{approved.length} ta</span>{pending.length>0&&<span style={{...S.badge,background:"#FEF3C7",color:"#92400E"}}>⏳ {pending.length} ta tasdiq kutmoqda</span>}</span>
        <span style={{fontSize:13,color:C.textMid,fontWeight:700,transform:open?"rotate(180deg)":"none",transition:"transform 0.2s"}}>▼</span>
      </button>
      {open&&(<>
      <p style={{margin:"0 0 14px",fontSize:12,color:C.textLight,lineHeight:1.5}}>Ular faqat o'zlariga berilgan login/parol bilan kirib, bitta testni tanlab, o'z Excel natijalarini yuklay oladi (saytdagi natijalar bilan birga bitta Rash hisobida qo'shiladi). Boshqa hech qanday ma'lumotni (o'quvchilar, testlar tarkibi va h.k.) ko'rmaydi. O'zlari ham "Kirish" sahifasidan ro'yxatdan o'ta oladi — lekin siz shu yerda tasdiqlamaguningizcha kira olmaydilar.</p>

      {pending.length>0 && (
        <div style={{marginBottom:16,padding:12,background:"#FFFBEB",border:"1.5px solid #FDE68A",borderRadius:10}}>
          <p style={{margin:"0 0 10px",fontSize:13,fontWeight:800,color:"#92400E"}}>⏳ Tasdiqlanishi kerak ({pending.length})</p>
          {pending.map(p=>(
            <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,padding:"8px 0",borderTop:"1px solid #FDE68A"}}>
              <div>
                <b style={{fontSize:13}}>{p.name}</b>
                <p style={{margin:"2px 0 0",fontSize:11.5,color:C.textLight}}>{p.contact?p.contact:"O'z-o'zidan ro'yxatdan o'tgan"}</p>
              </div>
              <div style={{display:"flex",gap:6,flexShrink:0}}>
                <button onClick={()=>approve(p.id)} style={{...S.btnSmall,background:C.successDark,padding:"6px 10px",fontSize:12}}>✅ Tasdiqlash</button>
                <button onClick={()=>del(p.id)} style={{...S.btnSmall,background:C.danger,padding:"6px 10px",fontSize:12}}>🗑</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {err&&<div style={S.err}>{err}</div>}
      <p style={{margin:"0 0 8px",fontSize:12,fontWeight:700,color:C.textMid}}>Yoki o'zingiz qo'shing:</p>
      <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:12}}>
        {[["Nomi","name","text","Markaz nomi (kirish uchun ham)"],["Parol","password",showPwd?"text":"password","min 4 ta belgi"]].map(([lbl,k,tp,ph])=>(
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
        <button onClick={add} style={{...S.btnSmall,margin:0,padding:"10px",fontSize:13,background:"#0891B2"}}>+ Qo'shish (darhol tasdiqlangan)</button>
      </div>
      {approved.length===0
        ? <p style={{color:C.textLight,fontSize:13,margin:0,textAlign:"center"}}>Hali tasdiqlangan hamkor markaz yo'q</p>
        : <table style={{...S.table,fontSize:13}}>
            <thead><tr>{["","Nomi","Amal"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>{approved.map(t=>(
              <tr key={t.id}>
                <td style={S.td}>
                  {t.logo
                    ? <img src={t.logo} alt="" style={{width:28,height:28,borderRadius:6,objectFit:"cover"}}/>
                    : <div style={{width:28,height:28,borderRadius:6,background:"#CFFAFE",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🤝</div>}
                </td>
                <td style={S.td}><b>{t.name}</b></td>
                <td style={S.td}>
                  <button onClick={()=>del(t.id)} style={{...S.btnSmall,background:C.danger,padding:"4px 10px",fontSize:12}}>🗑 O'chirish</button>
                </td>
              </tr>
            ))}</tbody>
          </table>
      }
      </>)}
    </div>
  );
}



// ===== ADMIN: UMUMIY RASH NATIJALARI (sayt + barcha o'quv markazlar fayllari) =====
function RaschCombinedView({ tests, onExport, onRecalc, busyId, version }) {
  const [store, setStore] = useState(() => db.get("raschResults") || []);
  const [uploads, setUploads] = useState(() => db.get("partnerUploads") || []);
  const [testId, setTestId] = useState(() => {
    const st = db.get("raschResults") || [];
    return st.length ? String([...st].sort((a, b) => b.calculatedAt - a.calculatedAt)[0].id) : "";
  });
  const [query, setQuery] = useState("");
  const [source, setSource] = useState("");
  const [page, setPage] = useState(0);
  const [confirm, setConfirm] = useState(null);
  const [publishing, setPublishing] = useState(false);
  const [flash, setFlash] = useState(null);
  const [testsNow, setTestsNow] = useState(() => db.get("tests") || []);
  const [pdfReady, setPdfReady] = useState(null);
  const [pdfBusy, setPdfBusy] = useState(null);
  const [pdfCfg, setPdfCfg] = useState(() => { const { logoSrc, adSrc, ...rest } = getPdfSettings(); return rest; });
  const [overallPdfBusy, setOverallPdfBusy] = useState(false);
  const [zipProgress, setZipProgress] = useState(null); // {done,total}
  const allCentersZip = async () => {
    if (!summary || zipProgress) return;
    setZipProgress({ done: 0, total: 1 });
    try {
      const z = await buildAllCentersZip({ summary, uploads: db.get("partnerUploads") || [], onProgress: (done, total) => setZipProgress({ done, total }) });
      setPdfReady({ ...z, title: `🗂 ZIP tayyor: ${z.count} ta markaz + umumiy hisobot` });
    } catch (e) { setFlash("❌ " + (e.message || "ZIP yaratilmadi")); setTimeout(() => setFlash(null), 5000); }
    setZipProgress(null);
  };
  const overallPdf = async () => {
    if (!summary) return;
    setOverallPdfBusy(true);
    try { setPdfReady(await buildOverallReportPdf({ testName: summary.testName, rows: summary.rows || [], itemShare: summary.itemShare, itemN: summary.itemN })); }
    catch (e) { setFlash("❌ " + (e.message || "PDF yaratilmadi")); setTimeout(() => setFlash(null), 5000); }
    setOverallPdfBusy(false);
  };
  // Rasm yuklash: hisobotga joylash uchun kichraytirib JPEG qilinadi (sozlamalar hajmi kichik bo'lsin)
  const pickPdfImage = (key, maxSide) => (e) => {
    const file = e.target.files[0]; e.target.value = "";
    if (!file || !file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const im = new Image();
      im.onload = () => {
        const k = Math.min(1, maxSide / Math.max(im.naturalWidth, im.naturalHeight));
        const cv = document.createElement("canvas");
        cv.width = Math.round(im.naturalWidth * k); cv.height = Math.round(im.naturalHeight * k);
        const ctx = cv.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height); ctx.drawImage(im, 0, 0, cv.width, cv.height);
        setPdfCfg(p => ({ ...p, [key]: cv.toDataURL("image/jpeg", 0.8) }));
      };
      im.src = ev.target.result;
    };
    reader.readAsDataURL(file);
  };
  const savePdfCfg = () => {
    db.set("appSettings", [...(db.get("appSettings") || []).filter(x => x.id !== "pdf"), { ...pdfCfg, id: "pdf" }]);
    setFlash("✅ PDF hisobot sozlamalari saqlandi."); setTimeout(() => setFlash(null), 3000);
  };
  const adminCenterPdf = async (f) => {
    const u = uploads.find(x => x.id === f.id);
    if (!u) return;
    setPdfBusy(f.id);
    try {
      const pdf = await buildCenterReportPdf({ centerName: f.source, testName: summary?.testName || test?.name || "", rows: u.rows || [], stats: u.stats });
      setPdfReady(pdf);
    } catch (e) { setFlash("❌ " + (e.message || "PDF yaratilmadi")); setTimeout(() => setFlash(null), 5000); }
    setPdfBusy(null);
  };
  const refresh = () => { setStore(db.get("raschResults") || []); setUploads(db.get("partnerUploads") || []); setTestsNow(db.get("tests") || []); };
  useEffect(() => { refresh(); }, [busyId, version]);
  useEffect(() => {
    window.addEventListener("firestore-sync", refresh);
    return () => window.removeEventListener("firestore-sync", refresh);
  }, []);
  const PAGE = 100;
  const eligible = tests.filter(t => testEligibleForRasch(t));
  const test = tests.find(t => String(t.id) === String(testId));
  const summary = store.find(x => String(x.id) === String(testId));
  const testUploads = uploads.filter(u => String(u.testId) === String(testId));
  const newFiles = summary ? testUploads.filter(u => u.uploadedAt > summary.calculatedAt).length : testUploads.length;
  const rows = useMemo(() => summary ? [...summary.rows].sort((a, b) => (b.ball ?? -999) - (a.ball ?? -999)).map((r, i) => ({ ...r, _pos: i + 1 })) : [], [summary]);
  const sources = useMemo(() => [...new Set(rows.map(r => r.source || ""))].filter(Boolean).sort(), [rows]);
  const q = query.trim().toLowerCase();
  const filtered = rows.filter(r => (!source || r.source === source) && (!q || (r.name || "").toLowerCase().includes(q) || (r.group || "").toLowerCase().includes(q)));
  const pages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const pg = Math.min(page, pages - 1);
  const shown = filtered.slice(pg * PAGE, (pg + 1) * PAGE);
  const withBall = filtered.filter(r => typeof r.ball === "number");
  const avg = withBall.length ? withBall.reduce((a, r) => a + r.ball, 0) / withBall.length : null;
  const gradeCounts = {};
  withBall.forEach(r => { gradeCounts[r.daraja] = (gradeCounts[r.daraja] || 0) + 1; });
  const deleteFile = (f) => setConfirm({
    message: `"${f.source}" fayli (${f.rows} ta o'quvchi) umumiy hisobdan olib tashlansinmi? Qolganlar qayta hisoblanadi.`,
    onConfirm: () => {
      db.set("partnerUploads", (db.get("partnerUploads") || []).filter(u => u.id !== f.id));
      if (test) calculateRaschCombined(test, test.raschSettings || DEFAULT_RASCH_SETTINGS);
      setConfirm(null); refresh();
    },
  });
  const partnerFiles = testUploads.filter(u => u.partnerId !== "admin");
  const publishedAt = testsNow.find(t => String(t.id) === String(testId))?.raschPublishedAt || null;
  const unpublished = partnerFiles.filter(u => !u.publishedRows).length;
  const newSincePublish = !!summary && (!publishedAt || summary.calculatedAt > publishedAt);
  const publish = () => setConfirm({
    message: `"${test.name}" bo'yicha umumiy hisob yangilanib, natijalar ${new Set(partnerFiles.map(u => u.partnerId)).size} ta o'quv markaziga va saytda topshirgan o'quvchilarga bir vaqtda ochiladi. Davom etilsinmi?`,
    confirmLabel: "E'lon qilish", danger: false,
    onConfirm: () => {
      setConfirm(null); setPublishing(true);
      setTimeout(() => {
        const r = publishRaschToPartners(test, test.raschSettings || DEFAULT_RASCH_SETTINGS);
        setPublishing(false); refresh();
        setFlash(`📢 Natijalar e'lon qilindi: ${r.publishedCenters} ta o'quv markazi va saytdagi ${r.publishedSite} ta o'quvchi (umumiy hisobda ${r.allRows.length} ta).`);
        setTimeout(() => setFlash(null), 6000);
      }, 30);
    },
  });
  const pager = (
    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", margin: "10px 0" }}>
      <button onClick={() => setPage(0)} disabled={pg === 0} style={{ ...S.btnGhost, padding: "6px 10px", opacity: pg === 0 ? 0.4 : 1 }}>«</button>
      <button onClick={() => setPage(pg - 1)} disabled={pg === 0} style={{ ...S.btnGhost, padding: "6px 12px", opacity: pg === 0 ? 0.4 : 1 }}>← Oldingi</button>
      <span style={{ fontSize: 13, color: C.textMid }}><b style={{ color: C.text }}>{pg + 1}</b> / {pages} sahifa • {filtered.length} ta</span>
      <button onClick={() => setPage(pg + 1)} disabled={pg >= pages - 1} style={{ ...S.btnGhost, padding: "6px 12px", opacity: pg >= pages - 1 ? 0.4 : 1 }}>Keyingi →</button>
      <button onClick={() => setPage(pages - 1)} disabled={pg >= pages - 1} style={{ ...S.btnGhost, padding: "6px 10px", opacity: pg >= pages - 1 ? 0.4 : 1 }}>»</button>
    </div>
  );
  return (
    <div>
      {confirm && <ConfirmModal message={confirm.message} confirmLabel={confirm.confirmLabel} danger={confirm.danger} onConfirm={confirm.onConfirm} onCancel={() => setConfirm(null)} />}
      {flash && <div style={{ position: "fixed", top: 16, left: "50%", transform: "translateX(-50%)", zIndex: 99999, background: C.successDark, color: "white", padding: "12px 20px", borderRadius: 12, fontWeight: 700, fontSize: 13, boxShadow: "0 6px 20px rgba(0,0,0,0.2)", maxWidth: "90%", textAlign: "center" }}>{flash}</div>}
      <PdfReadyModal pdf={pdfReady} onClose={() => setPdfReady(null)} />
      <h3 style={{ margin: "0 0 6px" }}>🏆 Umumiy natijalar</h3>
      <p style={{ margin: "0 0 14px", color: C.textMid, fontSize: 13, lineHeight: 1.5 }}>Saytda topshirganlar va o'quv markazlar yuklagan barcha Excel fayllar bitta umumiy Rash hisobida baholanadi. Saytda yangi o'quvchi test topshirsa, u ham avtomatik shu hisobga qo'shiladi. Bu natijalarni faqat siz ko'rasiz. O'quv markazlari ham, saytda topshirgan o'quvchilar ham Rash natijasini siz "📢 E'lon qilish"ni bosganingizdan keyin, hammasi bir vaqtda ko'radi.</p>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 14 }}>
        <select value={testId} onChange={e => { setTestId(e.target.value); setPage(0); setSource(""); setQuery(""); }} style={{ ...S.input, marginBottom: 0, maxWidth: 340, flex: "1 1 220px" }}>
          <option value="">— Testni tanlang —</option>
          {eligible.map(t => <option key={t.id} value={t.id}>{t.name}{store.some(x => x.id === t.id) ? "" : " (hali hisoblanmagan)"}</option>)}
        </select>
        {test && <button onClick={() => onRecalc(test)} disabled={busyId === test.id} style={{ ...S.btnSmall, background: "#6D28D9", padding: "10px 16px", opacity: busyId === test.id ? 0.6 : 1 }}>{busyId === test.id ? "⏳ Hisoblanmoqda..." : "🔄 Qayta hisoblash"}</button>}
        {summary && partnerFiles.some(u => (u.rows || []).length) && <button onClick={allCentersZip} disabled={!!zipProgress} style={{ ...S.btnSmall, background: "#0E7490", padding: "10px 16px", opacity: zipProgress ? 0.7 : 1 }}>{zipProgress ? `⏳ ${zipProgress.done}/${zipProgress.total} PDF tayyorlanmoqda...` : "🗂 Barcha markazlar (ZIP)"}</button>}
        {summary && rows.length > 0 && <button onClick={overallPdf} disabled={overallPdfBusy} style={{ ...S.btnSmall, background: "#0891B2", padding: "10px 16px", opacity: overallPdfBusy ? 0.6 : 1 }}>{overallPdfBusy ? "⏳ PDF tayyorlanmoqda..." : "📄 PDF (umumiy)"}</button>}
        {summary && rows.length > 0 && <button onClick={() => { const ex = buildRaschCalcExport(filtered); if (ex) onExport({ ...ex, filename: `${summary.testName}_umumiy_natijalar.xlsx` }); }} style={{ ...S.btnSmall, background: C.successDark, padding: "10px 16px" }}>📥 Excel</button>}
      </div>

      {test && summary && (
        <div style={{ ...S.card, padding: "12px 14px", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", background: (unpublished || newSincePublish) ? "#EFF6FF" : C.successLight, border: `1.5px solid ${(unpublished || newSincePublish) ? "#BFDBFE" : "#86EFAC"}` }}>
          <div style={{ fontSize: 13, lineHeight: 1.5 }}>
            <b>{summary.siteCount || 0} ta sayt o'quvchisi</b> va <b>{new Set(partnerFiles.map(u => u.partnerId)).size} ta o'quv markazi</b> hisobda.{" "}
            {(unpublished || newSincePublish)
              ? <span style={{ color: "#1D4ED8" }}>{publishedAt ? "Oxirgi e'londan keyin yangi natijalar qo'shilgan — ular hali e'lon qilinmagan." : "Natijalar hali e'lon qilinmagan: o'quvchilar va markazlar Rash ballini ko'rmayapti."}</span>
              : <span style={{ color: C.successDark }}>Hammasiga e'lon qilingan: {new Date(publishedAt).toLocaleString("uz-UZ")}.</span>}
          </div>
          <button onClick={publish} disabled={publishing} style={{ ...S.btnSmall, background: "#1D4ED8", padding: "10px 16px", opacity: publishing ? 0.6 : 1, whiteSpace: "nowrap" }}>{publishing ? "⏳ E'lon qilinmoqda..." : "📢 E'lon qilish"}</button>
        </div>
      )}
      {test && !summary && <div style={{ ...S.card, padding: 20, textAlign: "center", color: C.textMid, fontSize: 13 }}>Bu test hali umumiy hisobda hisoblanmagan. "🔄 Qayta hisoblash" tugmasini bosing.</div>}
      {!test && <div style={S.empty}>Natijalarni ko'rish uchun testni tanlang</div>}

      {summary && (
        <>
          {newFiles > 0 && <div style={{ ...S.card, padding: "10px 14px", marginBottom: 12, background: "#FFFBEB", border: "1.5px solid #FDE68A", color: "#92400E", fontSize: 13, fontWeight: 600 }}>⏳ Oxirgi hisobdan keyin {newFiles} ta yangi fayl kelgan — "🔄 Qayta hisoblash"ni bosing.</div>}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(130px,1fr))", gap: 10, marginBottom: 12 }}>
            {[["O'quvchilar", withBall.length], ["Saytdan", filtered.filter(r => r.source === "Sayt").length], ["Fayllardan", filtered.filter(r => r.source !== "Sayt").length], ["O'rtacha ball", avg != null ? avg.toFixed(1) : "—"]].map(([l, v]) => (
              <div key={l} style={{ ...S.card, padding: "12px 14px" }}>
                <div style={{ fontSize: 22, fontWeight: 900, color: "#6D28D9", fontVariantNumeric: "tabular-nums" }}>{v}</div>
                <div style={{ fontSize: 12, color: C.textMid }}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
            {["A+", "A", "B+", "B", "C+", "C", "NC"].map(g => {
              const dc = g === "NC" ? C.danger : (g === "C" || g === "C+") ? C.warning : C.successDark;
              const cnt = gradeCounts[g] || 0;
              // Foiz — tanlangan (filtrlangan) ishtirokchilar soniga nisbatan
              const pc = withBall.length ? (cnt / withBall.length * 100) : 0;
              return <span key={g} style={{ ...S.badge, background: dc + "1A", color: dc, fontVariantNumeric: "tabular-nums" }}>{g}: {cnt} <span style={{ fontWeight: 600, opacity: 0.8 }}>({pc.toFixed(1)}%)</span></span>;
            })}
          </div>
          <details style={{ ...S.card, padding: "10px 14px", marginBottom: 12 }}>
            <summary style={{ cursor: "pointer", fontWeight: 700, fontSize: 13 }}>📂 Hisobga kirgan fayllar ({(summary.files || []).length}) • oxirgi hisob: {new Date(summary.calculatedAt).toLocaleString("uz-UZ")}</summary>
            <div style={{ marginTop: 8 }}>
              {(summary.files || []).length === 0 && <p style={{ margin: 0, fontSize: 13, color: C.textLight }}>Faqat saytdagi natijalar hisoblangan.</p>}
              {(summary.files || []).map(f => (
                <div key={f.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "7px 0", borderTop: `1px solid ${C.border}`, fontSize: 13 }}>
                  <span><b>{f.source}</b> — {f.rows} ta o'quvchi <span style={{ color: C.textLight }}>({new Date(f.uploadedAt).toLocaleDateString("uz-UZ")})</span></span>
                  <span style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                    {f.partnerId !== "admin" && <button onClick={() => adminCenterPdf(f)} disabled={pdfBusy === f.id} style={{ ...S.btnSmall, background: "#0891B2", padding: "4px 10px", fontSize: 12, opacity: pdfBusy === f.id ? 0.6 : 1 }}>{pdfBusy === f.id ? "⏳" : "📄 PDF"}</button>}
                    <button onClick={() => deleteFile(f)} style={{ ...S.btnSmall, background: C.danger, padding: "4px 10px", fontSize: 12 }}>🗑 Olib tashlash</button>
                  </span>
                </div>
              ))}
            </div>
          </details>
          <details style={{ ...S.card, padding: "10px 14px", marginBottom: 12 }}>
            <summary style={{ cursor: "pointer", fontWeight: 700, fontSize: 13 }}>🧾 PDF hisobot sozlamalari (markazlarga beriladigan hisobotdagi ma'lumotlar)</summary>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 10, marginTop: 10 }}>
              {[["orgName", "Tashkilotchi nomi (imzoda)"], ["telegram", "Telegram kanal (tepada)"], ["phone", "Telefon"], ["subtitle", "Sarlavha ostidagi yozuv"], ["adTitle", "Reklama bo'limi sarlavhasi"], ["socialTelegram", "Telegram (pastda)"], ["socialInstagram", "Instagram"], ["socialYoutube", "YouTube"]].map(([k, l]) => (
                <label key={k} style={{ fontSize: 11.5, color: C.textMid, fontWeight: 600 }}>{l}
                  <input id={"pdfcfg_" + k} value={pdfCfg[k] || ""} onChange={e => setPdfCfg(p => ({ ...p, [k]: e.target.value }))} style={{ ...S.input, marginTop: 4, marginBottom: 0, fontSize: 13 }} />
                </label>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 12, marginTop: 12 }}>
              {[["logo", "Logotip (sarlavha chap tomonida)", 360], ["adImage", "Reklama rasmi (umumiy hisobot oxirida)", 1200]].map(([k, l, max]) => {
                const src = pdfCfg[k] === undefined ? PDF_DEFAULT_IMAGES[k === "logo" ? "logo" : "ad"] : pdfCfg[k];
                return (
                  <div key={k} style={{ border: `1.5px dashed ${C.border}`, borderRadius: 10, padding: 10 }}>
                    <p style={{ margin: "0 0 8px", fontSize: 11.5, color: C.textMid, fontWeight: 600 }}>{l}</p>
                    {src ? <img src={src} alt="" style={{ maxHeight: 70, maxWidth: "100%", display: "block", marginBottom: 8, borderRadius: 6 }} /> : <p style={{ margin: "0 0 8px", fontSize: 12, color: C.textLight }}>Rasm qo'yilmaydi</p>}
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      <label style={{ ...S.btnSmall, background: C.primary, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>Yuklash<input type="file" accept="image/*" onChange={pickPdfImage(k, max)} style={{ display: "none" }} /></label>
                      {src && <button onClick={() => setPdfCfg(p => ({ ...p, [k]: "" }))} style={{ ...S.btnSmall, background: C.danger, padding: "5px 10px", fontSize: 12 }}>Olib tashlash</button>}
                      {pdfCfg[k] !== undefined && <button onClick={() => setPdfCfg(p => { const n = { ...p }; delete n[k]; return n; })} style={{ ...S.btnGhost, padding: "5px 10px", fontSize: 12 }}>Standart</button>}
                    </div>
                  </div>
                );
              })}
            </div>
            <button onClick={savePdfCfg} style={{ ...S.btnSmall, background: C.primary, marginTop: 10, padding: "9px 16px" }}>💾 Saqlash</button>
          </details>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <input value={query} onChange={e => { setQuery(e.target.value); setPage(0); }} placeholder="🔍 Ism yoki o'quv markazi bo'yicha qidirish" style={{ ...S.input, marginBottom: 0, maxWidth: 340, flex: "1 1 220px" }} />
            <select value={source} onChange={e => { setSource(e.target.value); setPage(0); }} style={{ ...S.input, marginBottom: 0, maxWidth: 260, flex: "1 1 180px" }}>
              <option value="">Barcha manbalar</option>
              {sources.map(x => <option key={x} value={x}>{x}</option>)}
            </select>
          </div>
          {pager}
          <div style={{ overflowX: "auto" }}>
            <table style={S.table}>
              <thead><tr>{["O'rin", "F.I.Sh", "O'quv markazi", "Manba", "To'g'ri", "Algebra", "Geometriya", "Umumiy ball", "Baho", "Foiz", "BMBA"].map(h => <th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>{shown.map((r, i) => {
                const dg = r.daraja; const dc = dg === "NC" ? C.danger : (dg === "C" || dg === "C+") ? C.warning : C.successDark;
                return (
                  <tr key={r._pos} style={{ background: i % 2 === 0 ? C.card : "#FAFBFF" }}>
                    <td style={S.td}>{r._pos}</td>
                    <td style={S.td}>{r.name}</td>
                    <td style={S.td}>{r.group || "-"}</td>
                    <td style={{ ...S.td, fontSize: 12, color: C.textMid }}>{r.source || "-"}</td>
                    <td style={S.td}>{r.correct ?? "-"}</td>
                    <td style={S.td}>{r.alg != null ? r.alg.toFixed(1) : "-"}</td>
                    <td style={S.td}>{r.geo != null ? r.geo.toFixed(1) : "-"}</td>
                    <td style={S.td}><b style={{ color: "#6D28D9" }}>{r.ball != null ? r.ball.toFixed(1) : "-"}</b></td>
                    <td style={S.td}>{dg ? <span style={{ ...S.badge, background: dc + "22", color: dc, fontWeight: 800 }}>{dg}</span> : "-"}</td>
                    <td style={S.td}>{r.foiz != null ? r.foiz.toFixed(1) + "%" : "-"}</td>
                    <td style={S.td}>{r.bmba != null ? r.bmba.toFixed(1) : "-"}</td>
                  </tr>
                );
              })}</tbody>
            </table>
          </div>
          {pages > 1 && pager}
        </>
      )}
    </div>
  );
}

function AdminPanel({ onLogout, isFullAdmin=true, teacherInfo=null }) {
  const roleName = isFullAdmin ? "Admin" : (teacherInfo?.name || "O'qituvchi");
  const [tab,setTab]=useState("tests");
  const [tests,setTests]=useState([]); const [users,setUsers]=useState([]); const [results,setResults]=useState([]);
  const [creating,setCreating]=useState(false); const [editing,setEditing]=useState(null);
  const [openResultTest,setOpenResultTest]=useState(null); // "Natijalar" bo'limida ochilgan test
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
  const [raschCalcPage,setRaschCalcPage]=useState(0);
  const [raschCalcQuery,setRaschCalcQuery]=useState("");
  const RASCH_PAGE_SIZE=100;
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
            setRaschCalcPage(0); setRaschCalcQuery("");
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
    if (!testEligibleForRasch(test)) { setRaschDoneMsg(`⚠️ "${test.name}" — Rash modeli faqat aynan ${RASCH_REQUIRED_ITEMS} ta savol/banddan iborat testlar uchun ishlaydi.`); setTimeout(()=>setRaschDoneMsg(null),6000); return; }
    setRaschCalcAttaching(true);
    setTimeout(() => {
      const settings = test.raschSettings || DEFAULT_RASCH_SETTINGS;
      const src = raschCalcRawRows.filter(x => x.items || x.itemsStr || typeof x.ball === "number");
      saveRaschUpload({ id: Date.now(), partnerId: "admin", fileName: raschCalcFileName || "Excel fayl", testId: test.id, testName: test.name, uploadedAt: Date.now(), status: "pending", rawRows: packUploadRows(src) });
      const r = calculateRaschCombined(test, settings);
      reload();
      let msg = `💾 "${test.name}" — sayt va barcha yuklangan fayllar (${r.filesCount} ta) bitta umumiy hisobda qayta hisoblandi: jami ${r.allRows.length} ta o'quvchi, ${r.matched} tasi saytdagi profiliga yozildi. Natija "🏆 Umumiy natijalar" bo'limida.`;
      setRaschDoneMsg(msg);
      setRaschCalcAttaching(false);
      setTimeout(()=>setRaschDoneMsg(null), 8000);
    }, 30);
  };

  const triggerRaschUpload = (test) => {
    if (!testEligibleForRasch(test)) { setRaschDoneMsg(`⚠️ "${test.name}" — Rash modeli faqat aynan ${RASCH_REQUIRED_ITEMS} ta savol/banddan iborat testlar uchun ishlaydi. Bu testda ${testTotalItems(test)} ta bor, shuning uchun fayl yuklab bo'lmaydi.`); setTimeout(()=>setRaschDoneMsg(null),6000); return; }
    setRaschUploadTarget(test); requestAnimationFrame(()=>raschFileInputRef.current?.click());
  };
  const handleRaschFile = (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file || !raschUploadTarget) return;
    const test = raschUploadTarget;
    if (!testEligibleForRasch(test)) { setRaschDoneMsg(`⚠️ "${test.name}" — ${RASCH_REQUIRED_ITEMS} talik test emas, Rashda hisoblanmaydi.`); setRaschUploadTarget(null); setTimeout(()=>setRaschDoneMsg(null),6000); return; }
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
          saveRaschUpload({ id: Date.now(), partnerId: "admin", fileName: file.name, testId: test.id, testName: test.name, uploadedAt: Date.now(), status: "pending", rawRows: packUploadRows(rows) });
          const r = calculateRaschCombined(test, settings);
          reload();
          let msg = `📤 "${test.name}" — sayt va barcha yuklangan fayllar (${r.filesCount} ta) bitta umumiy Rash hisobida qayta ishlandi: jami ${r.allRows.length} ta o'quvchi, ${r.matched} tasi saytdagi profiliga yozildi. Natija "🏆 Umumiy natijalar" bo'limida.`;
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
    if (!testEligibleForRasch(test)) { setRaschDoneMsg(`⚠️ "${test.name}" — Rash modeli faqat aynan ${RASCH_REQUIRED_ITEMS} ta savol/banddan iborat testlar uchun ishlaydi. Bu testda ${testTotalItems(test)} ta bor.`); setTimeout(()=>setRaschDoneMsg(null),6000); return; }
    setRaschModal({ test, settings: { ...(test.raschSettings || DEFAULT_RASCH_SETTINGS) } });
  };
  // "🏆 Umumiy natijalar"dagi "🔄 Qayta hisoblash" — sozlamalar oynasisiz, darhol hisoblaydi
  const recalcRaschNow = (test) => {
    if (!testEligibleForRasch(test)) return;
    setRaschBusyId(test.id);
    setTimeout(() => {
      const r = calculateRaschCombined(test, test.raschSettings || DEFAULT_RASCH_SETTINGS);
      setRaschBusyId(null);
      reload(); setRaschVersion(v=>v+1);
      setRaschDoneMsg(r.allRows.length>0
        ? `🔄 "${test.name}" qayta hisoblandi: jami ${r.allRows.length} ta o'quvchi.`
        : `"${test.name}" bo'yicha hali natija yo'q.`);
      setTimeout(()=>setRaschDoneMsg(null), 4000);
    }, 30);
  };
  const runRasch = () => {
    const { test, settings } = raschModal;
    setRaschBusyId(test.id);
    setTimeout(() => { // UI qotib qolmasligi uchun keyingi tikda hisoblaymiz
      const r = calculateRaschCombined(test, settings);
      setRaschBusyId(null);
      setRaschModal(null);
      reload();
      const totalCount = r.matched + (r.fileOnly?.length || 0);
      let msg = r.allRows.length>0
        ? `🎯 "${test.name}" — jami ${r.allRows.length} ta o'quvchi bitta umumiy hisobda baholandi, ${r.matched} tasi saytdagi profiliga yozildi.`
        : `"${test.name}" bo'yicha hali hech kim test topshirmagan va hamkor markazlardan ham fayl kelmagan.`;
      if (r.partnerBatchesProcessed>0) msg += ` ${r.partnerCentersCount} ta hamkor markazning fayli ham shu hisobga qo'shildi.`;
      if (r.allRows.length>0) msg += ` Natija "🏆 Umumiy natijalar" bo'limida.`;
      setRaschDoneMsg(msg);
      setTimeout(()=>setRaschDoneMsg(null), 5500);
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

  // ===== Saytda topshirilgan natijalarni umumiy Rash hisobiga AVTOMATIK qo'shish =====
  // Admin/o'qituvchi paneli ochiq turganda har 5 soniyada (va boshqa qurilmadan yangi ma'lumot
  // kelganda) tekshiriladi: biror 55 talik test bo'yicha saytda yangi natija paydo bo'lgan
  // (yoki o'chirilgan) bo'lsa — shu test uchun umumiy hisob (sayt + barcha markaz fayllari)
  // qayta hisoblanadi. Hisob o'quvchining qurilmasida emas, shu yerda bajariladi.
  // O'quv markazlari baribir faqat admin "📢 e'lon qilish"ni bosgandan keyin ko'radi.
  const [raschVersion,setRaschVersion]=useState(0);
  // Eski ochiq parollarni fon rejimida xeshlash (faqat to'liq admin panelida, bir marta)
  useEffect(()=>{ if(isFullAdmin && countPlainPasswords()>0) migratePlainPasswords().then(()=>reload()).catch(e=>console.error("[Parol migratsiyasi]",e)); },[]);
  const autoRaschBusy=useRef(false);
  useEffect(()=>{
    const check=()=>{
      if(autoRaschBusy.current) return;
      const ts=db.get("tests")||[], rs=db.get("results")||[], sums=db.get("raschResults")||[];
      const stale=ts.filter(t=>{
        if(!testEligibleForRasch(t)) return false;
        const site=rs.filter(r=>r.testId===t.id);
        const sum=sums.find(x=>x.id===t.id);
        if(!sum) return site.length>0;
        const siteInSum=(sum.rows||[]).filter(r=>r.source==="Sayt"||/saytda ham bor/.test(r.source||"")).length;
        return siteInSum!==site.length || site.some(r=>r.id>sum.calculatedAt);
      });
      if(!stale.length) return;
      autoRaschBusy.current=true;
      setTimeout(()=>{
        try { stale.forEach(t=>calculateRaschCombined(t, t.raschSettings||DEFAULT_RASCH_SETTINGS)); }
        catch(e){ console.error("[Rasch auto]", e); }
        autoRaschBusy.current=false;
        reload(); setRaschVersion(v=>v+1);
      },50);
    };
    check();
    const iv=setInterval(check,5000);
    window.addEventListener("firestore-sync",check);
    return ()=>{ clearInterval(iv); window.removeEventListener("firestore-sync",check); };
  },[]);
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
  const toggleActive=(id)=>{db.set("tests",(db.get("tests")||[]).map(t=>t.id===id?{...t,active:!t.active,startedAt:!t.active?Date.now():null,scheduledAt:null,everActivated:t.everActivated||!t.active}:t));reload();};
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
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: C.textMid, lineHeight: 1.5 }}>"<b>{raschModal.test.name}</b>" testini topshirgan barcha o'quvchilarning natijasi Rash modeli asosida qayta baholanadi: har bir savolga qiyinligiga qarab vazn beriladi, Algebra (2/3) va Geometriya (1/3) alohida hisoblanadi — "RASH_30000" Excel shabloni bilan bir xil. Baho chegaralarini xohlasangiz o'zgartiring.</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
              {[
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
      {isFullAdmin && adminUsesDefaultPassword() && (
        <div style={{background:C.dangerLight,borderBottom:`1.5px solid #FECACA`,color:C.danger,padding:"10px 20px",fontSize:13,fontWeight:700,display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <span>⚠️ Admin paroli standart holatda. Xavfsizlik uchun uni darhol almashtiring.</span>
          <button onClick={()=>setTab("users")} style={{...S.btnSmall,background:C.danger,padding:"6px 12px",fontSize:12}}>Almashtirish</button>
        </div>
      )}

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

      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,display:"flex",padding:"0 16px",overflowX:"auto"}}>
        {[["tests","📋 Testlar"],["users","👥 O'quvchilar"],["results","📊 Natijalar"],["combined","🏆 Umumiy natijalar"]].map(([t,l])=>(
          <button key={t} onClick={()=>{setTab(t);setCreating(false);setEditing(null);setOpenResultTest(null);}} style={{padding:"14px 18px",background:"none",border:"none",cursor:"pointer",color:tab===t?C.primary:C.textMid,fontWeight:tab===t?800:500,borderBottom:tab===t?`3px solid ${C.primary}`:"3px solid transparent",fontSize:14,whiteSpace:"nowrap",flexShrink:0}}>{l}</button>
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
            {!(creating||editing)&&(<>{tests.map(test=>{
              const tr=results.filter(r=>r.testId===test.id);
              const testFiles=(db.get("partnerUploads")||[]).filter(u=>u.testId===test.id);
              const pendingPartners=testFiles.length;
              const unpublishedPartners=testFiles.filter(u=>u.partnerId!=="admin"&&!u.publishedRows).length;
              const endAt=test.startedAt?new Date(test.startedAt+test.duration*60000).toLocaleTimeString("uz-UZ",{hour:"2-digit",minute:"2-digit"}):null;
              return (
                <div key={test.id} style={{...S.card,padding:18,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                    <div>
                      <h4 style={{margin:"0 0 5px",fontSize:16}}>{test.name}</h4>
                      <p style={{margin:"0 0 8px",color:C.textMid,fontSize:13}}>{testTotalItems(test)} savol • {test.duration} daqiqa</p>
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
                        
                        {unpublishedPartners>0&&<span onClick={()=>setTab("combined")} style={{...S.badge,background:"#FEF3C7",color:"#92400E",cursor:"pointer"}}>📢 {unpublishedPartners} ta markaz natijasi e'lon qilinmagan</span>}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      <button onClick={()=>toggleActive(test.id)} style={{...S.btnSmall,background:test.active?C.danger:C.success}}>{test.active?"⛔ To'xtatish":"✅ Faollashtirish"}</button>
                      <button onClick={()=>{setEditing(test);setCreating(false);}} style={{...S.btnSmall,background:C.primary}}>✏️ Tahrirlash</button>
                      <button onClick={()=>handleRegrade(test)} disabled={regradingId===test.id} style={{...S.btnSmall,background:regradingId===test.id?"#94A3B8":"#7C3AED"}}>{regradingId===test.id?"⏳ Baholanmoqda...":"🔄 Qayta baholash"}</button>
                      
                      <button onClick={()=>setExportModal(buildExcelExport(test,results,users))} style={{...S.btnSmall,background:C.successDark}}>📥 Excel</button>
                      <button onClick={()=>triggerRaschUpload(test)} disabled={!testEligibleForRasch(test)||raschUploading} title={!testEligibleForRasch(test)?`Faqat ${RASCH_REQUIRED_ITEMS} talik testlar uchun`:undefined} style={{...S.btnSmall,background:!testEligibleForRasch(test)?"#CBD5E1":"#0891B2",opacity:raschUploading?0.6:1}}>{raschUploading&&raschUploadTarget?.id===test.id?"⏳ Yuklanmoqda...":"📤 Natija yuklash"}</button>
                      <button onClick={()=>deleteTest(test.id)} style={{...S.btnSmall,background:C.danger}}>🗑️</button>
                    </div>
                  </div>
                </div>
              );
            })}
            {tests.length===0&&<div style={S.empty}>Hali test yaratilmagan</div>}</>)}
          </div>
        )}

        {tab==="users"&&(
          <div>
            {/* Teacher manager — only full admin sees this */}
            {isFullAdmin && <SecurityPanel />}
            {isFullAdmin && <TeacherManager />}
            {isFullAdmin && <PartnerManager />}
            <GroupManager />
            <h3 style={{marginBottom:16}}>O'quvchilar ({users.length})</h3>
            <div style={{overflowX:"auto"}}>
              <table style={S.table}>
                <thead><tr>{["#","Ism","Familiya","Guruh","ID / Telefon","Testlar","Amal"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>{users.map((u,i)=>(
                  <tr key={u.phone} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                    <td style={S.td}>{i+1}</td><td style={S.td}>{u.firstName}</td><td style={S.td}>{u.lastName}</td>
                    <td style={S.td}>{u.group||<span style={{color:C.textLight}}>—</span>}</td>
                    <td style={S.td}><code style={{color:C.primary,background:C.primaryLight,padding:"2px 6px",borderRadius:4}}>{u.phone}</code></td>
                    <td style={S.td}>{results.filter(r=>r.userPhone===u.phone).length}</td>
                    <td style={S.td}>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={()=>{
                          const np = prompt("Yangi parol kiriting ("+u.firstName+" "+u.lastName+" uchun):");
                          if(np&&np.length>=4){
                            hashPassword(np).then(h=>{
                              const us=(db.get("users")||[]).map(x=>x.phone===u.phone?{...x,password:h}:x);
                              db.set("users",us); reload();
                              alert("Parol yangilandi!");
                            }).catch(e=>alert(e.message));
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
            {openResultTest===null && results.length>0 && <p style={{margin:"0 0 12px",fontSize:13,color:C.textMid}}>Natijalarni ko'rish uchun testni tanlang.</p>}
            {openResultTest===null && <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:12}}>
              {tests.map(test=>{
                const tr=results.filter(r=>r.testId===test.id);
                if(!tr.length) return null;
                const total=testTotalItems(test)||1;
                const avg=Math.round(tr.reduce((a,r)=>a+(r.totalScore||0),0)/tr.length/total*100);
                const last=Math.max(...tr.map(r=>r.id));
                return (
                  <button key={test.id} onClick={()=>setOpenResultTest(test.id)} style={{...S.card,padding:16,textAlign:"left",cursor:"pointer",display:"flex",flexDirection:"column",gap:8,font:"inherit",color:C.text}}>
                    <span style={{fontWeight:800,fontSize:15,color:C.primary}}>{test.name}</span>
                    <span style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                      <span style={{...S.badge,background:C.primaryLight,color:C.primary}}>👥 {tr.length} ta topshirdi</span>
                      <span style={{...S.badge,background:avg>=70?C.successLight:avg>=50?C.warningLight:C.dangerLight,color:avg>=70?C.successDark:avg>=50?C.warning:C.danger}}>o'rtacha {avg}%</span>
                    </span>
                    <span style={{fontSize:12,color:C.textLight}}>{testTotalItems(test)} savol • oxirgi: {new Date(last).toLocaleDateString("uz-UZ")}</span>
                    <span style={{fontSize:12.5,color:C.primary,fontWeight:700}}>Ko'rish →</span>
                  </button>
                );
              })}
            </div>}
            {openResultTest!==null && tests.map(test=>{
              if(test.id!==openResultTest) return null;
              const tr=results.filter(r=>r.testId===test.id);
              return (
                <div key={test.id} style={{marginBottom:24}}>
                  <div style={{display:"flex",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:10}}>
                    <button onClick={()=>setOpenResultTest(null)} style={{...S.btnGhost,padding:"7px 12px"}}>← Barcha testlar</button>
                    <h4 style={{margin:0,color:C.primary,fontSize:16}}>{test.name}</h4>
                    <span style={{...S.badge,background:C.primaryLight,color:C.primary}}>{tr.length} ta topshirdi</span>
                  </div>
                  {tr.length===0&&<div style={S.empty}>Bu testni hali hech kim topshirmagan</div>}
                  <div style={{overflowX:"auto"}}>
                    <table style={S.table}>
                      <thead><tr>{["#","F.I.O","Guruh","Ball","Foiz",...(testEligibleForRasch(test)?["Rash ball","Daraja"]:[]),"Vaqt","Sana"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>{tr.sort((a,b)=>b.totalScore-a.totalScore).map((r,i)=>{
                        const u=users.find(u=>u.phone===r.userPhone);
                        const pct=Math.round((r.totalScore/testTotalItems(test))*100);
                        const dGrade=r.rasch?.daraja;
                        const dColor=dGrade==="NC"?C.danger:dGrade==="C"||dGrade==="C+"?C.warning:C.successDark;
                        return (
                          <tr key={r.id} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                            <td style={S.td}>{i+1}</td>
                            <td style={S.td}>{u?`${u.firstName} ${u.lastName}`:r.userPhone}</td>
                            <td style={S.td}>{u?.group||"-"}</td>
                            <td style={S.td}><b style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger}}>{r.totalScore}</b>/{testTotalItems(test)}</td>
                            <td style={S.td}><span style={{color:pct>=70?C.successDark:pct>=50?C.warning:C.danger,fontWeight:700}}>{pct}%</span></td>
                            {testEligibleForRasch(test)&&<td style={S.td}>{r.rasch?<b style={{color:"#6D28D9"}}>{r.rasch.ball!=null?r.rasch.ball.toFixed(1):"—"}</b>:<span style={{color:C.textLight}}>—</span>}</td>}
                            {testEligibleForRasch(test)&&<td style={S.td}>{dGrade?<span style={{...S.badge,background:dColor+"22",color:dColor,fontWeight:800}}>{dGrade}</span>:<span style={{color:C.textLight}}>—</span>}</td>}
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

        {tab==="combined"&&(
          <RaschCombinedView tests={tests} onExport={setExportModal} onRecalc={recalcRaschNow} busyId={raschBusyId} version={raschVersion}/>
        )}

        {tab==="raschcalc"&&(
          <div>
            <h3 style={{margin:"0 0 14px"}}>🎯 Rash modeli — Excel fayldan hisoblash</h3>
            

            <div style={{...S.card,padding:16,marginBottom:16}}>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <button onClick={()=>raschCalcFileInputRef.current?.click()} disabled={raschCalcBusy} style={{...S.btnSmall,background:"#6D28D9",padding:"11px 20px",fontSize:13,opacity:raschCalcBusy?0.6:1}}>{raschCalcBusy?"⏳ Hisoblanmoqda...":"📤 Excel faylni yuklash"}</button>
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
                    {tests.filter(t=>testEligibleForRasch(t)).map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
                  </select>
                  <button onClick={attachRaschCalcToTest} disabled={!raschCalcAttachTestId||raschCalcAttaching} style={{...S.btnSmall,background:raschCalcAttachTestId?"#6D28D9":"#CBD5E1",padding:"11px 20px",fontSize:13,opacity:raschCalcAttaching?0.6:1}}>{raschCalcAttaching?"⏳ Saqlanmoqda...":"💾 Ushbu testga saqlash"}</button>
                </div>
                <p style={{margin:"8px 0 0",fontSize:11,color:C.textLight}}>Tanlangan test bo'yicha saytda topshirganlar + shu fayl birgalikda bitta Rash hisobida qayta hisoblanadi va mos o'quvchilarning profiliga yoziladi. Faqat aynan {RASCH_REQUIRED_ITEMS} talik testlar ro'yxatda ko'rinadi.</p>
              </div>
            )}

            {raschCalcRows&&raschCalcRows.length>0&&(()=>{
              // Reyting o'rni butun ro'yxat bo'yicha; qidiruv va sahifalash faqat ko'rinishni o'zgartiradi
              const ranked=[...raschCalcRows].sort((a,b)=>(b.ball??-999)-(a.ball??-999)).map((r,i)=>({...r,_pos:i+1}));
              const q=raschCalcQuery.trim().toLowerCase();
              const filtered=q?ranked.filter(r=>(r.name||"").toLowerCase().includes(q)||(r.group||"").toLowerCase().includes(q)):ranked;
              const pages=Math.max(1,Math.ceil(filtered.length/RASCH_PAGE_SIZE));
              const page=Math.min(raschCalcPage,pages-1);
              const shown=filtered.slice(page*RASCH_PAGE_SIZE,(page+1)*RASCH_PAGE_SIZE);
              const pager=(
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",margin:"10px 0"}}>
                  <button onClick={()=>setRaschCalcPage(0)} disabled={page===0} style={{...S.btnGhost,padding:"6px 10px",opacity:page===0?0.4:1}}>«</button>
                  <button onClick={()=>setRaschCalcPage(page-1)} disabled={page===0} style={{...S.btnGhost,padding:"6px 12px",opacity:page===0?0.4:1}}>← Oldingi</button>
                  <span style={{fontSize:13,color:C.textMid}}><b style={{color:C.text}}>{page+1}</b> / {pages} sahifa • {filtered.length} ta natija</span>
                  <button onClick={()=>setRaschCalcPage(page+1)} disabled={page>=pages-1} style={{...S.btnGhost,padding:"6px 12px",opacity:page>=pages-1?0.4:1}}>Keyingi →</button>
                  <button onClick={()=>setRaschCalcPage(pages-1)} disabled={page>=pages-1} style={{...S.btnGhost,padding:"6px 10px",opacity:page>=pages-1?0.4:1}}>»</button>
                </div>
              );
              return (
              <div>
                <input value={raschCalcQuery} onChange={e=>{setRaschCalcQuery(e.target.value);setRaschCalcPage(0);}} placeholder="🔍 Ism yoki o'quv markazi bo'yicha qidirish" style={{...S.input,maxWidth:380,marginBottom:0}}/>
                {pager}
              <div style={{overflowX:"auto"}}>
                <table style={S.table}>
                  <thead><tr>{["O'rin","F.I.Sh","O'quv markazi","To'g'ri","Algebra","Geometriya","Umumiy ball","Baho","Foiz","BMBA"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>{shown.map((r,i)=>{
                    const dg=r.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                    return (
                      <tr key={r._pos} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                        <td style={S.td}>{r._pos}</td>
                        <td style={S.td}>{r.name}</td>
                        <td style={S.td}>{r.group||"-"}</td>
                        <td style={S.td}>{r.correct??"-"}</td>
                        <td style={S.td}>{r.alg!=null?r.alg.toFixed(1):"-"}</td>
                        <td style={S.td}>{r.geo!=null?r.geo.toFixed(1):"-"}</td>
                        <td style={S.td}><b style={{color:"#6D28D9"}}>{r.ball!=null?r.ball.toFixed(1):"-"}</b></td>
                        <td style={S.td}>{dg?<span style={{...S.badge,background:dc+"22",color:dc,fontWeight:800}}>{dg}</span>:"-"}</td>
                        <td style={S.td}>{r.foiz!=null?r.foiz.toFixed(1)+"%":"-"}</td>
                        <td style={S.td}>{r.bmba!=null?r.bmba.toFixed(1):"-"}</td>
                      </tr>
                    );
                  })}</tbody>
                </table>
              </div>
                {pages>1&&pager}
              </div>
              );
            })()}
          </div>
        )}
      </div>
    </div>
  );
}

// ===== STUDENT DASHBOARD =====
// ===== HAMKOR MARKAZ PANELI =====
// Juda cheklangan: faqat testlar nomini ko'radi (tanlash uchun), Excel faylini
// yuklaydi, va shu bitta amal — saytdagi natijalar bilan birgalikda Rash modeliga
// qo'shiladi. O'quvchilar ro'yxati, boshqa hamkorlar, testlar tarkibi, statistika —
// hech biri ko'rinmaydi.
function PartnerPanel({ partnerInfo, onLogout }) {
  const [tab, setTab] = useState("upload");
  const [tests, setTests] = useState(() => db.get("tests") || []);
  const [selectedTestId, setSelectedTestId] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);
  const [err, setErr] = useState(null);
  const fileInputRef = useRef(null);
  const [uploads, setUploads] = useState(() => (db.get("partnerUploads")||[]).filter(u=>u.partnerId===partnerInfo?.id));
  const [printBatch, setPrintBatch] = useState(null); // PDF (chop etish) uchun tanlangan yuklama
  const [logo, setLogo] = useState(() => (db.get("partners")||[]).find(p=>p.id===partnerInfo?.id)?.logo || null);
  const [logoErr, setLogoErr] = useState(null);
  const logoInputRef = useRef(null);
  const [docModal, setDocModal] = useState(null); // {type:"pdf"|"latex", url?, source?, name}
  const [answersModal, setAnswersModal] = useState(null); // {test}

  const reloadUploads = () => setUploads((db.get("partnerUploads")||[]).filter(u=>u.partnerId===partnerInfo?.id));

  const handleLogoFile = (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file || !partnerInfo?.id) return;
    if (!file.type.startsWith("image/")) { setLogoErr("Faqat rasm fayl (PNG/JPG) yuklang!"); return; }
    if (file.size > 2*1024*1024) { setLogoErr("Rasm hajmi 2MB dan oshmasin!"); return; }
    setLogoErr(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      const dataUrl = ev.target.result;
      db.set("partners", (db.get("partners")||[]).map(p=>p.id===partnerInfo.id?{...p,logo:dataUrl}:p));
      setLogo(dataUrl);
    };
    reader.onerror = () => setLogoErr("Rasmni o'qishda xatolik yuz berdi.");
    reader.readAsDataURL(file);
  };
  const removeLogo = () => {
    db.set("partners", (db.get("partners")||[]).map(p=>p.id===partnerInfo.id?{...p,logo:null}:p));
    setLogo(null);
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file || !selectedTestId) return;
    const test = tests.find(t => t.id === Number(selectedTestId) || t.id === selectedTestId);
    if (!test) return;
    if (!(test.active || test.everActivated)) { setErr("Bu test hali faollashtirilmagan — natija yuklab bo'lmaydi."); return; }
    if (!testEligibleForRasch(test)) { setErr(`Rash modeli faqat aynan ${RASCH_REQUIRED_ITEMS} ta savol/banddan iborat testlar uchun ishlaydi. Bu testda ${testTotalItems(test)} ta bor, shuning uchun fayl yuklab bo'lmaydi.`); return; }
    setBusy(true); setMsg(null); setErr(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const { rows } = parseUploadedResultsSheet(wb);
        if (!rows.length) {
          setErr("Faylda F.I.O / Ism ustuni topilmadi. Fayl tuzilishini tekshiring.");
        } else {
          // MUHIM: bu yerda Rash HISOBLANMAYDI — fayl faqat "kutilmoqda" holatida
          // saqlanadi. Sayt va barcha hamkor markazlarning natijalari bir joyga
          // yig'ilgach, admin/o'qituvchi "🎯 Rash modeli" tugmasini bosgandagina
          // hammasi BIRGALIKDA bitta hisobga solinadi.
          // Fayl saqlanadi (shu markazning shu testga oldingi fayli bo'lsa — almashtiriladi) va
          // DARHOL sayt + BARCHA markazlar fayllari bilan birga bitta umumiy Rash hisobi qayta
          // hisoblanadi. Natija admin panelidagi "🏆 Umumiy natijalar" bo'limida ko'rinadi.
          const usable = rows.filter(x => x.items || typeof x.ball === "number");
          if (!usable.length) { setErr(`Faylda ${RASCH_REQUIRED_ITEMS} ta javob ustuni (1–35, 36(a)–45(b)) topilmadi. Fayl tuzilishini tekshiring.`); setBusy(false); return; }
          const replaced = saveRaschUpload({ id: Date.now(), partnerId: partnerInfo.id, partnerName: partnerInfo.name, testId: test.id, testName: test.name, uploadedAt: Date.now(), status: "pending", rawRows: packUploadRows(usable) });
          calculateRaschCombined(test, test.raschSettings || DEFAULT_RASCH_SETTINGS);
          reloadUploads();
          setMsg(`✅ Natijangiz qabul qilindi (${usable.length} ta o'quvchi, "${test.name}")${replaced?" — oldingi faylingiz almashtirildi":""}. Tez orada Rash modelida hisoblanadi.`);
        }
      } catch {
        setErr("Faylni o'qib bo'lmadi. .xlsx formatida ekanligini tekshiring.");
      }
      setBusy(false);
    };
    reader.onerror = () => { setBusy(false); setErr("Faylni o'qishda xatolik yuz berdi."); };
    reader.readAsArrayBuffer(file);
  };

  const [pdfReady, setPdfReady] = useState(null);
  const [pdfBusyId, setPdfBusyId] = useState(null);
  const printPDF = async (batch) => {
    setPdfBusyId(batch.id); setErr(null);
    try {
      const centerName = (db.get("partners")||[]).find(p=>p.id===partnerInfo?.id)?.name || partnerInfo?.name || "O'quv markazi";
      const pdf = await buildCenterReportPdf({ centerName, testName: batch.testName, rows: batch.publishedRows || [], stats: batch.publishedStats });
      setPdfReady(pdf);
    } catch (e) { setErr(e.message || "PDF yaratishda xatolik yuz berdi."); setTab("results"); }
    setPdfBusyId(null);
  };

  // Statistika: barcha yuklamalar bo'yicha umumlashtirilgan ko'rsatkichlar
  const allRows = uploads.flatMap(u => (u.publishedRows||[]).map(r => ({ ...r, testName: u.testName })));
  const withBall = allRows.filter(r => typeof r.ball === "number");
  const avgBall = withBall.length ? (withBall.reduce((a,b)=>a+b.ball,0)/withBall.length) : null;
  const darajaCounts = {};
  withBall.forEach(r => { const d = r.daraja||"—"; darajaCounts[d] = (darajaCounts[d]||0)+1; });
  const DARAJA_ORDER = ["NC","C","C+","B","B+","A","A+"];
  const darajaColor = d => d==="NC"?C.danger:(d==="C"||d==="C+")?C.warning:C.successDark;

  return (
    <div style={{ minHeight: "100vh", background: C.bg }}>
      {/* Chop etish (PDF) uchun maxsus uslub — faqat #partner-print-area ko'rinadi */}
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #partner-print-area, #partner-print-area * { visibility: visible; }
          #partner-print-area { position: absolute; left: 0; top: 0; width: 100%; }
        }
      `}</style>
      {printBatch && (
        <div id="partner-print-area" style={{ padding: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 4 }}>
            {logo && <img src={logo} alt="logo" style={{ width: 48, height: 48, borderRadius: 8, objectFit: "cover" }}/>}
            <h2 style={{ margin: 0 }}>{partnerInfo?.name}</h2>
          </div>
          <p style={{ margin: "0 0 4px", color: "#555" }}>Test: {printBatch.testName}</p>
          <p style={{ margin: "0 0 16px", color: "#555" }}>Sana: {new Date(printBatch.uploadedAt).toLocaleDateString("uz-UZ")}</p>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr>{["#","F.I.O","To'g'ri","BALL","Daraja"].map(h=>(
              <th key={h} style={{ border: "1px solid #999", padding: "6px 8px", background: "#eee", textAlign: "left" }}>{h}</th>
            ))}</tr></thead>
            <tbody>{[...(printBatch.publishedRows||[])].sort((a,b)=>(b.ball??-999)-(a.ball??-999)).map((r,i)=>(
              <tr key={i}>
                <td style={{ border: "1px solid #999", padding: "6px 8px" }}>{i+1}</td>
                <td style={{ border: "1px solid #999", padding: "6px 8px" }}>{r.name}</td>
                <td style={{ border: "1px solid #999", padding: "6px 8px" }}>{r.correct??"-"}/{r.total??"-"}</td>
                <td style={{ border: "1px solid #999", padding: "6px 8px" }}>{r.ball!=null?r.ball.toFixed(1):"-"}</td>
                <td style={{ border: "1px solid #999", padding: "6px 8px" }}>{r.daraja||"-"}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}

      <div style={{ background: "linear-gradient(135deg,#0891B2,#0E7490)", padding: "18px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div onClick={()=>logoInputRef.current?.click()} title="Logotipni o'zgartirish" style={{
            width: 44, height: 44, borderRadius: 12, background: "rgba(255,255,255,0.15)",
            display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer",
            overflow: "hidden", flexShrink: 0, border: "1.5px solid rgba(255,255,255,0.35)",
          }}>
            {logo ? <img src={logo} alt="logo" style={{ width: "100%", height: "100%", objectFit: "cover" }}/> : <span style={{ fontSize: 22 }}>🤝</span>}
          </div>
          <input ref={logoInputRef} type="file" accept="image/*" onChange={handleLogoFile} style={{ display: "none" }}/>
          <div>
            <p style={{ margin: 0, color: "white", fontWeight: 800, fontSize: 17 }}>{partnerInfo?.name || "Hamkor markaz"}</p>
            <p style={{ margin: 0, fontSize: 11.5 }}>
              <span onClick={()=>logoInputRef.current?.click()} style={{ color: "rgba(255,255,255,0.85)", cursor: "pointer", textDecoration: "underline" }}>{logo ? "Logotipni o'zgartirish" : "Logotip yuklash"}</span>
              {logo && <span onClick={removeLogo} style={{ color: "rgba(255,255,255,0.85)", cursor: "pointer", textDecoration: "underline", marginLeft: 8 }}>O'chirish</span>}
            </p>
          </div>
        </div>
        <button onClick={onLogout} style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 10, color: "white", padding: "8px 14px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Chiqish</button>
      </div>
      {logoErr && <div style={{ background:"#FEF2F2", color:"#991B1B", padding:"8px 20px", fontSize:12.5, fontWeight:600 }}>⚠️ {logoErr}</div>}

      <div style={{ display: "flex", background: "white", borderBottom: `1.5px solid ${C.border}` }}>
        {[["upload","📤 Yuklash"],["view","🧾 Testlar"],["results","📊 Natijalarim"],["stats","📈 Statistika"]].map(([t,l])=>(
          <button key={t} onClick={()=>{setTab(t); if(t!=="upload") reloadUploads();}} style={{
            flex:1, padding:"13px 4px", background: tab===t?"#ECFEFF":"transparent",
            border:"none", borderBottom: tab===t?"3px solid #0891B2":"3px solid transparent",
            color: tab===t?"#0891B2":C.textMid, fontWeight: tab===t?800:600, fontSize:13, cursor:"pointer",
          }}>{l}</button>
        ))}
      </div>

      <div style={{ padding: 20, maxWidth: 560, margin: "0 auto" }}>

        {tab==="upload" && (
          <div style={{ ...S.card, padding: 18 }}>
            <p style={{ margin: "0 0 14px", fontSize: 13, color: C.textMid, lineHeight: 1.6 }}>Tegishli testni tanlab, o'z markazingizning Excel natijalar faylini yuklaysiz. Natija shu testni saytda topshirganlar bilan birgalikda bitta Rash hisobida qayta hisoblanadi.</p>

            <label style={S.label}>Test tanlang</label>
            <select value={selectedTestId} onChange={e=>setSelectedTestId(e.target.value)} style={S.input}>
              <option value="">— Testni tanlang —</option>
              {tests.filter(t=>(t.active||t.everActivated)&&testEligibleForRasch(t)).map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
            </select>
            <p style={{margin:"-6px 0 12px",fontSize:11,color:C.textLight}}>Faqat kamida bir marta faollashtirilgan VA aynan {RASCH_REQUIRED_ITEMS} ta savol/banddan iborat testlar ro'yxatda ko'rinadi.</p>

            <button onClick={()=>fileInputRef.current?.click()} disabled={!selectedTestId || busy}
              style={{ width: "100%", padding: "13px", borderRadius: 10, border: "none", cursor: (!selectedTestId||busy)?"not-allowed":"pointer",
                background: (!selectedTestId||busy) ? "#CBD5E1" : "#0891B2", color: "white", fontWeight: 800, fontSize: 14, marginTop: 6 }}>
              {busy ? "⏳ Yuklanmoqda..." : "📤 Excel faylni yuklash"}
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }}/>

            {msg && <div style={{ marginTop: 14, padding: "12px 14px", borderRadius: 10, background: "#ECFDF5", color: "#065F46", fontSize: 13, fontWeight: 600, lineHeight: 1.5 }}>{msg}</div>}
            {err && <div style={{ marginTop: 14, padding: "12px 14px", borderRadius: 10, background: "#FEF2F2", color: "#991B1B", fontSize: 13, fontWeight: 600 }}>⚠️ {err}</div>}
          </div>
        )}

        {tab==="view" && (
          <div>
            <p style={{margin:"0 0 14px",fontSize:12.5,color:C.textLight,lineHeight:1.5}}>Testning savol variantini (PDF/LaTeX) faollashtirilgach ko'rishingiz mumkin, lekin uni topshira olmaysiz. Test tugagach (admin to'xtatgach), to'g'ri javoblar shu yerda ochiladi.</p>
            {tests.length===0 && <div style={{...S.card,padding:24,textAlign:"center",color:C.textLight,fontSize:13}}>Hozircha testlar yo'q</div>}
            {tests.map(test=>{
              const langs = availableDocLangs(test);
              const activated = !!(test.active || test.everActivated); // hech bo'lmasa bir marta faollashtirilganmi
              // Vaqti tugagan bo'lsa ham (admin hali "to'xtatish" tugmasini bosmagan bo'lsa-da),
              // test tugagan hisoblanadi — shunda javoblar ko'rinishi kerak.
              const timeExpired = !!(test.startedAt && test.duration && Date.now() > test.startedAt + test.duration*60000);
              const ended = activated && (!test.active || timeExpired);
              return (
                <div key={test.id} style={{...S.card,padding:16,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:10,flexWrap:"wrap"}}>
                    <div>
                      <h4 style={{margin:"0 0 3px",fontSize:14}}>{test.name}</h4>
                      <p style={{margin:0,fontSize:11.5,color:C.textLight}}>{testTotalItems(test)} savol • {!activated?"⏸ Hali faollashtirilmagan":ended?"✅ Tugagan":"🟢 Faol"}</p>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:10}}>
                    {!activated
                      ? <span style={{...S.badge,background:"#F1F5F9",color:C.textLight,fontSize:11}}>Test hali faollashtirilmagan — ko'rish mumkin emas</span>
                      : (langs.length>0
                          ? <DocLangButtons test={test} onOpen={setDocModal} small/>
                          : <span style={{fontSize:11.5,color:C.textLight}}>Bu testga hujjat (PDF/LaTeX) biriktirilmagan</span>)}
                    {ended
                      ? <button onClick={()=>setAnswersModal({test})} style={{...S.btnSmall,background:C.successDark,padding:"6px 12px",fontSize:12}}>✅ To'g'ri javoblar</button>
                      : activated && <span style={{...S.badge,background:"#F1F5F9",color:C.textLight,fontSize:11}}>Javoblar test tugagach ochiladi</span>}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {tab==="results" && (
          <div>
            <PdfReadyModal pdf={pdfReady} onClose={()=>setPdfReady(null)}/>
            {err && <div style={{ marginBottom: 12, padding: "12px 14px", borderRadius: 10, background: "#FEF2F2", color: "#991B1B", fontSize: 13, fontWeight: 600 }}>⚠️ {err}</div>}
            {uploads.length===0 && <div style={{...S.card,padding:24,textAlign:"center",color:C.textLight,fontSize:13}}>Hali hech qanday fayl yuklanmagan</div>}
            {uploads.slice().sort((a,b)=>b.uploadedAt-a.uploadedAt).map(batch=>{
              // Admin e'lon qilmaguncha markaz hech qanday ball ko'rmaydi — faqat "qabul qilindi" yozuvi
              const isPending = !batch.publishedRows;
              const rows = batch.publishedRows || [];
              const nUploaded = (batch.rawRows || []).length;
              return (
                <div key={batch.id} style={{...S.card,padding:16,marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
                    <div>
                      <h4 style={{margin:"0 0 2px",fontSize:14,color:"#0891B2"}}>{batch.testName}</h4>
                      <p style={{margin:0,fontSize:11,color:C.textLight}}>Yuklangan: {new Date(batch.uploadedAt).toLocaleString("uz-UZ")} • {nUploaded} o'quvchi{batch.publishedAt?` • E'lon qilingan: ${new Date(batch.publishedAt).toLocaleString("uz-UZ")}`:""}</p>
                    </div>
                    {!isPending && <button onClick={()=>printPDF(batch)} disabled={pdfBusyId===batch.id} style={{...S.btnSmall,background:"#0891B2",padding:"7px 12px",fontSize:12,opacity:pdfBusyId===batch.id?0.6:1}}>{pdfBusyId===batch.id?"⏳ Tayyorlanmoqda...":"📄 PDF qilib yuklab olish"}</button>}
                  </div>
                  {isPending ? (
                    <div style={{padding:"18px 14px",borderRadius:10,background:"#ECFDF5",border:"1.5px solid #A7F3D0",textAlign:"center"}}>
                      <div style={{fontSize:28,marginBottom:6}}>✅</div>
                      <p style={{margin:0,fontSize:14,fontWeight:800,color:"#065F46"}}>Natijangiz qabul qilindi</p>
                      <p style={{margin:"4px 0 0",fontSize:13,color:"#047857"}}>Tez orada Rash modelida hisoblanadi.</p>
                    </div>
                  ) : (
                  <div style={{overflowX:"auto"}}>
                    <table style={S.table}>
                      <thead><tr>{["O'rin","F.I.O","To'g'ri","Algebra","Geometriya","BALL","Daraja"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>{[...rows].sort((a,b)=>(b.ball??-999)-(a.ball??-999)).map((r,i)=>{
                        const dc = darajaColor(r.daraja);
                        return (
                          <tr key={i} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                            <td style={S.td}>{r.rank ?? i+1}</td>
                            <td style={S.td}>{r.name}</td>
                            <td style={S.td}>{r.correct??"-"}/{r.total??"-"}</td>
                            <td style={S.td}>{r.alg!=null?r.alg.toFixed(1):"-"}</td>
                            <td style={S.td}>{r.geo!=null?r.geo.toFixed(1):"-"}</td>
                            <td style={S.td}><b style={{color:"#0891B2"}}>{r.ball!=null?r.ball.toFixed(1):"-"}</b></td>
                            <td style={S.td}>{r.daraja?<span style={{...S.badge,background:dc+"22",color:dc,fontWeight:800}}>{r.daraja}</span>:"-"}</td>
                          </tr>
                        );
                      })}</tbody>
                    </table>
                  </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {tab==="stats" && (
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
              <div style={{...S.card,padding:16,textAlign:"center"}}>
                <div style={{fontSize:28,fontWeight:900,color:"#0891B2"}}>{withBall.length}</div>
                <div style={{fontSize:11.5,color:C.textMid}}>Jami o'quvchi</div>
              </div>
              <div style={{...S.card,padding:16,textAlign:"center"}}>
                <div style={{fontSize:28,fontWeight:900,color:"#0891B2"}}>{avgBall!=null?avgBall.toFixed(1):"-"}</div>
                <div style={{fontSize:11.5,color:C.textMid}}>O'rtacha BALL</div>
              </div>
            </div>
            <div style={{...S.card,padding:16,marginBottom:16}}>
              <p style={{margin:"0 0 12px",fontWeight:800,fontSize:13}}>Darajalar bo'yicha taqsimot</p>
              {withBall.length===0
                ? <p style={{color:C.textLight,fontSize:13,margin:0,textAlign:"center"}}>Ma'lumot yo'q</p>
                : DARAJA_ORDER.filter(d=>darajaCounts[d]).map(d=>{
                    const cnt = darajaCounts[d]||0;
                    const pct = Math.round(cnt/withBall.length*100);
                    const dc = darajaColor(d);
                    return (
                      <div key={d} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                        <span style={{width:32,fontWeight:800,fontSize:12,color:dc}}>{d}</span>
                        <div style={{flex:1,background:"#F1F5F9",borderRadius:6,height:16,overflow:"hidden"}}>
                          <div style={{width:`${pct}%`,height:"100%",background:dc}}/>
                        </div>
                        <span style={{width:56,textAlign:"right",fontSize:12,color:C.textMid}}>{cnt} ({pct}%)</span>
                      </div>
                    );
                  })
              }
            </div>
            <div style={{...S.card,padding:16}}>
              <p style={{margin:"0 0 10px",fontWeight:800,fontSize:13}}>Testlar bo'yicha</p>
              {uploads.length===0
                ? <p style={{color:C.textLight,fontSize:13,margin:0,textAlign:"center"}}>Ma'lumot yo'q</p>
                : uploads.map(b=>{
                    const bAvg = (b.publishedRows||[]).filter(r=>typeof r.ball==="number");
                    const avg = bAvg.length ? (bAvg.reduce((a,x)=>a+x.ball,0)/bAvg.length) : null;
                    return (
                      <div key={b.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderTop:`1px solid ${C.border}`,fontSize:12.5}}>
                        <span>{b.testName}</span>
                        <span style={{color:C.textMid}}>{b.publishedRows?`${b.publishedRows.length} ta`:"⏳ hisoblanmoqda"} • o'rt. {avg!=null?avg.toFixed(1):"-"}</span>
                      </div>
                    );
                  })
              }
            </div>
          </div>
        )}

      </div>

      {/* Test hujjatini (PDF/LaTeX) ko'rish — faqat ko'rish, topshirish yo'q */}
      {docModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",zIndex:9999,display:"flex",flexDirection:"column"}}>
          <div style={{background:C.card,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${C.border}`}}>
            <span style={{fontWeight:700,fontSize:16}}>{docModal.type==="pdf"?"📄":"∑"} {docModal.name||"Test varianti"}</span>
            <button onClick={()=>setDocModal(null)} style={{...S.btnDanger,padding:"8px 14px"}}>✕ Yopish</button>
          </div>
          {docModal.type==="pdf"
            ? <PdfViewer url={docModal.url} persistKey={"partner_pdf_"+docModal.url?.slice(-8)}/>
            : <ScrollPersistDiv persistKey={"partner_latex_scroll_"+(docModal.id||"x")} style={{flex:1,overflowY:"auto",background:"white"}}><LatexDocViewer source={docModal.source} images={docModal.images}/></ScrollPersistDiv>
          }
        </div>
      )}

      {/* Test tugagandan keyin — to'g'ri javoblar */}
      {answersModal&&(()=>{
        const test = answersModal.test;
        const closedQs = test.questions.filter(q=>q.type==="closed");
        const openQs = test.questions.filter(q=>q.type==="open");
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",zIndex:9999,display:"flex",flexDirection:"column"}}>
            <div style={{background:C.successDark,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontWeight:700,fontSize:15,color:"white"}}>✅ To'g'ri javoblar: {test.name}</span>
              <button onClick={()=>setAnswersModal(null)} style={{...S.btnSmall,background:"rgba(255,255,255,0.2)",color:"white"}}>✕ Yopish</button>
            </div>
            <div style={{flex:1,overflowY:"auto",background:C.bg,padding:16}}>
              {closedQs.length>0&&(
                <div style={{...S.card,padding:16,marginBottom:14}}>
                  <h4 style={{margin:"0 0 10px",color:C.primary,fontSize:14}}>Yopiq savollar</h4>
                  <div style={{overflowX:"auto"}}>
                    <table style={S.table}>
                      <thead><tr>{["#","To'g'ri javob"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                      <tbody>{closedQs.map((q,i)=>(
                        <tr key={i} style={{background:i%2===0?C.card:"#FAFBFF"}}>
                          <td style={S.td}>{i+1}</td>
                          <td style={{...S.td,fontWeight:700,color:C.successDark}}>{q.correctAnswer||"—"}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                </div>
              )}
              {openQs.map((q,ri)=>{
                const idx = closedQs.length+ri;
                return (
                  <div key={idx} style={{...S.card,padding:16,marginBottom:12}}>
                    <h4 style={{margin:"0 0 10px",color:C.warning,fontSize:14}}>Savol {idx+1}</h4>
                    {q.subParts?.length>0
                      ? q.subParts.map((sp,si)=>(
                          <div key={si} style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",padding:"8px 10px",borderRadius:8,background:C.successLight,marginBottom:6}}>
                            <span style={{color:C.warning,fontWeight:700,minWidth:34}}>{idx+1}{sp.label})</span>
                            <KatexSpan latex={toLatex(sp.answer||"")} fontSize={17}/>
                          </div>
                        ))
                      : <div style={{padding:"8px 10px",borderRadius:8,background:C.successLight}}><KatexSpan latex={toLatex(q.correctAnswer||"")} fontSize={17}/></div>
                    }
                  </div>
                );
              })}
            </div>
          </div>
        );
      })()}

    </div>
  );
}


// O'quvchi guruhga qo'shiladigan oyna: guruhlar ro'yxati KO'RSATILMAYDI — o'quvchi nomini
// o'zi yozadi, shunday guruh ochilgan bo'lsa qo'shiladi.
function GroupPicker({ user, onSave, onClose }) {
  const [name, setName] = useState("");
  const [err, setErr] = useState("");
  const join = () => {
    if (!name.trim()) { setErr("Guruh nomini yozing."); return; }
    const g = findOpenGroup(name);
    if (!g) { setErr("Bunday guruh topilmadi. Guruh nomini o'qituvchingizdan aniq so'rab, qayta yozing."); return; }
    onSave(g);
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }} onClick={onClose}>
      <div style={{ ...S.card, padding: 22, maxWidth: 400, width: "100%" }} onClick={e => e.stopPropagation()}>
        <h3 style={{ margin: "0 0 4px", fontSize: 18 }}>👥 Guruhga qo'shilish</h3>
        <p style={{ margin: "0 0 14px", color: C.textMid, fontSize: 13, lineHeight: 1.5 }}>O'qituvchingiz aytgan guruh nomini yozing.{user.group ? <> Hozirgi guruhingiz: <b>{user.group}</b>.</> : null}</p>
        {err && <div style={S.err}>{err}</div>}
        <label style={S.label} htmlFor="group_join_name">Guruh nomi</label>
        <input id="group_join_name" value={name} autoFocus onChange={e => { setName(e.target.value); setErr(""); }} onKeyDown={e => e.key === "Enter" && join()} placeholder="Masalan: 10-A" style={S.input} />
        <button onClick={join} style={S.btnPrimary}>Qo'shilish</button>
        <button onClick={onClose} style={{ ...S.btnGhost, width: "100%", marginTop: 8 }}>Bekor qilish</button>
      </div>
    </div>
  );
}

// ===== ADMIN: XAVFSIZLIK =====
function SecurityPanel() {
  const [open, setOpen] = useState(adminUsesDefaultPassword());
  const [f, setF] = useState({ cur: "", n1: "", n2: "" });
  const [msg, setMsg] = useState(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [plain, setPlain] = useState(() => countPlainPasswords());
  const [isDefault, setIsDefault] = useState(() => adminUsesDefaultPassword());
  useEffect(() => { const t = setInterval(() => setPlain(countPlainPasswords()), 2000); return () => clearInterval(t); }, []);
  const change = async () => {
    setErr(""); setMsg(null);
    if (!f.cur || !f.n1) { setErr("Barcha maydonlarni to'ldiring!"); return; }
    if (f.n1.length < 8) { setErr("Yangi parol kamida 8 ta belgidan iborat bo'lsin!"); return; }
    if (!/[A-Za-z]/.test(f.n1) || !/\d/.test(f.n1)) { setErr("Yangi parolda ham harf, ham raqam bo'lsin!"); return; }
    if (f.n1 !== f.n2) { setErr("Yangi parollar mos kelmadi!"); return; }
    setBusy(true);
    try {
      if (!(await verifyPassword(getAdminPasswordHash(), f.cur))) { setErr("Joriy parol noto'g'ri!"); setBusy(false); return; }
      const h = await hashPassword(f.n1);
      db.set("appSettings", [...(db.get("appSettings") || []).filter(x => x.id !== "admin"), { id: "admin", passwordHash: h, changedAt: Date.now() }]);
      saveSession({ role: "admin", v: sessionVerifier(h) }); // shu qurilmada kirgan holda qolasiz, boshqalarda chiqib ketadi
      setF({ cur: "", n1: "", n2: "" }); setIsDefault(false);
      setMsg("✅ Admin paroli almashtirildi. Boshqa qurilmalardagi admin sessiyalari bekor qilindi.");
    } catch (e) { setErr(e.message || "Xatolik yuz berdi."); }
    setBusy(false);
  };
  return (
    <div style={{ ...S.card, padding: 16, marginBottom: 18, border: `2px solid ${isDefault ? C.danger : C.successDark}` }}>
      <button onClick={() => setOpen(o => !o)} aria-expanded={open} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, width: "100%", background: "none", border: "none", padding: 0, cursor: "pointer", textAlign: "left", marginBottom: open ? 14 : 0 }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <b style={{ fontSize: 15, color: isDefault ? C.danger : C.successDark }}>🔐 Xavfsizlik</b>
          {isDefault && <span style={{ ...S.badge, background: C.dangerLight, color: C.danger }}>Standart admin paroli!</span>}
          {plain > 0 && <span style={{ ...S.badge, background: C.warningLight, color: "#92400E" }}>⏳ {plain} ta parol shifrlanmoqda</span>}
        </span>
        <span style={{ fontSize: 13, color: C.textMid, fontWeight: 700, transform: open ? "rotate(180deg)" : "none", transition: "transform 0.2s" }}>▼</span>
      </button>
      {open && (<>
        {isDefault && <div style={{ ...S.err, lineHeight: 1.5 }}>Admin paroli hali standart holatda. U avvalgi versiyalarda kod ichida ochiq yozilgan edi, shuning uchun uni hozir almashtiring.</div>}
        <p style={{ margin: "0 0 10px", fontSize: 13, fontWeight: 700, color: C.text }}>Admin parolini almashtirish</p>
        {err && <div style={S.err}>{err}</div>}
        {msg && <div style={{ background: C.successLight, color: C.successDark, borderRadius: 10, padding: "10px 14px", fontSize: 13, marginBottom: 12, fontWeight: 600 }}>{msg}</div>}
        {[["cur", "Joriy parol"], ["n1", "Yangi parol (kamida 8 ta belgi, harf va raqam)"], ["n2", "Yangi parolni takrorlang"]].map(([k, l]) => (
          <div key={k}>
            <label style={S.label} htmlFor={"sec_" + k}>{l}</label>
            <input id={"sec_" + k} type="password" autoComplete={k === "cur" ? "current-password" : "new-password"} value={f[k]} onChange={e => setF(p => ({ ...p, [k]: e.target.value }))} onKeyDown={e => e.key === "Enter" && change()} style={S.input} />
          </div>
        ))}
        <button onClick={change} disabled={busy} style={{ ...S.btnSmall, background: C.primary, padding: "10px 18px", opacity: busy ? 0.6 : 1 }}>{busy ? "⏳ Saqlanmoqda..." : "🔐 Parolni almashtirish"}</button>
        <div style={{ marginTop: 16, paddingTop: 12, borderTop: `1px solid ${C.border}`, fontSize: 12.5, color: C.textMid, lineHeight: 1.6 }}>
          <b style={{ color: C.text }}>Parollar holati:</b> {plain > 0 ? `${plain} ta eski parol hozir fon rejimida shifrlanmoqda (sahifani yopmang).` : "barcha parollar shifrlangan (xesh) holda saqlanadi."}
        </div>
      </>)}
    </div>
  );
}

// ===== ADMIN: GURUHLARNI OCHISH =====
function GroupManager() {
  const [open, setOpen] = useState(false);
  const [list, setList] = useState(() => {
    // Birinchi marta: avval o'quvchilar yozib qo'ygan guruhlar ham ochilgan guruh sifatida saqlanadi
    if (!(db.get("groups") || []).length) {
      const legacy = [...new Set((db.get("users") || []).map(u => (u.group || "").trim()).filter(Boolean))];
      if (legacy.length) ensureGroups(legacy);
    }
    return db.get("groups") || [];
  });
  const [name, setName] = useState("");
  const [err, setErr] = useState("");
  const [confirm, setConfirm] = useState(null);
  const users = db.get("users") || [];
  const count = (g) => users.filter(u => normGroupName(u.group) === normGroupName(g)).length;
  const add = () => {
    const nm = name.trim().replace(/\s+/g, " ");
    if (!nm) { setErr("Guruh nomini yozing!"); return; }
    if ((db.get("groups") || []).some(g => normGroupName(g.name) === normGroupName(nm))) { setErr("Bu guruh allaqachon ochilgan!"); return; }
    ensureGroups([nm]); setList(db.get("groups") || []); setName(""); setErr("");
  };
  const del = (g) => setConfirm({
    message: `"${g.name}" guruhi yopilsinmi? Yangi o'quvchilar unga qo'shila olmaydi. Hozir guruhdagi ${count(g.name)} ta o'quvchi guruhida qoladi.`,
    confirmLabel: "Yopish",
    onConfirm: () => { db.set("groups", (db.get("groups") || []).filter(x => x.id !== g.id)); setList(db.get("groups") || []); setConfirm(null); },
  });
  const sorted = [...list].sort((a, b) => a.name.localeCompare(b.name));
  return (
    <div style={{ ...S.card, padding: 16, marginBottom: 18, border: "2px solid #7C3AED" }}>
      {confirm && <ConfirmModal message={confirm.message} confirmLabel={confirm.confirmLabel} onConfirm={confirm.onConfirm} onCancel={() => setConfirm(null)} />}
      <button onClick={() => setOpen(o => !o)} aria-expanded={open} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, width: "100%", background: "none", border: "none", padding: 0, cursor: "pointer", textAlign: "left", marginBottom: open ? 14 : 0 }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}><b style={{ fontSize: 15, color: "#7C3AED" }}>👥 Guruhlar</b><span style={{ ...S.badge, background: "#EDE9FE", color: "#7C3AED" }}>{list.length} ta</span></span>
        <span style={{ fontSize: 13, color: C.textMid, fontWeight: 700, transform: open ? "rotate(180deg)" : "none", transition: "transform 0.2s" }}>▼</span>
      </button>
      {open && (<>
        <p style={{ margin: "0 0 12px", fontSize: 12, color: C.textLight, lineHeight: 1.5 }}>O'quvchilar guruhlar ro'yxatini ko'rmaydi: ular guruh nomini o'zi yozadi va faqat shu yerda ochilgan guruhga qo'shila oladi. Guruh nomini o'quvchilarga o'zingiz ayting. Testni biror guruhga mo'ljallasangiz, u guruh ham avtomatik ochiladi.</p>
        {err && <div style={S.err}>{err}</div>}
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <input id="group_new_name" value={name} onChange={e => { setName(e.target.value); setErr(""); }} onKeyDown={e => e.key === "Enter" && add()} placeholder="Yangi guruh nomi (masalan 10-A)" style={{ ...S.input, margin: 0, fontSize: 13 }} />
          <button onClick={add} style={{ ...S.btnSmall, background: "#7C3AED", whiteSpace: "nowrap" }}>+ Ochish</button>
        </div>
        {sorted.length === 0
          ? <p style={{ color: C.textLight, fontSize: 13, margin: 0, textAlign: "center" }}>Hali guruh ochilmagan</p>
          : <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
              {sorted.map(g => (
                <span key={g.id} style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "6px 8px 6px 12px", borderRadius: 20, border: `1.5px solid ${C.border}`, background: "white", fontSize: 13, fontWeight: 600 }}>
                  {g.name}<span style={{ color: C.textLight, fontWeight: 500 }}>{count(g.name)} o'quvchi</span>
                  <button onClick={() => del(g)} title="Guruhni yopish" style={{ width: 22, height: 22, borderRadius: "50%", border: "none", background: C.dangerLight, color: C.danger, cursor: "pointer", fontSize: 12, lineHeight: 1 }}>✕</button>
                </span>
              ))}
            </div>}
      </>)}
    </div>
  );
}

function StudentDashboard({ user, onLogout, onUserUpdate }) {
  const [groupPicker,setGroupPicker]=useState(false);
  const saveGroup = (g) => {
    const users = db.get("users") || [];
    const next = users.map(u => u.phone === user.phone ? { ...u, group: g } : u);
    db.set("users", next);
    const me = next.find(u => u.phone === user.phone);
    if (me && onUserUpdate) onUserUpdate(me);
    setGroupPicker(false);
  };
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
      {groupPicker&&<GroupPicker user={user} onSave={saveGroup} onClose={()=>setGroupPicker(false)}/>}
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
          <div><p style={{margin:0,color:"white",fontWeight:700}}>{user.firstName} {user.lastName}</p><p onClick={()=>setGroupPicker(true)} style={{margin:0,color:"rgba(255,255,255,0.8)",fontSize:12,cursor:"pointer"}}>{user.group?`👥 ${user.group} ✏️`:"👥 Guruhga qo'shilish"}</p></div>
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
            {!user.group&&(
              <button onClick={()=>setGroupPicker(true)} style={{...S.btnSmall,background:C.primary,padding:"10px 18px",marginBottom:18}}>👥 Guruhga qo'shilish</button>
            )}
            {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now&&visibleForMe(t)).length>0&&(
              <div style={{marginBottom:24}}>
                <h3 style={{marginBottom:16}}>🗓️ Rejalashtirilgan Testlar</h3>
                {tests.filter(t=>!t.active&&t.scheduledAt&&t.scheduledAt>now&&visibleForMe(t)).map(test=>(
                  <div key={test.id} style={{...S.card,padding:18,marginBottom:12,border:`1.5px dashed ${C.primary}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                      <div>
                        <h4 style={{margin:"0 0 5px",fontSize:16}}>{test.name}</h4>
                        <p style={{margin:"0 0 6px",color:C.textMid,fontSize:13}}>{testTotalItems(test)} savol • {test.duration} daqiqa</p>
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
                      <p style={{margin:"0 0 6px",color:C.textMid,fontSize:13}}>{testTotalItems(test)} savol • {test.duration} daqiqa</p>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                        {timeInfo&&<span style={{...S.badge,background:expired?C.dangerLight:C.warningLight,color:expired?C.danger:C.warning,fontSize:12}}>⏱ {timeInfo}</span>}
                        <DocLangButtons test={test} onOpen={setDocModal} small/>
                      </div>
                    </div>
                    {myRes?(
                      <div style={{textAlign:"right"}}>
                        <p style={{margin:"0 0 4px",color:C.successDark,fontWeight:700}}>✅ Topshirildi</p>
                        <p style={{margin:0,color:C.textMid,fontSize:13}}>{myRes.totalScore}/{testTotalItems(test)} ball</p>
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
              const pct=Math.round((r.totalScore/testTotalItems(test))*100);
              const canView=test.showAnswersAfter==="immediate"||r.canViewAnswers;
              return (
                <div key={r.id} style={{...S.card,padding:18,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}>
                    <div><h4 style={{margin:"0 0 4px",fontSize:15}}>{test.name}</h4><p style={{margin:0,color:C.textMid,fontSize:12}}>{new Date(r.id).toLocaleDateString("uz-UZ")}</p>
                      {!r.raschPublished&&testEligibleForRasch(test)&&<p style={{margin:"5px 0 0",fontSize:11,color:C.textLight}}>🎯 Rash natijasi tez orada e'lon qilinadi</p>}
                      {r.raschPublished&&testEligibleForRasch(test)&&(()=>{ const dg=r.raschPublished.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                        return <p style={{margin:"5px 0 0",display:"inline-flex",alignItems:"center",gap:6}}><span style={{fontSize:11,color:C.textMid}}>🎯 Rash:</span><b style={{color:"#6D28D9",fontSize:13}}>{r.raschPublished.ball!=null?r.raschPublished.ball.toFixed(1):"—"}</b><span style={{background:dc+"22",color:dc,padding:"2px 8px",borderRadius:6,fontSize:11,fontWeight:800}}>{dg}</span></p>; })()}
                    </div>
                    <div style={{textAlign:"right"}}>
                      <p style={{margin:"0 0 4px",fontWeight:800,fontSize:20,color:pct>=70?C.successDark:pct>=50?C.warning:C.danger}}>{r.totalScore}<span style={{color:C.textMid,fontWeight:400,fontSize:14}}>/{testTotalItems(test)}</span></p>
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
                    <p style={{margin:0,color:C.textMid,fontSize:12}}>{testTotalItems(test)} savol • {availableDocLangs(test).map(l=>l.label).join(" / ")}</p>
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
                              {r?.raschPublished&&testEligibleForRasch(t)
                                ? (()=>{ const dg=r.raschPublished.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
                                  return <span style={{display:"inline-flex",alignItems:"center",gap:5}}><b style={{color:"#6D28D9",fontSize:13}}>{r.raschPublished.ball!=null?r.raschPublished.ball.toFixed(1):"—"}</b><span style={{background:dc+"22",color:dc,padding:"2px 7px",borderRadius:6,fontSize:11,fontWeight:800}}>{dg}</span></span>; })()
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
  const pct=Math.round((result.totalScore/testTotalItems(test))*100);
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
          <div style={{color:C.textMid,fontSize:18,margin:"4px 0 12px"}}>/ {testTotalItems(test)} ball ({pct}%)</div>
          <div style={{background:C.bg,borderRadius:999,height:14,overflow:"hidden",maxWidth:400,margin:"0 auto"}}>
            <div style={{width:`${pct}%`,height:"100%",background:pct>=70?C.success:pct>=50?C.warning:C.danger,borderRadius:999,transition:"width 1.5s"}}/>
          </div>
        </div>
        {result.raschPublished&&testEligibleForRasch(test)&&(()=>{ const dg=result.raschPublished.daraja; const dc=dg==="NC"?C.danger:(dg==="C"||dg==="C+")?C.warning:C.successDark;
          return (
            <div style={{...S.card,padding:20,textAlign:"center",marginBottom:20,background:"linear-gradient(135deg,#F5F3FF,#FFFFFF)",border:"1.5px solid #DDD6FE"}}>
              <p style={{margin:"0 0 8px",color:"#6D28D9",fontSize:13,fontWeight:800}}>🎯 RASH MODELI BO'YICHA BAHOLASH</p>
              <div style={{display:"flex",justifyContent:"center",alignItems:"baseline",gap:14,flexWrap:"wrap"}}>
                <div><div style={{fontSize:40,fontWeight:900,color:"#6D28D9",lineHeight:1}}>{result.raschPublished.ball!=null?result.raschPublished.ball.toFixed(1):"—"}</div><div style={{fontSize:11,color:C.textMid}}>Rash ball</div></div>
                <div><span style={{background:dc+"22",color:dc,padding:"6px 16px",borderRadius:10,fontSize:20,fontWeight:900}}>{dg}</span><div style={{fontSize:11,color:C.textMid,marginTop:4}}>Daraja</div></div>
              </div>
              {(result.raschPublished.alg!=null||result.raschPublished.bmba!=null||result.raschPublished.rank!=null)&&(
                <div style={{display:"flex",justifyContent:"center",gap:18,flexWrap:"wrap",marginTop:14,fontSize:12.5,color:C.textMid}}>
                  {result.raschPublished.alg!=null&&<span>Algebra: <b style={{color:C.text}}>{result.raschPublished.alg.toFixed(1)}</b></span>}
                  {result.raschPublished.geo!=null&&<span>Geometriya: <b style={{color:C.text}}>{result.raschPublished.geo.toFixed(1)}</b></span>}
                  {result.raschPublished.foiz!=null&&<span>Foiz: <b style={{color:C.text}}>{result.raschPublished.foiz.toFixed(1)}%</b></span>}
                  {result.raschPublished.bmba!=null&&<span>BMBA: <b style={{color:C.text}}>{result.raschPublished.bmba.toFixed(1)}</b></span>}
                  {result.raschPublished.rank!=null&&<span>Reyting o'rni: <b style={{color:C.text}}>{result.raschPublished.rank}</b></span>}
                </div>
              )}
              <p style={{margin:"10px 0 0",fontSize:11,color:C.textLight}}>Savollar qiyinligiga qarab vaznlangan, Algebra va Geometriya alohida hisoblangan Rash modeli. Hisoblangan sana: {new Date(result.raschPublished.calculatedAt).toLocaleDateString("uz-UZ")}</p>
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
    const out = () => { clearSession(); return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null, isPartner: false, partnerInfo: null }; };
    if (!s) return { page: "login", user: null, isAdmin: false, isTeacher: false, teacherInfo: null, isPartner: false, partnerInfo: null };
    if (!s.exp || s.exp < Date.now()) return out();
    // Sessiyadagi tasdiq (v) hisobning HOZIRGI parol xeshiga mos kelishi shart
    const ok = (stored) => (s.v || "") === sessionVerifier(stored);
    if (s.role === "admin") {
      if (!ok(getAdminPasswordHash())) return out();
      return { page: "admin", user: null, isAdmin: true, isTeacher: false, teacherInfo: null, isPartner: false, partnerInfo: null };
    }
    if (s.role === "teacher") {
      const t = (db.get("teachers")||[]).find(x => x.id === s.teacherId);
      if (!t || !ok(t.password)) return out();
      return { page: "teacher", user: null, isAdmin: false, isTeacher: true, teacherInfo: t, isPartner: false, partnerInfo: null };
    }
    if (s.role === "partner") {
      const p = (db.get("partners")||[]).find(x => x.id === s.partnerId);
      if (!p || !p.approved || !ok(p.password)) return out();
      return { page: "partner", user: null, isAdmin: false, isTeacher: false, teacherInfo: null, isPartner: true, partnerInfo: p };
    }
    if (s.role === "student") {
      const u = (db.get("users")||[]).find(x => x.phone === s.phone);
      if (!u || !ok(u.password)) return out();
      return { page: "student", user: u, isAdmin: false, isTeacher: false, teacherInfo: null, isPartner: false, partnerInfo: null };
    }
    return out();
  };
  const initState = initFromSession();
  const [page,setPage]=useState(initState.page);
  const [user,setUser]=useState(initState.user);
  const [isAdmin,setIsAdmin]=useState(initState.isAdmin);
  const [isTeacher,setIsTeacher]=useState(initState.isTeacher);
  const [teacherInfo,setTeacherInfo]=useState(initState.teacherInfo);
  const [isPartner,setIsPartner]=useState(initState.isPartner);
  const [partnerInfo,setPartnerInfo]=useState(initState.partnerInfo);
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
  if(page==="partner"&&isPartner) return <PartnerPanel partnerInfo={partnerInfo} onLogout={()=>{setIsPartner(false);setPartnerInfo(null);setPage("login");clearSession();}}/>;
  if(page==="student"&&user) return <StudentDashboard user={user} onUserUpdate={setUser} onLogout={()=>{setUser(null);setPage("login");clearSession();}}/>;
  if(page==="register") return <RegisterPage onDone={u=>{setUser(u);setPage("student");saveSession({role:"student",phone:u.phone,v:sessionVerifier(u.password)});}} onLogin={()=>setPage("login")}/>;
  if(page==="partnerRegister") return <PartnerRegisterPage onLogin={()=>setPage("login")}/>;
  return <LoginPage
    onLogin={u=>{setUser(u);setPage("student");saveSession({role:"student",phone:u.phone,v:sessionVerifier(u.password)});}}
    onRegister={()=>setPage("register")}
    onAdmin={(fullAdmin, teacher)=>{
      if(fullAdmin){ setIsAdmin(true); setPage("admin"); saveSession({role:"admin",v:sessionVerifier(getAdminPasswordHash())}); }
      else { setIsTeacher(true); setTeacherInfo(teacher); setPage("teacher"); saveSession({role:"teacher",teacherId:teacher.id,v:sessionVerifier(teacher.password)}); }
    }}
    onPartner={(partner)=>{ setIsPartner(true); setPartnerInfo(partner); setPage("partner"); saveSession({role:"partner",partnerId:partner.id,v:sessionVerifier(partner.password)}); }}
    onPartnerRegister={()=>setPage("partnerRegister")}
  />;
}
