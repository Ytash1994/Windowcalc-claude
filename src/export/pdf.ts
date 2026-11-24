import { PDFDocument, PDFPage, PDFFont, rgb, RGB } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import arRegSrc from "../assets/fonts/NotoSansArabic-Regular.ttf?inline";
import arBoldSrc from "../assets/fonts/NotoSansArabic-Bold.ttf?inline";
import symRegSrc from "../assets/fonts/DejaVuSans.ttf?inline";
import symBoldSrc from "../assets/fonts/DejaVuSans-Bold.ttf?inline";
import { SIG_NAME, SIG_PHONE } from "../engine/calc";
import type { AreaTotalItem, ExportStructUnit } from "../types";

export interface PdfExportInfo {
  project: string;
  type: string;
  unit: string;
  area: string;
  areaUnitLabel?: string;
  date: string;
  count: string;
  notes: string[];
  struct: ExportStructUnit[];
  table: string[][];
  dup?: [string, string, string, string][];
  dupOther?: number;
  areaTotals?: AreaTotalItem[];
}

interface FontBundle {
  arReg: Uint8Array;
  arBold: Uint8Array;
  symReg: Uint8Array;
  symBold: Uint8Array;
}

let cachedFontBytes: FontBundle | null = null;

function decodeBase64Font(dataUri: string): Uint8Array {
  const b64 = dataUri.includes(",") ? dataUri.split(",")[1] : dataUri;
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) {
    bytes[i] = bin.charCodeAt(i);
  }
  return bytes;
}

async function resolveFontBytes(src: string): Promise<Uint8Array> {
  if (src.startsWith("data:")) {
    return decodeBase64Font(src);
  }
  const res = await fetch(src);
  if (!res.ok) throw new Error(`Failed to fetch font: ${src}`);
  return new Uint8Array(await res.arrayBuffer());
}

async function getFontBundle(customLoader?: (name: string) => Promise<Uint8Array>): Promise<FontBundle> {
  if (cachedFontBytes) return cachedFontBytes;
  if (customLoader) {
    const [arReg, arBold, symReg, symBold] = await Promise.all([
      customLoader("NotoSansArabic-Regular.ttf"),
      customLoader("NotoSansArabic-Bold.ttf"),
      customLoader("DejaVuSans.ttf"),
      customLoader("DejaVuSans-Bold.ttf"),
    ]);
    cachedFontBytes = { arReg, arBold, symReg, symBold };
    return cachedFontBytes;
  }
  const [arReg, arBold, symReg, symBold] = await Promise.all([
    resolveFontBytes(arRegSrc),
    resolveFontBytes(arBoldSrc),
    resolveFontBytes(symRegSrc),
    resolveFontBytes(symBoldSrc),
  ]);
  cachedFontBytes = { arReg, arBold, symReg, symBold };
  return cachedFontBytes;
}

function hexToRgb(hex: string): RGB {
  const clean = hex.replace("#", "");
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return rgb(r, g, b);
}

function hslToRgb(h: number, s: number, l: number): RGB {
  s /= 100;
  l /= 100;
  const a = s * Math.min(l, 1 - l);
  const f = (n: number) => {
    const k = (n + h / 30) % 12;
    return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
  };
  return rgb(f(0), f(8), f(4));
}

// Sanitize symbols outside DejaVu/NotoArabic (e.g., hourglass/emojis) so subsetting never crashes
function sanitizeForPdf(input: string): string {
  return String(input == null ? "" : input)
    .replace(/[\u2066\u2067\u2068\u2069\u200E\u200F]/g, "")
    .replace(/⏳/g, "")
    .replace(/[^\u0020-\u007E\u00A0-\u00FF\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF\u2013\u2014\u2026\u2212\u00D7\u00F7\u00B2\u2248✓✕]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const AR_LETTER_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
const LTR_CORE_RE = /[A-Za-z0-9٠-٩]/;

function swapMirror(ch: string): string {
  if (ch === "(") return ")";
  if (ch === ")") return "(";
  if (ch === "[") return "]";
  if (ch === "]") return "[";
  return ch;
}

interface VisualToken {
  text: string;
  useArabicFont: boolean;
}

interface BidiSegment {
  dir: "rtl" | "ltr" | "neutral";
  raw: string;
}

/**
 * Converts logical RTL-first Arabic + embedded LTR numbers/symbols into exact
 * left-to-right visual drawing chunks, keeping Arabic shaping intact in NotoSansArabic
 * while routing digits and punctuation (+, ×, |, ², (), —, −, /, …) to DejaVuSans.
 */
function toVisualTokens(rawInput: string, baseRtl = true): VisualToken[] {
  const s = sanitizeForPdf(rawInput);
  if (!s) return [];

  // Fast check: pure LTR (numbers, dates, phone numbers)
  if (!AR_LETTER_RE.test(s)) {
    return [{ text: s, useArabicFont: false }];
  }

  // Segment into RTL (Arabic letters + spaces/،/؟ between Arabic), LTR (+,-,digits,Latin,decimals), and Neutral
  const segs: BidiSegment[] = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (AR_LETTER_RE.test(ch)) {
      let j = i + 1;
      while (j < s.length) {
        const c = s[j];
        if (AR_LETTER_RE.test(c) || c === "ـ" || c === "،" || c === "؟") {
          j++;
        } else if (c === " ") {
          // Include space inside Arabic segment if followed by Arabic letter
          let k = j + 1;
          while (k < s.length && s[k] === " ") k++;
          if (k < s.length && AR_LETTER_RE.test(s[k])) {
            j = k;
          } else {
            break;
          }
        } else {
          break;
        }
      }
      segs.push({ dir: "rtl", raw: s.slice(i, j) });
      i = j;
    } else if (LTR_CORE_RE.test(ch) || ((ch === "+" || ch === "-") && i + 1 < s.length && LTR_CORE_RE.test(s[i + 1]))) {
      let j = i + 1;
      while (j < s.length) {
        const c = s[j];
        if (LTR_CORE_RE.test(c) || c === "." || c === "%" || c === "²") {
          j++;
        } else if ((c === "-" || c === "/" || c === ":" || c === ",") && j + 1 < s.length && LTR_CORE_RE.test(s[j + 1])) {
          j++;
        } else {
          break;
        }
      }
      segs.push({ dir: "ltr", raw: s.slice(i, j) });
      i = j;
    } else {
      let j = i + 1;
      while (j < s.length && !AR_LETTER_RE.test(s[j]) && !LTR_CORE_RE.test(s[j])) {
        if ((s[j] === "+" || s[j] === "-") && j + 1 < s.length && LTR_CORE_RE.test(s[j + 1])) break;
        j++;
      }
      segs.push({ dir: "neutral", raw: s.slice(i, j) });
      i = j;
    }
  }

  // Reverse neutral substrings if inside RTL context and mirror brackets
  const tokens: VisualToken[] = [];
  const orderedSegs = baseRtl ? segs.slice().reverse() : segs;

  for (const seg of orderedSegs) {
    if (seg.dir === "rtl") {
      // Split Arabic characters/spaces from Arabic-unsupported punctuation like ؟ or space-only just in case
      // NotoSansArabic supports Arabic letters + spaces + Arabic comma (،) + Arabic question mark (؟)
      tokens.push({ text: seg.raw, useArabicFont: true });
    } else if (seg.dir === "ltr") {
      tokens.push({ text: seg.raw, useArabicFont: false });
    } else {
      const mirrored = baseRtl
        ? [...seg.raw].reverse().map(swapMirror).join("")
        : seg.raw;
      tokens.push({ text: mirrored, useArabicFont: false });
    }
  }

  return tokens;
}

interface DualFont {
  ar: PDFFont;
  sym: PDFFont;
}

function measureVisualWidth(tokens: VisualToken[], fonts: DualFont, size: number): number {
  let total = 0;
  for (const tok of tokens) {
    const f = tok.useArabicFont ? fonts.ar : fonts.sym;
    total += f.widthOfTextAtSize(tok.text, size);
  }
  return total;
}

function measureTextWidth(text: string, fonts: DualFont, size: number, baseRtl = true): number {
  return measureVisualWidth(toVisualTokens(text, baseRtl), fonts, size);
}

function drawTextVisual(
  page: PDFPage,
  text: string,
  x: number,
  yTopCenter: number,
  align: "left" | "center" | "right",
  fonts: DualFont,
  size: number,
  color: RGB,
  pageHeight: number,
  baseRtl = true
) {
  const tokens = toVisualTokens(text, baseRtl);
  if (!tokens.length) return;
  const totalW = measureVisualWidth(tokens, fonts, size);
  let curX = align === "center" ? x - totalW / 2 : align === "right" ? x - totalW : x;
  // Center vertically using cap/x-height approximation
  const baselineY = pageHeight - (yTopCenter + size * 0.32);

  for (const tok of tokens) {
    const f = tok.useArabicFont ? fonts.ar : fonts.sym;
    const w = f.widthOfTextAtSize(tok.text, size);
    page.drawText(tok.text, {
      x: curX,
      y: baselineY,
      size,
      font: f,
      color,
    });
    curX += w;
  }
}

function wrapTextLines(text: string, maxWidth: number, fonts: DualFont, size: number): string[] {
  const clean = sanitizeForPdf(text);
  if (!clean) return [""];
  const words = clean.split(" ");
  const lines: string[] = [];
  let current = "";
  for (const w of words) {
    const candidate = current ? `${current} ${w}` : w;
    if (!current || measureTextWidth(candidate, fonts, size) <= maxWidth) {
      current = candidate;
    } else {
      lines.push(current);
      current = w;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function truncateText(text: string, maxWidth: number, fonts: DualFont, size: number): string {
  let s = sanitizeForPdf(text);
  if (measureTextWidth(s, fonts, size) <= maxWidth) return s;
  while (s.length > 1 && measureTextWidth(`${s}…`, fonts, size) > maxWidth) {
    s = s.slice(0, -1);
  }
  return `${s}…`;
}

function drawRectTop(
  page: PDFPage,
  x: number,
  yTop: number,
  w: number,
  h: number,
  fill: RGB,
  borderColor: RGB | undefined,
  borderWidth: number,
  pageHeight: number
) {
  page.drawRectangle({
    x,
    y: pageHeight - (yTop + h),
    width: w,
    height: h,
    color: fill,
    borderColor,
    borderWidth: borderColor ? borderWidth : 0,
  });
}

function drawCellText(
  page: PDFPage,
  text: string,
  cx: number,
  yTop: number,
  h: number,
  maxWidth: number,
  fonts: DualFont,
  size: number,
  color: RGB,
  pageHeight: number,
  maxLines = 2
) {
  const clean = sanitizeForPdf(text);
  if (!clean) return;
  let lines = wrapTextLines(clean, maxWidth, fonts, size);
  let usedSize = size;
  if (lines.length > maxLines) {
    usedSize = Math.max(6.5, size * 0.78);
    lines = wrapTextLines(clean, maxWidth, fonts, usedSize);
    if (lines.length > maxLines) {
      lines = lines.slice(0, maxLines);
      lines[maxLines - 1] = truncateText(lines[maxLines - 1], maxWidth, fonts, usedSize);
    }
  } else {
    lines = lines.map((ln) =>
      measureTextWidth(ln, fonts, usedSize) > maxWidth ? truncateText(ln, maxWidth, fonts, usedSize) : ln
    );
  }
  const lh = usedSize * 1.22;
  lines.forEach((ln, idx) => {
    const cy = yTop + h / 2 + (idx - (lines.length - 1) / 2) * lh;
    drawTextVisual(page, ln, cx, cy, "center", fonts, usedSize, color, pageHeight);
  });
}

export async function buildPdf(
  info: PdfExportInfo,
  customFontLoader?: (name: string) => Promise<Uint8Array>
): Promise<Blob> {
  // A4 landscape in PDF points: 841.89 x 595.28
  const W = 841.89;
  const H = 595.28;
  const M = 22;
  const FOOTER_Y = H - 30;
  const CONTENT_BOTTOM = H - 38;

  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  pdfDoc.setTitle(sanitizeForPdf(info.project || `نتائج ${info.type}`));
  pdfDoc.setProducer("حاسبة القياسات");

  const bundle = await getFontBundle(customFontLoader);
  const arReg = await pdfDoc.embedFont(bundle.arReg, { subset: true });
  const arBold = await pdfDoc.embedFont(bundle.arBold, { subset: true });
  const symReg = await pdfDoc.embedFont(bundle.symReg, { subset: true });
  const symBold = await pdfDoc.embedFont(bundle.symBold, { subset: true });

  const fReg: DualFont = { ar: arReg, sym: symReg };
  const fBold: DualFont = { ar: arBold, sym: symBold };

  // Build table columns from right to left (RTL layout matching Excel sheetView rightToLeft)
  interface PdfCol {
    w: number;
    t?: string;
    pw: number;
    l: number;
    r: number;
    cx: number;
  }
  const cols: PdfCol[] = [];
  [
    ["رقم", 0.8],
    ["طول الشباك", 1.05],
    ["عرض الشباك", 1.05],
  ].forEach(([t, w]) => cols.push({ w: w as number, t: t as string, pw: 0, l: 0, r: 0, cx: 0 }));

  info.struct.forEach((u) =>
    u.subs.forEach((sub) =>
      cols.push({
        w: u.label === "العدد" && !u.group ? 0.78 : u.group ? (sub.length > 10 ? 1.32 : 1) : 1.32,
        pw: 0,
        l: 0,
        r: 0,
        cx: 0,
      })
    )
  );

  const sumW = cols.reduce((a, c) => a + c.w, 0);
  const uw = (W - 2 * M) / sumW;
  let xr = W - M;
  cols.forEach((c) => {
    c.pw = c.w * uw;
    c.r = xr;
    c.l = xr - c.pw;
    c.cx = (c.r + c.l) / 2;
    xr = c.l;
  });

  const hf = Math.max(7.2, Math.min(10.5, uw / 4.6));
  const nf = Math.max(7.6, Math.min(11.8, uw / 3.6));

  const TY = 98;
  const H1 = 23;
  const H2 = 26;
  const RH = 22;
  const TABLE_TOP = TY + H1 + H2;
  const rpp = Math.max(1, Math.floor((CONTENT_BOTTOM - TABLE_TOP) / RH));
  const rows = info.table;
  const tp = Math.max(1, Math.ceil(rows.length / rpp));

  const DARK_BG = hexToRgb("#1f2937");
  const WHITE = rgb(1, 1, 1);
  const INK = hexToRgb("#111827");
  const MUTED = hexToRgb("#6b7280");
  const FX_HEAD_BG = hexToRgb("#e5e7eb");
  const EVEN_BG = hexToRgb("#f3f4f6");
  const BORDER_GRAY = hexToRgb("#b8c0c8");
  const GOLD_TEXT = hexToRgb("#7a5200");
  const GOLD_SIG = hexToRgb("#9a6b00");
  const GOLD_BG = hexToRgb("#fff1c7");
  const GOLD_BD = hexToRgb("#b8860b");
  const GREEN_BG = hexToRgb("#d1e7dd");
  const GREEN_BD = hexToRgb("#198754");
  const GREEN_TX = hexToRgb("#0f5132");

  // Pre-build pages sequentially; first generate main table pages, then append summary sections flowing smoothly
  const pages: PDFPage[] = [];

  const drawHeaderAndFooter = (page: PDFPage, pageNum: number, totalPagesPlaceholder?: number) => {
    // Top dark header box
    drawRectTop(page, M, 17, W - 2 * M, 58, DARK_BG, undefined, 0, H);

    // Project Title (top-right inside header)
    const projTitle = info.project || "نتائج قياسات الشبابيك والواجهات";
    let titleSize = 19;
    const maxTitleW = W - 2 * M - 180;
    while (titleSize > 12 && measureTextWidth(projTitle, fBold, titleSize) > maxTitleW) {
      titleSize -= 1;
    }
    drawTextVisual(page, projTitle, W - M - 14, 37, "right", fBold, titleSize, WHITE, H);

    // Subtitle metadata line
    const parts = [
      `النوع: ${info.type}`,
      `الوحدة: ${info.unit}`,
      ...(info.area ? [`المساحة: ${info.area}`] : []),
      `التاريخ: ${info.date}`,
      `عدد الشبابيك: ${info.count}`,
    ];
    const subStr = parts.join("  |  ");
    let subSize = 11.5;
    const maxSubW = W - 2 * M - 165;
    while (subSize > 7.5 && measureTextWidth(subStr, fReg, subSize) > maxSubW) {
      subSize -= 0.5;
    }
    drawTextVisual(page, subStr, W - M - 14, 60, "right", fReg, subSize, hexToRgb("#cbd5e1"), H);

    // App brand top-left inside header
    drawTextVisual(page, "حاسبة القياسات", M + 14, 36, "left", fBold, 14, hexToRgb("#93c5fd"), H);

    // Notes line above table
    if (info.notes.length) {
      const notesLine = truncateText(info.notes.join("   |   "), W - 2 * M, fReg, 9.5);
      drawTextVisual(page, notesLine, W - M, TY - 11, "right", fReg, 9.5, MUTED, H);
    }

    // Footer divider and fixed signature
    page.drawLine({
      start: { x: M, y: H - FOOTER_Y },
      end: { x: W - M, y: H - FOOTER_Y },
      thickness: 1,
      color: hexToRgb("#d1d5db"),
    });
    drawTextVisual(page, SIG_NAME, W - M, H - 15, "right", fBold, 12, GOLD_SIG, H);
    drawTextVisual(page, info.date, W / 2, H - 15, "center", fReg, 9.5, MUTED, H, false);
    drawTextVisual(page, SIG_PHONE, M, H - 15, "left", fBold, 12, GOLD_SIG, H, false);

    void pageNum;
    void totalPagesPlaceholder;
  };

  const addNewPage = (): PDFPage => {
    const p = pdfDoc.addPage([W, H]);
    pages.push(p);
    drawHeaderAndFooter(p, pages.length);
    return p;
  };

  const drawTableHeaders = (page: PDFPage, yStart: number) => {
    for (let i = 0; i < 3; i++) {
      const c = cols[i];
      drawRectTop(page, c.l, yStart, c.pw, H1 + H2, FX_HEAD_BG, MUTED, 1, H);
      drawCellText(page, c.t || "", c.cx, yStart, H1 + H2, c.pw - 5, fBold, hf, INK, H);
    }
    let ci = 3;
    info.struct.forEach((u) => {
      const n = u.subs.length;
      const a = cols[ci];
      const b = cols[ci + n - 1];
      const hb = hslToRgb(u.hue, 70, 82);
      const bc = hslToRgb(u.hue, 50, 38);
      if (u.group) {
        drawRectTop(page, b.l, yStart, a.r - b.l, H1, hb, bc, 1, H);
        drawCellText(page, u.name, (a.r + b.l) / 2, yStart, H1, a.r - b.l - 5, fBold, hf, INK, H);
        for (let i = 0; i < n; i++) {
          const c = cols[ci + i];
          drawRectTop(page, c.l, yStart + H1, c.pw, H2, hb, bc, 1, H);
          drawCellText(page, u.subs[i], c.cx, yStart + H1, H2, c.pw - 4, fBold, hf, INK, H);
        }
      } else {
        drawRectTop(page, a.l, yStart, a.pw, H1 + H2, hb, bc, 1, H);
        drawCellText(page, u.label, a.cx, yStart, H1 + H2, a.pw - 5, fBold, hf, INK, H);
      }
      ci += n;
    });
  };

  const metaHues: number[] = [];
  info.struct.forEach((u) => u.subs.forEach(() => metaHues.push(u.hue)));

  let curPage: PDFPage = addNewPage();
  let curY = TABLE_TOP;

  for (let pi = 0; pi < tp; pi++) {
    if (pi > 0) {
      curPage = addNewPage();
    }
    drawTableHeaders(curPage, TY);
    const slice = rows.slice(pi * rpp, (pi + 1) * rpp);
    slice.forEach((row, k) => {
      const gi = pi * rpp + k;
      const even = gi % 2 === 1;
      const y = TABLE_TOP + k * RH;
      row.forEach((v, i) => {
        const c = cols[i];
        if (i < 3) {
          drawRectTop(curPage, c.l, y, c.pw, RH, even ? EVEN_BG : WHITE, BORDER_GRAY, 0.6, H);
          drawCellText(
            curPage,
            String(v === "" ? "" : v),
            c.cx,
            y,
            RH,
            c.pw - 4,
            fBold,
            nf,
            INK,
            H,
            1
          );
        } else {
          const h = metaHues[i - 3];
          drawRectTop(
            curPage,
            c.l,
            y,
            c.pw,
            RH,
            hslToRgb(h, even ? 75 : 85, even ? 90 : 95),
            hslToRgb(h, 35, 68),
            0.6,
            H
          );
          drawCellText(
            curPage,
            String(v === "" ? "" : v),
            c.cx,
            y,
            RH,
            c.pw - 4,
            fReg,
            nf,
            INK,
            H,
            2
          );
        }
      });
      curY = y + RH;
    });
  }

  // Optional Area Totals Summary (when user explicitly enabled area fields)
  if (info.areaTotals && info.areaTotals.length > 0) {
    const AT_ROW_H = 21;
    const neededFirst = 14 + 24 + 22 + Math.min(info.areaTotals.length, 2) * AT_ROW_H;
    if (curY + neededFirst > CONTENT_BOTTOM) {
      curPage = addNewPage();
      curY = TY;
    } else {
      curY += 14;
    }

    const fw = W - 2 * M;
    const atColsRatio = [2.6, 1.5, 1.5];
    const atSum = atColsRatio.reduce((a, b) => a + b, 0);
    let ax = W - M;
    const ac = atColsRatio.map((r) => {
      const pw = (r / atSum) * fw;
      const o = { l: ax - pw, r: ax, pw, cx: ax - pw / 2 };
      ax -= pw;
      return o;
    });

    const drawAreaHeader = (p: PDFPage, y0: number, cont = false) => {
      drawRectTop(p, M, y0, fw, 24, GREEN_BG, GREEN_BD, 1, H);
      drawTextVisual(
        p,
        `ملخص إجمالي المساحات (لطلب المواد)${cont ? " (تابع)" : ""}`,
        W - M - 10,
        y0 + 12,
        "right",
        fBold,
        11.5,
        GREEN_TX,
        H
      );
      const heads = ["البند", `المساحة (${info.areaUnitLabel || ""})`, "المساحة (م²)"];
      heads.forEach((t, i) => {
        drawRectTop(p, ac[i].l, y0 + 24, ac[i].pw, 22, FX_HEAD_BG, MUTED, 1, H);
        drawCellText(p, t, ac[i].cx, y0 + 24, 22, ac[i].pw - 6, fBold, 10, INK, H);
      });
      return y0 + 46;
    };

    curY = drawAreaHeader(curPage, curY, false);
    info.areaTotals.forEach((item, idx) => {
      if (curY + AT_ROW_H > CONTENT_BOTTOM) {
        curPage = addNewPage();
        curY = drawAreaHeader(curPage, TY, true);
      }
      const bg = idx % 2 ? EVEN_BG : WHITE;
      const vals = [item.label, item.totalInUnit.toFixed(3), item.totalSqM.toFixed(3)];
      vals.forEach((val, ci) => {
        drawRectTop(curPage, ac[ci].l, curY, ac[ci].pw, AT_ROW_H, bg, BORDER_GRAY, 0.6, H);
        drawCellText(curPage, val, ac[ci].cx, curY, AT_ROW_H, ac[ci].pw - 6, fBold, 10, INK, H, 1);
      });
      curY += AT_ROW_H;
    });
  }

  // Duplicate Summary Table
  const dupL = info.dup || [];
  if (dupL.length > 0) {
    const SR = 21;
    const minNeeded = 14 + 24 + 22 + Math.min(dupL.length, 2) * SR;
    if (curY + minNeeded > CONTENT_BOTTOM) {
      curPage = addNewPage();
      curY = TY;
    } else {
      curY += 14;
    }

    const fw = W - 2 * M;
    const fx = [1.1, 1.1, 1, 5.8];
    const ft = fx.reduce((a, b) => a + b, 0);
    let xx = W - M;
    const dc = fx.map((w) => {
      const pw = (w / ft) * fw;
      const o = { r: xx, l: xx - pw, pw, cx: xx - pw / 2 };
      xx -= pw;
      return o;
    });

    const drawDupHeader = (p: PDFPage, y0: number, cont = false) => {
      drawRectTop(p, M, y0, fw, 24, GOLD_BG, GOLD_BD, 1, H);
      drawTextVisual(
        p,
        `ملخص الشبابيك المتشابهة${cont ? " (تابع)" : ""}`,
        W - M - 10,
        y0 + 12,
        "right",
        fBold,
        12,
        GOLD_TEXT,
        H
      );
      drawTextVisual(p, `${dupL.length} قياس مكرر`, M + 10, y0 + 12, "left", fReg, 10, GOLD_TEXT, H);
      ["طول الشباك", "عرض الشباك", "العدد", "أرقام الشبابيك"].forEach((t, i) => {
        drawRectTop(p, dc[i].l, y0 + 24, dc[i].pw, 22, FX_HEAD_BG, MUTED, 1, H);
        drawCellText(p, t, dc[i].cx, y0 + 24, 22, dc[i].pw - 6, fBold, 10, INK, H);
      });
      return y0 + 46;
    };

    curY = drawDupHeader(curPage, curY, false);
    dupL.forEach((r, idx) => {
      if (curY + SR > CONTENT_BOTTOM) {
        curPage = addNewPage();
        curY = drawDupHeader(curPage, TY, true);
      }
      const bg = idx % 2 ? EVEN_BG : WHITE;
      r.forEach((v, ci) => {
        drawRectTop(curPage, dc[ci].l, curY, dc[ci].pw, SR, bg, BORDER_GRAY, 0.6, H);
        drawCellText(
          curPage,
          String(v),
          dc[ci].cx,
          curY,
          SR,
          dc[ci].pw - 8,
          ci < 3 ? fBold : fReg,
          ci === 3 ? 9.2 : 10.5,
          INK,
          H,
          1
        );
      });
      curY += SR;
    });

    if (info.dupOther && info.dupOther > 0) {
      if (curY + 18 > CONTENT_BOTTOM) {
        curPage = addNewPage();
        curY = TY;
      }
      drawTextVisual(
        curPage,
        `وباقي الشبابيك (${info.dupOther}) قياساتها مختلفة وغير مكررة`,
        W - M,
        curY + 10,
        "right",
        fReg,
        9.5,
        MUTED,
        H
      );
    }
  }

  // Stamp total page count on every page
  const totalPages = pages.length;
  pages.forEach((p, idx) => {
    drawTextVisual(
      p,
      `صفحة ${idx + 1} من ${totalPages}`,
      M + 14,
      60,
      "left",
      fReg,
      10,
      hexToRgb("#94a3b8"),
      H
    );
  });

  const pdfBytes = await pdfDoc.save();
  return new Blob([new Uint8Array(pdfBytes)], { type: "application/pdf" });
}
