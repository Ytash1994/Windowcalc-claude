import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import appIconUrl from "./assets/app-icon.png";
import Splash from "./components/Splash";
import {
  AR_DIGITS,
  BASE_REL,
  BUILTIN_GROUPS,
  BUILTIN_TYPE_NAMES,
  CKEY,
  DEF_TYPE_DEFS,
  FONT_KEY,
  GN,
  KEY,
  MAXCM,
  MAXQ,
  PAL_KEY,
  PKEY,
  RKEY,
  THEME_KEY,
  TPL,
  UAREA,
  UAS,
  UF,
  UFULL,
  UN,
  WELCOME_KEY,
  buildResults,
  clean,
  createDefaultState,
  createRuntimeCatalog,
  cUnit,
  defaultLayout,
  defRel,
  effRel,
  fdim,
  fmt,
  fmt6,
  isCid,
  isFieldSelected,
  isGroupUnit,
  mfixValue,
  num,
  pL,
  pW,
  qOf,
  relParen,
  relSig,
  rndOwn,
  rowErr,
  rowWarn,
  shownV,
  syncLayoutWithCustom,
  validLayout,
} from "./engine/calc";
import type { RuntimeCatalog } from "./engine/calc";
import { buildXlsx } from "./export/xlsx";
import type { XlsxSettingsConfig } from "./export/xlsx";
import { parseImportedWorkbook, readXlsxSheets } from "./export/xlsxImport";
import type { ParsedWorkbook } from "./export/xlsxImport";
import { buildPdf } from "./export/pdf";
import type {
  AppState,
  PendingMode,
  RelConfig,
  RoundDir,
  SavedProject,
  StepNum,
  UnitCode,
  WindowRow,
} from "./types";

const PALS: [string, string, string][] = [
  ["mouse", "جرافيت", "#66708a"],
  ["amoled", "AMOLED أسود", "#000000"],
  ["sunlit", "نهاري (تحت الشمس)", "#0b4fa8"],
  ["titanium", "تيتانيوم ونحاس", "#D97745"],
  ["crimson", "عنّابي ياقوتي", "#9F1239"],
  ["blue", "أزرق", "#0f6fd6"],
  ["navy", "كحلي ملكي", "#1e3a8a"],
  ["emerald", "زمردي", "#047857"],
  ["slate", "فحمي وذهبي", "#2b3440"],
  ["indigo", "نيلي", "#5a49d8"],
  ["coffee", "قهوة", "#8a5a3b"],
  ["cyan", "سماوي نيون", "#22E5FF"],
  ["violet", "بنفسجي نيون", "#B28DFF"],
  ["flame", "برتقالي ناري", "#FF6B3D"],
];

const LEGACY_PAL_MAP: Record<string, string> = {
  bolt: "blue",
  teal: "emerald",
  forest: "emerald",
  steel: "navy",
  lime: "cyan",
  mint: "cyan",
  neon: "cyan",
  gold: "slate",
  amber: "flame",
};

const DARK_ONLY_PALS = new Set(["mouse", "amoled", "cyan", "violet", "flame"]);

const FONT_LEVELS: { level: number; scale: number; label: string }[] = [
  { level: 1, scale: 0.88, label: "صغير" },
  { level: 2, scale: 1, label: "عادي" },
  { level: 3, scale: 1.12, label: "كبير" },
  { level: 4, scale: 1.24, label: "كبير جداً" },
  { level: 5, scale: 1.36, label: "الأكبر" },
];

const SNAP_KEYS = [
  "sel",
  "inc",
  "dec",
  "rel",
  "rnd",
  "gr",
  "custom",
  "cn",
  "utypes",
  "layout",
  "cascade",
] as const;

function TypeIcon({ iconKey }: { iconKey: "a" | "b" | "c" | "d" | "e" }) {
  if (iconKey === "a") {
    return (
      <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
        <rect x="6" y="7" width="52" height="11" rx="2" />
        <path d="M14 12.5h36" />
        <rect x="6" y="18" width="52" height="39" rx="3" />
        <path d="M23 18v39M41 18v39" />
        <path d="M12 37h5M29 37h6M47 37h5" />
      </svg>
    );
  }
  if (iconKey === "b") {
    return (
      <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
        <rect x="4" y="14" width="56" height="38" rx="3" />
        <path d="M32 14v38" />
        <path d="M12 33h10M18 28l-6 5 6 5M52 33H42M46 28l6 5-6 5" />
      </svg>
    );
  }
  if (iconKey === "c") {
    return (
      <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
        <rect x="6" y="8" width="52" height="12" rx="2" />
        <path d="M14 14h36" />
        <rect x="6" y="20" width="52" height="36" rx="3" />
        <path d="M32 20v36" />
        <path d="M16 40h8M40 40h8" />
      </svg>
    );
  }
  if (iconKey === "d") {
    return (
      <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
        <rect x="6" y="12" width="52" height="40" rx="3" />
        <path d="M23 12v40M41 12v40" />
        <path d="M12 32h5M29 32h6M47 32h5" />
      </svg>
    );
  }
  return (
    <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="8" y="9" width="48" height="46" rx="3" />
      <path d="M20 32h24M32 20v24" />
    </svg>
  );
}

function withTrailingEmpty(rows: WindowRow[]): WindowRow[] {
  const r = rows[rows.length - 1];
  if (!r || r.L || r.W || r.p) return [...rows, { n: "", L: "", W: "" }];
  return rows;
}

function hydrateState(rawObj: unknown): AppState {
  const s = createDefaultState();
  if (rawObj && typeof rawObj === "object" && !Array.isArray(rawObj)) {
    const o = rawObj as Record<string, unknown>;
    if (Array.isArray(o.utypes)) {
      s.utypes = o.utypes
        .filter((u): u is { name: string } => Boolean(u && typeof u === "object" && typeof (u as { name?: unknown }).name === "string"))
        .map((u) => ({ name: clean(u.name).trim() }))
        .filter((u) => u.name && !BUILTIN_TYPE_NAMES.includes(u.name));
    }
    if (typeof o.type === "string" && (DEF_TYPE_DEFS[o.type] || s.utypes.some((u) => u.name === o.type))) {
      s.type = o.type;
    }
    if (Array.isArray(o.rows)) {
      const rr = o.rows
        .filter((r): r is Record<string, unknown> => Boolean(r && typeof r === "object"))
        .slice(0, 300)
        .map((r) => ({
          n: clean(String(r.n || "")).slice(0, 60),
          L: String(r.L == null ? "" : r.L).slice(0, 20),
          W: String(r.W == null ? "" : r.W).slice(0, 20),
          p: (["L", "W", "B"].includes(String(r.p)) ? String(r.p) : "") as PendingMode,
          q: String(r.q == null ? "" : r.q).replace(/[^0-9]/g, "").slice(0, 4),
        }));
      if (rr.length) s.rows = rr;
    }
    if (o.gr && typeof o.gr === "object") {
      const g = o.gr as { s?: unknown; d?: unknown };
      if (Number(g.s) > 0) {
        s.gr = {
          s: Number(g.s),
          d: (["n", "d", "u"].includes(String(g.d)) ? String(g.d) : "n") as RoundDir,
        };
      }
    }
    if (o.sel && typeof o.sel === "object") s.sel = o.sel as Record<string, boolean>;
    if (o.inc && typeof o.inc === "object") s.inc = o.inc as Record<string, string>;
    if (o.dec && typeof o.dec === "object") s.dec = o.dec as Record<string, string>;
    if (o.rel && typeof o.rel === "object") s.rel = o.rel as AppState["rel"];
    if (o.rnd && typeof o.rnd === "object") s.rnd = o.rnd as AppState["rnd"];
    if (o.custom && typeof o.custom === "object") s.custom = o.custom as AppState["custom"];
    s.cn = Number(o.cn) || 0;
    if (["m", "cm", "mm"].includes(String(o.unit))) s.unit = o.unit as UnitCode;
    s.cascade = o.cascade !== false;
    s.view = o.view === "table" ? "table" : "cards";
    s.layout = validLayout(o.layout) || defaultLayout();
    const svStep = Number(o.step);
    if ([1, 2, 3, 4].includes(svStep)) s.step = svStep as StepNum;
  }
  return s;
}

function loadInitialState(): AppState {
  let raw: Record<string, unknown> = {};
  try {
    const o = JSON.parse(localStorage.getItem(KEY) || "null");
    if (o && typeof o === "object") raw = o;
  } catch {
    // ignore
  }
  try {
    const r = JSON.parse(localStorage.getItem(RKEY) || "null");
    if (r && typeof r === "object" && !Array.isArray(r)) raw.rel = r;
  } catch {
    // ignore
  }
  try {
    const c = JSON.parse(localStorage.getItem(CKEY) || "null");
    if (c && typeof c === "object") {
      if (c.custom && typeof c.custom === "object") raw.custom = c.custom;
      if (c.cn) raw.cn = Math.max(Number(raw.cn) || 0, Number(c.cn) || 0);
      if (Array.isArray(c.utypes)) raw.utypes = c.utypes;
    }
  } catch {
    // ignore
  }

  const s = hydrateState(raw);
  Object.keys(s.custom).forEach((n) => {
    if (!Array.isArray(s.custom[n])) s.custom[n] = [];
    s.custom[n].forEach((c) => {
      c.name = clean(c.name);
      c.group = clean(c.group);
      s.cn = Math.max(s.cn, Number(String(c.id).slice(1)) || 0);
    });
  });
  s.layout = syncLayoutWithCustom(s);
  s.rows = withTrailingEmpty(s.rows);

  // Validate saved step so Step 3/4 doesn't reopen empty
  if (s.step >= 3) {
    const check = buildResults(s);
    if (!check.data.length) {
      s.step = 2;
    } else if (s.step === 3) {
      s.step = 4;
    }
  }
  return s;
}

function loadSavedProjects(): SavedProject[] {
  try {
    const raw = JSON.parse(localStorage.getItem(PKEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return raw.filter((p): p is SavedProject => Boolean(p && p.id && p.name && p.data));
  } catch {
    return [];
  }
}

const dstr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

interface ConfirmConfig {
  title: string;
  message: string;
  confirmLabel: string;
  cancelLabel?: string;
  danger?: boolean;
  onConfirm: () => void;
}

interface PendingChoice {
  r: WindowRow;
  i: number;
  ml: boolean;
  mw: boolean;
  draftL: string;
  draftW: string;
}

function ScaledDecimalInput({
  storedCm,
  scaleFactor,
  placeholder,
  zeroAsEmpty = false,
  unsigned = false,
  onCommitCm,
}: {
  storedCm: string;
  scaleFactor: number;
  placeholder?: string;
  zeroAsEmpty?: boolean;
  unsigned?: boolean;
  onCommitCm: (nextCm: string) => void;
}) {
  const formatDisplay = useCallback(
    (rawCm: string) => {
      if (rawCm === "" || rawCm == null) return "";
      const n = num(rawCm);
      if (n === null || isNaN(n)) return rawCm;
      const val = Math.round((n / scaleFactor) * 1e6) / 1e6;
      if (zeroAsEmpty && val === 0) return "";
      return String(val);
    },
    [scaleFactor, zeroAsEmpty]
  );

  const [focused, setFocused] = useState(false);
  const [draft, setDraft] = useState(() => formatDisplay(storedCm));
  const lastSentRef = useRef<string | null>(null);
  const prevScaleRef = useRef(scaleFactor);

  useEffect(() => {
    const scaleChanged = prevScaleRef.current !== scaleFactor;
    prevScaleRef.current = scaleFactor;
    if (!focused || scaleChanged || storedCm !== lastSentRef.current) {
      setDraft(formatDisplay(storedCm));
      if (scaleChanged || storedCm !== lastSentRef.current) {
        lastSentRef.current = null;
      }
    }
  }, [storedCm, scaleFactor, focused, formatDisplay]);

  return (
    <input
      inputMode="decimal"
      placeholder={placeholder}
      value={focused ? draft : formatDisplay(storedCm)}
      onFocus={() => {
        lastSentRef.current = null;
        setDraft(formatDisplay(storedCm));
        setFocused(true);
      }}
      onBlur={() => {
        setFocused(false);
        lastSentRef.current = null;
        setDraft(formatDisplay(storedCm));
      }}
      onChange={(ev) => {
        const cleaned = ev.target.value
          .replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)))
          .replace(/[٫،,]/g, ".")
          .replace(unsigned ? /[^0-9.]/g : /[^0-9.\-+]/g, "");
        setDraft(cleaned);
        if (cleaned === "" || cleaned === "." || cleaned === "-" || cleaned === "+") {
          const fallback = zeroAsEmpty ? "0" : "";
          lastSentRef.current = fallback;
          onCommitCm(fallback);
          return;
        }
        const parsed = num(cleaned);
        if (parsed !== null && !isNaN(parsed)) {
          const cmStr = String(Math.round(parsed * scaleFactor * 1e6) / 1e6);
          lastSentRef.current = cmStr;
          onCommitCm(cmStr);
        } else {
          lastSentRef.current = cleaned;
          onCommitCm(cleaned);
        }
      }}
    />
  );
}

export default function App() {
  const [showSplash, setShowSplash] = useState(true);
  const [state, setState] = useState<AppState>(() => loadInitialState());
  const [projects, setProjects] = useState<SavedProject[]>(() => loadSavedProjects());
  const [pjName, setPjName] = useState<string>("");
  const [newProjName, setNewProjName] = useState<string>("");
  const [editingProjId, setEditingProjId] = useState<string | null>(null);
  const [editingProjName, setEditingProjName] = useState<string>("");

  // Theme, Palette & Font Scale
  const [pal, setPal] = useState<string>(() => {
    try {
      const rawSaved = localStorage.getItem(PAL_KEY);
      const saved = rawSaved && LEGACY_PAL_MAP[rawSaved] ? LEGACY_PAL_MAP[rawSaved] : rawSaved;
      return saved && PALS.some((p) => p[0] === saved) ? saved : "mouse";
    } catch {
      return "mouse";
    }
  });
  const [themeMode, setThemeMode] = useState<"light" | "dark">(() => {
    try {
      return localStorage.getItem(THEME_KEY) === "light" ? "light" : "dark";
    } catch {
      return "dark";
    }
  });
  const [palBarOpen, setPalBarOpen] = useState(false);
  const [fontOpen, setFontOpen] = useState(false);
  const [fontLevel, setFontLevel] = useState<number>(() => {
    try {
      const v = Number(localStorage.getItem(FONT_KEY));
      return v >= 1 && v <= 5 ? v : 2;
    } catch {
      return 2;
    }
  });

  // UI Popups, Modals & Banners
  const [welOpen, setWelOpen] = useState(false);
  const [saveOpen, setSaveOpen] = useState(false);
  const [altOpen, setAltOpen] = useState(false);
  const [altTextOpen, setAltTextOpen] = useState(false);
  const [copyByNumOpen, setCopyByNumOpen] = useState(false);
  const [copyNumVal, setCopyNumVal] = useState("");
  const [badgeModalIdx, setBadgeModalIdx] = useState<number | null>(null);
  const [badgeInputVal, setBadgeInputVal] = useState("");
  const [exitPromptOpen, setExitPromptOpen] = useState(false);
  const [confirmDlg, setConfirmDlg] = useState<ConfirmConfig | null>(null);

  // Pending check modal
  const [pendChoices, setPendChoices] = useState<PendingChoice[] | null>(null);
  const pendAfterRef = useRef<((latestRows?: WindowRow[]) => void) | null>(null);

  // Step 2 inline states
  const [editingNameIdx, setEditingNameIdx] = useState<number | null>(null);
  const [openPendingBarIdx, setOpenPendingBarIdx] = useState<number | null>(null);
  const [deletedUndo, setDeletedUndo] = useState<{ row: WindowRow; index: number } | null>(null);
  const [negHighlightedRows, setNegHighlightedRows] = useState<WeakSet<WindowRow>>(() => new WeakSet());
  const [pasteText, setPasteText] = useState("");

  // Step 1 Type/Rel editor
  const [typeEdit, setTypeEdit] = useState(false);
  const [relFor, setRelFor] = useState<string | null>(null);
  const [ntForm, setNtForm] = useState(false);
  const [ntName, setNtName] = useState("");
  const [typeRenameVal, setTypeRenameVal] = useState("");

  // Step 3 Customize
  const [advOpen, setAdvOpen] = useState(false);
  const [openK, setOpenK] = useState<string | null>(null);
  const [colG, setColG] = useState<Set<string>>(() => new Set());
  const [snap3, setSnap3] = useState<string | null>(null);
  const [relUndo, setRelUndo] = useState<string[]>([]);
  const undoTimerRef = useRef(0);
  const [dragUnit, setDragUnit] = useState<{ id: string; targetId: string | null; pos: "before" | "after" } | null>(null);
  const [dragKid, setDragKid] = useState<{
    unitId: string;
    fieldId: string;
    targetFieldId: string | null;
    pos: "before" | "after";
  } | null>(null);
  const dragCleanupRef = useRef<(() => void) | null>(null);

  // Step 4 Results
  const [rndOpen, setRndOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

  // Excel import modal state
  const [xlImportState, setXlImportState] = useState<{
    fileName: string;
    parsed: ParsedWorkbook;
    useRows: boolean;
    useSettings: boolean;
    unit: UnitCode;
    mode: "add" | "rep";
    err: string;
  } | null>(null);

  // Toast & Save/Open bars
  const [toastMsg, setToastMsg] = useState("");
  const [toastOn, setToastOn] = useState(false);
  const toastTimerRef = useRef<number>(0);
  const [openBar, setOpenBar] = useState<{ txt: string; web: boolean; uri?: string; mime: string } | null>(null);
  const [fallBarVisible, setFallBarVisible] = useState(false);
  const lastSavedBlobRef = useRef<{ blob: Blob; filename: string; mime: string } | null>(null);
  const bkFileRef = useRef<HTMLInputElement>(null);
  const xlFileRef = useRef<HTMLInputElement>(null);
  const rowsBoxRef = useRef<HTMLDivElement>(null);
  const copyNumInputRef = useRef<HTMLInputElement>(null);
  const badgeInputRef = useRef<HTMLInputElement>(null);
  const noSaveRef = useRef(false);

  const toast = useCallback((msg: string, ms = 2400) => {
    setToastMsg(msg);
    setToastOn(true);
    window.clearTimeout(toastTimerRef.current);
    toastTimerRef.current = window.setTimeout(() => setToastOn(false), ms);
  }, []);

  // Derived Runtime Catalog & Results
  const cat: RuntimeCatalog = useMemo(() => createRuntimeCatalog(state), [state]);
  const res = useMemo(() => buildResults(state, negHighlightedRows), [state, negHighlightedRows]);

  // Apply Theme, Palette & Font Scale to <html>
  useEffect(() => {
    const activeTheme = DARK_ONLY_PALS.has(pal) ? "dark" : themeMode;
    document.documentElement.setAttribute("data-pal", pal);
    document.documentElement.setAttribute("data-theme", activeTheme);
    try {
      localStorage.setItem(PAL_KEY, pal);
      localStorage.setItem(THEME_KEY, themeMode);
    } catch {
      // ignore
    }
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) {
      const h2 = getComputedStyle(document.documentElement).getPropertyValue("--h2").trim() || "#0b111f";
      meta.setAttribute("content", h2);
    }
  }, [pal, themeMode]);

  // Keep browser/standalone favicon synced with bundled app icon
  useEffect(() => {
    document
      .querySelectorAll<HTMLLinkElement>('link[rel="icon"], link[rel="apple-touch-icon"]')
      .forEach((lnk) => {
        lnk.href = appIconUrl;
      });
  }, []);

  useEffect(() => {
    const found = FONT_LEVELS.find((f) => f.level === fontLevel) || FONT_LEVELS[1];
    document.documentElement.style.setProperty("--fscale", String(found.scale));
    try {
      localStorage.setItem(FONT_KEY, String(found.level));
    } catch {
      // ignore
    }
  }, [fontLevel]);

  // Body wide table mode on Step 4 table view
  useEffect(() => {
    document.body.classList.toggle("wide4", state.step === 4 && state.view === "table");
  }, [state.step, state.view]);

  // Persist AppState & Projects to localStorage
  const persistState = useCallback((s: AppState) => {
    if (noSaveRef.current) return;
    try {
      localStorage.setItem(
        KEY,
        JSON.stringify({
          type: s.type,
          rows: s.rows,
          sel: s.sel,
          inc: s.inc,
          dec: s.dec,
          rel: s.rel,
          rnd: s.rnd,
          gr: s.gr,
          custom: s.custom,
          cn: s.cn,
          utypes: s.utypes,
          cascade: s.cascade,
          view: s.view,
          layout: s.layout,
          unit: s.unit,
          step: s.step,
        })
      );
      localStorage.setItem(RKEY, JSON.stringify(s.rel || {}));
      localStorage.setItem(CKEY, JSON.stringify({ custom: s.custom, cn: s.cn, utypes: s.utypes }));
    } catch {
      // ignore
    }
  }, []);

  useEffect(() => {
    const t = window.setTimeout(() => persistState(state), 220);
    return () => clearTimeout(t);
  }, [state, persistState]);

  useEffect(() => {
    const onHide = () => persistState(state);
    window.addEventListener("pagehide", onHide);
    return () => window.removeEventListener("pagehide", onHide);
  }, [state, persistState]);

  useEffect(() => {
    try {
      localStorage.setItem(PKEY, JSON.stringify(projects));
    } catch {
      // ignore
    }
  }, [projects]);

  // Show welcome modal on first visit after Splash closes
  useEffect(() => {
    if (showSplash) return;
    try {
      if (!localStorage.getItem(WELCOME_KEY)) {
        setWelOpen(true);
      }
    } catch {
      // ignore
    }
  }, [showSplash]);

  // Sync rename input when opening relFor
  useEffect(() => {
    if (relFor) setTypeRenameVal(relFor);
  }, [relFor]);

  // Keep Step 1 project name input synced with active pjName without blocking manual clear
  useEffect(() => {
    setNewProjName(pjName);
  }, [pjName]);

  // Keep latest UI state in refs for clean popstate back-button handling
  const stateRef = useRef(state);
  stateRef.current = state;
  const snap3Ref = useRef(snap3);
  snap3Ref.current = snap3;

  const snap3String = useCallback((s: AppState) => {
    const o: Record<string, unknown> = {};
    SNAP_KEYS.forEach((k) => {
      o[k] = s[k];
    });
    return JSON.stringify(o);
  }, []);

  const anyModalOpen = Boolean(
    welOpen ||
      saveOpen ||
      altOpen ||
      copyByNumOpen ||
      badgeModalIdx !== null ||
      xlImportState !== null ||
      exitPromptOpen ||
      confirmDlg ||
      pendChoices
  );
  const anyModalRef = useRef(anyModalOpen);
  anyModalRef.current = anyModalOpen;

  const closeTopModal = useCallback(() => {
    if (confirmDlg) {
      setConfirmDlg(null);
      return true;
    }
    if (exitPromptOpen) {
      setExitPromptOpen(false);
      return true;
    }
    if (xlImportState !== null) {
      setXlImportState(null);
      return true;
    }
    if (badgeModalIdx !== null) {
      setBadgeModalIdx(null);
      return true;
    }
    if (copyByNumOpen) {
      setCopyByNumOpen(false);
      return true;
    }
    if (pendChoices) {
      setPendChoices(null);
      pendAfterRef.current = null;
      return true;
    }
    if (altOpen) {
      setAltOpen(false);
      return true;
    }
    if (saveOpen) {
      setSaveOpen(false);
      return true;
    }
    if (welOpen) {
      setWelOpen(false);
      try {
        localStorage.setItem(WELCOME_KEY, "1");
      } catch {
        // ignore
      }
      return true;
    }
    return false;
  }, [confirmDlg, exitPromptOpen, xlImportState, badgeModalIdx, copyByNumOpen, pendChoices, altOpen, saveOpen, welOpen]);

  const closeTopModalRef = useRef(closeTopModal);
  closeTopModalRef.current = closeTopModal;

  // Browser / Android Hardware Back Button Handling + Exit Confirmation Popup
  useEffect(() => {
    // Anchor an exit-guard state at load so hardware Back on Step 1 triggers popstate cleanly
    window.history.replaceState({ appSentinel: true }, "");
    window.history.pushState({ step: stateRef.current.step }, "");

    const onPopState = () => {
      if (anyModalRef.current) {
        closeTopModalRef.current();
        window.history.pushState({ step: stateRef.current.step }, "");
        return;
      }

      const cur = stateRef.current;
      if (cur.step === 3 && snap3Ref.current && snap3String(cur) !== snap3Ref.current) {
        window.history.pushState({ step: 3 }, "");
        const snapshot = snap3Ref.current;
        setConfirmDlg({
          title: "تغييرات غير محفوظة",
          message: "لديك تغييرات لم تُحفظ في تخصيص الخانات. هل تريد إلغاءها والرجوع للنتائج؟",
          confirmLabel: "إلغاء التغييرات والرجوع",
          danger: true,
          onConfirm: () => {
            const parsed = JSON.parse(snapshot);
            setState((prev) => {
              const next = { ...prev, step: 4 as StepNum };
              SNAP_KEYS.forEach((k) => {
                (next as Record<string, unknown>)[k] = parsed[k];
              });
              next.layout = syncLayoutWithCustom(next);
              return next;
            });
            setSnap3(null);
            toast("تم إلغاء التغييرات");
          },
        });
        return;
      }

      if (cur.step === 4) {
        setState((prev) => ({ ...prev, step: 2 }));
        window.history.pushState({ step: 2 }, "");
        window.scrollTo(0, 0);
        return;
      }
      if (cur.step === 3) {
        setSnap3(null);
        setState((prev) => ({ ...prev, step: 4 }));
        window.history.pushState({ step: 4 }, "");
        window.scrollTo(0, 0);
        return;
      }
      if (cur.step === 2) {
        setState((prev) => ({ ...prev, step: 1 }));
        window.history.pushState({ step: 1 }, "");
        window.scrollTo(0, 0);
        return;
      }

      // Already at Step 1 and nothing left to go back to: show smooth custom exit warning popup
      setExitPromptOpen(true);
    };

    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, [snap3String, toast]);

  // Escape key closes topmost dialog
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && anyModalRef.current) {
        closeTopModalRef.current();
      }
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, []);

  // Pending-measurements check before advancing to Step 3/4/Save
  const findNeedChoice = useCallback(
    (rows: WindowRow[], prevChoices?: PendingChoice[] | null): PendingChoice[] => {
      const draftMap = new WeakMap<WindowRow, { draftL: string; draftW: string }>();
      prevChoices?.forEach((c) => {
        draftMap.set(c.r, { draftL: c.draftL, draftW: c.draftW });
      });
      let lf = -1;
      rows.forEach((r, i) => {
        if (num(r.L) !== null || num(r.W) !== null || r.p) lf = i;
      });
      const out: PendingChoice[] = [];
      rows.forEach((r, i) => {
        if (i > lf) return;
        const ml = num(r.L) === null && !pL(r);
        const mw = num(r.W) === null && !pW(r);
        if (ml || mw) {
          const prevDraft = draftMap.get(r);
          out.push({
            r,
            i,
            ml,
            mw,
            draftL: prevDraft ? prevDraft.draftL : pL(r) ? "" : r.L,
            draftW: prevDraft ? prevDraft.draftW : pW(r) ? "" : r.W,
          });
        }
      });
      return out;
    },
    []
  );

  const askPend = useCallback(
    (after: (latestRows?: WindowRow[]) => void, rowsOverride?: WindowRow[]) => {
      const list = findNeedChoice(rowsOverride || state.rows);
      if (!list.length) return false;
      pendAfterRef.current = after;
      setPendChoices(list);
      return true;
    },
    [findNeedChoice, state.rows]
  );

  // Navigation helper with all validation guards
  const goToStep = useCallback(
    (target: StepNum, skipPendCheck = false, stateOverride?: AppState, skipUnsavedCheck = false) => {
      const cur = stateOverride || state;
      if (!skipUnsavedCheck && cur.step === 3 && target !== 3 && snap3Ref.current && snap3String(cur) !== snap3Ref.current) {
        const snapshot = snap3Ref.current;
        setConfirmDlg({
          title: "تغييرات غير محفوظة",
          message: "لديك تغييرات لم تُحفظ. هل تريد إلغاءها والمتابعة؟",
          confirmLabel: "إلغاء التغييرات",
          danger: true,
          onConfirm: () => {
            const parsed = JSON.parse(snapshot);
            const reverted = { ...cur };
            SNAP_KEYS.forEach((k) => {
              (reverted as Record<string, unknown>)[k] = parsed[k];
            });
            reverted.layout = syncLayoutWithCustom(reverted);
            snap3Ref.current = null;
            setSnap3(null);
            setState(reverted);
            toast("تم إلغاء التغييرات");
            goToStep(target, skipPendCheck, reverted, true);
          },
        });
        return;
      }

      if (!skipPendCheck && target >= 3 && cur.step < 3) {
        if (
          askPend(
            (latestRows) => goToStep(target, true, latestRows ? { ...cur, rows: latestRows } : undefined),
            cur.rows
          )
        )
          return;
      }

      if (target >= 3) {
        const computed = buildResults(cur);
        if (computed.bad) {
          toast(computed.bad, 3500);
          return;
        }
        const badIdx = cur.rows.findIndex((r) => Boolean(rowErr(r, cur.unit, negHighlightedRows.has(r))));
        if (badIdx >= 0) {
          toast(`الشباك ${badIdx + 1}: ${rowErr(cur.rows[badIdx], cur.unit, negHighlightedRows.has(cur.rows[badIdx]))}`, 3500);
          setState((prev) => ({ ...prev, step: 2 }));
          window.setTimeout(() => {
            const el = rowsBoxRef.current?.querySelectorAll(".wcard")?.[badIdx] as HTMLElement | undefined;
            el?.scrollIntoView({ block: "center", behavior: "smooth" });
          }, 80);
          return;
        }
        if (!computed.data.length) {
          toast("اكتب طول وعرض شباك واحد على الأقل");
          setState((prev) => ({ ...prev, step: 2 }));
          return;
        }
        if (cur.step < 3 && computed.negList.length) {
          const ws = new WeakSet<WindowRow>();
          computed.data.forEach((d) => {
            if (d.neg) ws.add(d.row);
          });
          setNegHighlightedRows(ws);
          toast(`الشباك ${computed.negList.join("، ")} صغير جداً: ستظهر قياسات سالبة. كبّر القياس`, 4500);
          setState((prev) => ({ ...prev, step: 2 }));
          window.setTimeout(() => {
            const el = rowsBoxRef.current?.querySelectorAll(".wcard")?.[computed.negList[0] - 1] as HTMLElement | undefined;
            el?.scrollIntoView({ block: "center", behavior: "smooth" });
          }, 80);
          return;
        }
      }

      if (target === 3) {
        const snap = snap3String(cur);
        snap3Ref.current = snap;
        setSnap3(snap);
        setRelUndo([]);
        setOpenK(null);
      } else {
        snap3Ref.current = null;
        setSnap3(null);
      }

      if (target !== 1) {
        setRelFor(null);
        setNtForm(false);
      }

      setState((prev) => ({ ...prev, step: target }));
      window.scrollTo(0, 0);
    },
    [state, snap3String, askPend, negHighlightedRows, toast]
  );

  // ---------- Step 2 Actions ----------
  const updateRow = (idx: number, patch: Partial<WindowRow>) => {
    setState((prev) => {
      const rows = prev.rows.map((r, i) => (i === idx ? { ...r, ...patch } : r));
      const updated = withTrailingEmpty(rows);
      return { ...prev, rows: updated };
    });
  };

  const handleDeleteRow = (idx: number) => {
    const targetRow = state.rows[idx];
    const hasContent = Boolean(targetRow && (targetRow.L || targetRow.W || targetRow.n || targetRow.p || (targetRow.q && targetRow.q !== "1")));
    setState((prev) => {
      let rows: WindowRow[];
      if (prev.rows.length > 1) {
        rows = prev.rows.filter((_, i) => i !== idx);
      } else {
        rows = [{ n: "", L: "", W: "" }];
      }
      return { ...prev, rows: withTrailingEmpty(rows) };
    });
    if (hasContent) {
      setDeletedUndo({ row: { ...targetRow }, index: idx });
    } else {
      setDeletedUndo(null);
    }
    setOpenPendingBarIdx(null);
    setEditingNameIdx(null);
  };

  const handleUndoDelete = () => {
    if (!deletedUndo) return;
    const { row, index } = deletedUndo;
    setState((prev) => {
      let rows = prev.rows.slice();
      // If all we had was the empty placeholder row, replace it cleanly
      if (rows.length === 1 && !rows[0].L && !rows[0].W && !rows[0].n && !rows[0].p) {
        rows = [row];
      } else {
        const insertAt = Math.min(index, rows.length);
        rows.splice(insertAt, 0, row);
      }
      return { ...prev, rows: withTrailingEmpty(rows) };
    });
    setDeletedUndo(null);
    toast("تمت استعادة الشباك المحذوف ✓");
  };

  const handleAddEmptyRow = () => {
    setDeletedUndo(null);
    setState((prev) => {
      const lastRow = prev.rows[prev.rows.length - 1];
      const rows =
        lastRow && (lastRow.L || lastRow.W || lastRow.p)
          ? [...prev.rows, { n: "", L: "", W: "" }]
          : prev.rows;
      return { ...prev, rows };
    });
    window.setTimeout(() => {
      const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
      const lastCard = cards?.[cards.length - 1] as HTMLElement | undefined;
      lastCard?.scrollIntoView({ block: "center", behavior: "smooth" });
      const lInp = lastCard?.querySelector(".L") as HTMLInputElement | null;
      lInp?.focus();
    }, 60);
  };

  const handleDupLast = () => {
    let lastValid = -1;
    state.rows.forEach((r, k) => {
      if (num(r.L) !== null && num(r.W) !== null && !rowErr(r, state.unit, negHighlightedRows.has(r))) {
        lastValid = k;
      }
    });
    if (lastValid < 0) {
      toast("اكتب شباكاً واحداً أولاً ثم انسخه");
      return;
    }
    const r = state.rows[lastValid];
    const copy: WindowRow = { n: "", L: r.L, W: r.W, q: r.q || "", p: r.p || "" };
    setDeletedUndo(null);
    setState((prev) => {
      const rows = prev.rows.slice();
      rows.splice(lastValid + 1, 0, copy);
      return { ...prev, rows: withTrailingEmpty(rows) };
    });
    window.setTimeout(() => {
      const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
      const card = cards?.[lastValid + 1] as HTMLElement | undefined;
      card?.scrollIntoView({ block: "center", behavior: "smooth" });
      const qInp = card?.querySelector(".Q") as HTMLInputElement | null;
      qInp?.focus();
    }, 60);
    toast("تم نسخ الشباك السابق. عدّل ما يختلف فيه");
  };

  const openCopyByNumber = () => {
    const validCount = state.rows.filter((r) => r.L || r.W || r.p || r.n).length;
    if (!validCount) {
      toast("اكتب شباكاً واحداً أولاً ثم انسخه");
      return;
    }
    setCopyNumVal("");
    setCopyByNumOpen(true);
    window.setTimeout(() => copyNumInputRef.current?.focus(), 60);
  };

  const copyTargetIdx = useMemo(() => {
    const parsed = num(copyNumVal);
    if (parsed === null || isNaN(parsed) || !Number.isInteger(parsed)) return -1;
    const idx = parsed - 1;
    if (idx < 0 || idx >= state.rows.length) return -1;
    const r = state.rows[idx];
    if (!r.L && !r.W && !r.p && !r.n) return -1;
    return idx;
  }, [copyNumVal, state.rows]);

  const confirmCopyByNumber = () => {
    if (copyTargetIdx < 0) {
      toast("أدخل رقم شباك صحيحاً من القائمة");
      return;
    }
    const src = state.rows[copyTargetIdx];
    const copy: WindowRow = {
      n: src.n,
      L: src.L,
      W: src.W,
      q: src.q || "",
      p: src.p || "",
    };
    setDeletedUndo(null);
    let insertedIdx = 0;
    setState((prev) => {
      const cleaned = prev.rows.filter((r, i) => i < prev.rows.length - 1 || r.L || r.W || r.p || r.n);
      insertedIdx = cleaned.length;
      return { ...prev, rows: withTrailingEmpty([...cleaned, copy]) };
    });
    setCopyByNumOpen(false);
    window.setTimeout(() => {
      const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
      const card = cards?.[insertedIdx] as HTMLElement | undefined;
      card?.scrollIntoView({ block: "center", behavior: "smooth" });
    }, 60);
    toast(`تم نسخ شباك رقم ${copyTargetIdx + 1} ✓`);
  };

  const openBadgeModal = (idx: number) => {
    const r = state.rows[idx];
    setBadgeInputVal(r?.n || String(idx + 1));
    setBadgeModalIdx(idx);
    window.setTimeout(() => {
      badgeInputRef.current?.focus();
      badgeInputRef.current?.select();
    }, 60);
  };

  const confirmBadgeAction = () => {
    if (badgeModalIdx === null) return;
    const idx = badgeModalIdx;
    const rawTrim = clean(badgeInputVal).trim();
    const normDigits = rawTrim.replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));
    if (/^[-+]?[0-9]+([.,٫][0-9]+)?$/.test(normDigits) && !/^[0-9]+$/.test(normDigits)) {
      toast("اكتب رقماً صحيحاً موجباً، أو اسماً يحتوي حروفاً", 3500);
      return;
    }
    if (/^[0-9]+$/.test(normDigits)) {
      const targetNum = parseInt(normDigits, 10);
      if (targetNum < 1 || targetNum > 300) {
        toast("اكتب رقماً بين 1 و300");
        return;
      }
      if (targetNum === idx + 1) {
        setBadgeModalIdx(null);
        return;
      }
      const nextRows = [...state.rows];
      while (nextRows.length < targetNum) {
        nextRows.push({ n: "", L: "", W: "" });
      }
      const moved = nextRows[idx];
      const other = nextRows[targetNum - 1];
      nextRows[targetNum - 1] = moved;
      nextRows[idx] = other;
      const finalRows = withTrailingEmpty(nextRows);
      setState((prev) => ({ ...prev, rows: finalRows }));
      setBadgeModalIdx(null);
      setDeletedUndo(null);
      setOpenPendingBarIdx(null);
      setEditingNameIdx(null);
      const hasOther =
        num(other.L) !== null || num(other.W) !== null || Boolean(other.p) || Boolean(other.n);
      toast(
        hasOther
          ? `انتقل الشباك إلى الرقم ${targetNum} وتبدّل مع الشباك الذي كان فيه (صار رقم ${idx + 1})`
          : `انتقل الشباك إلى الرقم ${targetNum}`
      );
      window.setTimeout(() => {
        const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
        const card = cards?.[targetNum - 1] as HTMLElement | undefined;
        card?.scrollIntoView({ block: "center", behavior: "smooth" });
      }, 60);
    } else {
      const newName = rawTrim.slice(0, 60);
      updateRow(idx, { n: newName });
      setBadgeModalIdx(null);
      toast(newName ? `تمت تسمية الشباك #${idx + 1}: ${newName}` : `تم مسح اسم الشباك #${idx + 1}`);
    }
  };

  const isRowBlank = (r: WindowRow) => !r.L && !r.W && !r.p && !r.n;
  const middleEmptyCount = useMemo(
    () => state.rows.filter((r, i) => i < state.rows.length - 1 && isRowBlank(r)).length,
    [state.rows]
  );

  const handleCleanEmptyRows = () => {
    const kept = state.rows.filter((r, i) => i === state.rows.length - 1 || !isRowBlank(r));
    const removed = state.rows.length - kept.length;
    if (!removed) return;
    setDeletedUndo(null);
    setOpenPendingBarIdx(null);
    setEditingNameIdx(null);
    setState((prev) => ({ ...prev, rows: withTrailingEmpty(kept) }));
    toast(`تم حذف ${removed} ${removed === 1 ? "صف فارغ" : "صفوف فارغة"} 🗑`);
  };

  const handleRemovePendingRows = (indices: Set<number>) => {
    const filtered = state.rows.filter((_, i) => !indices.has(i));
    const updatedRows = withTrailingEmpty(filtered.length ? filtered : [{ n: "", L: "", W: "" }]);
    setDeletedUndo(null);
    setOpenPendingBarIdx(null);
    setEditingNameIdx(null);
    setState((prev) => ({ ...prev, rows: updatedRows }));
    const rem = findNeedChoice(updatedRows, pendChoices);
    if (rem.length) {
      setPendChoices(rem);
      toast(`تم حذف ${indices.size === 1 ? "الصف" : `${indices.size} صفوف`} 🗑`);
    } else {
      const cb = pendAfterRef.current;
      pendAfterRef.current = null;
      setPendChoices(null);
      const hasAnyData = updatedRows.some((r) => Boolean(r.L || r.W || r.p));
      if (hasAnyData && cb) {
        window.setTimeout(() => cb(updatedRows), 100);
      } else {
        setState((prev) => ({ ...prev, rows: updatedRows, step: 2 }));
        toast("تم حذف الصفوف الفارغة — أدخل قياساً للمتابعة");
      }
    }
  };

  const handleUnitChange = (nu: UnitCode) => {
    if (nu === state.unit) return;
    const k = UF[state.unit] / UF[nu];
    const cv = (v: string, dim = 1) => {
      const n = num(v);
      return n === null || isNaN(n) ? v : fmt6(n * Math.pow(k, dim));
    };
    const hadAny = state.rows.some((r) => r.L || r.W);
    setState((prev) => {
      const nextInc: Record<string, string> = {};
      const nextDec: Record<string, string> = {};
      Object.keys(prev.inc).forEach((x) => {
        nextInc[x] = cv(prev.inc[x], fdim(cat, x));
      });
      Object.keys(prev.dec).forEach((x) => {
        nextDec[x] = cv(prev.dec[x], fdim(cat, x));
      });
      return {
        ...prev,
        unit: nu,
        rows: prev.rows.map((r) => ({ ...r, L: cv(r.L, 1), W: cv(r.W, 1) })),
        inc: nextInc,
        dec: nextDec,
      };
    });
    toast(`الوحدة الآن: ${UN[nu]}${hadAny ? " — حُوّلت القياسات المدخلة تلقائياً" : ""}`);
  };

  const handleClearAll = () => {
    if (!state.rows.some((r) => r.L || r.W || r.n || r.p)) {
      toast("لا توجد قياسات لمسحها");
      return;
    }
    setConfirmDlg({
      title: "مسح كل القياسات",
      message: "هل تريد مسح كل قياسات الشبابيك الحالية؟",
      confirmLabel: "مسح الكل",
      danger: true,
      onConfirm: () => {
        setDeletedUndo(null);
        setState((prev) => ({ ...prev, rows: [{ n: "", L: "", W: "" }] }));
        toast("تم مسح كل القياسات");
      },
    });
  };

  const handlePaste = () => {
    const out: { L: number; W: number; n: string }[] = [];
    pasteText.split(/\r?\n/).forEach((line) => {
      const parts = line
        .trim()
        .split(/[\t ;]+/)
        .filter(Boolean);
      if (parts.length < 2) return;
      const L = num(parts[0]);
      const W = num(parts[1]);
      if (L !== null && !isNaN(L) && L > 0 && W !== null && !isNaN(W) && W > 0) {
        const namePart = clean(parts.slice(2).join(" ")).trim().slice(0, 60);
        out.push({ L, W, n: namePart });
      }
    });
    if (!out.length) {
      toast("لم أجد أسطراً بالشكل: طول عرض");
      return;
    }
    setDeletedUndo(null);
    setState((prev) => {
      const kept = prev.rows.filter((r) => r.L || r.W || r.p || r.n);
      const added: WindowRow[] = out.map((item) => ({
        n: item.n,
        L: fmt(item.L, prev.unit),
        W: fmt(item.W, prev.unit),
      }));
      return { ...prev, rows: withTrailingEmpty([...kept, ...added]) };
    });
    setPasteText("");
    toast(`تمت إضافة ${out.length} شباك`);
  };

  // ---------- Projects Management ----------
  const handleSaveCurrentProject = () => {
    const name = clean(newProjName).trim();
    if (!name) {
      toast("اكتب اسم المشروع أولاً");
      return;
    }
    const { raw: _ignored, ...rest } = state;
    void _ignored;
    const nowIso = new Date().toISOString();
    setProjects((prev) => {
      const existing = prev.find((p) => p.name === name);
      if (existing) {
        return prev.map((p) => (p.id === existing.id ? { ...p, updatedAt: nowIso, data: rest } : p));
      }
      return [{ id: `pj_${Date.now()}`, name, updatedAt: nowIso, data: rest }, ...prev];
    });
    setPjName(name);
    setNewProjName(name);
    toast(`تم حفظ مشروع «${name}» ✓`);
  };

  const handleOpenProject = (proj: SavedProject) => {
    const hydrated = hydrateState(proj.data);
    // Preserve newer custom user types created after this project was saved
    state.utypes.forEach((u) => {
      if (!hydrated.utypes.some((x) => x.name === u.name)) {
        hydrated.utypes.push(u);
        if (state.custom[u.name] && !hydrated.custom[u.name]) {
          hydrated.custom[u.name] = state.custom[u.name].map((c) => ({ ...c }));
        }
        if (state.rel[u.name] && !hydrated.rel[u.name]) {
          hydrated.rel[u.name] = { ...state.rel[u.name] };
        }
        if (state.rnd[u.name] && !hydrated.rnd[u.name]) {
          hydrated.rnd[u.name] = { ...state.rnd[u.name] };
        }
        (state.custom[u.name] || []).forEach((c) => {
          if (c.id in state.sel && !(c.id in hydrated.sel)) {
            hydrated.sel[c.id] = state.sel[c.id];
          }
        });
      }
    });
    hydrated.cn = Math.max(hydrated.cn, state.cn || 0);
    hydrated.layout = syncLayoutWithCustom(hydrated);
    hydrated.rows = withTrailingEmpty(hydrated.rows);
    hydrated.step = hydrated.rows.some((r) => r.L || r.W || r.p) ? 2 : 1;
    setDeletedUndo(null);
    setNegHighlightedRows(new WeakSet());
    setPjName(proj.name);
    setState(hydrated);
    toast(`تم فتح مشروع «${proj.name}»`);
  };

  const handleNewBlankProject = () => {
    const startFresh = () => {
      setDeletedUndo(null);
      setNegHighlightedRows(new WeakSet());
      setPjName("");
      setState((prev) => ({
        ...createDefaultState(),
        utypes: prev.utypes,
        custom: prev.custom,
        cn: prev.cn,
        rel: prev.rel,
        sel: prev.sel,
        layout: prev.layout,
        unit: prev.unit,
        step: 2,
      }));
      toast("بدأ مشروع جديد فارغ");
    };
    if (state.rows.some((r) => r.L || r.W || r.p || r.n)) {
      setConfirmDlg({
        title: "بدء مشروع جديد",
        message: "سيتم تفريغ قياسات المشروع الحالي (احفظه أولاً إذا أردت الرجوع إليه). متابعة؟",
        confirmLabel: "مشروع جديد",
        onConfirm: startFresh,
      });
    } else {
      startFresh();
    }
  };

  // ---------- Step 3 & Relation Editor Helpers ----------
  const pushUndo3 = (force = false) => {
    const j = snap3String(state);
    if (relUndo.length && relUndo[relUndo.length - 1] === j) return;
    const now = Date.now();
    if (!force && now - undoTimerRef.current < 1500) return;
    undoTimerRef.current = now;
    setRelUndo((prev) => {
      const next = [...prev, j];
      return next.length > 40 ? next.slice(next.length - 40) : next;
    });
  };

  const cuv = (a: number) => fmt(a / UF[state.unit], state.unit);

  const srcOptions = (k: string, typeName: string) => {
    const t = cat.TYPES[typeName];
    if (!t) return ["LEN", "WID"];
    return [
      "LEN",
      "WID",
      ...cat.ORDER.slice(0, cat.ORDER.indexOf(k)).filter(
        (x) => x !== "E" && t.fields.includes(x)
      ),
    ];
  };

  const setRelPatch = (k: string, patch: Partial<RelConfig>) => {
    setState((prev) => {
      const curOv = (prev.rel && prev.rel[prev.type]) || {};
      const cur = curOv[k] || defRel(prev.type, k);
      const updated: RelConfig = { ...cur, ...patch };
      const typeRel = { ...curOv };
      if (relSig(updated) === relSig(defRel(prev.type, k))) {
        delete typeRel[k];
      } else {
        typeRel[k] = updated;
      }
      return {
        ...prev,
        rel: { ...prev.rel, [prev.type]: typeRel },
      };
    });
  };

  const setRndPatch = (k: string, patch: Partial<{ s: number; d: RoundDir }>) => {
    setState((prev) => {
      const typeRnd = { ...((prev.rnd && prev.rnd[prev.type]) || {}) };
      const prevCfg = typeRnd[k] || { s: 0, d: "n" as RoundDir };
      const n = { ...prevCfg, ...patch };
      if (!(Number(n.s) > 0)) delete typeRnd[k];
      else typeRnd[k] = n;
      return { ...prev, rnd: { ...prev.rnd, [prev.type]: typeRnd } };
    });
  };

  const addCustomField = (group = "") => {
    let createdId = "";
    setState((prev) => {
      const cn = (prev.cn || 0) + 1;
      createdId = `c${cn}`;
      const list = prev.custom[prev.type] || [];
      const nextCustom = {
        ...prev.custom,
        [prev.type]: [
          ...list,
          { id: createdId, name: `خانة جديدة ${list.length + 1}`, group },
        ],
      };
      const next: AppState = {
        ...prev,
        cn,
        custom: nextCustom,
        sel: { ...prev.sel, [createdId]: true },
      };
      next.layout = syncLayoutWithCustom(next);
      return next;
    });
    window.setTimeout(() => {
      setOpenK(createdId);
      toast("أُضيفت خانة جديدة. اكتب اسمها وحدّد علاقتها");
    }, 30);
  };

  const delCustomField = (id: string) => {
    const o = (state.rel && state.rel[state.type]) || {};
    const dep = Object.keys(o).find((k) => o[k].s === id || o[k].s2 === id);
    if (dep) {
      toast(`لا يمكن الحذف: خانة «${cat.LABEL[dep]}» مبنية عليها`);
      return;
    }
    setConfirmDlg({
      title: "حذف الخانة",
      message: "هل تريد حذف هذه الخانة نهائياً؟",
      confirmLabel: "حذف",
      danger: true,
      onConfirm: () => {
        setState((prev) => {
          const typeName = prev.type;
          const nextCustom = {
            ...prev.custom,
            [typeName]: (prev.custom[typeName] || []).filter((c) => c.id !== id),
          };
          const nextRelType = { ...((prev.rel && prev.rel[typeName]) || {}) };
          delete nextRelType[id];
          const nextRndType = { ...((prev.rnd && prev.rnd[typeName]) || {}) };
          delete nextRndType[id];
          const nextSel = { ...prev.sel };
          delete nextSel[id];
          const nextInc = { ...prev.inc };
          delete nextInc[id];
          const nextDec = { ...prev.dec };
          delete nextDec[id];
          const next: AppState = {
            ...prev,
            custom: nextCustom,
            rel: { ...prev.rel, [typeName]: nextRelType },
            rnd: { ...prev.rnd, [typeName]: nextRndType },
            sel: nextSel,
            inc: nextInc,
            dec: nextDec,
          };
          next.layout = syncLayoutWithCustom(next);
          return next;
        });
        setOpenK(null);
      },
    });
  };

  const visUnits = useMemo(() => {
    const t = cat.TYPES[state.type];
    if (!t) return [];
    return state.layout.units.filter((u) =>
      (state.layout.kids[u] || []).some((k) => t.fields.includes(k))
    );
  }, [cat.TYPES, state.type, state.layout]);

  const moveUnit = (unitId: string, d: -1 | 1) => {
    const t = cat.TYPES[state.type];
    if (!t) return;
    pushUndo3(true);
    setState((prev) => {
      const arr = prev.layout.units.slice();
      const i = arr.indexOf(unitId);
      let j = i + d;
      while (
        j >= 0 &&
        j < arr.length &&
        !(prev.layout.kids[arr[j]] || []).some((k) => t.fields.includes(k))
      ) {
        j += d;
      }
      if (j < 0 || j >= arr.length) return prev;
      [arr[i], arr[j]] = [arr[j], arr[i]];
      return { ...prev, layout: { ...prev.layout, units: arr } };
    });
  };

  const moveKid = (fieldId: string, d: -1 | 1) => {
    const t = cat.TYPES[state.type];
    if (!t) return;
    pushUndo3(true);
    setState((prev) => {
      const unitId = prev.layout.units.find((x) => (prev.layout.kids[x] || []).includes(fieldId));
      if (!unitId) return prev;
      const arr = (prev.layout.kids[unitId] || []).slice();
      const i = arr.indexOf(fieldId);
      let j = i + d;
      while (j >= 0 && j < arr.length && !t.fields.includes(arr[j])) j += d;
      if (j < 0 || j >= arr.length) return prev;
      [arr[i], arr[j]] = [arr[j], arr[i]];
      return {
        ...prev,
        layout: { ...prev.layout, kids: { ...prev.layout.kids, [unitId]: arr } },
      };
    });
  };

  const reorderUnit = (draggedId: string, targetId: string, pos: "before" | "after") => {
    if (draggedId === targetId) return;
    pushUndo3(true);
    setState((prev) => {
      const arr = prev.layout.units.filter((u) => u !== draggedId);
      const tIdx = arr.indexOf(targetId);
      if (tIdx < 0) return prev;
      const insertIdx = pos === "after" ? tIdx + 1 : tIdx;
      arr.splice(insertIdx, 0, draggedId);
      return { ...prev, layout: { ...prev.layout, units: arr } };
    });
  };

  const reorderKid = (unitId: string, draggedFieldId: string, targetFieldId: string, pos: "before" | "after") => {
    if (draggedFieldId === targetFieldId) return;
    pushUndo3(true);
    setState((prev) => {
      const orig = prev.layout.kids[unitId] || [];
      const arr = orig.filter((k) => k !== draggedFieldId);
      const tIdx = arr.indexOf(targetFieldId);
      if (tIdx < 0) return prev;
      const insertIdx = pos === "after" ? tIdx + 1 : tIdx;
      arr.splice(insertIdx, 0, draggedFieldId);
      return {
        ...prev,
        layout: { ...prev.layout, kids: { ...prev.layout.kids, [unitId]: arr } },
      };
    });
  };

  useEffect(() => {
    return () => {
      dragCleanupRef.current?.();
    };
  }, []);

  const beginUnitDrag = (unitId: string) => {
    dragCleanupRef.current?.();
    document.body.classList.add("is-dragging");
    let latestTarget: string | null = null;
    let latestPos: "before" | "after" = "before";
    setDragUnit({ id: unitId, targetId: null, pos: "before" });

    const blockTouchScroll = (ev: TouchEvent) => {
      if (ev.cancelable) ev.preventDefault();
    };

    const onMove = (ev: PointerEvent) => {
      if (ev.clientY < 85) window.scrollBy(0, -14);
      else if (ev.clientY > window.innerHeight - 95) window.scrollBy(0, 14);

      const els = Array.from(document.querySelectorAll<HTMLElement>("[data-drag-unit]"));
      let bestId: string | null = null;
      let bestPos: "before" | "after" = "before";
      let bestDist = Infinity;

      els.forEach((el) => {
        const uid = el.dataset.dragUnit;
        if (!uid || uid === unitId) return;
        const rect = el.getBoundingClientRect();
        const mid = rect.top + rect.height / 2;
        const dist = Math.abs(ev.clientY - mid);
        if (dist < bestDist) {
          bestDist = dist;
          bestId = uid;
          bestPos = ev.clientY < mid ? "before" : "after";
        }
      });

      latestTarget = bestId;
      latestPos = bestPos;
      setDragUnit({ id: unitId, targetId: bestId, pos: bestPos });
    };

    const cleanup = () => {
      document.body.classList.remove("is-dragging");
      window.removeEventListener("touchmove", blockTouchScroll);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", cleanup);
      dragCleanupRef.current = null;
      setDragUnit(null);
    };

    const onUp = () => {
      const tid = latestTarget;
      const pos = latestPos;
      cleanup();
      if (tid && tid !== unitId) {
        reorderUnit(unitId, tid, pos);
      }
    };

    dragCleanupRef.current = cleanup;
    window.addEventListener("touchmove", blockTouchScroll, { passive: false });
    window.addEventListener("pointermove", onMove, { passive: true });
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", cleanup);
  };

  const beginKidDrag = (unitId: string, fieldId: string) => {
    dragCleanupRef.current?.();
    document.body.classList.add("is-dragging");
    let latestTarget: string | null = null;
    let latestPos: "before" | "after" = "before";
    setDragKid({ unitId, fieldId, targetFieldId: null, pos: "before" });

    const blockTouchScroll = (ev: TouchEvent) => {
      if (ev.cancelable) ev.preventDefault();
    };

    const onMove = (ev: PointerEvent) => {
      if (ev.clientY < 85) window.scrollBy(0, -14);
      else if (ev.clientY > window.innerHeight - 95) window.scrollBy(0, 14);

      const els = Array.from(document.querySelectorAll<HTMLElement>("[data-drag-kid-id]"));
      let bestId: string | null = null;
      let bestPos: "before" | "after" = "before";
      let bestDist = Infinity;

      els.forEach((el) => {
        if (el.dataset.dragKidUnit !== unitId) return;
        const fid = el.dataset.dragKidId;
        if (!fid || fid === fieldId) return;
        const rect = el.getBoundingClientRect();
        const mid = rect.top + rect.height / 2;
        const dist = Math.abs(ev.clientY - mid);
        if (dist < bestDist) {
          bestDist = dist;
          bestId = fid;
          bestPos = ev.clientY < mid ? "before" : "after";
        }
      });

      latestTarget = bestId;
      latestPos = bestPos;
      setDragKid({ unitId, fieldId, targetFieldId: bestId, pos: bestPos });
    };

    const cleanup = () => {
      document.body.classList.remove("is-dragging");
      window.removeEventListener("touchmove", blockTouchScroll);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", cleanup);
      dragCleanupRef.current = null;
      setDragKid(null);
    };

    const onUp = () => {
      const tid = latestTarget;
      const pos = latestPos;
      cleanup();
      if (tid && tid !== fieldId) {
        reorderKid(unitId, fieldId, tid, pos);
      }
    };

    dragCleanupRef.current = cleanup;
    window.addEventListener("touchmove", blockTouchScroll, { passive: false });
    window.addEventListener("pointermove", onMove, { passive: true });
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", cleanup);
  };

  const startCardDragGesture = (
    e: React.PointerEvent,
    onActivate: () => void,
    ignoreSelector: string
  ) => {
    if (e.button !== 0 && e.pointerType === "mouse") return;
    const target = e.target as HTMLElement | null;
    if (target && target.closest(ignoreSelector)) return;
    e.stopPropagation();

    const startX = e.clientX;
    const startY = e.clientY;
    const pointerId = e.pointerId;
    const isTouch = e.pointerType !== "mouse";
    let done = false;
    let holdTimer: number | null = null;

    const suppressNextClick = () => {
      const blocker = (cev: MouseEvent) => {
        cev.preventDefault();
        cev.stopPropagation();
      };
      window.addEventListener("click", blocker, { capture: true, once: true });
      window.setTimeout(() => {
        window.removeEventListener("click", blocker, { capture: true });
      }, 450);
    };

    const detachTemp = () => {
      if (done) return;
      done = true;
      if (holdTimer !== null) {
        window.clearTimeout(holdTimer);
        holdTimer = null;
      }
      window.removeEventListener("pointermove", onTempMove);
      window.removeEventListener("pointerup", detachTemp);
      window.removeEventListener("pointercancel", detachTemp);
    };

    const activate = () => {
      if (done) return;
      detachTemp();
      suppressNextClick();
      if (isTouch && typeof navigator !== "undefined" && "vibrate" in navigator) {
        try {
          navigator.vibrate(14);
        } catch {
          // ignore
        }
      }
      onActivate();
    };

    const onTempMove = (ev: PointerEvent) => {
      if (ev.pointerId !== pointerId) return;
      const dist = Math.hypot(ev.clientX - startX, ev.clientY - startY);
      if (isTouch) {
        if (dist > 10) detachTemp();
      } else if (dist > 5) {
        activate();
      }
    };

    window.addEventListener("pointermove", onTempMove, { passive: true });
    window.addEventListener("pointerup", detachTemp);
    window.addEventListener("pointercancel", detachTemp);

    if (isTouch) {
      holdTimer = window.setTimeout(activate, 175);
    }
  };

  const startUnitDrag = (e: React.PointerEvent, unitId: string) => {
    if (e.button !== 0 && e.pointerType === "mouse") return;
    e.preventDefault();
    e.stopPropagation();
    beginUnitDrag(unitId);
  };

  const startKidDrag = (e: React.PointerEvent, unitId: string, fieldId: string) => {
    if (e.button !== 0 && e.pointerType === "mouse") return;
    e.preventDefault();
    e.stopPropagation();
    beginKidDrag(unitId, fieldId);
  };

  // ---------- Step 1 Type Actions ----------
  const handleCreateType = () => {
    const v = clean(ntName).trim();
    if (!v) {
      toast("اكتب اسم النوع أولاً");
      return;
    }
    if (Object.prototype.hasOwnProperty.call(cat.TYPES, v) || cat.TN.includes(v)) {
      toast("يوجد نوع بهذا الاسم، اختر اسماً آخر");
      return;
    }
    setState((prev) => ({
      ...prev,
      utypes: [...prev.utypes, { name: v }],
      type: v,
    }));
    setRelFor(v);
    setOpenK(null);
    setNtForm(false);
    setNtName("");
    toast("تم إنشاء النوع. اضغط «＋ إضافة خانة» لإضافة خاناته");
  };

  const handleRenameType = () => {
    const old = state.type;
    const v = clean(typeRenameVal).trim();
    if (v === old) return;
    if (!v) {
      toast("اسم النوع لا يمكن أن يكون فارغاً");
      setTypeRenameVal(old);
      return;
    }
    if (Object.prototype.hasOwnProperty.call(cat.TYPES, v) || cat.TN.includes(v)) {
      toast("يوجد نوع بهذا الاسم");
      setTypeRenameVal(old);
      return;
    }
    setState((prev) => {
      const nextUtypes = prev.utypes.map((u) => (u.name === old ? { name: v } : u));
      const nextCustom = { ...prev.custom };
      if (nextCustom[old]) {
        nextCustom[v] = nextCustom[old];
        delete nextCustom[old];
      }
      const nextRel = { ...prev.rel };
      if (nextRel[old]) {
        nextRel[v] = nextRel[old];
        delete nextRel[old];
      }
      const nextRnd = { ...prev.rnd };
      if (nextRnd[old]) {
        nextRnd[v] = nextRnd[old];
        delete nextRnd[old];
      }
      const pre = `cg:${old}:`;
      const nextUnits = prev.layout.units.map((x) =>
        x.startsWith(pre) ? `cg:${v}:${x.slice(pre.length)}` : x
      );
      const nextKids: Record<string, string[]> = {};
      Object.keys(prev.layout.kids).forEach((x) => {
        if (x.startsWith(pre)) nextKids[`cg:${v}:${x.slice(pre.length)}`] = prev.layout.kids[x];
        else nextKids[x] = prev.layout.kids[x];
      });
      return {
        ...prev,
        type: v,
        utypes: nextUtypes,
        custom: nextCustom,
        rel: nextRel,
        rnd: nextRnd,
        layout: { units: nextUnits, kids: nextKids },
      };
    });
    setRelFor(v);
    toast("تم تغيير اسم النوع");
  };

  const handleDeleteUserType = (typeName: string) => {
    setConfirmDlg({
      title: "حذف النوع",
      message: `حذف نوع «${typeName}» وجميع خاناته نهائياً؟`,
      confirmLabel: "حذف النوع",
      danger: true,
      onConfirm: () => {
        setState((prev) => {
          const nextCustom = { ...prev.custom };
          delete nextCustom[typeName];
          const nextRel = { ...prev.rel };
          delete nextRel[typeName];
          const nextRnd = { ...prev.rnd };
          delete nextRnd[typeName];
          const next: AppState = {
            ...prev,
            utypes: prev.utypes.filter((u) => u.name !== typeName),
            custom: nextCustom,
            rel: nextRel,
            rnd: nextRnd,
            type: BUILTIN_TYPE_NAMES[0],
          };
          next.layout = syncLayoutWithCustom(next);
          return next;
        });
        setRelFor(null);
        setOpenK(null);
        toast("تم حذف النوع");
      },
    });
  };

  // ---------- Copy / Save / Share / Export ----------
  const copyText = useCallback(
    (txt: string, msg?: string) => {
      const done = () => toast(msg || "تم النسخ");
      const fb = () => {
        const a = document.createElement("textarea");
        a.value = txt;
        a.style.cssText = "position:fixed;opacity:0";
        document.body.appendChild(a);
        a.select();
        try {
          document.execCommand("copy");
        } catch {
          // ignore
        }
        a.remove();
        done();
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(done).catch(fb);
      } else {
        fb();
      }
    },
    [toast]
  );

  const webDownload = (blob: Blob, filename: string) => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  };

  const saveBlob = async (blob: Blob, filename: string, mime: string, mode?: "save" | "share") => {
    lastSavedBlobRef.current = { blob, filename, mime };
    const winAny = window as unknown as {
      Capacitor?: {
        isNativePlatform?: () => boolean;
        Plugins?: {
          Filesystem?: { writeFile: (o: unknown) => Promise<{ uri: string }> };
          Share?: { share: (o: unknown) => Promise<void> };
          SaveFile?: { save: (o: unknown) => Promise<{ uri?: string }>; open: (o: unknown) => Promise<void> };
        };
      };
    };
    const C = winAny.Capacitor;
    const native = Boolean(C && C.isNativePlatform && C.isNativePlatform() && C.Plugins);
    const isCancelErr = (e: unknown) =>
      Boolean(e && /cancel|abort|dismiss/i.test(String((e as Error).message || (e as Error).name || "")));

    const toB64 = (b: Blob) =>
      new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(String(r.result).split(",")[1]);
        r.onerror = reject;
        r.readAsDataURL(b);
      });

    if (mode === "share") {
      try {
        if (native && C?.Plugins?.Filesystem && C?.Plugins?.Share) {
          const w = await C.Plugins.Filesystem.writeFile({
            path: filename,
            data: await toB64(blob),
            directory: "CACHE",
          });
          await C.Plugins.Share.share({ title: filename, files: [w.uri], dialogTitle: "مشاركة الملف" });
          return;
        }
        const f = new File([blob], filename, { type: mime });
        if (navigator.canShare && navigator.canShare({ files: [f] })) {
          await navigator.share({ files: [f], title: filename });
          return;
        }
      } catch (e) {
        if (isCancelErr(e)) return;
      }
      toast("المشاركة غير متاحة هنا، سيتم الحفظ بدلاً منها", 3500);
    }

    if (native && C?.Plugins) {
      try {
        const b64 = await toB64(blob);
        if (C.Plugins.SaveFile) {
          try {
            const out = await C.Plugins.SaveFile.save({
              filename,
              mime: mime || "application/octet-stream",
              data: b64,
            });
            setOpenBar({
              txt: "✓ تم حفظ الملف في «التنزيلات / حاسبة الشبابيك»",
              web: false,
              uri: out?.uri,
              mime,
            });
            toast("✓ تم حفظ الملف في «التنزيلات / حاسبة الشبابيك»");
            return;
          } catch {
            // fallback to Share / Filesystem
          }
        }
        if (C.Plugins.Filesystem && C.Plugins.Share) {
          const w = await C.Plugins.Filesystem.writeFile({ path: filename, data: b64, directory: "CACHE" });
          await C.Plugins.Share.share({ title: filename, files: [w.uri], dialogTitle: "احفظ الملف أو أرسله" });
          return;
        }
      } catch (e) {
        if (isCancelErr(e)) return;
      }
      toast("تعذّر حفظ الملف على الجهاز", 3500);
      setAltOpen(true);
      return;
    }

    webDownload(blob, filename);
    setOpenBar({
      txt: "✓ تم تنزيل الملف — افتحه من إشعار المتصفح أو «التنزيلات»",
      web: true,
      mime,
    });
  };

  // File name: "<project name>_<date>", or "<window type>_<date>" when no project name is set
  const fnBase = () => {
    const clean = (s: string) =>
      String(s || "")
        .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 60);
    const label = clean(pjName) || clean(state.type) || "نتائج";
    return `${label}_${dstr()}`;
  };

  const exportNotes = () => {
    const n: string[] = [];
    if (res.adj) {
      n.push(`التعديلات: ${res.adj}${res.cascade ? " (تنتقل للخانات المشتقة)" : " (على الخانة نفسها فقط)"}`);
    }
    if (res.rndOn) {
      n.push(state.raw ? "الأرقام الأصلية بدون تقريب" : "الأرقام مقرَّبة حسب الإعدادات");
    }
    const nmod = Object.keys((state.rel && state.rel[state.type]) || {}).length;
    if (nmod) n.push(`علاقات معدّلة: ${nmod}`);
    return n;
  };

  const buildSettingsExportInfo = (ds: string): XlsxSettingsConfig => {
    const { raw: _r, rows: _rows, ...restData } = state;
    void _r;
    void _rows;
    const activeFields = (cat.TYPES[state.type]?.fields || []).slice();
    const fields = cat.ORDER.filter((k) => activeFields.includes(k)).map((k) => ({
      label: cat.LABEL[k] || k,
      rel: k === "E" ? relParen(cat, state, k) : relParen(cat, state, k).replace(/\u2066|\u2069/g, ""),
      inc: state.inc[k] ? `＋ ${state.inc[k]}` : "",
      dec: state.dec[k] ? `－ ${state.dec[k]}` : "",
      on: isFieldSelected(k, state.sel),
    }));
    return {
      info: [
        ["النوع", state.type],
        ["الوحدة", UFULL[state.unit]],
        ["المشروع", pjName || ""],
        ["التاريخ", ds],
        ["عدد الشبابيك", String(res.data.length)],
      ],
      fields,
      json: JSON.stringify({ v: 1, project: pjName || "", data: restData }),
    };
  };

  const handleExportXlsx = async (mode: "save" | "share") => {
    if (!res.table.length) {
      toast("لا توجد نتائج للتصدير");
      return;
    }
    try {
      const ds = dstr();
      let sub = `${pjName ? `المشروع: ${pjName}   |   ` : ""}نوع المنتج: ${res.type}   |   الوحدة: ${UFULL[state.unit]}${res.area ? `   |   المساحة: ${UAREA[state.unit]}` : ""}   |   التاريخ: \u2066${ds}\u2069   |   عدد الشبابيك: ${res.data.length}${res.hasQ ? ` (${res.qtotal} قطعة)` : ""}`;
      const notes = exportNotes();
      sub += notes.length ? `   |   ${notes.join("   |   ")}` : "   |   بدون تعديلات";
      const blob = buildXlsx({
        dup: res.dupRows,
        dupOther: res.dupOther,
        areaTotals: res.areaTotals,
        areaUnitLabel: UAS[state.unit],
        title: pjName ? `${pjName} — نتائج قياسات الشبابيك والواجهات` : "نتائج قياسات الشبابيك والواجهات",
        sub,
        struct: res.xstruct,
        table: res.xtable,
        settings: buildSettingsExportInfo(ds),
      });
      await saveBlob(
        blob,
        `${fnBase()}.xlsx`,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        mode
      );
    } catch {
      toast("تعذّر إنشاء ملف إكسل");
    }
  };

  // ---------- Excel Import (.xlsx) ----------
  const openXlsxImport = () => {
    if (xlFileRef.current) {
      xlFileRef.current.value = "";
      xlFileRef.current.click();
    }
  };

  const handleXlsxFileSelected = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    if (f.size > 8e6) {
      toast("الملف كبير جداً (الحد الأقصى 8 ميجابايت)", 4000);
      return;
    }
    try {
      const buf = await f.arrayBuffer();
      const sheets = await readXlsxSheets(buf);
      const parsed = parseImportedWorkbook(sheets);
      if (!parsed.rows && !parsed.settings) {
        toast("لم أجد طولاً وعرضاً ولا إعدادات في هذا الملف", 4500);
        return;
      }
      setXlImportState({
        fileName: f.name,
        parsed,
        useRows: Boolean(parsed.rows),
        useSettings: Boolean(parsed.settings) && !parsed.rows,
        unit: parsed.unit || state.unit,
        mode: "add",
        err: "",
      });
    } catch {
      toast("تعذّرت قراءة ملف Excel. تأكد أنه بصيغة .xlsx سليمة", 4500);
    }
  };

  const executeXlsxImport = () => {
    if (!xlImportState) return;
    const { parsed, useRows, useSettings, unit, mode } = xlImportState;
    if (!useRows && !useSettings) {
      setXlImportState((prev) => (prev ? { ...prev, err: "اختر شيئاً واحداً على الأقل للاستيراد" } : null));
      return;
    }

    const doApply = () => {
      const settingsObj = parsed.settings;
      const settingsUnit =
        settingsObj && (settingsObj.unit === "m" || settingsObj.unit === "cm" || settingsObj.unit === "mm")
          ? (settingsObj.unit as UnitCode)
          : state.unit;
      const targetUnit: UnitCode = useSettings ? settingsUnit : state.unit;
      const k = (UF[unit] || 1) / (UF[targetUnit] || 1);
      const cv = (v: number) => fmt6(v * k);

      let nextRows: WindowRow[] | null = null;
      if (useRows && parsed.rows) {
        const importedRows: WindowRow[] = parsed.rows.map((r) => {
          const p: PendingMode =
            r.l === null && r.w === null ? "B" : r.l === null ? "L" : r.w === null ? "W" : "";
          return {
            n: r.n || "",
            L: r.l === null ? "" : cv(r.l),
            W: r.w === null ? "" : cv(r.w),
            q: r.q || "",
            p,
          };
        });
        if (mode === "rep") {
          nextRows = importedRows.slice(0, 300);
        } else {
          const cur = state.rows.filter((r) => r.L || r.W || r.n || r.p);
          nextRows = cur.concat(importedRows).slice(0, 300);
        }
      }

      if (useSettings && settingsObj) {
        const existingRowsConverted =
          !useRows && targetUnit !== state.unit
            ? state.rows.map((r) => {
                const conv = (v: string) => {
                  const n = num(v);
                  return n === null || isNaN(n) ? v : fmt6(n * (UF[state.unit] / UF[targetUnit]));
                };
                return { ...r, L: conv(r.L), W: conv(r.W) };
              })
            : state.rows;
        const merged = {
          ...state,
          ...settingsObj,
          rows: nextRows || existingRowsConverted,
          step: useRows && nextRows && nextRows.length ? 2 : 1,
        };
        const hydrated = hydrateState(merged);
        hydrated.layout = syncLayoutWithCustom(hydrated);
        hydrated.rows = withTrailingEmpty(hydrated.rows);
        setDeletedUndo(null);
        setNegHighlightedRows(new WeakSet());
        setState(hydrated);
        if (parsed.projectName) setPjName(parsed.projectName);
        setXlImportState(null);
        toast(
          useRows && parsed.rows
            ? `تم استيراد ${parsed.rows.length} شباك والإعدادات بنجاح ✓`
            : "تم استيراد الإعدادات والعلاقات بنجاح ✓"
        );
        return;
      }

      if (nextRows) {
        setDeletedUndo(null);
        setNegHighlightedRows(new WeakSet());
        setState((prev) => ({
          ...prev,
          rows: withTrailingEmpty(nextRows!),
          step: 2,
        }));
        if (parsed.projectName && !pjName) setPjName(parsed.projectName);
        setXlImportState(null);
        toast(`تم استيراد ${parsed.rows?.length || 0} شباك بنجاح ✓`);
      }
    };

    if (useRows && mode === "rep" && state.rows.some((r) => r.L || r.W || r.p || r.n)) {
      setConfirmDlg({
        title: "تأكيد استبدال الشبابيك",
        message: "سيتم حذف شبابيكك الحالية واستبدالها بالمستورَدة من ملف Excel. هل أنت متأكد؟",
        confirmLabel: "نعم، استبدلها",
        danger: true,
        onConfirm: doApply,
      });
      return;
    }
    doApply();
  };

  const handleExportPdf = async (mode: "save" | "share") => {
    if (!res.table.length) {
      toast("لا توجد نتائج للتصدير");
      return;
    }
    toast("جارٍ تجهيز ملف PDF نصي خفيف…", 1800);
    try {
      const blob = await buildPdf({
        project: pjName,
        type: res.type,
        unit: UFULL[state.unit],
        area: res.area ? UAREA[state.unit] : "",
        areaUnitLabel: UAS[state.unit],
        date: dstr(),
        count: `${res.data.length}${res.hasQ ? ` (${res.qtotal} قطعة)` : ""}`,
        notes: exportNotes(),
        struct: res.xstruct,
        table: res.xtable,
        dup: res.dupRows,
        dupOther: res.dupOther,
        areaTotals: res.areaTotals,
      });
      await saveBlob(blob, `${fnBase()}.pdf`, "application/pdf", mode);
    } catch {
      toast("تعذّر إنشاء ملف PDF");
    }
  };

  const getTabLines = () => {
    const n = res.xc1.length;
    const pad = (r: string[]) => [...r, ...Array(Math.max(0, n - r.length)).fill("")];
    const a: string[][] = [res.xc1, res.xc2, ...res.xtable];
    if (res.areaTotals.length) {
      a.push(pad([]), pad(["ملخص إجمالي المساحات (لطلب المواد)"]));
      a.push(pad(["البند", `المساحة (${UAS[state.unit]})`, "المساحة (م²)"]));
      res.areaTotals.forEach((item) =>
        a.push(pad([item.label, item.totalInUnit.toFixed(3), item.totalSqM.toFixed(3)]))
      );
    }
    if (res.dups.length) {
      a.push(pad([]), pad(["ملخص الشبابيك المتشابهة"]), pad(["", "طول", "عرض", "العدد", "الأرقام"]));
      res.dupRows.forEach((r) => a.push(pad(["", r[0], r[1], r[2], r[3]])));
    }
    return a;
  };

  const getPlainTextResults = () => {
    if (!res.data.length) return "";
    let txt = `${pjName ? `${pjName} — ` : ""}نتائج: ${res.type}\n`;
    res.data.forEach((d) => {
      txt += `\nشباك ${d.n}${d.nm ? ` — ${d.nm}` : ""}  (${d.L === null ? "⏳" : fmt(d.L, state.unit)} × ${d.W === null ? "⏳" : fmt(d.W, state.unit)})${d.q > 1 ? `  العدد: ${d.q}` : ""}\n`;
      res.units.forEach((u) => {
        const vals = u.ks
          .map((k) => {
            const isA = fdim(cat, k) === 2;
            const sfx = isA ? ` (${UAS[state.unit]})` : "";
            const nm = u.group ? `${cat.SUB[k]}${sfx} ` : "";
            return `${nm}${fmt(shownV(state, k, d.res[k], isA), state.unit)}`;
          })
          .join(" ، ");
        txt += `• ${u.group ? u.title : cat.LABEL[u.ks[0]]}: ${vals}\n`;
      });
    });
    if (res.areaTotals.length) {
      txt += "\nملخص إجمالي المساحات:\n";
      res.areaTotals.forEach((item) => {
        txt += `• ${item.label}: ${item.totalSqM.toFixed(3)} م² (${item.totalInUnit.toFixed(3)} ${UAS[state.unit]})\n`;
      });
    }
    if (res.dups.length) {
      txt += "\nملخص الشبابيك المتشابهة:\n";
      res.dups.forEach((g) => {
        txt += `${fmt(g.L, state.unit)} × ${fmt(g.W, state.unit)}  =  ${g.cnt} شبابيك  (الأرقام: ${g.nums.join("، ")})\n`;
      });
    }
    return txt;
  };

  // Backup & Restore JSON
  const handleBackupSave = () => {
    persistState(state);
    const { raw: _ignored, ...cleanData } = state;
    void _ignored;
    const payload = {
      app: "winCalc",
      v: 2,
      date: new Date().toISOString(),
      project: pjName,
      projects,
      data: cleanData,
    };
    const blob = new Blob([JSON.stringify(payload, null, 1)], { type: "application/json" });
    saveBlob(blob, `نسخة_احتياطية_حاسبة_القياسات_${dstr()}.json`, "application/json", "save");
  };

  const handleBackupFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    if (f.size > 3e6) {
      toast("الملف كبير جداً", 4000);
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      let o: {
        app?: string;
        project?: string;
        projects?: unknown[];
        data?: Record<string, unknown>;
      } | null = null;
      try {
        o = JSON.parse(String(reader.result));
      } catch {
        o = null;
      }
      if (!o || o.app !== "winCalc" || !o.data || typeof o.data !== "object" || Array.isArray(o.data)) {
        toast("هذا الملف ليس نسخة احتياطية من التطبيق", 4000);
        return;
      }
      const dataObj = o.data;
      const importedProjects = Array.isArray(o.projects) ? o.projects : null;
      const importedProjectName = typeof o.project === "string" ? clean(o.project).trim() : "";
      setConfirmDlg({
        title: "استعادة النسخة الاحتياطية",
        message: "سيتم استبدال أنواعك وعلاقاتك ومشاريعك الحالية بالنسخة المحفوظة. هل تريد المتابعة؟",
        confirmLabel: "استعادة",
        onConfirm: () => {
          try {
            const restored = hydrateState(dataObj);
            restored.layout = syncLayoutWithCustom(restored);
            restored.rows = withTrailingEmpty(restored.rows);
            setDeletedUndo(null);
            setNegHighlightedRows(new WeakSet());
            setState(restored);
            setPjName(importedProjectName);
            if (importedProjects) {
              const validProjects = importedProjects.filter(
                (p): p is SavedProject =>
                  Boolean(
                    p &&
                      typeof p === "object" &&
                      (p as SavedProject).id &&
                      (p as SavedProject).name &&
                      (p as SavedProject).data
                  )
              );
              setProjects(validProjects);
            }
            toast("تمت استعادة النسخة الاحتياطية بنجاح ✓", 3000);
          } catch {
            toast("تعذّرت الاستعادة");
          }
        },
      });
    };
    reader.readAsText(f);
  };

  const openSaveDialog = (rowsOverride?: unknown) => {
    const validRows = Array.isArray(rowsOverride) ? (rowsOverride as WindowRow[]) : undefined;
    setOpenBar(null);
    setFallBarVisible(false);
    if (!validRows && askPend((latestRows) => openSaveDialog(latestRows))) return;
    const hasRows = validRows ? validRows.some((r) => Boolean(r.L || r.W || r.p)) : res.table.length > 0;
    if (!hasRows) {
      toast("لا توجد نتائج للحفظ");
      return;
    }
    setSaveOpen(true);
  };

  // Search normalized query for Step 4
  const normalizedSearch = searchQuery
    .trim()
    .toLowerCase()
    .replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));

  // Step 3 preview sample window
  const sampleWin = useMemo(
    () => res.data.find((x) => x.L !== null && x.W !== null) || res.data[0],
    [res.data]
  );

  // Render relation editor for a field key k
  const renderRelEditor = (k: string, ctx: "p" | "f") => {
    const e = effRel(state, k);
    const cu = isCid(k);
    const custFields = (state.custom && state.custom[state.type]) || [];
    const cc = cu ? custFields.find((x) => x.id === k) : null;
    const gl = [
      ...new Set([
        ...Array.from(BUILTIN_GROUPS),
        ...custFields.map((x) => x.group).filter(Boolean),
      ]),
    ];
    const dimK = fdim(cat, k);
    const dimS = fdim(cat, e.s);
    const dimS2 = e.s2 ? fdim(cat, e.s2) : 1;
    const scaleK = Math.pow(UF[state.unit], dimK);
    const scaleS = Math.pow(UF[state.unit], dimS);
    const scaleS2 = Math.pow(UF[state.unit], dimS2);
    const prevTxt = (val: number | "" | undefined) => {
      if (val === "" || val === undefined) return "—";
      const s = shownV(state, k, val, dimK === 2);
      const a = fmt(s, state.unit);
      const orig = fmt(val, state.unit);
      return a + (a !== orig ? ` (الأصلي ${orig})` : "");
    };

    if (k === "E") {
      return (
        <div className="redit" onPointerDown={(ev) => ev.stopPropagation()}>
          <div className="rrow">
            <div className="rfld">
              <span>القيمة الثابتة</span>
              <ScaledDecimalInput
                storedCm={e.a}
                scaleFactor={UF[state.unit]}
                unsigned
                onCommitCm={(cmVal) => {
                  pushUndo3();
                  setRelPatch("E", { a: cmVal });
                }}
              />
            </div>
          </div>
          <div className="rprev">= {relParen(cat, state, "E")}</div>
          <button
            type="button"
            className="btn sm w"
            onClick={() => {
              pushUndo3(true);
              setState((prev) => {
                const o = { ...(prev.rel[prev.type] || {}) };
                delete o.E;
                return { ...prev, rel: { ...prev.rel, [prev.type]: o } };
              });
            }}
          >
            ↩︎ إرجاع الأصلية
          </button>
        </div>
      );
    }

    const rc = (state.rnd && state.rnd[state.type] && state.rnd[state.type][k]) || { s: 0, d: "n" as RoundDir };
    const rOn = Number(rc.s) > 0;
    const opts = srcOptions(k, state.type);

    return (
      <div className="redit" onPointerDown={(ev) => ev.stopPropagation()}>
        {cu && cc && (
          <div className="rrow">
            <label>
              <span>اسم الخانة</span>
              <input
                value={cc.name}
                style={{ textAlign: "right", direction: "rtl" }}
                placeholder="مثال: طول"
                onChange={(ev) => {
                  const v = clean(ev.target.value);
                  setState((prev) => ({
                    ...prev,
                    custom: {
                      ...prev.custom,
                      [prev.type]: (prev.custom[prev.type] || []).map((x) =>
                        x.id === k ? { ...x, name: v } : x
                      ),
                    },
                  }));
                }}
              />
            </label>
            <label>
              <span>المجموعة (اختياري)</span>
              <input
                list={`gl_${k}`}
                value={cc.group || ""}
                style={{ textAlign: "right", direction: "rtl" }}
                placeholder="مثال: الحصيرة"
                onChange={(ev) => {
                  const v = clean(ev.target.value).trim();
                  setState((prev) => {
                    const next: AppState = {
                      ...prev,
                      custom: {
                        ...prev.custom,
                        [prev.type]: (prev.custom[prev.type] || []).map((x) =>
                          x.id === k ? { ...x, group: v } : x
                        ),
                      },
                    };
                    next.layout = syncLayoutWithCustom(next);
                    return next;
                  });
                }}
              />
              <datalist id={`gl_${k}`}>
                {gl.map((g) => (
                  <option key={g} value={g} />
                ))}
              </datalist>
            </label>
          </div>
        )}

        {ctx === "f" && (
          <>
            <div className="tpl-lab">قوالب سريعة (اضغط لتطبيقها ثم عدّل الرقم):</div>
            <div className="tpls">
              {TPL.map(([label, tplCfg], i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => {
                    pushUndo3(true);
                    setRelPatch(k, tplCfg);
                    toast("طُبّق القالب. عدّل الرقم المطلوب");
                  }}
                >
                  {label}
                </button>
              ))}
            </div>
          </>
        )}

        <div className="rrow">
          <label>
            <span>تُحسب من</span>
            <select
              value={e.s}
              onChange={(ev) => {
                pushUndo3(true);
                setRelPatch(k, { s: ev.target.value });
              }}
            >
              {opts.map((s) => (
                <option key={s} value={s}>
                  {s === "LEN" ? "الطول" : s === "WID" ? "العرض" : cat.LABEL[s]}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="rrow">
          <label>
            <span>العملية مع خانة ثانية</span>
            <select
              value={e.op || "+"}
              disabled={!e.s2}
              onChange={(ev) => {
                pushUndo3(true);
                setRelPatch(k, { op: ev.target.value as "+" | "-" | "*" | "/" });
              }}
            >
              <option value="+">＋ جمع</option>
              <option value="-">－ طرح (الفرق)</option>
              <option value="*">× ضرب (المساحة)</option>
              <option value="/">÷ قسمة</option>
            </select>
          </label>
          <label>
            <span>الخانة الثانية (اختياري)</span>
            <select
              value={e.s2 || ""}
              onChange={(ev) => {
                pushUndo3(true);
                setRelPatch(k, { s2: ev.target.value });
              }}
            >
              <option value="">— بدون —</option>
              {opts.map((x) => (
                <option key={x} value={x}>
                  {x === "LEN" ? "الطول" : x === "WID" ? "العرض" : cat.LABEL[x]}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="rrow">
          <div className="rfld">
            <span>{e.dv ? "÷ اقسم على" : "× اضرب في"}</span>
            <span className="sgn">
              <button
                type="button"
                className="op"
                aria-label="تبديل الضرب والقسمة"
                onClick={(ev) => {
                  ev.preventDefault();
                  ev.stopPropagation();
                  pushUndo3(true);
                  setRelPatch(k, { dv: !e.dv });
                }}
              >
                {e.dv ? "÷" : "×"}
              </button>
              <input
                inputMode="decimal"
                value={e.m}
                onChange={(ev) => {
                  pushUndo3();
                  setRelPatch(k, { m: ev.target.value.replace(/[eE]/g, "") });
                }}
              />
            </span>
            {e.dv && num(e.m) === 0 && (
              <small style={{ display: "block", marginTop: 4, color: "var(--bad, #dc2626)", fontWeight: 700 }}>
                ⚠️ لا يمكن القسمة على صفر — أدخل رقماً أكبر من 0
              </small>
            )}
          </div>
          <div className="rfld">
            <span>ثم أضف أو انقص</span>
            <span className="sgn">
              <button
                type="button"
                className={e.neg ? "neg" : "pos"}
                aria-label="تبديل الزائد والناقص"
                onClick={(ev) => {
                  ev.preventDefault();
                  ev.stopPropagation();
                  pushUndo3(true);
                  setRelPatch(k, { neg: !e.neg });
                }}
              >
                {e.neg ? "－" : "＋"}
              </button>
              <ScaledDecimalInput
                storedCm={e.a}
                scaleFactor={scaleK}
                unsigned
                onCommitCm={(cmVal) => {
                  pushUndo3();
                  setRelPatch(k, { a: cmVal });
                }}
              />
            </span>
          </div>
        </div>

        {e.s2 && (
          <div className="rrow">
            <div className="rfld">
              <span>تعديل «{e.s === "LEN" ? "الطول" : e.s === "WID" ? "العرض" : cat.LABEL[e.s]}» قبل العملية</span>
              <span className="sgn">
                <button
                  type="button"
                  className={e.n1 ? "neg" : "pos"}
                  aria-label="تبديل إشارة تعديل الخانة الأولى"
                  onClick={(ev) => {
                    ev.preventDefault();
                    ev.stopPropagation();
                    pushUndo3(true);
                    setRelPatch(k, { n1: !e.n1 });
                  }}
                >
                  {e.n1 ? "－" : "＋"}
                </button>
                <ScaledDecimalInput
                  storedCm={e.a1 || "0"}
                  scaleFactor={scaleS}
                  placeholder="0"
                  zeroAsEmpty
                  unsigned
                  onCommitCm={(cmVal) => {
                    pushUndo3();
                    setRelPatch(k, { a1: cmVal });
                  }}
                />
              </span>
            </div>
            <div className="rfld">
              <span>تعديل «{e.s2 === "LEN" ? "الطول" : e.s2 === "WID" ? "العرض" : cat.LABEL[e.s2]}» قبل العملية</span>
              <span className="sgn">
                <button
                  type="button"
                  className={e.n2 ? "neg" : "pos"}
                  aria-label="تبديل إشارة تعديل الخانة الثانية"
                  onClick={(ev) => {
                    ev.preventDefault();
                    ev.stopPropagation();
                    pushUndo3(true);
                    setRelPatch(k, { n2: !e.n2 });
                  }}
                >
                  {e.n2 ? "－" : "＋"}
                </button>
                <ScaledDecimalInput
                  storedCm={e.a2 || "0"}
                  scaleFactor={scaleS2}
                  placeholder="0"
                  zeroAsEmpty
                  unsigned
                  onCommitCm={(cmVal) => {
                    pushUndo3();
                    setRelPatch(k, { a2: cmVal });
                  }}
                />
              </span>
            </div>
          </div>
        )}

        <div className="rn-box">
          <div className="rn-lab">
            التقريب (اختياري)
            {!rOn && Number(state.gr.s) > 0 ? ` — «بدون» = يتبع التقريب العام (${state.gr.s})` : ""}
          </div>
          <div className="rn-seg">
            {[
              [0, "بدون"],
              [0.1, "0.1"],
              [0.5, "0.5"],
              [1, "1"],
            ].map(([v, t]) => (
              <button
                key={String(v)}
                type="button"
                className={(Number(rc.s) || 0) === v ? "on" : ""}
                onClick={() => {
                  pushUndo3(true);
                  setRndPatch(k, { s: Number(v) });
                }}
              >
                {t}
              </button>
            ))}
          </div>
          <div className="rn-lab">الاتجاه</div>
          <div className={`rn-seg dir${rOn ? "" : " off"}`}>
            {[
              ["n", "لأقرب"],
              ["d", "لأسفل"],
              ["u", "لأعلى"],
            ].map(([v, t]) => (
              <button
                key={v}
                type="button"
                disabled={!rOn}
                className={(rc.d || "n") === v ? "on" : ""}
                onClick={() => {
                  pushUndo3(true);
                  setRndPatch(k, { d: v as RoundDir });
                }}
              >
                {t}
              </button>
            ))}
          </div>
        </div>

        <div className="rprev">
          = {relParen(cat, state, k)}
          {sampleWin ? `   ←   ${prevTxt(sampleWin.res[k])}  (لأول شباك)` : ""}
        </div>
        <p className="rtip">
          ملاحظة: الأرقام داخل العلاقات (مثل − 9) بنفس وحدة القياس المختارة، وتتحول تلقائياً إذا غيّرت الوحدة.
        </p>

        {cu ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {cc && cc.group && (
              <button type="button" className="btn sm g" onClick={() => addCustomField(cc.group)}>
                ＋ خانة أخرى في مجموعة «{cc.group}»
              </button>
            )}
            <button type="button" className="btn sm r" onClick={() => delCustomField(k)}>
              🗑 حذف هذه الخانة
            </button>
          </div>
        ) : (
          <button
            type="button"
            className="btn sm w"
            onClick={() => {
              pushUndo3(true);
              setState((prev) => {
                const o = { ...(prev.rel[prev.type] || {}) };
                delete o[k];
                return { ...prev, rel: { ...prev.rel, [prev.type]: o } };
              });
            }}
          >
            ↩︎ إرجاع العلاقة الأصلية
          </button>
        )}
      </div>
    );
  };

  // Render field row in Step 3
  const renderFieldRow = (k: string, isHeadless: boolean, parentUnitId = "") => {
    const on = isFieldSelected(k, state.sel);
    const dep = Object.keys(BASE_REL)
      .concat(((state.custom && state.custom[state.type]) || []).map((c) => c.id))
      .filter((x) => {
        const rel = cat.REL[x];
        return rel && (rel[0] === k || rel[3] === k) && cat.TYPES[state.type]?.fields.includes(x);
      })
      .map((x) => cat.LABEL[x]);
    const val = sampleWin ? fmt(shownV(state, k, sampleWin.res[k], fdim(cat, k) === 2), state.unit) : "—";
    const rOwn = rndOwn(state, k);
    const isMod = !isCid(k) && Boolean((state.rel && state.rel[state.type] && state.rel[state.type][k]));
    const dirArrow = rOwn ? ({ n: "", d: " ↓", u: " ↑" }[rOwn.d] || "") : "";
    const isKidDragging = dragKid?.unitId === parentUnitId && dragKid?.fieldId === k;
    const isKidDropBefore =
      dragKid?.unitId === parentUnitId && dragKid?.targetFieldId === k && dragKid?.pos === "before";
    const isKidDropAfter =
      dragKid?.unitId === parentUnitId && dragKid?.targetFieldId === k && dragKid?.pos === "after";

    return (
      <div
        key={k}
        data-drag-kid-unit={!isHeadless ? parentUnitId : undefined}
        data-drag-kid-id={!isHeadless ? k : undefined}
        className={`fld${on ? "" : " off"}${isKidDragging ? " dragging" : ""}${isKidDropBefore ? " drop-before" : ""}${isKidDropAfter ? " drop-after" : ""}`}
      >
        {!isHeadless && (
          <div
            className="fl1"
            onPointerDown={(e) =>
              startCardDragGesture(
                e,
                () => beginKidDrag(parentUnitId, k),
                "button, input, select, textarea, a, .fl2"
              )
            }
          >
            <button
              type="button"
              className="drag-h"
              title="اسحب لإعادة الترتيب"
              aria-label="سحب لإعادة الترتيب"
              onPointerDown={(e) => startKidDrag(e, parentUnitId, k)}
            >
              <svg width="14" height="18" viewBox="0 0 14 18" aria-hidden="true">
                <circle cx="4" cy="3.5" r="1.7" fill="currentColor" />
                <circle cx="10" cy="3.5" r="1.7" fill="currentColor" />
                <circle cx="4" cy="9" r="1.7" fill="currentColor" />
                <circle cx="10" cy="9" r="1.7" fill="currentColor" />
                <circle cx="4" cy="14.5" r="1.7" fill="currentColor" />
                <circle cx="10" cy="14.5" r="1.7" fill="currentColor" />
              </svg>
            </button>
            <label>
              <input
                type="checkbox"
                checked={on}
                onChange={(e) => {
                  pushUndo3(true);
                  const checked = e.target.checked;
                  setState((prev) => ({ ...prev, sel: { ...prev.sel, [k]: checked } }));
                }}
              />
              <span>{cat.SUB[k]}</span>
              {fdim(cat, k) === 2 && <b className="mod">📐 مساحة</b>}
              {isCid(k) && <b className="mod">جديدة</b>}
            </label>
            <div className="o-dock">
              {on && val && <span className="chipv">{val}</span>}
              <span className="mvb">
                <button type="button" title="لأعلى" aria-label="لأعلى" onClick={() => moveKid(k, -1)}>
                  <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                    <path d="M5 1.2L9 6.8H1L5 1.2Z" fill="currentColor" />
                  </svg>
                </button>
                <button type="button" title="لأسفل" aria-label="لأسفل" onClick={() => moveKid(k, 1)}>
                  <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                    <path d="M5 6.8L1 1.2H9L5 6.8Z" fill="currentColor" />
                  </svg>
                </button>
              </span>
            </div>
          </div>
        )}
        <div className="fl2">
          <span className="rel2 advonly">
            <span>= {relParen(cat, state, k)}</span>{" "}
            {isMod && <b className="mod">معدّلة</b>}{" "}
            <button
              type="button"
              className="edb"
              onClick={() => setOpenK((prev) => (prev === k ? null : k))}
            >
              ✎ {openK === k ? "إغلاق" : "تغيير"}
            </button>{" "}
            {k !== "E" && (
              <button
                type="button"
                className={`edb rn-b${rOwn ? " on" : ""}`}
                onClick={() => setOpenK((prev) => (prev === k ? null : k))}
              >
                ≈ {rOwn ? `تقريب ${rOwn.s}${dirArrow}` : "تقريب"}
              </button>
            )}
            {state.cascade && dep.length > 0 && (
              <>
                <br />
                يؤثر على: {dep.join("، ")}
              </>
            )}
          </span>
          <span className="sgnw pl">
            <b className="sg">＋</b>
            <input
              inputMode="decimal"
              placeholder="0"
              disabled={!on}
              value={state.inc[k] || ""}
              onChange={(e) => {
                const v = e.target.value.replace(/[eE]/g, "");
                setState((prev) => ({ ...prev, inc: { ...prev.inc, [k]: v } }));
              }}
            />
          </span>
          <span className="sgnw mi">
            <b className="sg">－</b>
            <input
              inputMode="decimal"
              placeholder="0"
              disabled={!on}
              value={state.dec[k] || ""}
              onChange={(e) => {
                const v = e.target.value.replace(/[eE]/g, "");
                setState((prev) => ({ ...prev, dec: { ...prev.dec, [k]: v } }));
              }}
            />
          </span>
        </div>
        {openK === k && renderRelEditor(k, "f")}
      </div>
    );
  };

  const activeThemeEff = DARK_ONLY_PALS.has(pal) ? "dark" : themeMode;
  const curStepTab = state.step === 3 ? 4 : state.step;

  return (
    <>
      {showSplash && <Splash onDone={() => setShowSplash(false)} />}

      {/* الشريط العلوي */}
      <div className="top">
        <div className="top-head">
          <div className="top-title">
            <h1>
              <img src={appIconUrl} alt="" className="app-logo" />
              <span>حاسبة القياسات</span>
            </h1>
            <p>اكتب القياسات واحصل على كل المقاسات جاهزة</p>
          </div>
          <div className="top-acts">
            <button
              type="button"
              className={`thm fs${fontOpen ? " on" : ""}`}
              title="تكبير وتصغير الخط"
              aria-label="حجم الخط"
              onClick={() => {
                setFontOpen((o) => !o);
                setPalBarOpen(false);
              }}
            >
              أ±
            </button>
            <button
              type="button"
              className={`thm${palBarOpen ? " on" : ""}`}
              title="تغيير ألوان التطبيق"
              aria-label="تغيير الثيم"
              onClick={() => {
                setPalBarOpen((o) => !o);
                setFontOpen(false);
              }}
            >
              🎨
            </button>
            <button
              type="button"
              className="thm"
              title="كيف أستخدم التطبيق؟"
              aria-label="مساعدة"
              onClick={() => setWelOpen(true)}
            >
              ؟
            </button>
            <button
              type="button"
              className="thm"
              title={activeThemeEff === "dark" ? "التحويل إلى الوضع الفاتح" : "التحويل إلى الوضع الداكن"}
              aria-label="تبديل الوضع الداكن والفاتح"
              onClick={() => {
                if (DARK_ONLY_PALS.has(pal)) {
                  toast("هذا الثيم مصمم للوضع الداكن. اختر ثيماً آخر لتبديل الفاتح/الداكن", 2600);
                  return;
                }
                setThemeMode((m) => (m === "dark" ? "light" : "dark"));
              }}
            >
              {activeThemeEff === "dark" ? "☀️" : "🌙"}
            </button>
          </div>
        </div>

        {/* مؤشر حجم الخط (5 درجات) */}
        {fontOpen && (
          <div id="fontbar">
            <div className="fontbar-h">
              <b>حجم خط التطبيق</b>
              <span>الدرجة {fontLevel} من 5</span>
            </div>
            <div className="fontbar-dots">
              {FONT_LEVELS.map((fl) => (
                <button
                  key={fl.level}
                  type="button"
                  className={fontLevel === fl.level ? "on" : ""}
                  onClick={() => setFontLevel(fl.level)}
                >
                  <span>{fl.level}</span>
                  <small>{fl.label}</small>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* شريط الخطوات */}
        <div id="steps">
          {[
            { g: 1 as StepNum, num: 1, label: "النوع" },
            { g: 2 as StepNum, num: 2, label: "القياسات" },
            { g: 4 as StepNum, num: 3, label: "النتائج" },
          ].map((s) => (
            <button
              key={s.g}
              type="button"
              className={s.g === curStepTab ? "on" : s.g < curStepTab ? "done" : ""}
              onClick={() => {
                if (s.g !== state.step) {
                  goToStep(s.g);
                }
              }}
            >
              <i>{s.num}</i>
              <span>{s.label}</span>
            </button>
          ))}
        </div>

        {/* شريط الألوان */}
        {palBarOpen && (
          <div id="palbar">
            {PALS.map(([id, name, color]) => (
              <button
                key={id}
                type="button"
                className={pal === id ? "on" : ""}
                onClick={() => {
                  setPal(id);
                  if (id === "sunlit") setThemeMode("light");
                }}
              >
                <i style={{ background: color }} />
                <span className="pn">{name}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      <main>
        {/* ================= الخطوة 1: النوع والمشاريع ================= */}
        {state.step === 1 && (
          <section className="step">
            {/* المشاريع المحفوظة */}
            <details className="proj-sec">
              <summary>
                <span>📁 المشاريع المحفوظة {pjName ? `· ${pjName}` : ""}</span>
                <span className="proj-badge">{projects.length} مشروع</span>
              </summary>
              <div className="proj-save-row">
                <input
                  placeholder="اسم المشروع الحالي (مثال: فيلا أبو أحمد)"
                  value={newProjName}
                  maxLength={60}
                  onChange={(e) => setNewProjName(e.target.value)}
                />
                <button type="button" className="btn sm p" onClick={handleSaveCurrentProject}>
                  💾 حفظ المشروع
                </button>
                <button type="button" className="btn sm w" onClick={handleNewBlankProject}>
                  🆕 جديد
                </button>
                <button type="button" className="btn sm g" onClick={openXlsxImport}>
                  📊 استيراد Excel
                </button>
              </div>
              {projects.length > 0 ? (
                <div className="proj-list">
                  {projects.map((pj) => {
                    const winCnt = pj.data.rows.filter((r) => r.L || r.W || r.p).length;
                    return (
                      <div key={pj.id} className="proj-item">
                        {editingProjId === pj.id ? (
                          <div style={{ display: "flex", gap: 6, flex: 1, flexWrap: "wrap" }}>
                            <input
                              value={editingProjName}
                              onChange={(e) => setEditingProjName(e.target.value)}
                              style={{
                                flex: 1,
                                minWidth: 120,
                                minHeight: 44,
                                padding: "0 10px",
                                borderRadius: 12,
                                background: "var(--card)",
                              }}
                            />
                            <button
                              type="button"
                              className="btn sm g"
                              onClick={() => {
                                const nm = clean(editingProjName).trim();
                                if (!nm) return;
                                setProjects((prev) =>
                                  prev.map((x) => (x.id === pj.id ? { ...x, name: nm } : x))
                                );
                                if (pjName === pj.name) setPjName(nm);
                                setEditingProjId(null);
                              }}
                            >
                              ✓
                            </button>
                            <button
                              type="button"
                              className="btn sm"
                              onClick={() => setEditingProjId(null)}
                            >
                              ✕
                            </button>
                          </div>
                        ) : (
                          <>
                            <div className="proj-item-info">
                              <b>{pj.name}</b>
                              <small>
                                {pj.data.type} · {winCnt} شباك · {pj.updatedAt.slice(0, 10)}
                              </small>
                            </div>
                            <div className="proj-item-acts">
                              <button
                                type="button"
                                className="btn sm g"
                                onClick={() => handleOpenProject(pj)}
                              >
                                فتح
                              </button>
                              <button
                                type="button"
                                className="btn sm"
                                aria-label="إعادة تسمية"
                                onClick={() => {
                                  setEditingProjId(pj.id);
                                  setEditingProjName(pj.name);
                                }}
                              >
                                ✎
                              </button>
                              <button
                                type="button"
                                className="btn sm r"
                                aria-label="حذف المشروع"
                                onClick={() =>
                                  setConfirmDlg({
                                    title: "حذف المشروع",
                                    message: `هل تريد حذف مشروع «${pj.name}» من القائمة؟`,
                                    confirmLabel: "حذف",
                                    danger: true,
                                    onConfirm: () => {
                                      setProjects((prev) => prev.filter((x) => x.id !== pj.id));
                                      toast("تم حذف المشروع");
                                    },
                                  })
                                }
                              >
                                ✕
                              </button>
                            </div>
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
              ) : (
                <p className="rtip" style={{ marginTop: 10 }}>
                  احفظ مشاريعك هنا بأسمائها للرجوع السريع إليها لاحقاً.
                </p>
              )}
            </details>

            <h2>ما نوع المنتج؟</h2>
            <p className="lead">
              اضغط على النوع الذي تريده وسننتقل للخطوة التالية. لتعديل علاقات الخانات أو إضافة نوع جديد استخدم زر «⚙ تعديل الأنواع والعلاقات» في الأسفل.
            </p>

            <div>
              {cat.TN.map((n) => {
                const ed = typeEdit || Boolean(relFor) || ntForm;
                const cOv = Object.keys((state.rel && state.rel[n]) || {}).filter((k) => !isCid(k)).length;
                const uCnt = ((state.custom && state.custom[n]) || []).length;
                const isUser = state.utypes.some((u) => u.name === n);

                return (
                  <div key={n} className="twrap">
                    <button
                      type="button"
                      className={`tcard${n === state.type ? " sel" : ""}`}
                      onClick={() => {
                        setState((prev) => ({ ...prev, type: n }));
                        setRelFor(null);
                        setOpenK(null);
                        goToStep(2, false, { ...state, type: n });
                      }}
                    >
                      <TypeIcon iconKey={cat.TINFO[n]?.iconKey || "d"} />
                      <span>
                        <b>{n}</b>
                        <small>{cat.TINFO[n]?.sub || ""}</small>
                      </span>
                      <span className="tick">✓</span>
                    </button>

                    {ed && (
                      <button
                        type="button"
                        className={`tedit${relFor === n ? " on" : ""}`}
                        onClick={() => {
                          setState((prev) => ({ ...prev, type: n }));
                          setOpenK(null);
                          setRelFor((prev) => (prev === n ? null : n));
                        }}
                      >
                        ⚙ علاقات الخانات{" "}
                        {cOv > 0 && <b className="mod">{cOv} معدّلة</b>}{" "}
                        {uCnt > 0 && <b className="mod">{uCnt} جديدة</b>}{" "}
                        {relFor === n ? "▲" : "▼"}
                      </button>
                    )}

                    {relFor === n && (
                      <div className="rpanel">
                        <h3>علاقات: {n}</h3>
                        {isUser && (
                          <div className="rrow" style={{ marginTop: 8 }}>
                            <label>
                              <span>اسم النوع</span>
                              <input
                                value={typeRenameVal}
                                maxLength={60}
                                style={{ textAlign: "right", direction: "rtl" }}
                                onChange={(e) => setTypeRenameVal(e.target.value)}
                                onBlur={handleRenameType}
                              />
                            </label>
                          </div>
                        )}
                        <p className="rtip">
                          اضغط «✎ تغيير» لتعديل أي علاقة أو إضافة خانة مخصصة لهذا النوع.
                        </p>
                        {(cat.TYPES[n]?.fields || []).map((k) => {
                          const cust = isCid(k);
                          return (
                            <div key={k} className="fld">
                              <div className="fl1">
                                <span style={{ flex: 1, fontWeight: "bold" }}>
                                  {cat.LABEL[k]} {fdim(cat, k) === 2 && <b className="mod">📐 مساحة</b>}
                                  {cust && <b className="mod">جديدة</b>}
                                </span>
                                <button
                                  type="button"
                                  className="edb"
                                  onClick={() => setOpenK((prev) => (prev === k ? null : k))}
                                >
                                  ✎ {openK === k ? "إغلاق" : "تغيير"}
                                </button>
                              </div>
                              <div style={{ fontSize: 13.5, color: "var(--mut)", marginTop: 4 }}>
                                = {relParen(cat, state, k)}
                              </div>
                              {openK === k && renderRelEditor(k, "p")}
                            </div>
                          );
                        })}

                        <div className="rpf" style={{ flexWrap: "wrap" }}>
                          <button
                            type="button"
                            className="btn sm g"
                            onClick={() => addCustomField("")}
                          >
                            ＋ إضافة خانة
                          </button>
                          {isUser && (
                            <button
                              type="button"
                              className="btn sm r"
                              onClick={() => handleDeleteUserType(n)}
                            >
                              🗑 حذف النوع
                            </button>
                          )}
                          <button
                            type="button"
                            className="btn sm p"
                            onClick={() => {
                              setRelFor(null);
                              setOpenK(null);
                            }}
                          >
                            ✓ إغلاق
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}

              {ntForm ? (
                <div className="rpanel">
                  <h3>إضافة نوع جديد</h3>
                  <p className="rtip">
                    اكتب اسم النوع (مثل: شباك سحاب 3 مسارات). بعد الإنشاء تضيف خانات النوع وعلاقاتها.
                  </p>
                  <input
                    className="ntin"
                    placeholder="اسم النوع"
                    maxLength={60}
                    value={ntName}
                    onChange={(e) => setNtName(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        handleCreateType();
                      }
                    }}
                  />
                  <div className="rpf">
                    <button
                      type="button"
                      className="btn sm w"
                      onClick={() => {
                        setNtForm(false);
                        setNtName("");
                      }}
                    >
                      إلغاء
                    </button>
                    <button type="button" className="btn sm g" onClick={handleCreateType}>
                      إنشاء النوع ←
                    </button>
                  </div>
                </div>
              ) : (
                (typeEdit || Boolean(relFor)) && (
                  <button
                    type="button"
                    className="addnew"
                    onClick={() => {
                      setNtForm(true);
                      setRelFor(null);
                    }}
                  >
                    ＋ إضافة نوع جديد
                  </button>
                )
              )}

              <button
                type="button"
                className={`tedit tmode${typeEdit || Boolean(relFor) || ntForm ? " on" : ""}`}
                onClick={() => {
                  const ed = typeEdit || Boolean(relFor) || ntForm;
                  setTypeEdit(!ed);
                  setRelFor(null);
                  setNtForm(false);
                  setOpenK(null);
                }}
              >
                {typeEdit || Boolean(relFor) || ntForm
                  ? "✓ إنهاء تعديل الأنواع"
                  : "⚙ تعديل الأنواع والعلاقات"}
              </button>
            </div>

            <details className="tools">
              <summary>💾 نسخة احتياطية لأنواعك وعلاقاتك</summary>
              <p className="rtip">
                احفظ نسخة قبل تحديث التطبيق أو حذفه، ثم استعدها بعد التثبيت.
              </p>
              <div className="row">
                <button type="button" className="btn g sm" onClick={handleBackupSave}>
                  💾 حفظ نسخة احتياطية
                </button>
                <button
                  type="button"
                  className="btn w sm"
                  onClick={() => {
                    if (bkFileRef.current) {
                      bkFileRef.current.value = "";
                      bkFileRef.current.click();
                    }
                  }}
                >
                  📥 استعادة من ملف
                </button>
                <button type="button" className="btn p sm" onClick={openXlsxImport}>
                  📊 استيراد من Excel
                </button>
              </div>
              {typeof window !== "undefined" && window.location.protocol.startsWith("http") && (
                <div className="row">
                  <a
                    href="/window-calc.html"
                    download="حاسبة_القياسات.html"
                    className="btn p sm"
                    style={{ width: "100%", textDecoration: "none" }}
                  >
                    📲 تنزيل نسخة ملف HTML واحد (للتشغيل بدون إنترنت)
                  </a>
                </div>
              )}
              <input
                ref={bkFileRef}
                type="file"
                accept=".json,application/json"
                hidden
                onChange={handleBackupFile}
              />
            </details>
          </section>
        )}

        {/* ================= الخطوة 2: القياسات ================= */}
        {state.step === 2 && (
          <section className="step">
            <div className="chip">
              <span>النوع:</span>
              <b>{state.type}</b>
              <button type="button" onClick={() => goToStep(1)}>
                تغيير
              </button>
            </div>

            <h2>اكتب قياسات الشبابيك</h2>
            <p className="lead">
              اكتب الطول والعرض لكل شباك، وتُضاف خانة جديدة تلقائياً. (اضغط على رقم الشباك إن أردت تسميته)
            </p>

            <div className="unitbar">
              <span>وحدة القياس:</span>
              {(["m", "cm", "mm"] as UnitCode[]).map((u) => (
                <button
                  key={u}
                  type="button"
                  className={state.unit === u ? "on" : ""}
                  onClick={() => handleUnitChange(u)}
                >
                  {UN[u]}
                </button>
              ))}
            </div>

            {state.unit === "m" && (
              <p id="uhint">
                💡 بالمتر: اكتب <b>150</b> فتصبح <b>1.50</b> تلقائياً. للأطوال 10 م فأكثر اكتب الفاصلة بنفسك (مثل 12.5).
              </p>
            )}

            <details className="how">
              <summary>❓ كيف أعرف الطول والعرض؟</summary>
              <svg viewBox="0 0 300 190" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round">
                <rect x="70" y="14" width="200" height="140" rx="4" />
                <path d="M170 14v140" strokeWidth="2" />
                <path d="M40 14v140M33 24l7-10 7 10M33 144l7 10 7-10" stroke="#e8590c" />
                <path d="M70 176h200M80 169l-10 7 10 7M260 169l10 7-10 7" stroke="#0f6fd6" />
                <text x="8" y="92" fill="#e8590c" stroke="none" fontSize="15" fontWeight="bold" transform="rotate(-90 14 92)">
                  الطول
                </text>
                <text x="170" y="188" fill="#0f6fd6" stroke="none" fontSize="15" fontWeight="bold" textAnchor="middle">
                  العرض
                </text>
              </svg>
              <p>
                <b>الطول</b> = الارتفاع (من فوق إلى تحت). <b>العرض</b> = من جنب إلى جنب.
              </p>
            </details>

            {!state.rows.some((r) => r.L || r.W || r.p) && (
              <div className="hintbox">
                لا تعرف من أين تبدأ؟{" "}
                <button
                  type="button"
                  onClick={() => {
                    setDeletedUndo(null);
                    setState((prev) => ({
                      ...prev,
                      rows: [
                        { n: "", L: "120", W: "150" },
                        { n: "", L: "100", W: "200" },
                        { n: "", L: "", W: "" },
                      ],
                    }));
                    toast("تم إضافة مثالين. اضغط «عرض النتائج»");
                  }}
                >
                  جرّب مثالاً
                </button>
                {" أو "}
                <button type="button" onClick={openXlsxImport}>
                  📊 استيراد من Excel
                </button>
              </div>
            )}

            {middleEmptyCount > 0 && (
              <div className="empty-clean-bar" role="region" aria-label="تنظيف الصفوف الفارغة">
                <span>
                  يوجد {middleEmptyCount} {middleEmptyCount === 1 ? "صف فارغ غير مستخدم" : "صفوف فارغة غير مستخدمة"}
                </span>
                <button type="button" onClick={handleCleanEmptyRows}>
                  🗑 حذف الفارغة
                </button>
              </div>
            )}

            <div ref={rowsBoxRef}>
              {state.rows.map((r, i) => {
                const isBlankMiddle = i < state.rows.length - 1 && isRowBlank(r);
                const err = rowErr(r, state.unit, negHighlightedRows.has(r));
                const warn = !err && rowWarn(r, state.unit);
                const pl = pL(r);
                const pw = pW(r);
                const pendBarVisible = openPendingBarIdx === i;
                const hasNameEdit = editingNameIdx === i;

                return (
                  <React.Fragment key={i}>
                    {/* زر التراجع يظهر فقط إذا تم الحذف وبجانب موضع الشباك المحذوف مباشرة */}
                    {deletedUndo && deletedUndo.index === i && (
                      <div className="undo-slot" role="status">
                        <span>
                          تم حذف الشباك {deletedUndo.row.n ? `(${deletedUndo.row.n})` : `#${deletedUndo.index + 1}`}
                        </span>
                        <div className="undo-slot-acts">
                          <button type="button" className="undo-btn" onClick={handleUndoDelete}>
                            ↶ تراجع
                          </button>
                          <button
                            type="button"
                            className="undo-dismiss"
                            aria-label="إغلاق التراجع"
                            onClick={() => setDeletedUndo(null)}
                          >
                            ✕
                          </button>
                        </div>
                      </div>
                    )}

                    <div className={`wcard${err ? " bad" : warn ? " warn" : ""}`}>
                      {/* صف علوي منفصل للشارة والاسم وزر الحذف ✕ (بعيد تماماً عن خانة العدد) */}
                      <div className="wcard-top">
                        <button
                          type="button"
                          className={`badge${r.n ? " cust" : ""}`}
                          title="اضغط لتغيير رقم الشباك (نقل) أو تسميته"
                          onClick={() => openBadgeModal(i)}
                        >
                          {i + 1}
                        </button>

                        {hasNameEdit ? (
                          <input
                            className="wname-input"
                            placeholder="اسم الشباك (مثلاً: مطبخ)"
                            maxLength={60}
                            value={r.n}
                            autoFocus
                            onChange={(e) => updateRow(i, { n: clean(e.target.value) })}
                            onBlur={() => setEditingNameIdx(null)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") {
                                e.preventDefault();
                                setEditingNameIdx(null);
                              }
                            }}
                          />
                        ) : (
                          <button
                            type="button"
                            className={`wname-btn${r.n ? " has" : ""}`}
                            onClick={() => setEditingNameIdx(i)}
                          >
                            {r.n || "تسمية الشباك (اختياري)"}
                          </button>
                        )}

                        <button
                          type="button"
                          className="del"
                          title={isBlankMiddle ? "حذف هذا الصف الفارغ" : "حذف الشباك"}
                          aria-label={isBlankMiddle ? `حذف الصف الفارغ ${i + 1}` : `حذف شباك ${i + 1}`}
                          onClick={() => handleDeleteRow(i)}
                        >
                          {isBlankMiddle ? "🗑" : "✕"}
                        </button>
                      </div>

                      <div className="wr">
                        <label>
                          <span>
                            الطول ({UN[state.unit]})
                            {pl && <em className="pnd-tag">⏳ غير جاهز</em>}
                          </span>
                          <input
                            className={`big L${pl ? " pnd" : ""}`}
                            inputMode="decimal"
                            enterKeyHint="next"
                            placeholder={pl ? "⏳ لاحقاً" : "—"}
                            title={pl && r.L ? `القيمة المحفوظة: ${r.L}` : undefined}
                            disabled={pl}
                            value={pl ? "" : r.L}
                            onChange={(e) => {
                              if (negHighlightedRows.has(r)) {
                                const ws = new WeakSet(
                                  state.rows.filter((x) => x !== r && negHighlightedRows.has(x))
                                );
                                setNegHighlightedRows(ws);
                              }
                              const v = mfixValue(e.target.value, state.unit);
                              updateRow(i, { L: v, ...(v ? { p: r.p === "B" ? "W" : r.p === "L" ? "" : r.p } : {}) });
                            }}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") {
                                e.preventDefault();
                                const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
                                const wEl = cards?.[i]?.querySelector(".W:not(:disabled)") as HTMLInputElement | null;
                                if (wEl) {
                                  wEl.focus();
                                } else {
                                  const nextL = cards?.[i + 1]?.querySelector(".L:not(:disabled), .W:not(:disabled)") as HTMLInputElement | null;
                                  nextL?.focus();
                                }
                              }
                            }}
                          />
                        </label>

                        <label>
                          <span>
                            العرض ({UN[state.unit]})
                            {pw && <em className="pnd-tag">⏳ غير جاهز</em>}
                          </span>
                          <input
                            className={`big W${pw ? " pnd" : ""}`}
                            inputMode="decimal"
                            enterKeyHint="next"
                            placeholder={pw ? "⏳ لاحقاً" : "—"}
                            title={pw && r.W ? `القيمة المحفوظة: ${r.W}` : undefined}
                            disabled={pw}
                            value={pw ? "" : r.W}
                            onChange={(e) => {
                              if (negHighlightedRows.has(r)) {
                                const ws = new WeakSet(
                                  state.rows.filter((x) => x !== r && negHighlightedRows.has(x))
                                );
                                setNegHighlightedRows(ws);
                              }
                              const v = mfixValue(e.target.value, state.unit);
                              updateRow(i, { W: v, ...(v ? { p: r.p === "B" ? "L" : r.p === "W" ? "" : r.p } : {}) });
                            }}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") {
                                e.preventDefault();
                                const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
                                const nextEl = cards?.[i + 1]?.querySelector(".L:not(:disabled), .W:not(:disabled)") as HTMLInputElement | null;
                                nextEl?.focus();
                              }
                            }}
                          />
                        </label>

                        <label className="qb">
                          <span>العدد</span>
                          <input
                            className="big Q"
                            inputMode="numeric"
                            enterKeyHint="next"
                            placeholder="1"
                            maxLength={4}
                            value={r.q || ""}
                            onChange={(e) => {
                              let rawVal = e.target.value.replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));
                              if (/[.,٫،]/.test(rawVal)) {
                                rawVal = rawVal.split(/[.,٫،]/)[0] || "1";
                                toast("العدد يجب أن يكون رقماً صحيحاً");
                              }
                              let v = rawVal.replace(/[^0-9]/g, "").slice(0, 4);
                              if (v === "0" || /^0+$/.test(v)) {
                                if (v !== "") toast("أقل عدد مسموح للشباك هو 1", 1800);
                                v = "";
                              } else {
                                v = v.replace(/^0+/, "");
                              }
                              if (Number(v) > MAXQ) v = String(MAXQ);
                              updateRow(i, { q: v });
                            }}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") {
                                e.preventDefault();
                                const cards = rowsBoxRef.current?.querySelectorAll(".wcard");
                                const nextEl = cards?.[i + 1]?.querySelector(".L:not(:disabled), .W:not(:disabled)") as HTMLInputElement | null;
                                nextEl?.focus();
                              }
                            }}
                          />
                        </label>
                      </div>

                      <div className="pbar">
                        <button
                          type="button"
                          className={`plink${r.p ? " on" : ""}`}
                          onClick={() => setOpenPendingBarIdx((cur) => (cur === i ? null : i))}
                        >
                          {r.p === "L"
                            ? "⏳ الطول غير جاهز (تعديل)"
                            : r.p === "W"
                              ? "⏳ العرض غير جاهز (تعديل)"
                              : r.p === "B"
                                ? "⏳ الطول والعرض غير جاهز (تعديل)"
                                : "⏳ قياس غير جاهز؟"}
                        </button>

                        {pendBarVisible && (
                          <div className="pchips">
                            <small>المقاس غير الجاهز:</small>
                            {(
                              [
                                ["L", "الطول غير جاهز"],
                                ["W", "العرض غير جاهز"],
                                ["B", "كلاهما غير جاهز"],
                              ] as [PendingMode, string][]
                            ).map(([mode, lbl]) => (
                              <button
                                key={mode}
                                type="button"
                                className={r.p === mode ? "on" : ""}
                                onClick={() => {
                                  const nextP: PendingMode = r.p === mode ? "" : mode;
                                  updateRow(i, { p: nextP });
                                  if (!nextP) setOpenPendingBarIdx(null);
                                }}
                              >
                                {lbl}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>

                      {err && <div className="wmsg">{err}</div>}
                    </div>
                  </React.Fragment>
                );
              })}

              {/* إذا كان الحذف بعد آخر بطاقة متبقية */}
              {deletedUndo && deletedUndo.index >= state.rows.length && (
                <div className="undo-slot" role="status">
                  <span>
                    تم حذف الشباك {deletedUndo.row.n ? `(${deletedUndo.row.n})` : `#${deletedUndo.index + 1}`}
                  </span>
                  <div className="undo-slot-acts">
                    <button type="button" className="undo-btn" onClick={handleUndoDelete}>
                      ↶ تراجع
                    </button>
                    <button
                      type="button"
                      className="undo-dismiss"
                      aria-label="إغلاق التراجع"
                      onClick={() => setDeletedUndo(null)}
                    >
                      ✕
                    </button>
                  </div>
                </div>
              )}
            </div>

            <div className="addrow">
              <button type="button" className="add" onClick={handleAddEmptyRow}>
                ＋ إضافة شباك
              </button>
              <button type="button" className="add clr" onClick={handleClearAll}>
                🗑 تصفير الكل
              </button>
            </div>

            <div className="addrow" style={{ marginTop: 10 }}>
              <button
                type="button"
                className="add"
                style={{ border: "2px solid var(--pri)" }}
                onClick={handleDupLast}
              >
                ⧉ نسخ الشباك السابق
              </button>
              <button
                type="button"
                className="add"
                style={{ border: "2px solid var(--pri)" }}
                onClick={openCopyByNumber}
              >
                🔢 نسخ شباك محدد
              </button>
            </div>

            <details className="tools" style={{ marginTop: 12 }}>
              <summary>📋 أدوات: لصق قياسات · استيراد Excel · مسح الكل</summary>
              <textarea
                placeholder="الصق هنا: كل سطر فيه طول ثم عرض (أو طول عرض اسم)"
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
              />
              <div className="row">
                <button type="button" className="btn g sm" onClick={handlePaste}>
                  إضافة المُلصق
                </button>
                <button type="button" className="btn p sm" onClick={openXlsxImport}>
                  📊 استيراد من ملف Excel
                </button>
                <button type="button" className="btn r sm" onClick={handleClearAll}>
                  🗑 مسح كل الشبابيك
                </button>
              </div>
            </details>
          </section>
        )}

        {/* ================= الخطوة 3: تخصيص الخانات ================= */}
        {state.step === 3 && (
          <section className="step">
            <div className="chip">
              <span>النوع:</span>
              <b>{state.type}</b>
              <button type="button" onClick={() => goToStep(1)}>
                تغيير
              </button>
            </div>

            <div className="h2r">
              <h2>تخصيص الخانات</h2>
              <button
                type="button"
                className="ic ok"
                title="حفظ"
                onClick={() => {
                  snap3Ref.current = null;
                  setSnap3(null);
                  toast("تم حفظ التغييرات");
                  goToStep(4, false, undefined, true);
                }}
              >
                ✓
              </button>
              <button
                type="button"
                className="ic no"
                title="إلغاء"
                onClick={() => goToStep(4)}
              >
                ✕
              </button>
            </div>

            <p className="lead">
              كل خانة تُحسب بعلاقة مكتوبة تحت اسمها. يمكنك إظهار أو إخفاء أي خانة أو إعادة ترتيبها أو تعديل معادلتها.
            </p>

            {sampleWin && (
              <div className="base">
                <div className="bv">
                  <small>الطول</small>
                  <b>{sampleWin.L === null ? "—" : fmt(sampleWin.L, state.unit)}</b>
                </div>
                <div className="bv">
                  <small>العرض</small>
                  <b>{sampleWin.W === null ? "—" : fmt(sampleWin.W, state.unit)}</b>
                </div>
                <p>
                  القياسان الأساسيان (من أول شباك). كل الخانات تُحسب منهما، والأرقام الملوّنة أدناه هي نتيجة كل خانة لهذا الشباك.
                </p>
              </div>
            )}

            <div className="leg">
              <span>
                <i style={{ background: "#d9f3e2", color: "#14532d" }}>زيادة ＋</i> تُضاف للقيمة
              </span>
              <span>
                <i style={{ background: "#fde3e1", color: "#a02a22" }}>نقصان －</i> تُنقص منها
              </span>
            </div>

            <div className="seg cust-seg" role="tablist" aria-label="وضع تخصيص الخانات">
              <button
                type="button"
                role="tab"
                aria-selected={!advOpen}
                className={!advOpen ? "on" : ""}
                onClick={() => setAdvOpen(false)}
              >
                ↕ اختيار وترتيب
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={advOpen}
                className={advOpen ? "on" : ""}
                onClick={() => setAdvOpen(true)}
              >
                ⚙ علاقات وتعديل
              </button>
            </div>

            <div className={advOpen ? "" : "noadv"}>
              <div className="row" style={{ margin: "0 0 10px" }}>
                <button
                  type="button"
                  className="btn sm w"
                  onClick={() => {
                    if (!relUndo.length) {
                      toast("لا يوجد ما يُتراجع عنه");
                      return;
                    }
                    const lastSnap = JSON.parse(relUndo[relUndo.length - 1]);
                    setRelUndo((prev) => prev.slice(0, -1));
                    setState((prev) => {
                      const next = { ...prev };
                      SNAP_KEYS.forEach((k) => {
                        (next as Record<string, unknown>)[k] = lastSnap[k];
                      });
                      next.layout = syncLayoutWithCustom(next);
                      return next;
                    });
                    toast("تم التراجع عن آخر تعديل");
                  }}
                >
                  ↶ تراجع
                </button>
                <button
                  type="button"
                  className="btn sm g"
                  onClick={() => {
                    if (colG.size) setColG(new Set());
                    else setColG(new Set(visUnits));
                  }}
                >
                  {colG.size ? "▾ فتح الكل" : "◂ طيّ الكل"}
                </button>
                <button
                  type="button"
                  className="btn sm g"
                  onClick={() => {
                    pushUndo3(true);
                    const t = cat.TYPES[state.type];
                    if (!t) return;
                    setState((prev) => {
                      const nextSel = { ...prev.sel };
                      t.fields.forEach((k) => {
                        nextSel[k] = true;
                      });
                      return { ...prev, sel: nextSel };
                    });
                  }}
                >
                  تحديد الكل
                </button>
                <button
                  type="button"
                  className="btn sm w"
                  onClick={() => {
                    pushUndo3(true);
                    const t = cat.TYPES[state.type];
                    if (!t) return;
                    setState((prev) => {
                      const nextSel = { ...prev.sel };
                      t.fields.forEach((k) => {
                        nextSel[k] = false;
                      });
                      return { ...prev, sel: nextSel };
                    });
                  }}
                >
                  إلغاء الكل
                </button>
                <button
                  type="button"
                  className="btn sm w advonly"
                  onClick={() => {
                    pushUndo3(true);
                    setState((prev) => ({ ...prev, inc: {}, dec: {} }));
                  }}
                >
                  تصفير التعديلات
                </button>
                <button
                  type="button"
                  className="btn sm w advonly"
                  onClick={() => {
                    pushUndo3(true);
                    setState((prev) => {
                      const cur = { ...(prev.rel[prev.type] || {}) };
                      Object.keys(cur).forEach((k) => {
                        if (!isCid(k)) delete cur[k];
                      });
                      return { ...prev, rel: { ...prev.rel, [prev.type]: cur } };
                    });
                  }}
                >
                  إرجاع العلاقات الأصلية
                </button>
                <button
                  type="button"
                  className="btn sm w advonly"
                  onClick={() => {
                    pushUndo3(true);
                    setState((prev) => {
                      const next = { ...prev, layout: defaultLayout() };
                      next.layout = syncLayoutWithCustom(next);
                      return next;
                    });
                  }}
                >
                  الترتيب الافتراضي
                </button>
              </div>

              <p className="rtip" style={{ marginBottom: 10 }}>
                ✋ لإعادة الترتيب: اسحب أي بطاقة من <b>⋮⋮</b> أو بالضغط والسحب من <b>أي مكان فيها</b>، أو استخدم الأسهم.
              </p>

              {visUnits.map((id) => {
                const t = cat.TYPES[state.type];
                const ks = (state.layout.kids[id] || []).filter((k) => t?.fields.includes(k));
                const collapsed = colG.has(id);
                const groupMode = isGroupUnit(id);
                const isUnitDragging = dragUnit?.id === id;
                const isUnitDropBefore = dragUnit?.targetId === id && dragUnit?.pos === "before";
                const isUnitDropAfter = dragUnit?.targetId === id && dragUnit?.pos === "after";
                const unitDragCls = `${isUnitDragging ? " dragging" : ""}${isUnitDropBefore ? " drop-before" : ""}${isUnitDropAfter ? " drop-after" : ""}`;

                const unitGripSvg = (
                  <svg width="14" height="18" viewBox="0 0 14 18" aria-hidden="true">
                    <circle cx="4" cy="3.5" r="1.7" fill="currentColor" />
                    <circle cx="10" cy="3.5" r="1.7" fill="currentColor" />
                    <circle cx="4" cy="9" r="1.7" fill="currentColor" />
                    <circle cx="10" cy="9" r="1.7" fill="currentColor" />
                    <circle cx="4" cy="14.5" r="1.7" fill="currentColor" />
                    <circle cx="10" cy="14.5" r="1.7" fill="currentColor" />
                  </svg>
                );

                const unitCollapseIcon = (
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 14 14"
                    aria-hidden="true"
                    style={{
                      transform: collapsed ? "rotate(90deg)" : "rotate(0deg)",
                      transition: "transform .16s ease",
                    }}
                  >
                    <path
                      d="M3.2 5.2L7 9L10.8 5.2"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2.2"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                );

                if (groupMode) {
                  const selectedCount = ks.filter((k) => isFieldSelected(k, state.sel)).length;
                  const allChecked = selectedCount === ks.length;
                  const partial = selectedCount > 0 && selectedCount < ks.length;

                  return (
                    <div
                      key={id}
                      data-drag-unit={id}
                      className={`ur${unitDragCls}`}
                      style={{ ["--h" as string]: cat.HUE[id] ?? 0 }}
                      onPointerDown={(e) =>
                        startCardDragGesture(
                          e,
                          () => beginUnitDrag(id),
                          ".fld, .fl2, button, input, select, textarea, a"
                        )
                      }
                    >
                      <div className="uh">
                        <button
                          type="button"
                          className="drag-h"
                          title="اسحب لإعادة الترتيب"
                          aria-label="سحب لإعادة الترتيب"
                          onPointerDown={(e) => startUnitDrag(e, id)}
                        >
                          {unitGripSvg}
                        </button>
                        <label>
                          <input
                            type="checkbox"
                            checked={allChecked}
                            ref={(el) => {
                              if (el) el.indeterminate = partial;
                            }}
                            onChange={(e) => {
                              pushUndo3(true);
                              const ch = e.target.checked;
                              setState((prev) => {
                                const nextSel = { ...prev.sel };
                                ks.forEach((k) => {
                                  nextSel[k] = ch;
                                });
                                return { ...prev, sel: nextSel };
                              });
                            }}
                          />
                          <b>{GN(id)}</b>
                        </label>
                        <div className="o-dock">
                          <span className="mvb">
                            <button type="button" title="لأعلى" aria-label="لأعلى" onClick={() => moveUnit(id, -1)}>
                              <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                                <path d="M5 1.2L9 6.8H1L5 1.2Z" fill="currentColor" />
                              </svg>
                            </button>
                            <button type="button" title="لأسفل" aria-label="لأسفل" onClick={() => moveUnit(id, 1)}>
                              <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                                <path d="M5 6.8L1 1.2H9L5 6.8Z" fill="currentColor" />
                              </svg>
                            </button>
                          </span>
                        </div>
                        <button
                          type="button"
                          className="chev"
                          aria-label="طيّ/فتح"
                          onClick={() =>
                            setColG((prev) => {
                              const next = new Set(prev);
                              if (next.has(id)) next.delete(id);
                              else next.add(id);
                              return next;
                            })
                          }
                        >
                          {unitCollapseIcon}
                        </button>
                      </div>
                      {!collapsed && <div>{ks.map((k) => renderFieldRow(k, false, id))}</div>}
                    </div>
                  );
                }

                const k = ks[0];
                if (!k) return null;
                const on = isFieldSelected(k, state.sel);
                const isA = fdim(cat, k) === 2;
                const val = sampleWin ? fmt(shownV(state, k, sampleWin.res[k], isA), state.unit) : "—";

                return (
                  <div
                    key={id}
                    data-drag-unit={id}
                    className={`ur${unitDragCls}`}
                    style={{ ["--h" as string]: cat.HUE[id] ?? 0 }}
                    onPointerDown={(e) =>
                      startCardDragGesture(
                        e,
                        () => beginUnitDrag(id),
                        ".fld, .fl2, button, input, select, textarea, a"
                      )
                    }
                  >
                    <div className="uh">
                      <button
                        type="button"
                        className="drag-h"
                        title="اسحب لإعادة الترتيب"
                        aria-label="سحب لإعادة الترتيب"
                        onPointerDown={(e) => startUnitDrag(e, id)}
                      >
                        {unitGripSvg}
                      </button>
                      <label>
                        <input
                          type="checkbox"
                          checked={on}
                          onChange={(e) => {
                            pushUndo3(true);
                            const checked = e.target.checked;
                            setState((prev) => ({ ...prev, sel: { ...prev.sel, [k]: checked } }));
                          }}
                        />
                        <b>{cat.LABEL[k]}</b>
                        {isA && <b className="mod">📐 مساحة</b>}
                        {isCid(k) && <b className="mod">جديدة</b>}
                      </label>
                      <div className="o-dock">
                        {on && val && <span className="chipv">{val}</span>}
                        <span className="mvb">
                          <button type="button" title="لأعلى" aria-label="لأعلى" onClick={() => moveUnit(id, -1)}>
                            <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                              <path d="M5 1.2L9 6.8H1L5 1.2Z" fill="currentColor" />
                            </svg>
                          </button>
                          <button type="button" title="لأسفل" aria-label="لأسفل" onClick={() => moveUnit(id, 1)}>
                            <svg width="10" height="8" viewBox="0 0 10 8" aria-hidden="true">
                              <path d="M5 6.8L1 1.2H9L5 6.8Z" fill="currentColor" />
                            </svg>
                          </button>
                        </span>
                      </div>
                      <button
                        type="button"
                        className="chev"
                        aria-label="طيّ/فتح"
                        onClick={() =>
                          setColG((prev) => {
                            const next = new Set(prev);
                            if (next.has(id)) next.delete(id);
                            else next.add(id);
                            return next;
                          })
                        }
                      >
                        {unitCollapseIcon}
                      </button>
                    </div>
                    {!collapsed && renderFieldRow(k, true)}
                  </div>
                );
              })}

              <div className="sec">
                <label className="opt">
                  <input
                    type="checkbox"
                    checked={state.cascade}
                    onChange={(e) => {
                      pushUndo3(true);
                      setState((prev) => ({ ...prev, cascade: e.target.checked }));
                    }}
                  />
                  <span>
                    <b>تطبيق التعديل على الخانات المرتبطة أيضاً</b>
                    <br />
                    <small style={{ color: "var(--mut)" }}>
                      إذا غيّرت خانة تتغير الخانات المبنية عليها تلقائياً.
                    </small>
                  </span>
                </label>
              </div>

              <button type="button" className="addnew advonly" onClick={() => addCustomField("")}>
                ＋ إضافة خانة جديدة
              </button>
            </div>
          </section>
        )}

        {/* ================= الخطوة 4: النتائج ================= */}
        {state.step === 4 && (
          <section className="step">
            <div className="chip">
              <span>النوع:</span>
              <b>{state.type}</b>
              <button type="button" onClick={() => goToStep(1)}>
                تغيير
              </button>
            </div>

            <h2>النتائج ✅</h2>

            <div className="sum">
              <span>
                {res.data.length} شباك{res.hasQ ? ` · ${res.qtotal} قطعة` : ""}
              </span>
              <span>{res.keys.length} خانة</span>
              {res.areaTotals.map((item) => (
                <span key={item.key} className="atot">
                  إجمالي {item.label}: <b>{item.totalSqM.toFixed(2)}</b> م²
                </span>
              ))}
              {res.adj && <span>فيه تعديلات</span>}
              {res.rndOn && (
                <button
                  type="button"
                  className={`rn-tog${state.raw ? " on" : ""}`}
                  onClick={() => setState((prev) => ({ ...prev, raw: !prev.raw }))}
                >
                  {state.raw ? "✓ الأرقام الأصلية" : "عرض الأرقام الأصلية"}
                </button>
              )}
            </div>

            <div className="tb3">
              <button
                type="button"
                id="bRnd"
                className={`btn${Number(state.gr.s) > 0 || rndOpen ? " on" : ""}`}
                onClick={() => setRndOpen((o) => !o)}
              >
                {Number(state.gr.s) > 0 ? `≈ تقريب: ${state.gr.s} ✓` : "≈ تقريب"}
              </button>
              <button type="button" id="bAdv" className="btn" onClick={() => goToStep(3)}>
                ⚙️ تخصيص
              </button>
              <button
                type="button"
                id="bAdv2"
                className="btn"
                onClick={() => {
                  setAdvOpen(true);
                  goToStep(3);
                }}
              >
                🛠 متقدمة
              </button>
            </div>

            {rndOpen && res.data.length > 0 && res.keys.length > 0 && (
              <div className="rn-gbar">
                <div className="rn-gt">
                  <b>التقريب لكل النتائج</b>
                  <small>
                    {Number(state.gr.s) > 0
                      ? "يشمل خانات الطول والعرض (ما عدا المساحات وما له تقريب خاص)"
                      : "اختياري. الأرقام تظهر كما هي"}
                  </small>
                </div>
                <div className="rn-seg">
                  {[
                    [0, "بدون"],
                    [0.1, "0.1"],
                    [0.5, "0.5"],
                    [1, "1"],
                  ].map(([v, t]) => (
                    <button
                      key={String(v)}
                      type="button"
                      className={(Number(state.gr.s) || 0) === v ? "on" : ""}
                      onClick={() =>
                        setState((prev) => ({ ...prev, gr: { s: Number(v), d: prev.gr.d || "n" } }))
                      }
                    >
                      {t}
                    </button>
                  ))}
                </div>
                <div
                  className={`rn-seg dir${Number(state.gr.s) > 0 ? "" : " off"}`}
                  style={{ marginTop: 8 }}
                >
                  {[
                    ["n", "لأقرب"],
                    ["d", "لأسفل"],
                    ["u", "لأعلى"],
                  ].map(([v, t]) => (
                    <button
                      key={v}
                      type="button"
                      disabled={!(Number(state.gr.s) > 0)}
                      className={(state.gr.d || "n") === v ? "on" : ""}
                      onClick={() =>
                        setState((prev) => ({ ...prev, gr: { ...prev.gr, d: v as RoundDir } }))
                      }
                    >
                      {t}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {res.data.length > 0 && (
              <div id="srow">
                <button
                  type="button"
                  id="sBtn"
                  className={searchOpen ? "on" : ""}
                  title="بحث برقم الشباك"
                  aria-label="بحث"
                  onClick={() => {
                    if (searchOpen) {
                      setSearchOpen(false);
                      setSearchQuery("");
                    } else {
                      setSearchOpen(true);
                    }
                  }}
                >
                  🔍
                </button>
                {searchOpen && (
                  <>
                    <input
                      id="rsearch"
                      type="search"
                      inputMode="search"
                      placeholder="رقم الشباك أو اسمه"
                      autoFocus
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                    <button
                      type="button"
                      id="sX"
                      aria-label="إغلاق البحث"
                      onClick={() => {
                        setSearchOpen(false);
                        setSearchQuery("");
                      }}
                    >
                      ✕
                    </button>
                  </>
                )}
              </div>
            )}

            <div className="seg">
              <button
                type="button"
                className={state.view === "cards" ? "on" : ""}
                onClick={() => setState((prev) => ({ ...prev, view: "cards" }))}
              >
                بطاقات
              </button>
              <button
                type="button"
                className={state.view === "table" ? "on" : ""}
                onClick={() => setState((prev) => ({ ...prev, view: "table" }))}
              >
                جدول
              </button>
            </div>

            <div id="unote">
              📏 الوحدة: <b>{UFULL[state.unit]}</b>
              {res.area && (
                <>
                  {" "}
                  &nbsp;·&nbsp; المساحة: <b>{UAREA[state.unit]}</b>
                </>
              )}
              {res.pend.length > 0 && (
                <div className="pnote">
                  ⏳ قياسات غير جاهزة في الشباك {res.pend.join("، ")}: خاناتها فارغة وتُعبّأ يدوياً لاحقاً
                </div>
              )}
            </div>

            {!res.keys.length ? (
              <div className="empty">
                لم تختر أي خانة.
                <br />
                اضغط «تخصيص» واختر ما تريد.
              </div>
            ) : !res.data.length ? (
              <div className="empty">لا توجد قياسات. ارجع للخطوة 2.</div>
            ) : (
              <>
                {state.view === "cards" ? (
                  <div>
                    {res.data
                      .filter(
                        (d) =>
                          !normalizedSearch ||
                          `${d.n} ${d.nm}`.toLowerCase().includes(normalizedSearch)
                      )
                      .map((d) => (
                        <div key={d.n} className="rc">
                          <div className="rc-h">
                            <b>
                              شباك {d.n}
                              {d.nm && <small className="nmx"> · {d.nm}</small>}
                              {d.q > 1 && <small className="qx">×{d.q}</small>}
                            </b>
                            <span>
                              {d.L === null ? "⏳" : fmt(d.L, state.unit)} ×{" "}
                              {d.W === null ? "⏳" : fmt(d.W, state.unit)}
                            </span>
                          </div>
                          {res.units.map((u) => (
                            <div key={u.id} className="blk" style={{ ["--h" as string]: u.hue }}>
                              <div className="bn">{u.group ? u.title : cat.LABEL[u.ks[0]]}</div>
                              <div className="vals">
                                {u.ks.map((k) => {
                                  const isA = fdim(cat, k) === 2;
                                  const o = fmt(d.res[k], state.unit);
                                  const v = fmt(shownV(state, k, d.res[k], isA), state.unit);
                                  const ch = v !== o;
                                  const subLbl = cat.SUB[k] + (isA ? ` (${UAS[state.unit]})` : "");
                                  return (
                                    <button
                                      key={k}
                                      type="button"
                                      className="v"
                                      onClick={() => {
                                        if (v !== "") copyText(v, `تم نسخ ${v}`);
                                      }}
                                    >
                                      {u.group && <small>{subLbl}</small>}
                                      <b>{v === "" ? "⏳" : v}</b>
                                      {ch && (
                                        <i className="rn-og" title="القيمة الأصلية">
                                          {o}
                                        </i>
                                      )}
                                    </button>
                                  );
                                })}
                              </div>
                            </div>
                          ))}
                        </div>
                      ))}
                  </div>
                ) : (
                  <>
                    <div className="res">
                      <table>
                        <thead>
                          <tr>
                            <th className="fx" rowSpan={2}>
                              رقم
                            </th>
                            <th className="fx" rowSpan={2}>
                              طول
                            </th>
                            <th className="fx" rowSpan={2}>
                              عرض
                            </th>
                            {res.hasN && (
                              <th className="fx" rowSpan={2}>
                                الاسم
                              </th>
                            )}
                            {res.units.map((u) =>
                              u.group ? (
                                <th
                                  key={u.id}
                                  className="c gs ge"
                                  style={{ ["--h" as string]: u.hue }}
                                  colSpan={u.ks.length}
                                >
                                  {u.title}
                                </th>
                              ) : (
                                <th
                                  key={u.id}
                                  className="c gs ge"
                                  style={{ ["--h" as string]: u.hue }}
                                  rowSpan={2}
                                >
                                  {cat.LABEL[u.ks[0]] + (fdim(cat, u.ks[0]) === 2 ? ` (${UAS[state.unit]})` : "")}
                                </th>
                              )
                            )}
                          </tr>
                          <tr>
                            {res.units.flatMap((u) =>
                              u.group
                                ? u.ks.map((k, idx) => (
                                    <th
                                      key={k}
                                      className={`c${idx === 0 ? " gs" : ""}${idx === u.ks.length - 1 ? " ge" : ""}`}
                                      style={{ ["--h" as string]: u.hue }}
                                    >
                                      {cat.SUB[k] + (fdim(cat, k) === 2 ? ` (${UAS[state.unit]})` : "")}
                                    </th>
                                  ))
                                : []
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {res.table.map((row, ri) => {
                            const d = res.data[ri];
                            if (
                              normalizedSearch &&
                              !`${d.n} ${d.nm}`.toLowerCase().includes(normalizedSearch)
                            ) {
                              return null;
                            }
                            return (
                              <tr key={d.n}>
                                <td className="fx">
                                  {row[0]}
                                  {d.q > 1 && <small className="qx"> ×{d.q}</small>}
                                </td>
                                <td className="fx">{row[1] === "" ? "⏳" : row[1]}</td>
                                <td className="fx">{row[2] === "" ? "⏳" : row[2]}</td>
                                {res.hasN && <td className="fx nmc">{d.nm}</td>}
                                {res.units.flatMap((u) =>
                                  u.ks.map((k, idx) => {
                                    const colOffset = 3 + res.keys.indexOf(k);
                                    const v = row[colOffset];
                                    const orig = fmt(d.res[k], state.unit);
                                    return (
                                      <td
                                        key={k}
                                        className={`c${idx === 0 ? " gs" : ""}${idx === u.ks.length - 1 ? " ge" : ""}`}
                                        style={{ ["--h" as string]: u.hue, cursor: v === "" ? "default" : "pointer" }}
                                        title={v === "" ? undefined : "اضغط لنسخ الرقم"}
                                        onClick={() => {
                                          if (v !== "") copyText(v, `تم نسخ ${v}`);
                                        }}
                                      >
                                        {v === "" ? "⏳" : v}
                                        {orig !== v && <i className="rn-og">{orig}</i>}
                                      </td>
                                    );
                                  })
                                )}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    <p className="rtip">↔ اسحب الجدول يميناً ويساراً لعرض باقي الخانات بالكامل.</p>
                  </>
                )}

                {/* ملخص إجمالي المساحات (يظهر فقط عند تفعيل خانات المساحة من التخصيص) */}
                {res.areaTotals.length > 0 && (
                  <div className="area-sum">
                    <div className="dup-h">
                      <b>📐 ملخص إجمالي المساحات (لطلب المواد)</b>
                      <span>إجمالي القطع: {res.qtotal}</span>
                    </div>
                    {res.areaTotals.map((item) => (
                      <div key={item.key} className="dup-r">
                        <span style={{ fontWeight: 800, flex: 1 }}>{item.label}</span>
                        <span className="dup-c">{item.totalSqM.toFixed(3)} م²</span>
                        {state.unit !== "m" && (
                          <small>
                            بالوحدة المختارة: {item.totalInUnit.toFixed(2)} {UAS[state.unit]}
                          </small>
                        )}
                      </div>
                    ))}
                  </div>
                )}

                {/* ملخص الشبابيك المتشابهة (150×100 منفصلة عن 100×150) */}
                {res.dups.length > 0 && (
                  <div className="dup">
                    <div className="dup-h">
                      <b>🔁 ملخص الشبابيك المتشابهة</b>
                      <span>
                        {res.dups.length} قياس مكرر · {res.dups.reduce((a, g) => a + g.cnt, 0)} شباك
                      </span>
                    </div>
                    {res.dups.map((g, idx) => (
                      <div key={idx} className="dup-r">
                        <span className="dup-s">
                          {fmt(g.L, state.unit)} × {fmt(g.W, state.unit)}
                        </span>
                        <span className="dup-c">{g.cnt} شبابيك</span>
                        <small>الأرقام: {g.nums.join("، ")}</small>
                      </div>
                    ))}
                    {res.dupOther > 0 && (
                      <div className="dup-o">
                        وباقي الشبابيك ({res.dupOther}) قياساتها مختلفة وغير مكررة
                      </div>
                    )}
                  </div>
                )}

                <p className="lead" style={{ marginTop: 10 }}>
                  💡 اضغط على أي رقم في البطاقات أو الجدول لنسخه.
                </p>
              </>
            )}
          </section>
        )}
      </main>

      {/* الشريط السفلي */}
      {state.step > 1 && (
        <div id="bar">
          {state.step === 4 && (
            <>
              <button type="button" className="btn back s4" onClick={() => goToStep(2)}>
                → القياسات
              </button>
              <button type="button" className="btn p" onClick={() => openSaveDialog()}>
                💾 حفظ ومشاركة
              </button>
            </>
          )}
          {state.step === 3 && (
            <>
              <button
                type="button"
                className="btn cx"
                style={{ flex: 1 }}
                onClick={() => goToStep(4)}
              >
                ✕ إلغاء
              </button>
              <button
                type="button"
                className="btn p"
                style={{ flex: 1 }}
                onClick={() => {
                  snap3Ref.current = null;
                  setSnap3(null);
                  toast("تم حفظ التغييرات");
                  goToStep(4, false, undefined, true);
                }}
              >
                ✓ تم
              </button>
            </>
          )}
          {state.step === 2 && (
            <>
              <button type="button" className="btn back" onClick={() => goToStep(1)}>
                → النوع
              </button>
              <button type="button" className="btn p" onClick={() => goToStep(4)}>
                عرض النتائج ←
              </button>
            </>
          )}
        </div>
      )}

      {/* خلفية النوافذ */}
      {(pendChoices || saveOpen || altOpen) && (
        <div
          id="dshade"
          onClick={() => {
            setPendChoices(null);
            pendAfterRef.current = null;
            setSaveOpen(false);
            setAltOpen(false);
          }}
        />
      )}

      {/* نافذة تأكيد القياسات الناقصة */}
      {pendChoices && (() => {
        const emptyPendIndices = pendChoices
          .filter((x) => isRowBlank(x.r))
          .map((x) => x.i);

        const getChoiceState = (x: PendingChoice) => {
          const dL = pL(x.r) ? null : num(x.draftL);
          const dW = pW(x.r) ? null : num(x.draftW);
          const f = UF[state.unit];
          let inputErr = "";
          if ((dL !== null && isNaN(dL)) || (dW !== null && isNaN(dW))) {
            inputErr = "اكتب أرقاماً فقط";
          } else if (
            (dL !== null && (!isFinite(dL) || dL <= 0)) ||
            (dW !== null && (!isFinite(dW) || dW <= 0))
          ) {
            inputErr = "القياس يجب أن يكون أكبر من صفر";
          } else if ((dL !== null && dL * f > MAXCM) || (dW !== null && dW * f > MAXCM)) {
            inputErr = "القياس كبير جداً (الحد الأقصى 50 متراً)";
          }
          const isComplete =
            !inputErr && (pL(x.r) || dL !== null) && (pW(x.r) || dW !== null);
          return { dL, dW, inputErr, isComplete };
        };

        const applyChoiceToRow = (row: WindowRow, x: PendingChoice): WindowRow => ({
          ...row,
          L: pL(x.r) ? row.L : x.draftL,
          W: pW(x.r) ? row.W : x.draftW,
        });

        const commitAndAdvance = (updatedRows: WindowRow[]) => {
          setState((prev) => ({ ...prev, rows: updatedRows }));
          const rem = findNeedChoice(updatedRows, pendChoices);
          if (rem.length) {
            setPendChoices(rem);
          } else {
            const cb = pendAfterRef.current;
            pendAfterRef.current = null;
            setPendChoices(null);
            if (cb) window.setTimeout(() => cb(updatedRows), 100);
          }
        };

        const saveSingleRow = (x: PendingChoice) => {
          const { inputErr, isComplete } = getChoiceState(x);
          if (inputErr) {
            toast(inputErr);
            return;
          }
          if (!isComplete) {
            toast("أدخل الطول والعرض المطلوبين أولاً");
            return;
          }
          const updatedRows = state.rows.map((row, idx) =>
            idx === x.i ? applyChoiceToRow(row, x) : row
          );
          commitAndAdvance(updatedRows);
        };

        const saveAllCompleteDrafts = () => {
          const byIdx = new Map<number, PendingChoice>();
          for (const x of pendChoices) {
            const { dL, dW, inputErr } = getChoiceState(x);
            if (inputErr) {
              toast(`شباك ${x.i + 1}: ${inputErr}`);
              return;
            }
            if ((!pL(x.r) && dL !== null) || (!pW(x.r) && dW !== null)) {
              byIdx.set(x.i, x);
            }
          }
          if (!byIdx.size) {
            toast("أدخل قياساً واحداً على الأقل أو اختر تركه فارغاً/حذفه");
            return;
          }
          const updatedRows = state.rows.map((row, idx) => {
            const item = byIdx.get(idx);
            return item ? applyChoiceToRow(row, item) : row;
          });
          commitAndAdvance(updatedRows);
        };

        const hasAnyDraft = pendChoices.some((x) => {
          const { dL, dW, inputErr } = getChoiceState(x);
          return !inputErr && ((x.ml && dL !== null) || (x.mw && dW !== null));
        });

        return (
          <div className="dlg">
            <div style={{ fontSize: 19, fontWeight: "bold" }}>⚠️ تأكد قبل المتابعة</div>
            <p style={{ margin: "6px 0 0", color: "var(--mut)", fontSize: 15 }}>
              أدخل الطول والعرض مباشرة هنا، أو اتركه فارغاً عمداً، أو احذفه:
            </p>
            {emptyPendIndices.length >= 2 && (
              <button
                type="button"
                className="pq-bulk"
                onClick={() => handleRemovePendingRows(new Set(emptyPendIndices))}
              >
                🗑 حذف جميع الصفوف الفارغة ({emptyPendIndices.length})
              </button>
            )}
            {pendChoices.map((x) => {
              const { dL, dW, inputErr, isComplete } = getChoiceState(x);
              const nm = x.ml && x.mw ? "الطول والعرض" : x.ml ? "الطول" : "العرض";
              return (
                <div key={x.i} className="pq">
                  <b>
                    شباك {x.i + 1}
                    {x.r.n ? ` · ${x.r.n}` : ""} —{" "}
                    {isComplete ? (
                      <em className="ok">✓ جاهز</em>
                    ) : (
                      <em>{nm} فارغ</em>
                    )}
                  </b>

                  <div className="pq-inps">
                    <label className={x.ml && dL === null ? "need" : ""}>
                      <span>الطول ({UN[state.unit]})</span>
                      <input
                        inputMode="decimal"
                        enterKeyHint="next"
                        disabled={pL(x.r)}
                        placeholder={pL(x.r) ? "⏳ غير جاهز" : "أدخل الطول"}
                        value={pL(x.r) ? "" : x.draftL}
                        onChange={(e) => {
                          const v = mfixValue(e.target.value, state.unit);
                          setPendChoices((prev) =>
                            prev ? prev.map((it) => (it.i === x.i ? { ...it, draftL: v } : it)) : null
                          );
                        }}
                      />
                    </label>
                    <label className={x.mw && dW === null ? "need" : ""}>
                      <span>العرض ({UN[state.unit]})</span>
                      <input
                        inputMode="decimal"
                        enterKeyHint="done"
                        disabled={pW(x.r)}
                        placeholder={pW(x.r) ? "⏳ غير جاهز" : "أدخل العرض"}
                        value={pW(x.r) ? "" : x.draftW}
                        onChange={(e) => {
                          const v = mfixValue(e.target.value, state.unit);
                          setPendChoices((prev) =>
                            prev ? prev.map((it) => (it.i === x.i ? { ...it, draftW: v } : it)) : null
                          );
                        }}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && isComplete) {
                            e.preventDefault();
                            saveSingleRow(x);
                          }
                        }}
                      />
                    </label>
                  </div>

                  {inputErr && (
                    <div className="wmsg" style={{ marginBottom: 8 }}>
                      {inputErr}
                    </div>
                  )}

                  <div className="pq-acts">
                    <button
                      type="button"
                      className="w1"
                      onClick={() => saveSingleRow(x)}
                    >
                      ✓ حفظ القياس
                    </button>
                    <button
                      type="button"
                      className="w2"
                      onClick={() => {
                        if (inputErr) {
                          toast(inputErr);
                          return;
                        }
                        const a = pL(x.r) || dL === null;
                        const w = pW(x.r) || dW === null;
                        const newP: PendingMode = a && w ? "B" : a ? "L" : "W";
                        const updatedRows = state.rows.map((row, idx) =>
                          idx === x.i ? { ...applyChoiceToRow(row, x), p: newP } : row
                        );
                        commitAndAdvance(updatedRows);
                      }}
                    >
                      ⏳ اتركه فارغاً عمداً
                    </button>
                  </div>
                  <button
                    type="button"
                    className="w3"
                    onClick={() => handleRemovePendingRows(new Set([x.i]))}
                  >
                    🗑 لا أحتاجه (احذفه)
                  </button>
                </div>
              );
            })}
            {hasAnyDraft && (
              <button
                type="button"
                className="btn p sm"
                style={{ marginTop: 12, justifyContent: "center" }}
                onClick={saveAllCompleteDrafts}
              >
                ✓ متابعة بالقياسات المدخلة
              </button>
            )}
            <button
              type="button"
              className="btn sm"
              style={{ marginTop: 8, justifyContent: "center" }}
              onClick={() => {
                setPendChoices(null);
                pendAfterRef.current = null;
              }}
            >
              رجوع للقياسات
            </button>
          </div>
        );
      })()}

      {/* نافذة الحفظ والمشاركة */}
      {saveOpen && (
        <div className="dlg">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <b style={{ fontSize: 19 }}>💾 حفظ ومشاركة النتائج</b>
            <button
              type="button"
              className="btn sm"
              style={{ width: "auto", margin: 0 }}
              onClick={() => setSaveOpen(false)}
            >
              ✕
            </button>
          </div>

          <label className="svl">
            <span>
              اسم المشروع <small>(يُكتب في الترويسة ويُسمّى به الملف)</small>
            </span>
            <input
              id="pjName"
              maxLength={60}
              placeholder="مثال: فيلا أبو أحمد"
              autoComplete="off"
              value={pjName}
              onChange={(e) => setPjName(e.target.value)}
            />
          </label>

          <div className="svinfo">
            تلقائياً في الترويسة: {res.type} · <bdi dir="ltr">{dstr()}</bdi> · {res.data.length} شباك
            {res.hasQ ? ` (${res.qtotal} قطعة)` : ""}
          </div>

          <div className="svgrid">
            <button
              type="button"
              className="btn p"
              onClick={() => {
                setSaveOpen(false);
                handleExportPdf("save");
              }}
            >
              📄 حفظ PDF
            </button>
            <button
              type="button"
              className="btn p"
              onClick={() => {
                setSaveOpen(false);
                handleExportXlsx("save");
              }}
            >
              📊 حفظ إكسل
            </button>
            <button
              type="button"
              className="btn g"
              onClick={() => {
                setSaveOpen(false);
                handleExportPdf("share");
              }}
            >
              📤 مشاركة PDF
            </button>
            <button
              type="button"
              className="btn g"
              onClick={() => {
                setSaveOpen(false);
                handleExportXlsx("share");
              }}
            >
              📤 مشاركة إكسل
            </button>
          </div>

          <details className="svmore">
            <summary>خيارات أخرى</summary>
            <button
              type="button"
              className="btn g"
              onClick={() => {
                if (res.table.length) {
                  copyText(
                    getTabLines()
                      .map((r) => r.join("\t"))
                      .join("\n"),
                    "تم النسخ. الصقها في إكسل"
                  );
                }
              }}
            >
              📋 نسخ للإكسل
            </button>
            <button
              type="button"
              className="btn g"
              onClick={async () => {
                const txt = getPlainTextResults();
                if (!txt) return;
                if (navigator.share) {
                  try {
                    await navigator.share({ text: txt });
                    return;
                  } catch {
                    // fallback copy
                  }
                }
                copyText(txt, "تم نسخ النتائج كنص");
              }}
            >
              💬 إرسال النتائج كنص
            </button>
            <button
              type="button"
              className="btn w"
              onClick={() => {
                if (!res.table.length) return;
                const csv =
                  "\uFEFF" +
                  getTabLines()
                    .map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
                    .join("\r\n");
                saveBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), `${fnBase()}.csv`, "text/csv", "save");
              }}
            >
              📄 ملف نصي CSV
            </button>
            <button
              type="button"
              className="btn w"
              onClick={() => {
                setSaveOpen(false);
                setAltOpen(true);
              }}
            >
              🆘 لم يُحفظ الملف؟ جرّب طرقاً بديلة
            </button>
          </details>
        </div>
      )}

      {/* نافذة الطرق البديلة */}
      {altOpen && (
        <div className="dlg">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <b style={{ fontSize: 19 }}>🆘 طرق بديلة للحصول على النتائج</b>
            <button
              type="button"
              className="btn sm"
              style={{ width: "auto", margin: 0 }}
              onClick={() => setAltOpen(false)}
            >
              ✕
            </button>
          </div>
          <p className="lead">بعض المتصفحات تمنع حفظ الملفات تلقائياً. جرّب إحدى هذه الطرق:</p>
          <button
            type="button"
            className="btn p"
            onClick={async () => {
              if (!lastSavedBlobRef.current) {
                toast("احفظ الملف أولاً ثم افتحه");
                return;
              }
              setAltOpen(false);
              await saveBlob(
                lastSavedBlobRef.current.blob,
                lastSavedBlobRef.current.filename,
                lastSavedBlobRef.current.mime,
                "save"
              );
            }}
          >
            📂 إعادة تنزيل الملف
          </button>
          <button
            type="button"
            className="btn g"
            onClick={() => {
              if (res.table.length) {
                copyText(
                  getTabLines()
                    .map((r) => r.join("\t"))
                    .join("\n"),
                  "تم النسخ. الصقها في إكسل أو Google Sheets"
                );
              }
            }}
          >
            📋 نسخ الجدول ثم لصقه في إكسل
          </button>
          <button
            type="button"
            className="btn g"
            onClick={() => {
              const t = getPlainTextResults();
              if (t) {
                const a = document.createElement("a");
                a.href = `https://wa.me/?text=${encodeURIComponent(t)}`;
                a.target = "_blank";
                a.rel = "noopener";
                document.body.appendChild(a);
                a.click();
                a.remove();
              }
            }}
          >
            💬 إرسال النتائج بالواتساب
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => setAltTextOpen((v) => !v)}
          >
            👁 عرض النتائج كنص لنسخه يدوياً
          </button>
          {altTextOpen && <textarea readOnly value={getPlainTextResults()} />}
        </div>
      )}

      {/* نافذة نسخ شباك محدد حسب رقمه */}
      {copyByNumOpen && (
        <div className="pop-modal" onClick={() => setCopyByNumOpen(false)}>
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <div className="cnhead">
              <span className="cnico">
                <svg viewBox="0 0 48 48" width="32" height="32" aria-hidden="true">
                  <rect
                    x="13"
                    y="13"
                    width="26"
                    height="26"
                    rx="6"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="3"
                    opacity=".55"
                  />
                  <rect
                    x="7"
                    y="7"
                    width="26"
                    height="26"
                    rx="6"
                    fill="currentColor"
                    opacity=".28"
                    stroke="currentColor"
                    strokeWidth="3.5"
                  />
                  <path
                    d="M20 12v16M12 20h16"
                    stroke="currentColor"
                    strokeWidth="4"
                    strokeLinecap="round"
                  />
                </svg>
              </span>
              <div>
                <h2 style={{ margin: 0 }}>نسخ شباك محدد</h2>
                <p className="wl" style={{ margin: "2px 0 0" }}>
                  اكتب رقم الشباك (من 1 إلى {state.rows.reduce((acc, r, idx) => (r.L || r.W || r.p || r.n ? idx + 1 : acc), 0)})، وسيُضاف نسخة منه في آخر القائمة.
                </p>
              </div>
            </div>

            <input
              ref={copyNumInputRef}
              type="tel"
              inputMode="numeric"
              placeholder="رقم الشباك (مثال: 3)"
              maxLength={3}
              value={copyNumVal}
              onChange={(e) =>
                setCopyNumVal(
                  e.target.value
                    .replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)))
                    .replace(/[^0-9]/g, "")
                )
              }
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  confirmCopyByNumber();
                }
              }}
              style={{
                width: "100%",
                height: 54,
                textAlign: "center",
                fontSize: 22,
                fontWeight: "bold",
                borderRadius: 16,
                background: "var(--bg)",
                marginTop: 8,
                boxShadow: "var(--cl-press)",
              }}
            />

            {copyNumVal.trim() !== "" && (
              copyTargetIdx >= 0 ? (
                <div className="cnprev">
                  <b>شباك {copyTargetIdx + 1}</b>
                  {state.rows[copyTargetIdx].n && <small> · {state.rows[copyTargetIdx].n}</small>}
                  {qOf(state.rows[copyTargetIdx]) > 1 && (
                    <small> — العدد {qOf(state.rows[copyTargetIdx])}</small>
                  )}
                  <div className="cnv">
                    <span>
                      <small>الطول ({UN[state.unit]})</small>
                      {pL(state.rows[copyTargetIdx]) ? "⏳" : state.rows[copyTargetIdx].L || "⏳"}
                    </span>
                    <span>
                      <small>العرض ({UN[state.unit]})</small>
                      {pW(state.rows[copyTargetIdx]) ? "⏳" : state.rows[copyTargetIdx].W || "⏳"}
                    </span>
                  </div>
                </div>
              ) : (
                <div className="cnprev bad">
                  {(() => {
                    const pNum = num(copyNumVal.trim());
                    const idxCheck = pNum !== null && !isNaN(pNum) && Number.isInteger(pNum) ? pNum - 1 : -1;
                    if (idxCheck >= 0 && idxCheck < state.rows.length && isRowBlank(state.rows[idxCheck])) {
                      return `الصف رقم ${copyNumVal.trim()} فارغ (لا يحتوي على قياسات)`;
                    }
                    return `لا يوجد شباك مكتمل بالرقم ${copyNumVal.trim()}`;
                  })()}
                </div>
              )
            )}

            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button
                type="button"
                className="btn sm"
                style={{ flex: 1 }}
                onClick={() => setCopyByNumOpen(false)}
              >
                إلغاء
              </button>
              <button
                type="button"
                className="btn sm p"
                style={{ flex: 1 }}
                onClick={confirmCopyByNumber}
              >
                نسخ الشباك ✓
              </button>
            </div>
          </div>
        </div>
      )}

      {/* نافذة استيراد من Excel */}
      {xlImportState && (
        <div className="pop-modal" onClick={() => setXlImportState(null)}>
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <h2 style={{ marginBottom: 6 }}>📊 استيراد من Excel</h2>
            <p className="wl" style={{ marginTop: 0 }}>
              الملف: <b>{xlImportState.fileName}</b>
              {xlImportState.parsed.projectName ? ` · المشروع: ${xlImportState.parsed.projectName}` : ""}
              . اختر ما تريد أخذه منه:
            </p>

            {xlImportState.parsed.rows && (
              <>
                <label className="xlc">
                  <input
                    type="checkbox"
                    checked={xlImportState.useRows}
                    onChange={(ev) =>
                      setXlImportState((prev) =>
                        prev ? { ...prev, useRows: ev.target.checked, err: "" } : null
                      )
                    }
                  />
                  <span>
                    📏 القياسات: <b>{xlImportState.parsed.rows.length}</b> شباك
                    {(() => {
                      const ready = xlImportState.parsed.rows!.filter(
                        (r) => r.l !== null && r.w !== null
                      ).length;
                      const pendCnt = xlImportState.parsed.rows!.length - ready;
                      const parts: string[] = [];
                      if (pendCnt > 0) parts.push(`منها ${pendCnt} غير جاهز`);
                      if (xlImportState.parsed.hasNames) parts.push("شاملة الأسماء");
                      return parts.length ? ` (${parts.join(" · ")})` : "";
                    })()}
                  </span>
                </label>

                {xlImportState.useRows && (
                  <div className="xlsub">
                    <label className="xlu-lbl">
                      وحدة القياسات في الملف
                      <select
                        value={xlImportState.unit}
                        onChange={(ev) =>
                          setXlImportState((prev) =>
                            prev ? { ...prev, unit: ev.target.value as UnitCode } : null
                          )
                        }
                      >
                        {(["m", "cm", "mm"] as UnitCode[]).map((u) => (
                          <option key={u} value={u}>
                            {UFULL[u]}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="xlr">
                      <input
                        type="radio"
                        name="xlmode"
                        checked={xlImportState.mode === "add"}
                        onChange={() =>
                          setXlImportState((prev) => (prev ? { ...prev, mode: "add" } : null))
                        }
                      />
                      <span>إضافتها بعد شبابيكي الحالية</span>
                    </label>
                    <label className="xlr">
                      <input
                        type="radio"
                        name="xlmode"
                        checked={xlImportState.mode === "rep"}
                        onChange={() =>
                          setXlImportState((prev) => (prev ? { ...prev, mode: "rep" } : null))
                        }
                      />
                      <span>استبدال شبابيكي الحالية بها</span>
                    </label>
                  </div>
                )}
              </>
            )}

            {xlImportState.parsed.settings && (
              <label className="xlc">
                <input
                  type="checkbox"
                  checked={xlImportState.useSettings}
                  onChange={(ev) =>
                    setXlImportState((prev) =>
                      prev ? { ...prev, useSettings: ev.target.checked, err: "" } : null
                    )
                  }
                />
                <span>
                  ⚙ الإعدادات: النوع «{String(xlImportState.parsed.settings.type || "")}» والعلاقات
                  والتخصيص <small>(تستبدل علاقاتك الحالية)</small>
                </span>
              </label>
            )}

            {xlImportState.err && <div className="xl-err">{xlImportState.err}</div>}

            <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
              <button
                type="button"
                className="btn sm"
                style={{ flex: 1 }}
                onClick={() => setXlImportState(null)}
              >
                إلغاء
              </button>
              <button
                type="button"
                className="btn sm p"
                style={{ flex: 1 }}
                onClick={executeXlsxImport}
              >
                استيراد ✓
              </button>
            </div>
          </div>
        </div>
      )}

      <input
        ref={xlFileRef}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        hidden
        onChange={handleXlsxFileSelected}
      />

      {/* نافذة نقل أو تسمية الشباك عند الضغط على رقمه */}
      {badgeModalIdx !== null && state.rows[badgeModalIdx] && (
        <div className="pop-modal" onClick={() => setBadgeModalIdx(null)}>
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <h2>🔢 نقل أو تسمية شباك #{badgeModalIdx + 1}</h2>
            <p className="wl">
              اكتب رقماً (1–300) لنقل الشباك إليه (يتبدّل مع الشباك الموجود فيه)، أو اكتب اسماً لتسميته (يظهر في عمود «الاسم» ويُبحث به):
            </p>
            <input
              ref={badgeInputRef}
              type="text"
              placeholder="رقم للانتقال (مثل: 3) أو اسم (مثل: مطبخ)"
              maxLength={60}
              value={badgeInputVal}
              onChange={(e) => setBadgeInputVal(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  confirmBadgeAction();
                }
              }}
              style={{
                width: "100%",
                height: 54,
                textAlign: "center",
                fontSize: 18,
                fontWeight: "bold",
                borderRadius: 16,
                background: "var(--bg)",
                marginTop: 12,
                padding: "0 12px",
                boxShadow: "var(--cl-press)",
              }}
            />

            <div className="svinfo" style={{ marginTop: 12 }}>
              {(() => {
                const rawTrim = clean(badgeInputVal).trim();
                const normDigits = rawTrim.replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)));
                if (/^[-+]?[0-9]+([.,٫][0-9]+)?$/.test(normDigits) && !/^[0-9]+$/.test(normDigits)) {
                  return "⚠️ اكتب رقماً صحيحاً موجباً للنقل، أو اسماً يحتوي حروفاً";
                }
                if (/^[0-9]+$/.test(normDigits)) {
                  const N = parseInt(normDigits, 10);
                  if (N < 1 || N > 300) return "⚠️ الرقم المسموح بين 1 و300";
                  if (N === badgeModalIdx + 1) return "الرقم الحالي للشباك — اكتب رقماً آخر لنقله أو اسماً لتسميته";
                  const targetRow = state.rows[N - 1];
                  const hasOther =
                    targetRow &&
                    (num(targetRow.L) !== null ||
                      num(targetRow.W) !== null ||
                      Boolean(targetRow.p) ||
                      Boolean(targetRow.n));
                  return hasOther
                    ? `سيتم نقل الشباك #${badgeModalIdx + 1} إلى الرقم #${N} وتبديله مع الشباك الموجود فيه (${targetRow.L || "⏳"} × ${targetRow.W || "⏳"})`
                    : `سيتم نقل الشباك #${badgeModalIdx + 1} إلى الرقم #${N}`;
                }
                return rawTrim
                  ? `سيتم حفظ اسم الشباك: «${rawTrim.slice(0, 60)}»`
                  : "اتركه فارغاً لمسح اسم الشباك";
              })()}
            </div>

            <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
              <button
                type="button"
                className="btn sm"
                style={{ flex: 1 }}
                onClick={() => setBadgeModalIdx(null)}
              >
                إلغاء
              </button>
              {state.rows[badgeModalIdx].n && (
                <button
                  type="button"
                  className="btn sm w"
                  onClick={() => {
                    updateRow(badgeModalIdx, { n: "" });
                    setBadgeModalIdx(null);
                    toast(`تم مسح اسم الشباك #${badgeModalIdx + 1}`);
                  }}
                >
                  مسح الاسم
                </button>
              )}
              <button
                type="button"
                className="btn sm p"
                style={{ flex: 1 }}
                onClick={confirmBadgeAction}
              >
                تطبيق ✓
              </button>
            </div>
          </div>
        </div>
      )}

      {/* نافذة الترحيب والمساعدة */}
      {welOpen && (
        <div
          id="wel"
          onClick={() => {
            setWelOpen(false);
            try {
              localStorage.setItem(WELCOME_KEY, "1");
            } catch {
              // ignore
            }
          }}
        >
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <h2>👋 أهلاً بك</h2>
            <p className="wl">احسب كل مقاسات الشباك أو الواجهة من الطول والعرض فقط، في ثلاث خطوات:</p>
            <ol>
              <li>
                <b>اختر النوع</b> (مثل شباك 3 مجاري نظامي أو شباك 2.5 نظامي).
              </li>
              <li>
                <b>اكتب الطول والعرض</b> لكل شباك، واختر الوحدة (م / سم / مم). الطول = الارتفاع، والعرض = من جنب لجنب.
              </li>
              <li>
                <b>النتائج</b> تظهر جاهزة: انسخ أي رقم بالضغط عليه، أو احفظها PDF نصي خفيف أو Excel.
              </li>
            </ol>
            <p className="wn">
              💡 من «تخصيص» تقدر تخفي أو تُظهر خانات <b>المساحة</b> أو تضيف زيادة أو نقصاناً، ومن «متقدمة» تغيّر العلاقات وتضيف خانة جديدة.
            </p>
            <button
              type="button"
              className="btn p"
              onClick={() => {
                setWelOpen(false);
                try {
                  localStorage.setItem(WELCOME_KEY, "1");
                } catch {
                  // ignore
                }
              }}
            >
              ابدأ ✓
            </button>
            <p className="wh">تجد هذا الشرح دائماً من زر «؟» أعلى الصفحة.</p>
          </div>
        </div>
      )}

      {/* نافذة تأكيد عام (بديل مخصص لـ confirm) */}
      {confirmDlg && (
        <div className="pop-modal" onClick={() => setConfirmDlg(null)}>
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <h2>{confirmDlg.title}</h2>
            <p className="wl" style={{ marginTop: 6 }}>
              {confirmDlg.message}
            </p>
            <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
              <button
                type="button"
                className="btn sm"
                style={{ flex: 1 }}
                onClick={() => setConfirmDlg(null)}
              >
                {confirmDlg.cancelLabel || "تراجع"}
              </button>
              <button
                type="button"
                className={`btn sm ${confirmDlg.danger ? "r" : "p"}`}
                style={{ flex: 1 }}
                onClick={() => {
                  const fn = confirmDlg.onConfirm;
                  setConfirmDlg(null);
                  fn();
                }}
              >
                {confirmDlg.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* نافذة تحذير الخروج عند الضغط على زر الرجوع في الشاشة الأولى */}
      {exitPromptOpen && (
        <div
          className="pop-modal"
          onClick={() => {
            setExitPromptOpen(false);
            window.history.pushState({ step: state.step }, "");
          }}
        >
          <div className="wbox" onClick={(e) => e.stopPropagation()}>
            <h2>⚠️ هل تريد الخروج؟</h2>
            <p className="wl" style={{ marginTop: 6 }}>
              بياناتك وقياساتك محفوظة تلقائياً على جهازك. هل تريد الخروج من التطبيق؟
            </p>
            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                type="button"
                className="btn sm p"
                style={{ flex: 1 }}
                onClick={() => {
                  setExitPromptOpen(false);
                  window.history.pushState({ step: state.step }, "");
                }}
              >
                البقاء في التطبيق
              </button>
              <button
                type="button"
                className="btn sm r"
                style={{ flex: 1 }}
                onClick={() => {
                  setExitPromptOpen(false);
                  window.history.back();
                }}
              >
                خروج
              </button>
            </div>
          </div>
        </div>
      )}

      {/* شريط إشعار التنزيل */}
      {openBar && (
        <div id="openbar" className={openBar.web ? "web" : ""}>
          <span id="obTxt">{openBar.txt}</span>
          <button
            type="button"
            id="obGo"
            onClick={async () => {
              const winAny = window as unknown as {
                Capacitor?: {
                  isNativePlatform?: () => boolean;
                  Plugins?: {
                    SaveFile?: { open: (o: unknown) => Promise<void> };
                  };
                };
              };
              const C = winAny.Capacitor;
              if (openBar && !openBar.web && openBar.uri && C?.Plugins?.SaveFile?.open) {
                try {
                  await C.Plugins.SaveFile.open({ uri: openBar.uri, mime: openBar.mime });
                  setOpenBar(null);
                  return;
                } catch {
                  // fallback below
                }
              }
              if (lastSavedBlobRef.current) {
                if (C?.isNativePlatform?.()) {
                  await saveBlob(
                    lastSavedBlobRef.current.blob,
                    lastSavedBlobRef.current.filename,
                    lastSavedBlobRef.current.mime,
                    "share"
                  );
                } else {
                  webDownload(lastSavedBlobRef.current.blob, lastSavedBlobRef.current.filename);
                }
              }
              setOpenBar(null);
            }}
          >
            📂 فتح الملف
          </button>
          {openBar.web && (
            <button
              type="button"
              id="obAlt"
              onClick={() => {
                setOpenBar(null);
                setAltOpen(true);
              }}
            >
              طرق بديلة
            </button>
          )}
          <button type="button" id="obClose" aria-label="إغلاق" onClick={() => setOpenBar(null)}>
            ✕
          </button>
        </div>
      )}

      {fallBarVisible && (
        <div id="fallbar">
          <span>لم يبدأ حفظ الملف؟</span>
          <button
            type="button"
            onClick={() => {
              setFallBarVisible(false);
              setAltOpen(true);
            }}
          >
            طرق بديلة
          </button>
        </div>
      )}

      <div id="toast" className={toastOn ? "on" : ""}>
        {toastMsg}
      </div>
    </>
  );
}
