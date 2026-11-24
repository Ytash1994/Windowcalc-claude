import type {
  AppState,
  AreaTotalItem,
  BuildResult,
  ComputedWindow,
  CustomField,
  DuplicateGroup,
  ExportStructUnit,
  LayoutConfig,
  RelConfig,
  ResultUnit,
  RoundConfig,
  UnitCode,
  WindowRow,
} from "../types";

export const AR_DIGITS = "٠١٢٣٤٥٦٧٨٩";
export const MAXCM = 5000;
export const MAXQ = 9999;

export const KEY = "winCalcV3";
export const RKEY = "winCalcRelV1";
export const CKEY = "winCalcCustomV1";
export const PKEY = "winCalcProjectsV1";
export const THEME_KEY = "winCalcTheme";
export const PAL_KEY = "winCalcPal";
export const FONT_KEY = "winCalcFontScale";
export const WELCOME_KEY = "winCalcWelcome";

export const SIG_NAME = 'يوسف القرعان " أبو فارس "';
export const SIG_PHONE = "+962795467235";

export const BASE_FIELDS: [string, string, string][] = [
  ["M", "حلق", "الطول (جنب الحلق)"],
  ["N", "حلق", "العرض (أرضية الحلق)"],
  ["L", "", "زرفيل سكين"],
  ["K", "", "أرضية درفة"],
  ["I", "منخل", "طول"],
  ["J", "منخل", "عرض"],
  ["G", "زجاج", "طول"],
  ["H", "زجاج", "عرض"],
  ["AG", "زجاج", "مساحة الزجاج"],
  ["FF", "", "F"],
  ["D", "", "U"],
  ["E", "صندوق", "طول"],
  ["F", "صندوق", "عرض"],
  ["A", "", "عامود"],
  ["KL", "كشفة", "طول"],
  ["KW", "كشفة", "عرض"],
  ["B", "حصيرة", "طول"],
  ["C", "حصيرة", "عرض"],
  ["AW", "", "مساحة الشباك"],
];

export const AREA_KEYS = new Set(["AW", "AG"]);

export type RelTuple = [string, number, number | null, string?, string?, number?, number?];

export const BASE_REL: Record<string, RelTuple> = {
  M: ["LEN", 1, null],
  N: ["WID", 1, -2.6],
  L: ["M", 1, -7.8],
  K: ["WID", 0.5, 1],
  I: ["L", 1, 0],
  J: ["K", 1, 0],
  G: ["L", 1, -9.3],
  H: ["K", 1, -7.8],
  AG: ["G", 1, 0, "H", "*"],
  FF: ["M", 1, -5.5],
  D: ["M", 1, 6.5],
  F: ["WID", 1, 0],
  A: ["WID", 1, -9.5],
  KL: ["LEN", 1, 9.2],
  KW: ["WID", 1, 9.2],
  B: ["M", 1, 10],
  C: ["WID", 1, -11.5],
  AW: ["LEN", 1, 0, "WID", "*", 10, 10],
};

export const BASE_DIV: Record<string, number> = { K: 2 };

export const BASE_ORDER = [
  "M",
  "N",
  "L",
  "K",
  "I",
  "J",
  "G",
  "H",
  "FF",
  "D",
  "E",
  "F",
  "A",
  "KL",
  "KW",
  "B",
  "C",
  "AG",
  "AW",
];

export interface TypeDef {
  mAdd: number;
  eConst: number | null;
  fields: string[];
  rel?: Record<string, RelTuple>;
}

export const BUILTIN_TYPE_NAMES = [
  "شباك 3 مجاري نظامي",
  "شباك 2.5 نظامي",
  "واجهات 2.5 بدون صندوق",
  "شباك 3 مجاري بدون صندوق",
];

export const DEF_TYPE_DEFS: Record<string, TypeDef> = {
  "شباك 3 مجاري نظامي": {
    mAdd: -29.5,
    eConst: 29.5,
    fields: ["M", "N", "L", "K", "I", "J", "G", "H", "AG", "FF", "D", "E", "F", "A", "KL", "KW", "B", "C", "AW"],
  },
  "شباك 2.5 نظامي": {
    mAdd: -29.5,
    eConst: 29.5,
    fields: ["M", "N", "L", "K", "I", "J", "G", "H", "AG", "D", "E", "F", "A", "KL", "KW", "B", "C", "AW"],
    rel: {
      L: ["M", 1, -7.5],
      K: ["N", 0.5, -0.1],
      H: ["K", 1, -5.8],
    },
  },
  "واجهات 2.5 بدون صندوق": {
    mAdd: 0,
    eConst: null,
    fields: ["M", "N", "L", "K", "I", "J", "G", "H", "AG", "FF", "KL", "KW", "AW"],
  },
  "شباك 3 مجاري بدون صندوق": {
    mAdd: 0,
    eConst: null,
    fields: ["M", "N", "L", "K", "I", "J", "G", "H", "AG", "FF", "AW"],
  },
};

export const TYPE_SUBS: Record<string, { sub: string; iconKey: "a" | "b" | "c" | "d" | "e" }> = {
  // The four built-in types show only their name (empty sub = no description line)
  "شباك 3 مجاري نظامي": { sub: "", iconKey: "a" },
  "شباك 2.5 نظامي": { sub: "", iconKey: "c" },
  "واجهات 2.5 بدون صندوق": { sub: "", iconKey: "b" },
  "شباك 3 مجاري بدون صندوق": { sub: "", iconKey: "d" },
};

export const BASE_HUE: Record<string, number> = {
  حلق: 48,
  L: 0,
  K: 305,
  منخل: 28,
  زجاج: 188,
  FF: 240,
  D: 145,
  صندوق: 95,
  A: 275,
  كشفة: 212,
  حصيرة: 342,
  AW: 170,
};

export const CHUE = [165, 15, 295, 95, 345, 200];

export const UF: Record<UnitCode, number> = { m: 100, cm: 1, mm: 0.1 };
export const UN: Record<UnitCode, string> = { m: "م", cm: "سم", mm: "مم" };
export const UFULL: Record<UnitCode, string> = {
  m: "متر (م)",
  cm: "سانتي ميتر (سم)",
  mm: "ملي ميتر (مم)",
};
export const UAREA: Record<UnitCode, string> = {
  m: "متر مربع (م²)",
  cm: "سانتي ميتر مربع (سم²)",
  mm: "ملي ميتر مربع (مم²)",
};
export const UAS: Record<UnitCode, string> = { m: "م²", cm: "سم²", mm: "مم²" };

export const TPL: [string, Partial<RelConfig>][] = [
  ["الطول − رقم", { s: "LEN", s2: "", m: "1", dv: false, neg: true, a: "10", op: "+" }],
  ["العرض − رقم", { s: "WID", s2: "", m: "1", dv: false, neg: true, a: "10", op: "+" }],
  ["الطول + رقم", { s: "LEN", s2: "", m: "1", dv: false, neg: false, a: "10", op: "+" }],
  ["العرض + رقم", { s: "WID", s2: "", m: "1", dv: false, neg: false, a: "10", op: "+" }],
  ["الطول ÷ 2", { s: "LEN", s2: "", m: "2", dv: true, neg: false, a: "0", op: "+" }],
  ["العرض ÷ 2", { s: "WID", s2: "", m: "2", dv: true, neg: false, a: "0", op: "+" }],
  ["الطول × العرض", { s: "LEN", s2: "WID", m: "1", dv: false, neg: false, a: "0", op: "*", n1: false, a1: "0", n2: false, a2: "0" }],
  ["الطول + العرض", { s: "LEN", s2: "WID", m: "1", dv: false, neg: false, a: "0", op: "+", n1: false, a1: "0", n2: false, a2: "0" }],
];

export const r6 = (x: number) => Math.round(x * 1e6) / 1e6;
export const clean = (v: unknown) => String(v == null ? "" : v).replace(/[<>&"]/g, "");
export const isCid = (id: string) => /^c\d+$/.test(id);
export const isCu = (id: string) => isCid(id) || String(id).startsWith("cg:");
export const GN = (id: string) =>
  String(id).startsWith("cg:") ? String(id).split(":").slice(2).join(":") : id;
export const BUILTIN_GROUPS = new Set(["حلق", "منخل", "زجاج", "صندوق", "كشفة", "حصيرة"]);
export const cUnit = (typeName: string, c: CustomField) =>
  c.group ? (BUILTIN_GROUPS.has(c.group) ? c.group : `cg:${typeName}:${c.group}`) : c.id;

/** Strict numeric parser: rejects scientific notation like 1e3, hex, Infinity */
export function num(s: unknown): number | null {
  const str = String(s == null ? "" : s)
    .trim()
    .replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)))
    .replace(/[٫،,]/g, ".")
    .replace(/\s/g, "");
  if (str === "" || str === "+" || str === "-" || str === ".") return null;
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(str)) return NaN;
  const v = Number(str);
  return isNaN(v) || !isFinite(v) ? NaN : v;
}

export function mfixValue(raw: string, unit: UnitCode): string {
  if (unit !== "m") return raw.replace(/[eE]/g, "");
  let v = raw
    .replace(/[٠-٩]/g, (d) => String(AR_DIGITS.indexOf(d)))
    .replace(/[٫،,]/g, ".")
    .replace(/[^0-9.]/g, "");
  const i = v.indexOf(".");
  if (i >= 0) v = v.slice(0, i + 1) + v.slice(i + 1).replace(/\./g, "");
  else if (v.length > 1) v = v[0] + "." + v.slice(1);
  return v;
}

export function fmt(x: number | "" | null | undefined, unit: UnitCode = "cm"): string {
  if (x === "" || x === null || x === undefined) return "";
  const q = unit === "m" ? 1000 : 100;
  const r = Math.round(r6(x) * q) / q;
  return String(r);
}

export const fmt6 = (x: number) => String(Math.round(x * 1e6) / 1e6);

export function defaultLayout(): LayoutConfig {
  const units: string[] = [];
  const kids: Record<string, string[]> = {};
  BASE_FIELDS.forEach(([k, g]) => {
    const id = g || k;
    if (!kids[id]) {
      kids[id] = [];
      units.push(id);
    }
    kids[id].push(k);
  });
  return { units, kids };
}

const sameSet = (a: unknown, b: string[]) =>
  Array.isArray(a) && a.length === b.length && b.every((x) => a.includes(x));

export function validLayout(o: unknown): LayoutConfig | null {
  const d = defaultLayout();
  if (!o || typeof o !== "object") return null;
  const obj = o as { units?: unknown; kids?: Record<string, unknown> };
  if (!Array.isArray(obj.units) || !obj.kids || typeof obj.kids !== "object") return null;

  // Migrate older v11 layouts gracefully by inserting new area units/kids if missing
  const units = obj.units.filter((u): u is string => typeof u === "string");
  const kids: Record<string, string[]> = {};
  Object.keys(obj.kids).forEach((k) => {
    const arr = obj.kids![k];
    if (Array.isArray(arr)) kids[k] = arr.filter((x): x is string => typeof x === "string");
  });

  d.units.forEach((u) => {
    if (!units.includes(u)) units.push(u);
    const cur = kids[u] || [];
    d.kids[u].forEach((k) => {
      if (!cur.includes(k)) cur.push(k);
    });
    kids[u] = cur;
  });

  if (
    sameSet(
      units.filter((u) => !isCu(u)),
      d.units
    ) &&
    d.units.every((u) => sameSet((kids[u] || []).filter((k) => !isCid(k)), d.kids[u]))
  ) {
    return { units, kids };
  }
  return null;
}

export function isFieldSelected(k: string, sel: Record<string, boolean>): boolean {
  return sel[k] !== false;
}

export function createDefaultState(): AppState {
  return {
    type: BUILTIN_TYPE_NAMES[0],
    rows: [{ n: "", L: "", W: "" }],
    sel: {},
    inc: {},
    dec: {},
    rel: {},
    rnd: {},
    gr: { s: 0, d: "n" },
    raw: false,
    custom: {},
    cn: 0,
    utypes: [],
    cascade: true,
    layout: defaultLayout(),
    view: "cards",
    unit: "cm",
    step: 1,
  };
}

export interface RuntimeCatalog {
  ORDER: string[];
  REL: Record<string, RelTuple>;
  TYPES: Record<string, TypeDef>;
  LABEL: Record<string, string>;
  SUB: Record<string, string>;
  HUE: Record<string, number>;
  DIV: Record<string, number>;
  TN: string[];
  TINFO: Record<string, { sub: string; iconKey: "a" | "b" | "c" | "d" | "e" }>;
}

const sgv = (n?: boolean, v?: string) => {
  const x = num(v);
  return x === null || isNaN(x) ? 0 : (n ? -1 : 1) * Math.abs(x);
};

export function createRuntimeCatalog(state: AppState): RuntimeCatalog {
  const ORDER = BASE_ORDER.slice();
  const REL: Record<string, RelTuple> = {};
  Object.keys(BASE_REL).forEach((k) => {
    REL[k] = BASE_REL[k].slice() as RelTuple;
  });

  // Apply built-in per-type formula differences before user overrides
  const typeRelDefs = DEF_TYPE_DEFS[state.type]?.rel;
  if (typeRelDefs) {
    Object.keys(typeRelDefs).forEach((k) => {
      REL[k] = typeRelDefs[k].slice() as RelTuple;
    });
  }

  const TYPES: Record<string, TypeDef> = {};
  Object.keys(DEF_TYPE_DEFS).forEach((n) => {
    TYPES[n] = {
      mAdd: DEF_TYPE_DEFS[n].mAdd,
      eConst: DEF_TYPE_DEFS[n].eConst,
      fields: DEF_TYPE_DEFS[n].fields.slice(),
      rel: DEF_TYPE_DEFS[n].rel,
    };
  });

  const LABEL: Record<string, string> = {};
  const SUB: Record<string, string> = {};
  BASE_FIELDS.forEach(([k, g, n]) => {
    SUB[k] = n;
    const gClean = g.replace(/^ال/, "");
    LABEL[k] = g ? (n.includes(gClean) ? n : `${g} - ${n}`) : n;
  });

  const HUE: Record<string, number> = { ...BASE_HUE };
  const DIV: Record<string, number> = { ...BASE_DIV };
  const TINFO: Record<string, { sub: string; iconKey: "a" | "b" | "c" | "d" | "e" }> = { ...TYPE_SUBS };

  const custList = (n?: string) => (state.custom && state.custom[n || state.type]) || [];

  state.utypes.forEach((u) => {
    if (!TYPES[u.name]) {
      TYPES[u.name] = { mAdd: 0, eConst: null, fields: [] };
    }
    TINFO[u.name] = {
      sub: `نوع أضفته أنت · ${custList(u.name).length} خانة`,
      iconKey: "e",
    };
  });

  const TN = [...BUILTIN_TYPE_NAMES, ...state.utypes.map((u) => u.name)];

  Object.keys(TYPES).forEach((n) => {
    const base = DEF_TYPE_DEFS[n]?.fields || [];
    TYPES[n].fields = base.concat(custList(n).map((c) => c.id));
  });

  const hues: Record<string, number> = {};
  let hi = 0;
  custList(state.type).forEach((c) => {
    const u = cUnit(state.type, c);
    if (!(u in hues)) hues[u] = CHUE[hi++ % CHUE.length];
    ORDER.push(c.id);
    REL[c.id] = ["LEN", 1, 0];
    const nm = clean(c.name).trim() || "خانة جديدة";
    const gp = clean(c.group).trim();
    SUB[c.id] = nm;
    LABEL[c.id] = gp ? `${gp} - ${nm}` : nm;
    HUE[u] = hues[u];
  });

  // Apply relation overrides for active type
  const ov = (state.rel && state.rel[state.type]) || {};
  const t = TYPES[state.type];
  if (t) {
    Object.keys(ov).forEach((k) => {
      const o = ov[k];
      if (!o) return;
      if (k === "E") {
        const c = num(o.a);
        if (t.eConst !== null && c !== null && !isNaN(c)) t.eConst = c;
        return;
      }
      if (!REL[k]) return;
      const idxK = ORDER.indexOf(k);
      const allowedSrc = (s: string) =>
        s === "LEN" ||
        s === "WID" ||
        (ORDER.indexOf(s) >= 0 && ORDER.indexOf(s) < idxK && s !== "E" && t.fields.includes(s));
      if (!allowedSrc(o.s)) return;
      let m = num(o.m);
      let a = num(o.a);
      if (m === null || isNaN(m)) m = 1;
      if (a === null || isNaN(a)) a = 0;
      delete DIV[k];
      if (o.dv && m !== 0) {
        DIV[k] = m;
        m = 1 / m;
      }
      a = (o.neg ? -1 : 1) * Math.abs(a);
      let s2 = "";
      let op: "+" | "-" | "*" | "/" = "+";
      if (o.s2) {
        if (!allowedSrc(o.s2)) return;
        s2 = o.s2;
        op = o.op && ["+", "-", "*", "/"].includes(o.op) ? o.op : "+";
      }
      const q1 = s2 ? sgv(o.n1, o.a1) : 0;
      const q2 = s2 ? sgv(o.n2, o.a2) : 0;
      if (k === "M") {
        REL.M = [o.s, m, null, s2 || undefined, op, q1, q2];
        t.mAdd = a;
      } else {
        REL[k] = [o.s, m, a, s2 || undefined, op, q1, q2];
      }
    });
  }

  return { ORDER, REL, TYPES, LABEL, SUB, HUE, DIV, TN, TINFO };
}

export function syncLayoutWithCustom(state: AppState): LayoutConfig {
  const L: LayoutConfig = {
    units: state.layout.units.slice(),
    kids: { ...state.layout.kids },
  };
  const want: Record<string, string[]> = {};
  Object.keys(state.custom || {}).forEach((n) => {
    (state.custom[n] || []).forEach((c) => {
      const u = cUnit(n, c);
      (want[u] = want[u] || []).push(c.id);
    });
  });
  L.units = L.units.filter((u) => !isCu(u) || Boolean(want[u]));
  Object.keys(L.kids).forEach((u) => {
    if (isCu(u) && !want[u]) {
      delete L.kids[u];
    } else {
      const wantedHere = want[u] || [];
      L.kids[u] = (L.kids[u] || []).filter((id) => !isCid(id) || wantedHere.includes(id));
    }
  });
  Object.keys(want).forEach((u) => {
    if (!L.units.includes(u)) L.units.push(u);
    const cur = (L.kids[u] || []).slice();
    want[u].forEach((id) => {
      if (!cur.includes(id)) cur.push(id);
    });
    L.kids[u] = cur;
  });
  return L;
}

export function fdim(cat: RuntimeCatalog, k: string, dp = 0): number {
  if (dp > 12 || k === "LEN" || k === "WID" || k === "E") return 1;
  const r = cat.REL[k];
  if (!r) return 1;
  const d1 = fdim(cat, r[0], dp + 1);
  if (r[3]) {
    const d2 = fdim(cat, r[3], dp + 1);
    return r[4] === "*" ? d1 + d2 : r[4] === "/" ? d1 - d2 : d1;
  }
  return d1;
}

export function isGroupUnit(id: string): boolean {
  return !BASE_FIELDS.some((f) => f[0] === id) && !isCid(id);
}

export function compute(
  cat: RuntimeCatalog,
  type: string,
  len: number,
  wid: number,
  adjust: Record<string, number> = {},
  cascade = true
): Record<string, number> {
  const t = cat.TYPES[type];
  const adj: Record<string, number> = {};
  const raw: Record<string, number> = {};
  const inp: Record<string, number> = { LEN: len, WID: wid };
  if (!t) return adj;

  for (const k of cat.ORDER) {
    let v: number;
    if (k === "E") {
      if (t.eConst === null) continue;
      raw[k] = t.eConst;
      v = t.eConst;
    } else {
      const rel = cat.REL[k];
      if (!rel) continue;
      const [s, m, aInit, s2, op, q1 = 0, q2 = 0] = rel;
      const a = (k === "M" ? t.mAdd : aInit) || 0;
      const g = (x: string): [number, number] =>
        x in inp ? [inp[x], inp[x]] : [raw[x] ?? 0, (cascade ? adj[x] : raw[x]) ?? 0];
      let [br, ba] = g(s);
      if (s2) {
        let [r2, b2] = g(s2);
        br += q1;
        ba += q1;
        r2 += q2;
        b2 += q2;
        const f = (x: number, y: number) =>
          op === "-" ? x - y : op === "*" ? x * y : op === "/" ? (y ? x / y : 0) : x + y;
        br = f(br, r2);
        ba = f(ba, b2);
      }
      raw[k] = br * m + a;
      v = ba * m + a;
    }
    adj[k] = r6(v + (adjust[k] || 0));
  }
  return adj;
}

export const pL = (r: WindowRow) => r.p === "L" || r.p === "B";
export const pW = (r: WindowRow) => r.p === "W" || r.p === "B";
export const qOf = (r: WindowRow) => {
  const v = num(r.q);
  return v && !isNaN(v) && v >= 1 ? Math.min(MAXQ, Math.floor(v)) : 1;
};

export function rowErr(r: WindowRow, unit: UnitCode, isNegRow = false): string {
  const L = pL(r) ? null : num(r.L);
  const W = pW(r) ? null : num(r.W);
  if (L === null && W === null) return "";
  if ((L === null && !pL(r)) || (W === null && !pW(r))) {
    return "اكتب الطول والعرض معاً أو علّم القياس «غير جاهز»";
  }
  if ((L !== null && isNaN(L)) || (W !== null && isNaN(W))) {
    return "اكتب أرقاماً فقط";
  }
  const f = UF[unit];
  if ((L !== null && (!isFinite(L) || L <= 0)) || (W !== null && (!isFinite(W) || W <= 0))) {
    return "القياس يجب أن يكون أكبر من صفر";
  }
  if ((L !== null && L * f > MAXCM) || (W !== null && W * f > MAXCM)) {
    return "القياس كبير جداً (الحد الأقصى 50 متراً)";
  }
  if (isNegRow) {
    return "الشباك صغير جداً لهذا النوع: ستظهر قياسات سالبة";
  }
  return "";
}

export function rowWarn(r: WindowRow, unit: UnitCode): boolean {
  const L = pL(r) ? null : num(r.L);
  const W = pW(r) ? null : num(r.W);
  const f = UF[unit];
  const o = (v: number | null) =>
    v !== null && !isNaN(v) && v > 0 && (v * f < 10 || v * f > 600);
  return o(L) || o(W);
}

export function depSet(cat: RuntimeCatalog, k: string, d = 0): Set<string> {
  const s = new Set<string>();
  if (d > 12 || k === "E") return s;
  if (k === "LEN") {
    s.add("L");
    return s;
  }
  if (k === "WID") {
    s.add("W");
    return s;
  }
  const r = cat.REL[k];
  if (!r) return s;
  [r[0], r[3]].forEach((x) => {
    if (x) depSet(cat, x, d + 1).forEach((v) => s.add(v));
  });
  return s;
}

export function rndOwn(state: AppState, k: string, type?: string): RoundConfig | null {
  const o = ((state.rnd || {})[type || state.type] || {})[k];
  return o && +o.s > 0 ? { s: +o.s, d: o.d || "n" } : null;
}

export function rndCfg(
  state: AppState,
  k: string,
  type?: string,
  isArea = AREA_KEYS.has(k)
): RoundConfig | null {
  return (
    rndOwn(state, k, type) ||
    (k !== "E" && !isArea && state.gr && +state.gr.s > 0
      ? { s: +state.gr.s, d: state.gr.d || "n" }
      : null)
  );
}

export function rndApply(v: number, c: RoundConfig | null): number {
  if (!c) return v;
  const q = v / c.s;
  const n =
    c.d === "u"
      ? Math.ceil(q - 1e-9)
      : c.d === "d"
        ? Math.floor(q + 1e-9)
        : Math.round(q + 1e-9);
  return r6(n * c.s);
}

export function shownV(state: AppState, k: string, v: number | "", isArea?: boolean): number | "" {
  if (v === "") return "";
  return state.raw ? v : rndApply(v, rndCfg(state, k, undefined, isArea));
}

export function defRel(type: string, k: string): RelConfig {
  if (isCid(k)) {
    return {
      s: "LEN",
      m: "1",
      neg: false,
      a: "0",
      dv: false,
      s2: "",
      op: "+",
      n1: false,
      a1: "0",
      n2: false,
      a2: "0",
    };
  }
  if (k === "E") {
    return { s: "", m: "1", neg: false, a: String(DEF_TYPE_DEFS[type]?.eConst ?? 0) };
  }
  const d = DEF_TYPE_DEFS[type]?.rel?.[k] || BASE_REL[k];
  const aRaw = k === "M" ? DEF_TYPE_DEFS[type]?.mAdd ?? 0 : d[2] ?? 0;
  const baseDiv = BASE_DIV[k];
  const q1Raw = d[5] ?? 0;
  const q2Raw = d[6] ?? 0;
  return {
    s: d[0],
    m: baseDiv ? String(baseDiv) : String(d[1]),
    neg: aRaw < 0,
    a: String(Math.abs(aRaw)),
    dv: Boolean(baseDiv),
    s2: d[3] || "",
    op: (d[4] as "+" | "-" | "*" | "/") || "+",
    n1: q1Raw < 0,
    a1: String(Math.abs(q1Raw)),
    n2: q2Raw < 0,
    a2: String(Math.abs(q2Raw)),
  };
}

export function effRel(state: AppState, k: string): RelConfig {
  const ov = (state.rel && state.rel[state.type]) || {};
  return ov[k] || defRel(state.type, k);
}

export function relSig(o: RelConfig): string {
  const m = num(o.m);
  const a = num(o.a);
  const a1 = num(o.a1);
  const a2 = num(o.a2);
  return [
    o.s,
    m === null || isNaN(m) ? 1 : m,
    o.neg ? 1 : 0,
    Math.abs(a === null || isNaN(a) ? 0 : a),
    o.dv ? 1 : 0,
    o.s2 || "",
    o.s2 ? o.op || "+" : "",
    o.s2 ? (o.n1 ? 1 : 0) : 0,
    o.s2 ? Math.abs(a1 === null || isNaN(a1) ? 0 : a1) : 0,
    o.s2 ? (o.n2 ? 1 : 0) : 0,
    o.s2 ? Math.abs(a2 === null || isNaN(a2) ? 0 : a2) : 0,
  ].join("|");
}

export function relParen(cat: RuntimeCatalog, state: AppState, k: string): string {
  const cuv = (a: number, dim = 1) => fmt(a / Math.pow(UF[state.unit], dim), state.unit);
  const t = cat.TYPES[state.type];
  if (k === "E") return t?.eConst !== null ? `ثابت = ${cuv(t.eConst)}` : "—";
  const rel = cat.REL[k];
  if (!rel) return "—";
  const [s, m, aInit, s2, op, q1 = 0, q2 = 0] = rel;
  const a = (k === "M" ? t?.mAdd : aInit) || 0;
  const srcName = (x: string) => (x === "LEN" ? "الطول" : x === "WID" ? "العرض" : cat.LABEL[x]);
  const adjT = (q: number, dim: number) =>
    q ? ` ${q > 0 ? "＋" : "−"} ${cuv(Math.abs(q), dim)}` : "";
  const dimK = fdim(cat, k);
  const dimS = fdim(cat, s);
  const dimS2 = s2 ? fdim(cat, s2) : 1;
  let nm = `(${srcName(s)})`;
  if (s2) {
    const opSym = op === "-" ? "−" : op === "*" ? "×" : op === "/" ? "÷" : "＋";
    nm = `(${srcName(s)}${adjT(q1, dimS)}) ${opSym} (${srcName(s2)}${adjT(q2, dimS2)})`;
    if (cat.DIV[k] || m !== 1 || a) nm = `(${nm})`;
  }
  if (cat.DIV[k]) nm += ` ÷ ${cat.DIV[k]}`;
  else if (m !== 1) nm += ` × ${m}`;
  if (a) nm += ` ${a > 0 ? "+" : "−"} ${cuv(Math.abs(a), dimK)}`;
  return nm;
}

export function buildResults(state: AppState, negSet?: WeakSet<WindowRow>): BuildResult {
  const cat = createRuntimeCatalog(state);
  const t = cat.TYPES[state.type] || { mAdd: 0, eConst: null, fields: [] };
  const units: ResultUnit[] = [];
  const keys: string[] = [];

  state.layout.units.forEach((id) => {
    const ks = (state.layout.kids[id] || []).filter(
      (k) => t.fields.includes(k) && isFieldSelected(k, state.sel)
    );
    if (ks.length) {
      units.push({
        id,
        title: GN(id),
        group: isGroupUnit(id),
        hue: cat.HUE[id] ?? 0,
        ks,
      });
      keys.push(...ks);
    }
  });

  const sfx = (k: string) => (fdim(cat, k) === 2 ? ` (${UAS[state.unit]})` : "");
  const LB = (k: string) => cat.LABEL[k] + sfx(k);
  const SB = (k: string) => cat.SUB[k] + sfx(k);

  const adjust: Record<string, number> = {};
  const adjTxt: string[] = [];
  let bad: string | null = null;

  keys.forEach((k) => {
    const a = num(state.inc[k]);
    const b = num(state.dec[k]);
    if (
      Number.isNaN(a) ||
      Number.isNaN(b) ||
      (a !== null && !isFinite(a)) ||
      (b !== null && !isFinite(b))
    ) {
      bad = `قيمة التعديل غير صحيحة في: ${cat.LABEL[k]}`;
      return;
    }
    if ((a !== null && Math.abs(a) > 1e5) || (b !== null && Math.abs(b) > 1e5)) {
      bad = `قيمة التعديل كبيرة جداً في: ${cat.LABEL[k]}`;
      return;
    }
    const v = Math.round(((a === null ? 0 : Math.abs(a)) - (b === null ? 0 : Math.abs(b))) * 1e6) / 1e6;
    if (v) {
      adjust[k] = v * Math.pow(UF[state.unit], fdim(cat, k));
      adjTxt.push(`${cat.LABEL[k]}: ${v > 0 ? "+" : ""}${v}`);
    }
  });

  const data: ComputedWindow[] = [];
  const negList: number[] = [];
  const negSeq: number[] = [];

  state.rows.forEach((r, i) => {
    const L = pL(r) ? null : num(r.L);
    const W = pW(r) ? null : num(r.W);
    if ((L === null && !pL(r)) || (W === null && !pW(r))) return;
    if (isNaN(L as number) || isNaN(W as number) || rowErr(r, state.unit, negSet?.has(r))) return;

    const f = UF[state.unit];
    const rsRaw = compute(cat, state.type, (L || 0) * f, (W || 0) * f, adjust, state.cascade);
    const rs: Record<string, number | ""> = {};
    keys.forEach((k) => {
      let val = rsRaw[k];
      if (typeof val === "number" && f !== 1) {
        val = r6(val / Math.pow(f, fdim(cat, k)));
      }
      const d = depSet(cat, k);
      if ((L === null && d.has("L")) || (W === null && d.has("W"))) {
        rs[k] = "";
      } else {
        rs[k] = val ?? "";
      }
    });

    const neg = keys.some((k) => typeof rs[k] === "number" && (rs[k] as number) < -1e-9);
    if (neg) {
      negList.push(i + 1);
      negSeq.push(data.length + 1);
    }
    data.push({
      n: String(data.length + 1),
      nm: String(r.n || ""),
      L,
      W,
      q: qOf(r),
      res: rs,
      neg,
      row: r,
    });
  });

  const struct: ExportStructUnit[] = units.map((u) => ({
    group: u.group,
    hue: u.hue,
    label: LB(u.ks[0]),
    name: u.title,
    subs: u.ks.map((k) => SB(k)),
  }));

  const c1 = ["رقم", "طول الشباك", "عرض الشباك"];
  const c2 = ["", "", ""];
  units.forEach((u) => {
    if (u.group) {
      c1.push(u.title, ...Array(u.ks.length - 1).fill(""));
      c2.push(...u.ks.map((k) => SB(k)));
    } else {
      c1.push(LB(u.ks[0]));
      c2.push("");
    }
  });

  const table = data.map((d) => [
    d.n,
    fmt(d.L, state.unit),
    fmt(d.W, state.unit),
    ...keys.map((k) => fmt(shownV(state, k, d.res[k], fdim(cat, k) === 2), state.unit)),
  ]);

  const rndOn = keys.some((k) => Boolean(rndCfg(state, k, undefined, fdim(cat, k) === 2)));
  const hasQ = data.some((d) => d.q > 1);
  const qtotal = data.reduce((a, d) => a + d.q, 0);
  const hasN = data.some((d) => Boolean(d.nm));

  let xtable = table;
  let xstruct = struct;
  let xc1 = c1;
  let xc2 = c2;

  if (hasN || hasQ) {
    const extra: { k: string; label: string; f: (d: ComputedWindow) => string }[] = [];
    if (hasN) extra.push({ k: "nm", label: "الاسم", f: (d) => d.nm });
    if (hasQ) extra.push({ k: "q", label: "العدد", f: (d) => String(d.q) });
    xtable = table.map((r, i) => [r[0], r[1], r[2], ...extra.map((e) => e.f(data[i])), ...r.slice(3)]);
    xstruct = [
      ...extra.map((e) => ({ group: false, hue: 210, label: e.label, name: e.label, subs: [""] })),
      ...struct,
    ];
    xc1 = [...c1.slice(0, 3), ...extra.map((e) => e.label), ...c1.slice(3)];
    xc2 = [...c2.slice(0, 3), ...extra.map(() => ""), ...c2.slice(3)];
  }

  // Group strictly by exact L × W (150×100 stays separate from 100×150)
  const gm = new Map<string, DuplicateGroup>();
  data.forEach((d) => {
    if (d.L === null || d.W === null) return;
    const k = `${fmt(d.L, state.unit)}|${fmt(d.W, state.unit)}`;
    let g = gm.get(k);
    if (!g) {
      g = { L: d.L, W: d.W, cnt: 0, nums: [] };
      gm.set(k, g);
    }
    g.cnt += d.q;
    g.nums.push(d.n);
  });

  const dups = [...gm.values()]
    .filter((g) => g.cnt > 1)
    .sort((a, b) => b.cnt - a.cnt || a.L - b.L || a.W - b.W);
  const dupRows: [string, string, string, string][] = dups.map((g) => [
    fmt(g.L, state.unit),
    fmt(g.W, state.unit),
    String(g.cnt),
    g.nums.join("، "),
  ]);
  const dupOther = gm.size - dups.length;
  const pend = data.filter((d) => d.L === null || d.W === null).map((d) => d.n);

  // Sum enabled area fields (multiplied by quantity q)
  const areaKeys = keys.filter((k) => fdim(cat, k) === 2);
  const areaTotals: AreaTotalItem[] = areaKeys.map((k) => {
    let totalInUnit = 0;
    data.forEach((d) => {
      const val = shownV(state, k, d.res[k], true);
      if (typeof val === "number" && val > 0) {
        totalInUnit += val * d.q;
      }
    });
    const toMeters = UF[state.unit] / 100;
    const totalSqM = r6(totalInUnit * toMeters * toMeters);
    return {
      key: k,
      label: cat.LABEL[k],
      totalInUnit: r6(totalInUnit),
      totalSqM,
    };
  });

  return {
    pend,
    dups,
    dupRows,
    dupOther,
    negList,
    negSeq,
    area: areaKeys.length > 0,
    areaTotals,
    hasN,
    units,
    keys,
    data,
    bad,
    rndOn,
    hasQ,
    qtotal,
    xtable,
    xstruct,
    xc1,
    xc2,
    c1,
    c2,
    table,
    struct,
    type: state.type,
    adj: adjTxt.join("  |  "),
    cascade: state.cascade,
  };
}
