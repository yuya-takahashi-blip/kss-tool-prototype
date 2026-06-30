import { getLayoutName, type PageType } from "./layout-labels";

export interface SelectedSlide {
  id: string;
  sectionKey: string;
  sectionName: string;
  type: PageType;
  slideIndexInSection: number;
}

export interface GeneratePptxOptions {
  title: string;
  clientName: string;
  date: string;
  companyLogo: string; // base64 data URL or "" for text fallback
  selectedSlides: SelectedSlide[];
}

// ─── Design tokens ────────────────────────────────────────────────────────────
const RED      = "C0392B";
const DARK     = "2C2C2C";
const MID      = "666666";
const LIGHT_BG = "F5F5F5";
const PH       = "DEDEDE";
const PH_B     = "C8C8C8";
const WHITE    = "FFFFFF";
const RED_SOFT = "FEF2F2";
const FONT     = "Arial";

// ─── Geometry  (LAYOUT_WIDE = 13.333 × 7.5 inches) ───────────────────────────
const SW       = 13.333;
const SH       = 7.5;
const M        = 0.5;
const HEADER_H = 0.8;
const FOOTER_Y = 6.9;
const CX       = M;
const CY       = HEADER_H + 0.22; // 1.02
const CW       = SW - M * 2;      // 12.333
const CH       = FOOTER_Y - CY - 0.16; // ~5.72

// ─── Slide type alias ─────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Slide = any;

// ─── Primitive helpers ────────────────────────────────────────────────────────

/** Rect shape + optional centered label (placeholder style) */
function box(
  sl: Slide,
  x: number, y: number, w: number, h: number,
  label = "", fill = PH, border = PH_B, labelColor = "AAAAAA"
) {
  sl.addShape("rect", { x, y, w, h, fill: { color: fill }, line: { color: border, width: 0.5 } });
  if (label) {
    sl.addText(label, { x, y, w, h, fontFace: FONT, fontSize: 9, color: labelColor, align: "center", valign: "middle" });
  }
}

/** addText wrapper with sensible defaults */
function txt(
  sl: Slide,
  text: string,
  x: number, y: number, w: number, h: number,
  opts: Record<string, unknown> = {}
) {
  sl.addText(text, { x, y, w, h, fontFace: FONT, color: DARK, fontSize: 10, valign: "top", ...opts });
}

/**
 * Arrow: Canvas-rendered PNG embedded as addImage.
 *
 * Root cause of all previous bleeding:
 *   pptxgenjs v4 internally shares state for non-rect shape types (preset geometries,
 *   rotate transforms, line shapes with endArrowType).  The bleed manifests as shapes
 *   from slide N appearing on slides N+1, N+2, etc.
 *
 * addImage embeds the arrow as a per-slide media file in the PPTX package.
 *   - Each call creates a distinct media entry → zero bleed possible.
 *   - canvas.toDataURL() is SYNCHRONOUS → no async refactoring needed.
 *   - Renders as true raster PNG: shaft (rect) + arrowhead (filled triangle polygon).
 *   - Fallback: plain rect shaft if canvas is unavailable (e.g. SSR context).
 *
 * Guard: only renders when sectionKey === "business_scheme".
 */
function drawArrow(sl: Slide, x: number, y: number, w: number, h: number, sectionKey: string) {
  if (sectionKey !== "business_scheme") return;

  if (typeof document !== "undefined") {
    const DPI  = 192;  // 2× screen density for crisp rendering
    const pixW = Math.max(Math.round(w * DPI), 4);
    const pixH = Math.max(Math.round(h * DPI), 4);
    const canvas = document.createElement("canvas");
    canvas.width  = pixW;
    canvas.height = pixH;
    const ctx = canvas.getContext("2d");
    if (ctx) {
      const mid    = pixH / 2;
      const shaftH = Math.max(Math.round(pixH * 0.28), 3);
      const headW  = Math.round(pixW * 0.42);
      const shaftW = pixW - headW + 2;
      ctx.fillStyle = `#${RED}`;
      // Shaft
      ctx.fillRect(0, Math.round(mid - shaftH / 2), shaftW, shaftH);
      // Arrowhead — filled triangle pointing right
      ctx.beginPath();
      ctx.moveTo(pixW - headW, 0);
      ctx.lineTo(pixW,         mid);
      ctx.lineTo(pixW - headW, pixH);
      ctx.closePath();
      ctx.fill();
      try {
        const dataUrl = canvas.toDataURL("image/png");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (sl as any).addImage({ data: dataUrl, x, y, w, h });
      } catch { /* ignore */ }
      return;
    }
  }

  // Fallback: shaft rect only (canvas unavailable in SSR context)
  const mid    = y + h / 2;
  const shaftH = Math.max(h * 0.28, 0.06);
  sl.addShape("rect", { x, y: mid - shaftH / 2, w, h: shaftH, fill: { color: RED }, line: { color: RED, width: 0 } });
}

/**
 * Solid circle using ● (U+25CF, BLACK CIRCLE).
 * Guard: only renders when sectionKey === "schedule".
 * This prevents addText("●") from appearing on other slides.
 */
function drawCircle(sl: Slide, cx: number, cy: number, r: number, sectionKey: string) {
  if (sectionKey !== "schedule") return;   // ← page guard

  sl.addText("●", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2,
    fontFace: FONT, fontSize: Math.round(r * 72 * 1.55),
    color: RED, align: "center", valign: "middle",
  });
}

// ─── Common header (every non-cover slide) ────────────────────────────────────
function addHeader(
  sl: Slide,
  sectionName: string,
  layoutName: string
) {
  // Full-width header band
  sl.addShape("rect", { x: 0, y: 0, w: SW, h: HEADER_H, fill: { color: LIGHT_BG }, line: { color: LIGHT_BG } });
  // Red left accent bar
  sl.addShape("rect", { x: 0, y: 0, w: 0.09, h: HEADER_H, fill: { color: RED }, line: { color: RED } });
  // Section name
  txt(sl, sectionName, 0.25, 0, SW - M - 3.0, HEADER_H, { fontSize: 16, bold: true, color: DARK, valign: "middle" });
  // Layout badge
  sl.addShape("rect", { x: SW - M - 2.8, y: 0.2, w: 2.8, h: 0.38, fill: { color: RED_SOFT }, line: { color: "FECACA", width: 0.5 } });
  txt(sl, layoutName, SW - M - 2.8, 0.2, 2.8, 0.38, { fontSize: 9, color: RED, align: "center", valign: "middle" });
}

// ─── Common footer (every non-cover slide) ────────────────────────────────────
function addFooter(
  sl: Slide,
  pageNum: number,
  totalPages: number,
  date: string,
  companyLogo: string
) {
  // Separator line
  sl.addShape("rect", { x: CX, y: FOOTER_Y, w: CW, h: 0.01, fill: { color: "E0E0E0" }, line: { color: "E0E0E0" } });

  const FY = FOOTER_Y + 0.1;
  const FH = 0.2;

  if (companyLogo) {
    // Logo image + date: group centered on full slide width
    const LOGO_W = 0.85;
    const LOGO_H = FH;
    const DATE_W = 0.75;
    const GAP    = 0.1;
    const totalW = LOGO_W + (date ? GAP + DATE_W : 0);
    const startX = (SW - totalW) / 2;
    try {
      sl.addImage({ data: companyLogo, x: startX, y: FY, w: LOGO_W, h: LOGO_H });
      if (date) {
        sl.addText(date, {
          x: startX + LOGO_W + GAP, y: FY, w: DATE_W, h: FH,
          fontFace: FONT, fontSize: 6.5, color: "B8B8B8", align: "left", valign: "middle",
        });
      }
    } catch {
      addFooterText(sl, date, FY, FH);
    }
  } else {
    addFooterText(sl, date, FY, FH);
  }

  // Page number (right-aligned, within slide)
  txt(sl, `${pageNum} / ${totalPages}`, SW - M - 1.8, FOOTER_Y + 0.06, 1.8, 0.32, {
    fontSize: 8, color: "BBBBBB", align: "right", valign: "middle",
  });
}

function addFooterText(sl: Slide, date: string, y: number, h: number) {
  // Single full-width, center-aligned text — guaranteed to be exactly centered
  sl.addText(
    date
      ? [
          { text: "KANSAI SUPER STUDIO", options: { bold: true } as Record<string, unknown> },
          { text: `　　${date}`, options: {} as Record<string, unknown> },
        ]
      : [{ text: "KANSAI SUPER STUDIO", options: { bold: true } as Record<string, unknown> }],
    { x: 0, y, w: SW, h, fontFace: FONT, fontSize: 6.5, color: "B8B8B8", align: "center", valign: "middle" }
  );
}

// ─── Cover slide ──────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function renderCover(pptx: any, title: string, clientName: string, date: string, companyLogo: string, coverType: PageType) {
  const sl = pptx.addSlide();

  if (coverType === "B" || coverType === "C") {
    // Custom cover: minimal — user edits in PowerPoint
    sl.addShape("rect", { x: 0, y: 0, w: SW, h: 0.15, fill: { color: RED }, line: { color: RED } });
    sl.addShape("rect", { x: 0, y: SH - 0.15, w: SW, h: 0.15, fill: { color: RED }, line: { color: RED } });
    return;
  }

  // Standard cover (A)
  sl.addShape("rect", { x: 0, y: 0, w: SW, h: 0.15, fill: { color: RED }, line: { color: RED } });
  sl.addShape("rect", { x: 0, y: SH - 0.15, w: SW, h: 0.15, fill: { color: RED }, line: { color: RED } });

  txt(sl, title || "企画書タイトル", M, 1.2, CW, 2.2, {
    fontSize: 36, bold: true, color: title ? DARK : "CCCCCC", align: "center", valign: "middle", wrap: true,
  });

  sl.addShape("rect", { x: SW / 2 - 2.5, y: 3.6, w: 5.0, h: 0.02, fill: { color: "E0E0E0" }, line: { color: "E0E0E0" } });

  txt(sl, clientName || "提案先会社名", M, 3.7, CW, 1.4, {
    fontSize: 26, bold: true, color: clientName ? DARK : "CCCCCC", align: "center", valign: "middle",
  });

  if (companyLogo) {
    const logoH = 0.55;
    const logoW = logoH * 3.5;
    const logoX = SW / 2 - logoW / 2 - 0.5;
    try { sl.addImage({ data: companyLogo, x: logoX, y: 5.5, w: logoW, h: logoH }); } catch { /* ignore */ }
    if (date) txt(sl, date, logoX + logoW + 0.2, 5.5, 2.5, logoH, { fontSize: 11, color: MID, valign: "middle" });
  } else {
    txt(sl, date ? `KANSAI SUPER STUDIO　${date}` : "KANSAI SUPER STUDIO", M, 5.5, CW, 0.55, {
      fontSize: 11, color: MID, align: "center", valign: "middle",
    });
  }
}

// ─── Section renderers ────────────────────────────────────────────────────────

function renderBrand(sl: Slide, type: PageType) {
  if (type === "A") {
    const leftW = 2.8;
    box(sl, CX, CY, leftW, 1.3, "ロゴ", PH, PH_B);
    box(sl, CX, CY + 1.5, leftW, CH - 1.5, "ブランド説明文", WHITE, "E0E0E0");
    box(sl, CX + leftW + 0.3, CY, CW - leftW - 0.3, CH, "キービジュアル", PH, PH_B);
  } else if (type === "B") {
    box(sl, CX, CY, CW, CH * 0.72, "世界観ビジュアル", PH, PH_B);
    box(sl, CX, CY + CH * 0.75, CW, CH * 0.25, "キャッチコピー・説明文", WHITE, "E0E0E0");
  } else {
    const colW = (CW - 0.4) / 3;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (colW + 0.2);
      box(sl, cx2, CY, colW, 1.6, ["特徴", "ターゲット", "強み"][i], PH, PH_B);
      box(sl, cx2, CY + 1.8, colW, CH - 1.8, "説明テキスト", WHITE, "E0E0E0");
    }
  }
}

function renderBusinessScheme(sl: Slide, type: PageType, sectionKey: string) {
  if (type === "A") {
    // 2 boxes + arrow — centered
    const bw = 4.4;
    const bh = 2.6;
    const aw = 1.0;
    const ah = 0.5;
    const totalSpan = bw * 2 + aw + 0.4;
    const sx = CX + (CW - totalSpan) / 2;
    const by = CY + (CH - bh) / 2;

    sl.addShape("rect", { x: sx,           y: by, w: bw, h: bh, fill: { color: WHITE }, line: { color: RED, width: 1.5 } });
    txt(sl, "自社",       sx,           by, bw, bh, { fontSize: 18, bold: true, align: "center", valign: "middle", color: DARK });

    drawArrow(sl, sx + bw + 0.2, by + bh / 2 - ah / 2, aw, ah, sectionKey);

    const b2x = sx + bw + aw + 0.4;
    sl.addShape("rect", { x: b2x,          y: by, w: bw, h: bh, fill: { color: WHITE }, line: { color: RED, width: 1.5 } });
    txt(sl, "パートナー企業", b2x,         by, bw, bh, { fontSize: 18, bold: true, align: "center", valign: "middle", color: DARK });

    txt(sl, "取引の流れ・役割分担", CX, CY + CH - 0.55, CW, 0.45, { fontSize: 10, color: MID, align: "center" });

  } else if (type === "B") {
    // 3 boxes horizontal — centered
    const bw = 3.4;
    const bh = 2.2;
    const aw = 0.55;
    const ah = 0.45;
    const totalSpan = bw * 3 + aw * 2 + 0.3 * 4;
    const sx = CX + (CW - totalSpan) / 2;
    const by = CY + (CH - bh) / 2;
    const labels = ["自社", "中間事業者", "エンドユーザー"];

    for (let i = 0; i < 3; i++) {
      const cx2 = sx + i * (bw + aw + 0.3);
      sl.addShape("rect", { x: cx2, y: by, w: bw, h: bh, fill: { color: WHITE }, line: { color: RED, width: 1.5 } });
      txt(sl, labels[i], cx2, by, bw, bh, { fontSize: 16, bold: true, align: "center", valign: "middle", color: DARK });
      if (i < 2) drawArrow(sl, cx2 + bw + 0.08, by + bh / 2 - ah / 2, aw, ah, sectionKey);
    }
    txt(sl, "3者間スキームの流れ", CX, CY + CH - 0.55, CW, 0.45, { fontSize: 10, color: MID, align: "center" });

  } else {
    // 4-step flow
    const steps = ["商品仕入れ", "プロモーション", "販売", "入金・精算"];
    const aw  = 0.32;
    const gap = aw + 0.06;
    const bw  = (CW - gap * 3) / 4;
    const bh  = 2.4;
    const by  = CY + (CH - bh) / 2;
    const cr  = 0.34;

    for (let i = 0; i < 4; i++) {
      const cx2 = CX + i * (bw + gap);
      sl.addShape("rect", { x: cx2, y: by, w: bw, h: bh, fill: { color: WHITE }, line: { color: RED, width: 1.2 } });
      // Step number square (rect replaces ellipse — avoids pptxgenjs v4 ellipse cross-slide bleed)
      sl.addShape("rect", { x: cx2 + bw / 2 - cr, y: by + 0.2, w: cr * 2, h: cr * 2, fill: { color: RED }, line: { color: RED } });
      txt(sl, String(i + 1), cx2 + bw / 2 - cr, by + 0.2, cr * 2, cr * 2, { fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(sl, steps[i], cx2, by + 1.05, bw, 1.1, { fontSize: 12, align: "center", valign: "middle", color: DARK, wrap: true });
      if (i < 3) drawArrow(sl, cx2 + bw + 0.04, by + bh / 2 - 0.2, aw, 0.4, sectionKey);
    }
  }
}

function renderCases(sl: Slide, type: PageType) {
  if (type === "A") {
    const imgW  = CW * 0.42;
    const textX = CX + imgW + 0.25;
    const textW = CW - imgW - 0.25;
    box(sl, CX,   CY,             imgW,  CH,              "導入事例イメージ",  PH,    PH_B);
    box(sl, textX, CY,             textW, 0.85,            "企業名・業種",      WHITE, "E0E0E0");
    box(sl, textX, CY + 1.0,       textW, CH - 1.0 - 0.85, "導入効果・コメント", WHITE, "E0E0E0");
    box(sl, textX, CY + CH - 0.8,  textW, 0.75,            "「お客様の声」",    WHITE, "E0E0E0");
  } else if (type === "B") {
    const colW = (CW - 0.4) / 3;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (colW + 0.2);
      box(sl, cx2, CY,           colW, CH * 0.45, `事例 ${i + 1} 画像`,  PH,    PH_B);
      box(sl, cx2, CY + CH * 0.48, colW, CH * 0.52, `事例 ${i + 1} テキスト`, WHITE, "E0E0E0");
    }
  } else {
    // Before / After
    const half = (CW - 0.35) / 2;
    sl.addShape("rect", { x: CX, y: CY, w: half, h: 0.55, fill: { color: RED },  line: { color: RED  } });
    txt(sl, "Before", CX, CY, half, 0.55, { fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(sl, CX, CY + 0.65, half, CH - 0.65, "課題・現状", WHITE, "E0E0E0");

    const ax = CX + half + 0.35;
    sl.addShape("rect", { x: ax, y: CY, w: half, h: 0.55, fill: { color: DARK }, line: { color: DARK } });
    txt(sl, "After", ax, CY, half, 0.55, { fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(sl, ax, CY + 0.65, half, CH - 0.65, "改善後の状態・効果", WHITE, "E0E0E0");
  }
}

function renderItems(sl: Slide, type: PageType) {
  // Items: image boxes, labels, descriptions only. No arrows, no red circles.
  if (type === "A") {
    const imgW  = CW * 0.52;
    const textX = CX + imgW + 0.25;
    const textW = CW - imgW - 0.25;
    box(sl, CX,   CY,          imgW,  CH,         "商品画像", PH,    PH_B);
    box(sl, textX, CY,          textW, 0.8,        "商品名",   WHITE, "E0E0E0");
    box(sl, textX, CY + 0.95,   textW, 0.65,       "価格",     WHITE, "E0E0E0");
    box(sl, textX, CY + 1.75,   textW, CH - 1.75,  "商品説明", WHITE, "E0E0E0");
  } else if (type === "B") {
    const gw = (CW - 0.25) / 2;
    const gh = (CH - 0.25) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) {
        box(sl, CX + c * (gw + 0.25), CY + r * (gh + 0.25), gw, gh, `商品 ${r * 2 + c + 1}`, PH, PH_B);
      }
    }
  } else {
    box(sl, CX,             CY,             CW,         CH * 0.58, "使用シーン画像", PH,    PH_B);
    box(sl, CX,             CY + CH * 0.62, CW * 0.55,  CH * 0.38, "商品説明",       WHITE, "E0E0E0");
    box(sl, CX + CW * 0.58, CY + CH * 0.62, CW * 0.42, CH * 0.38, "商品画像",       PH,    PH_B);
  }
}

function renderSchedule(sl: Slide, type: PageType, sectionKey: string) {
  if (type === "A") {
    // Horizontal timeline with milestone dots (ellipse — scoped to schedule type A only)
    const milestones = ["Phase 1", "Phase 2", "Phase 3", "Phase 4"];
    const mw = CW / milestones.length;
    sl.addShape("rect", { x: CX, y: CY + CH / 2 - 0.05, w: CW, h: 0.1, fill: { color: "DDDDDD" }, line: { color: "DDDDDD" } });
    milestones.forEach((m, i) => {
      const mx = CX + mw * i + mw / 2;
      // Milestone dot — schedule-only (guard inside drawCircle enforces this)
      drawCircle(sl, mx, CY + CH / 2, 0.25, sectionKey);
      const isAbove = i % 2 === 0;
      txt(sl, m, mx - mw / 2 + 0.1, isAbove ? CY + 0.2 : CY + CH / 2 + 0.45, mw - 0.2, 1.0, { fontSize: 11, bold: true, align: "center", valign: "middle", color: DARK });
      box(sl, mx - mw / 2 + 0.2, isAbove ? CY + CH / 2 + 0.45 : CY + 0.2, mw - 0.4, CH / 2 - 0.75, "タスク内容", WHITE, "E0E0E0");
    });

  } else if (type === "B") {
    // Gantt-style table
    const months = ["4月", "5月", "6月", "7月", "8月", "9月"];
    const tasks  = ["企画・設計", "開発", "テスト", "リリース"];
    const labelW = 2.4;
    const cellW  = (CW - labelW) / months.length;
    const rowH   = CH / (tasks.length + 1);
    box(sl, CX, CY, labelW, rowH, "タスク", LIGHT_BG, "E0E0E0", MID);
    months.forEach((m, i) => box(sl, CX + labelW + i * cellW, CY, cellW, rowH, m, LIGHT_BG, "E0E0E0", MID));
    tasks.forEach((task, row) => {
      const ry    = CY + (row + 1) * rowH;
      const start = row;
      box(sl, CX, ry, labelW, rowH, task, WHITE, "E0E0E0", DARK);
      box(sl, CX + labelW + start * cellW, ry + 0.1, cellW * 2.5, rowH - 0.2, "", RED, RED);
      months.forEach((_, i) => {
        if (i < start || i >= start + 2.5) {
          sl.addShape("rect", { x: CX + labelW + i * cellW, y: ry, w: cellW, h: rowH, fill: { color: WHITE }, line: { color: "E8E8E8", width: 0.5 } });
        }
      });
    });

  } else {
    // Phase steps — arrows between phases use drawArrow (rect-based)
    const phases = [
      { name: "準備期間", sub: "1〜2ヶ月" },
      { name: "試験運用", sub: "2〜3ヶ月" },
      { name: "本格展開", sub: "3〜6ヶ月" },
    ];
    const aw  = 0.3;
    const gap = aw + 0.2;
    const bw  = (CW - gap * 2) / 3;
    phases.forEach((p, i) => {
      const cx2          = CX + i * (bw + gap);
      const isHighlight  = i === 1;
      sl.addShape("rect", {
        x: cx2, y: CY, w: bw, h: 0.65,
        fill: { color: isHighlight ? RED : LIGHT_BG },
        line: { color: isHighlight ? RED : "CCCCCC", width: 0.5 },
      });
      txt(sl, `STEP ${i + 1}`, cx2, CY, bw, 0.65, { fontSize: 13, bold: true, align: "center", valign: "middle", color: isHighlight ? WHITE : MID });
      sl.addShape("rect", { x: cx2, y: CY + 0.8, w: bw, h: CH - 0.8, fill: { color: WHITE }, line: { color: "E0E0E0", width: 0.5 } });
      txt(sl, p.name, cx2, CY + 1.0, bw, 1.2, { fontSize: 14, bold: true, align: "center", valign: "top", color: DARK });
      txt(sl, p.sub,  cx2, CY + 2.4, bw, 0.6, { fontSize: 11, align: "center", valign: "middle", color: MID });
      box(sl, cx2 + 0.25, CY + 3.2, bw - 0.5, CH - 3.4, "主なタスク", WHITE, "E0E0E0");
      // Phase separator — thin vertical divider only; drawArrow is forbidden in schedule
      if (i < 2) sl.addShape("rect", { x: cx2 + bw + aw * 0.35, y: CY + 0.8, w: aw * 0.3, h: CH - 0.8, fill: { color: "DDDDDD" }, line: { color: "DDDDDD" } });
    });
  }
}

function renderPricing(sl: Slide, type: PageType) {
  // Pricing: tables and condition areas only. No arrows, no red circles.
  if (type === "A") {
    const rows   = ["掛け率", "最低発注数", "支払い条件", "契約期間", "専属条件", "備考"];
    const labelW = 2.6;
    const rowH   = CH / rows.length;
    rows.forEach((r, i) => {
      box(sl, CX,              CY + i * rowH, labelW,          rowH, r,   i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(sl, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    const plans = ["スタンダード", "スタンダード＋", "プレミアム"];
    const bw    = (CW - 0.4) / 3;
    plans.forEach((p, i) => {
      const cx2         = CX + i * (bw + 0.2);
      const isHighlight = i === 1;
      sl.addShape("rect", { x: cx2, y: CY, w: bw, h: 0.8, fill: { color: isHighlight ? RED : LIGHT_BG }, line: { color: isHighlight ? RED : "E0E0E0", width: 0.5 } });
      txt(sl, p, cx2, CY, bw, 0.8, { fontSize: 13, bold: true, align: "center", valign: "middle", color: isHighlight ? WHITE : DARK });
      box(sl, cx2, CY + 0.95, bw, CH - 0.95, "条件詳細", WHITE, isHighlight ? RED : "E0E0E0");
    });
  } else {
    sl.addShape("rect", { x: CX, y: CY, w: CW, h: 0.75, fill: { color: LIGHT_BG }, line: { color: "E0E0E0", width: 0.5 } });
    txt(sl, "ロイヤリティ条件", CX + 0.25, CY, CW - 0.5, 0.75, { fontSize: 15, bold: true, valign: "middle", color: DARK });
    const sections = ["基本ロイヤリティ率", "インセンティブ条件", "最低保証", "支払いスケジュール"];
    const rowH     = (CH - 0.85) / sections.length;
    sections.forEach((s, i) => {
      const sy = CY + 0.85 + i * rowH;
      // Thin red accent bar (rect, not ellipse/arrow)
      sl.addShape("rect", { x: CX, y: sy + 0.1, w: 0.05, h: rowH - 0.2, fill: { color: RED }, line: { color: RED } });
      txt(sl, s, CX + 0.18, sy, CW * 0.28, rowH, { fontSize: 11, bold: true, valign: "middle", color: DARK });
      box(sl, CX + CW * 0.3 + 0.12, sy + 0.1, CW * 0.7 - 0.12, rowH - 0.2, "内容", WHITE, "E0E0E0");
    });
  }
}

function renderCompany(sl: Slide, type: PageType) {
  // Company: tables and cards only. No arrows. Timeline dots (ellipse) only in type B.
  if (type === "A") {
    const fields = ["会社名", "代表者", "設立", "所在地", "事業内容", "資本金"];
    const labelW = 2.4;
    const rowH   = CH / fields.length;
    fields.forEach((f, i) => {
      box(sl, CX,              CY + i * rowH, labelW,          rowH, f,   i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(sl, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    // History timeline — ellipses are intentional timeline dots, only in this layout
    const labelW = 2.1;
    const infoH  = CH * 0.42;
    const rowH   = infoH / 3;
    ["会社名", "代表者", "所在地"].forEach((f, i) => {
      box(sl, CX,              CY + i * rowH, labelW,          rowH, f,  LIGHT_BG, "E0E0E0", MID);
      box(sl, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
    const tly = CY + infoH + 0.25;
    txt(sl, "沿革・主な実績", CX, tly, CW, 0.5, { fontSize: 13, bold: true, color: DARK });
    // Vertical timeline line
    sl.addShape("rect", { x: CX + 0.4, y: tly + 0.6, w: 0.05, h: CH - infoH - 0.9, fill: { color: "DDDDDD" }, line: { color: "DDDDDD" } });
    for (let i = 0; i < 3; i++) {
      const ey = tly + 0.65 + i * 1.0;
      // Timeline dot — rect replaces ellipse to prevent pptxgenjs v4 cross-slide bleeding
      sl.addShape("rect", { x: CX + 0.24, y: ey, w: 0.32, h: 0.32, fill: { color: RED }, line: { color: RED } });
      box(sl, CX + 0.8, ey, CW - 0.8, 0.75, `実績 ${i + 1}`, WHITE, "E0E0E0");
    }
  } else {
    // Strength cards — red accent bars only (no arrows, no ellipses)
    const bw = (CW - 0.4) / 3;
    ["強み・特徴", "主力サービス", "実績・数値"].forEach((item, i) => {
      const cx2 = CX + i * (bw + 0.2);
      sl.addShape("rect", { x: cx2, y: CY,        w: bw, h: 0.07, fill: { color: RED },     line: { color: RED } });
      sl.addShape("rect", { x: cx2, y: CY + 0.12, w: bw, h: 0.7,  fill: { color: LIGHT_BG }, line: { color: "E0E0E0", width: 0.5 } });
      txt(sl, item, cx2, CY + 0.12, bw, 0.7, { fontSize: 13, bold: true, align: "center", valign: "middle", color: DARK });
      box(sl, cx2, CY + 0.95, bw, CH - 0.95, "詳細内容", WHITE, "E0E0E0");
    });
  }
}

// ─── Content slide builder ────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildContentSlide(
  pptx: any,
  selectedSlide: SelectedSlide,
  pageNum: number,
  totalPages: number,
  date: string,
  companyLogo: string
) {
  const sl = pptx.addSlide();

  // ① White background — first shape, covers any inherited master elements
  sl.addShape("rect", { x: 0, y: 0, w: SW, h: SH, fill: { color: WHITE }, line: { color: WHITE } });

  const layoutName = getLayoutName(selectedSlide.sectionKey, selectedSlide.type);
  const sectionLabel =
    selectedSlide.slideIndexInSection > 1
      ? `${selectedSlide.sectionName}（${selectedSlide.slideIndexInSection}）`
      : selectedSlide.sectionName;

  console.log(`[PPTX] Slide ${pageNum}: section="${sectionLabel}" layout="${layoutName}"`);

  // ② Header (shared — header only, no page-specific shapes)
  addHeader(sl, sectionLabel, layoutName);

  // ③ Page-specific content (strictly isolated per sectionKey)
  switch (selectedSlide.sectionKey) {
    case "brand":           renderBrand(sl,           selectedSlide.type); break;
    case "business_scheme": renderBusinessScheme(sl,  selectedSlide.type, selectedSlide.sectionKey); break;
    case "cases":           renderCases(sl,           selectedSlide.type); break;
    case "items":           renderItems(sl,           selectedSlide.type); break;
    case "schedule":        renderSchedule(sl,        selectedSlide.type, selectedSlide.sectionKey); break;
    case "pricing":         renderPricing(sl,         selectedSlide.type); break;
    case "company":         renderCompany(sl,         selectedSlide.type); break;
    default:
      console.log(`[PPTX] Unknown sectionKey: "${selectedSlide.sectionKey}" — generic placeholder`);
      box(sl, CX, CY, CW, CH, selectedSlide.sectionName);
  }

  // ④ Footer (shared — footer only, no page-specific shapes)
  addFooter(sl, pageNum, totalPages, date, companyLogo);
}

// ─── Filename helper ──────────────────────────────────────────────────────────
function buildFileName(title: string, date: string): string {
  const safeTitle = title
    .replace(/[/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 40) || "企画書";

  const m = date.match(/(\d{4})[^\d]?(\d{1,2})[^\d]?(\d{1,2})/);
  if (m) return `${safeTitle}_${m[1]}${m[2].padStart(2, "0")}${m[3].padStart(2, "0")}.pptx`;

  const n = new Date();
  return `${safeTitle}_${n.getFullYear()}${String(n.getMonth() + 1).padStart(2, "0")}${String(n.getDate()).padStart(2, "0")}.pptx`;
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generatePptx(options: GeneratePptxOptions): Promise<void> {
  const { default: PptxGenJS } = await import("pptxgenjs");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pptx = new (PptxGenJS as any)();
  pptx.layout = "LAYOUT_WIDE"; // 13.333 × 7.5 inches (16:9)

  const greetingSlide = options.selectedSlides.find((s) => s.sectionKey === "greeting");
  const coverType: PageType = (greetingSlide?.type === "B" || greetingSlide?.type === "C") ? "B" : "A";

  const contentSlides = options.selectedSlides.filter((s) => s.sectionKey !== "greeting");
  const totalPages    = contentSlides.length + 1;

  console.log(`[PPTX] Generating ${totalPages} slides`);

  // Slide 1: Cover
  renderCover(pptx, options.title, options.clientName, options.date, options.companyLogo, coverType);

  // Slides 2+: Content (each fully isolated)
  contentSlides.forEach((s, i) => {
    buildContentSlide(pptx, s, i + 2, totalPages, options.date, options.companyLogo);
  });

  // Download
  const fileName    = buildFileName(options.title, options.date);
  const base64Data  = (await pptx.write("base64")) as string;
  const mime        = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const a           = document.createElement("a");
  a.href            = `data:${mime};base64,${base64Data}`;
  a.setAttribute("download", fileName);
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
