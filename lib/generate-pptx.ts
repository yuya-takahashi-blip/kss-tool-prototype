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
const RED = "C0392B";
const DARK = "2C2C2C";
const MID = "666666";
const LIGHT_BG = "F5F5F5";
const PH = "DEDEDE"; // placeholder fill
const PH_BORDER = "C8C8C8";
const WHITE = "FFFFFF";
const RED_SOFT = "FEF2F2";
const FONT = "Arial";

// ─── Geometry ─────────────────────────────────────────────────────────────────
const SW = 10; // slide width
const SH = 5.625; // slide height
const M = 0.35; // margin
const HEADER_H = 0.65;
const FOOTER_Y = 5.2;
const FOOTER_H = SH - FOOTER_Y;
const CX = M;
const CY = HEADER_H + 0.18;
const CW = SW - M * 2;
const CH = FOOTER_Y - CY - 0.12;

// ─── Helpers ──────────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Slide = any;

function box(
  slide: Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  label = "",
  fill = PH,
  border = PH_BORDER,
  labelColor = "AAAAAA"
) {
  slide.addShape("rect", {
    x,
    y,
    w,
    h,
    fill: { color: fill },
    line: { color: border, width: 0.5 },
  });
  if (label) {
    slide.addText(label, {
      x,
      y,
      w,
      h,
      fontSize: 8.5,
      color: labelColor,
      fontFace: FONT,
      align: "center",
      valign: "middle",
    });
  }
}

function txt(
  slide: Slide,
  text: string,
  x: number,
  y: number,
  w: number,
  h: number,
  opts: Record<string, unknown> = {}
) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: FONT,
    color: DARK,
    fontSize: 10,
    valign: "top",
    ...opts,
  });
}

function arrow(slide: Slide, x: number, y: number, w: number, h: number) {
  // Arrow as a thin red rectangle pointing right
  slide.addShape("rightArrow", {
    x,
    y,
    w,
    h,
    fill: { color: RED },
    line: { color: RED },
  });
}

// ─── Slide shell (header + footer) ────────────────────────────────────────────
function addShell(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pptx: any,
  slide: Slide,
  sectionName: string,
  layoutName: string,
  pageNum: number,
  totalPages: number,
  date: string
) {
  // Header band
  slide.addShape("rect", {
    x: 0, y: 0, w: SW, h: HEADER_H,
    fill: { color: LIGHT_BG }, line: { color: LIGHT_BG },
  });
  // Red left accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 0.07, h: HEADER_H,
    fill: { color: RED }, line: { color: RED },
  });
  // Section name
  txt(slide, sectionName, 0.2, 0, 6.5, HEADER_H, {
    fontSize: 15, bold: true, color: DARK, valign: "middle",
  });
  // Layout badge background
  slide.addShape("rect", {
    x: SW - M - 2.3, y: 0.15, w: 2.3, h: 0.35,
    fill: { color: RED_SOFT }, line: { color: "FECACA", width: 0.5 },
  });
  // Layout badge text
  txt(slide, layoutName, SW - M - 2.3, 0.15, 2.3, 0.35, {
    fontSize: 8.5, color: RED, align: "center", valign: "middle",
  });

  // Footer separator
  slide.addShape("rect", {
    x: M, y: FOOTER_Y, w: CW, h: 0.01,
    fill: { color: "E0E0E0" }, line: { color: "E0E0E0" },
  });
  // Footer left: KANSAI SUPER STUDIO + date
  const footerText = date ? `KANSAI SUPER STUDIO　${date}` : "KANSAI SUPER STUDIO";
  txt(slide, footerText, M, FOOTER_Y + 0.05, SW - M * 2 - 1.6, FOOTER_H - 0.05, {
    fontSize: 7, color: "BBBBBB", valign: "middle", align: "center",
  });
  // Footer page number
  txt(slide, `${pageNum} / ${totalPages}`, SW - M - 1.5, FOOTER_Y + 0.05, 1.5, FOOTER_H - 0.05, {
    fontSize: 7.5, color: "BBBBBB", align: "right", valign: "middle",
  });

  void pptx;
}

// ─── Cover slide ──────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function addCoverSlide(pptx: any, title: string, clientName: string, date: string, companyLogo: string) {
  const slide = pptx.addSlide();

  // Top red accent strip
  slide.addShape("rect", {
    x: 0, y: 0, w: SW, h: 0.12,
    fill: { color: RED }, line: { color: RED },
  });
  // Bottom red accent strip
  slide.addShape("rect", {
    x: 0, y: SH - 0.12, w: SW, h: 0.12,
    fill: { color: RED }, line: { color: RED },
  });

  // ── Title (largest) ──────────────────────────────────────────────────────
  txt(slide, title || "企画書タイトル", M, 0.5, CW, 1.6, {
    fontSize: 32,
    bold: true,
    color: title ? DARK : "CCCCCC",
    align: "center",
    valign: "middle",
    wrap: true,
  });

  // Thin separator below title
  slide.addShape("rect", {
    x: SW / 2 - 2.0, y: 2.25, w: 4.0, h: 0.015,
    fill: { color: "E0E0E0" }, line: { color: "E0E0E0" },
  });

  // ── Client name (second largest) ─────────────────────────────────────────
  txt(slide, clientName || "提案先名", M, 2.35, CW, 0.9, {
    fontSize: 22,
    bold: true,
    color: clientName ? DARK : "CCCCCC",
    align: "center",
    valign: "middle",
  });

  // ── KANSAI SUPER STUDIO + date on one line (small, lower area) ───────────
  const logoLine = date
    ? `KANSAI SUPER STUDIO　${date}`
    : "KANSAI SUPER STUDIO";

  if (companyLogo) {
    // Image logo (small, bottom area)
    slide.addImage({
      data: companyLogo,
      x: SW / 2 - 1.2,
      y: 3.55,
      w: 2.4,
      h: 0.45,
    });
    txt(slide, date || "", M, 4.05, CW, 0.35, {
      fontSize: 9, color: "AAAAAA", align: "center", valign: "middle",
    });
  } else {
    txt(slide, logoLine, M, 3.6, CW, 0.5, {
      fontSize: 10,
      color: MID,
      align: "center",
      valign: "middle",
    });
  }

  void pptx;
}

// ─── Layout renderers ─────────────────────────────────────────────────────────

function renderGreeting(slide: Slide, type: PageType) {
  if (type === "A") {
    // Large text area
    box(slide, CX, CY, CW, CH, "挨拶文", WHITE, "E0E0E0");
    txt(slide, "（挨拶文が入ります）", CX + 0.3, CY + 0.3, CW - 0.6, CH - 0.9, {
      fontSize: 12,
      color: "BBBBBB",
      valign: "top",
      italic: true,
    });
    // Signature line
    box(slide, CX + CW - 3.5, CY + CH - 0.7, 3.5, 0.6, "署名・日付", WHITE, "E0E0E0");
  } else if (type === "B") {
    // Photo placeholder + text
    box(slide, CX, CY, 2.4, CH, "担当者写真", PH, PH_BORDER);
    box(slide, CX + 2.6, CY, CW - 2.6, CH - 0.7, "コメント本文", WHITE, "E0E0E0");
    box(slide, CX + 2.6, CY + CH - 0.6, CW - 2.6, 0.55, "氏名・役職", WHITE, "E0E0E0");
  } else {
    // Large quote style
    txt(slide, "❝", CX, CY, CW, 0.7, {
      fontSize: 36,
      color: RED,
      align: "center",
    });
    box(slide, CX, CY + 0.6, CW, CH - 1.1, "メッセージ本文", WHITE, "E0E0E0");
    txt(slide, "— 代表取締役　氏名", CX, CY + CH - 0.45, CW, 0.4, {
      fontSize: 10,
      color: MID,
      align: "right",
      italic: true,
    });
  }
}

function renderBrand(slide: Slide, type: PageType) {
  if (type === "A") {
    // Logo + text + key visual
    box(slide, CX, CY, 2.0, 1.0, "ロゴ", PH, PH_BORDER);
    box(slide, CX, CY + 1.15, 2.0, CH - 1.15, "ブランド説明文", WHITE, "E0E0E0");
    box(slide, CX + 2.2, CY, CW - 2.2, CH, "キービジュアル", PH, PH_BORDER);
  } else if (type === "B") {
    // Full-width image
    box(slide, CX, CY, CW, CH * 0.72, "世界観ビジュアル", PH, PH_BORDER);
    box(slide, CX, CY + CH * 0.75, CW, CH * 0.25, "キャッチコピー・説明文", WHITE, "E0E0E0");
  } else {
    // 3 feature columns
    const colW = (CW - 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (colW + 0.15);
      box(slide, cx2, CY, colW, 1.2, ["特徴", "ターゲット", "強み"][i], PH, PH_BORDER);
      box(slide, cx2, CY + 1.35, colW, CH - 1.35, "説明テキスト", WHITE, "E0E0E0");
    }
  }
}

function renderScheme(slide: Slide, type: PageType) {
  if (type === "A") {
    // 2 boxes + arrow
    const bw = 3.0;
    const bh = 1.8;
    const by = CY + (CH - bh) / 2;
    box(slide, CX, by, bw, bh, "自社", WHITE, RED, DARK);
    txt(slide, "自社", CX, by, bw, bh, { fontSize: 14, bold: true, align: "center", valign: "middle", color: DARK });
    arrow(slide, CX + bw + 0.15, by + bh / 2 - 0.2, 0.7, 0.4);
    box(slide, CX + bw + 1.0, by, bw, bh, "パートナー企業", WHITE, RED, DARK);
    txt(slide, "パートナー企業", CX + bw + 1.0, by, bw, bh, { fontSize: 14, bold: true, align: "center", valign: "middle", color: DARK });
    // Caption
    txt(slide, "取引の流れ・役割分担", CX, CY + CH - 0.45, CW, 0.4, { fontSize: 9, color: MID, align: "center" });
  } else if (type === "B") {
    // 3 boxes horizontal
    const bw = 2.5;
    const bh = 1.6;
    const by = CY + (CH - bh) / 2;
    const labels = ["自社", "中間事業者", "エンドユーザー"];
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (bw + 0.55);
      box(slide, cx2, by, bw, bh, labels[i], WHITE, RED, DARK);
      txt(slide, labels[i], cx2, by, bw, bh, { fontSize: 13, bold: true, align: "center", valign: "middle", color: DARK });
      if (i < 2) arrow(slide, cx2 + bw + 0.05, by + bh / 2 - 0.2, 0.45, 0.4);
    }
    txt(slide, "3者間スキームの流れ", CX, CY + CH - 0.45, CW, 0.4, { fontSize: 9, color: MID, align: "center" });
  } else {
    // Flow: 4 steps
    const steps = ["商品仕入れ", "プロモーション", "販売", "入金・精算"];
    const bw = (CW - 0.6) / 4 - 0.1;
    const bh = 1.8;
    const by = CY + (CH - bh) / 2;
    for (let i = 0; i < 4; i++) {
      const cx2 = CX + i * (bw + 0.3);
      box(slide, cx2, by, bw, bh, "", WHITE, RED, DARK);
      // Step number circle
      slide.addShape("ellipse", { x: cx2 + bw / 2 - 0.25, y: by + 0.15, w: 0.5, h: 0.5, fill: { color: RED }, line: { color: RED } });
      txt(slide, String(i + 1), cx2 + bw / 2 - 0.25, by + 0.15, 0.5, 0.5, { fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
      txt(slide, steps[i], cx2, by + 0.8, bw, 0.9, { fontSize: 10, align: "center", valign: "middle", color: DARK });
      if (i < 3) arrow(slide, cx2 + bw + 0.02, by + bh / 2 - 0.18, 0.24, 0.36);
    }
  }
}

function renderCases(slide: Slide, type: PageType) {
  if (type === "A") {
    // Single large case
    box(slide, CX, CY, CW * 0.45, CH, "導入事例イメージ", PH, PH_BORDER);
    const rx = CX + CW * 0.45 + 0.2;
    const rw = CW * 0.55 - 0.2;
    box(slide, rx, CY, rw, 0.7, "企業名・業種", WHITE, "E0E0E0");
    box(slide, rx, CY + 0.85, rw, CH - 0.85 - 0.75, "導入効果・コメント", WHITE, "E0E0E0");
    box(slide, rx, CY + CH - 0.7, rw, 0.65, "「お客様の声」", WHITE, "E0E0E0");
  } else if (type === "B") {
    // 3 cards horizontal
    const cw2 = (CW - 0.3) / 3;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (cw2 + 0.15);
      box(slide, cx2, CY, cw2, CH * 0.45, `事例 ${i + 1} 画像`, PH, PH_BORDER);
      box(slide, cx2, CY + CH * 0.48, cw2, CH * 0.52, `事例 ${i + 1} テキスト`, WHITE, "E0E0E0");
    }
  } else {
    // Before / After
    const half = (CW - 0.25) / 2;
    // Before
    box(slide, CX, CY, half, 0.45, "", RED, RED);
    txt(slide, "Before", CX, CY, half, 0.45, { fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(slide, CX, CY + 0.5, half, CH - 0.5, "課題・現状", WHITE, "E0E0E0");
    // After
    const ax = CX + half + 0.25;
    box(slide, ax, CY, half, 0.45, "", DARK, DARK);
    txt(slide, "After", ax, CY, half, 0.45, { fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(slide, ax, CY + 0.5, half, CH - 0.5, "改善後の状態・効果", WHITE, "E0E0E0");
  }
}

function renderItems(slide: Slide, type: PageType) {
  if (type === "A") {
    // Large image + description
    box(slide, CX, CY, CW * 0.55, CH, "商品画像", PH, PH_BORDER);
    const rx = CX + CW * 0.55 + 0.2;
    const rw = CW * 0.45 - 0.2;
    box(slide, rx, CY, rw, 0.65, "商品名", WHITE, "E0E0E0");
    box(slide, rx, CY + 0.8, rw, 0.5, "価格", WHITE, "E0E0E0");
    box(slide, rx, CY + 1.45, rw, CH - 1.45, "商品説明", WHITE, "E0E0E0");
  } else if (type === "B") {
    // 2x2 grid
    const gw = (CW - 0.2) / 2;
    const gh = (CH - 0.2) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) {
        box(slide, CX + c * (gw + 0.2), CY + r * (gh + 0.2), gw, gh, `商品 ${r * 2 + c + 1}`, PH, PH_BORDER);
      }
    }
  } else {
    // Scene image top + product bottom right
    box(slide, CX, CY, CW, CH * 0.58, "使用シーン画像", PH, PH_BORDER);
    box(slide, CX, CY + CH * 0.62, CW * 0.55, CH * 0.38, "商品説明", WHITE, "E0E0E0");
    box(slide, CX + CW * 0.58, CY + CH * 0.62, CW * 0.42, CH * 0.38, "商品画像", PH, PH_BORDER);
  }
}

function renderSchedule(slide: Slide, type: PageType) {
  if (type === "A") {
    // Horizontal timeline
    const milestones = ["Phase 1", "Phase 2", "Phase 3", "Phase 4"];
    const mw = CW / milestones.length;
    // Timeline bar
    slide.addShape("rect", { x: CX, y: CY + CH / 2 - 0.04, w: CW, h: 0.08, fill: { color: "DDDDDD" }, line: { color: "DDDDDD" } });
    milestones.forEach((m, i) => {
      const mx = CX + mw * i + mw / 2;
      // Dot
      slide.addShape("ellipse", { x: mx - 0.2, y: CY + CH / 2 - 0.2, w: 0.4, h: 0.4, fill: { color: RED }, line: { color: RED } });
      // Label above/below alternating
      const isAbove = i % 2 === 0;
      txt(slide, m, mx - mw / 2 + 0.1, isAbove ? CY + 0.15 : CY + CH / 2 + 0.4, mw - 0.2, 0.9, { fontSize: 10, bold: true, align: "center", valign: "middle", color: DARK });
      box(slide, mx - mw / 2 + 0.15, isAbove ? CY + CH / 2 + 0.4 : CY + 0.15, mw - 0.3, CH / 2 - 0.65, "タスク内容", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    // Gantt-style table
    const months = ["4月", "5月", "6月", "7月", "8月", "9月"];
    const tasks = ["企画・設計", "開発", "テスト", "リリース"];
    const cellW = (CW - 1.8) / months.length;
    const rowH = (CH - 0.5) / (tasks.length + 1);
    // Header row
    box(slide, CX, CY, 1.8, rowH, "タスク", LIGHT_BG, "E0E0E0", MID);
    months.forEach((m, i) => box(slide, CX + 1.8 + i * cellW, CY, cellW, rowH, m, LIGHT_BG, "E0E0E0", MID));
    // Task rows
    tasks.forEach((task, row) => {
      const ry = CY + (row + 1) * rowH;
      box(slide, CX, ry, 1.8, rowH, task, WHITE, "E0E0E0", DARK);
      // Gantt bar (spans 2 cells)
      const start = row;
      box(slide, CX + 1.8 + start * cellW, ry + 0.08, cellW * 2.5, rowH - 0.16, "", RED, RED);
      // Remaining cells empty
      months.forEach((_, i) => {
        if (i < start || i >= start + 2.5) {
          slide.addShape("rect", { x: CX + 1.8 + i * cellW, y: ry, w: cellW, h: rowH, fill: { color: WHITE }, line: { color: "E8E8E8", width: 0.5 } });
        }
      });
    });
  } else {
    // Phase steps
    const phases = [
      { name: "Phase 1\n準備期間", sub: "1〜2ヶ月" },
      { name: "Phase 2\n試験運用", sub: "2〜3ヶ月" },
      { name: "Phase 3\n本格展開", sub: "3〜6ヶ月" },
    ];
    const bw = (CW - 0.4) / 3;
    phases.forEach((p, i) => {
      const cx2 = CX + i * (bw + 0.2);
      box(slide, cx2, CY, bw, 0.55, "", i === 1 ? RED : LIGHT_BG, "CCCCCC");
      txt(slide, `STEP ${i + 1}`, cx2, CY, bw, 0.55, { fontSize: 11, bold: true, align: "center", valign: "middle", color: i === 1 ? WHITE : MID });
      box(slide, cx2, CY + 0.65, bw, CH - 0.65, p.name, WHITE, "E0E0E0", DARK);
      txt(slide, p.name, cx2, CY + 0.8, bw, 1.0, { fontSize: 12, bold: true, align: "center", valign: "top", color: DARK });
      txt(slide, p.sub, cx2, CY + 2.0, bw, 0.5, { fontSize: 9, align: "center", valign: "middle", color: MID });
      box(slide, cx2 + 0.2, CY + 2.6, bw - 0.4, CH - 2.9, "主なタスク", WHITE, "E0E0E0");
    });
  }
}

function renderPricing(slide: Slide, type: PageType) {
  if (type === "A") {
    // Condition table
    const rows = ["掛け率", "最低発注数", "支払い条件", "契約期間", "専属条件", "備考"];
    const rowH = CH / rows.length;
    rows.forEach((r, i) => {
      box(slide, CX, CY + i * rowH, 2.0, rowH, r, i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(slide, CX + 2.1, CY + i * rowH, CW - 2.1, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    // 3 plan columns
    const plans = ["スタンダード", "スタンダード＋", "プレミアム"];
    const bw = (CW - 0.3) / 3;
    plans.forEach((p, i) => {
      const cx2 = CX + i * (bw + 0.15);
      const isHighlight = i === 1;
      box(slide, cx2, CY, bw, 0.65, "", isHighlight ? RED : LIGHT_BG, isHighlight ? RED : "E0E0E0");
      txt(slide, p, cx2, CY, bw, 0.65, { fontSize: 11, bold: true, align: "center", valign: "middle", color: isHighlight ? WHITE : DARK });
      box(slide, cx2, CY + 0.75, bw, CH - 0.75, "条件詳細", WHITE, isHighlight ? RED : "E0E0E0");
    });
  } else {
    // Royalty conditions — text-heavy
    box(slide, CX, CY, CW, 0.65, "", LIGHT_BG, "E0E0E0");
    txt(slide, "ロイヤリティ条件", CX + 0.2, CY, CW - 0.4, 0.65, { fontSize: 13, bold: true, valign: "middle", color: DARK });
    const sections2 = ["基本ロイヤリティ率", "インセンティブ条件", "最低保証", "支払いスケジュール"];
    const sh2 = (CH - 0.75) / sections2.length;
    sections2.forEach((s, i) => {
      const sy = CY + 0.75 + i * sh2;
      slide.addShape("rect", { x: CX, y: sy + 0.08, w: 0.04, h: sh2 - 0.16, fill: { color: RED }, line: { color: RED } });
      txt(slide, s, CX + 0.15, sy, CW * 0.3, sh2, { fontSize: 10, bold: true, valign: "middle", color: DARK });
      box(slide, CX + CW * 0.3 + 0.1, sy + 0.08, CW * 0.7 - 0.1, sh2 - 0.16, "内容", WHITE, "E0E0E0");
    });
  }
}

function renderCompany(slide: Slide, type: PageType) {
  if (type === "A") {
    // Basic info table
    const fields = ["会社名", "代表者", "設立", "所在地", "事業内容", "資本金"];
    const rowH = CH / fields.length;
    fields.forEach((f, i) => {
      box(slide, CX, CY + i * rowH, 1.8, rowH, f, i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(slide, CX + 1.9, CY + i * rowH, CW - 1.9, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    // Info + timeline
    const infoH = CH * 0.42;
    const fields = ["会社名", "代表者", "所在地"];
    const rowH = infoH / fields.length;
    fields.forEach((f, i) => {
      box(slide, CX, CY + i * rowH, 1.6, rowH, f, LIGHT_BG, "E0E0E0", MID);
      box(slide, CX + 1.7, CY + i * rowH, CW - 1.7, rowH, "—", WHITE, "E0E0E0");
    });
    // Timeline section
    const tly = CY + infoH + 0.2;
    txt(slide, "沿革・主な実績", CX, tly, CW, 0.4, { fontSize: 11, bold: true, color: DARK });
    slide.addShape("rect", { x: CX + 0.3, y: tly + 0.45, w: 0.04, h: CH - infoH - 0.7, fill: { color: "DDDDDD" }, line: { color: "DDDDDD" } });
    for (let i = 0; i < 3; i++) {
      const ey = tly + 0.5 + i * 0.75;
      slide.addShape("ellipse", { x: CX + 0.18, y: ey, w: 0.25, h: 0.25, fill: { color: RED }, line: { color: RED } });
      box(slide, CX + 0.6, ey, CW - 0.6, 0.6, `実績 ${i + 1}`, WHITE, "E0E0E0");
    }
  } else {
    // Strength cards
    const items = ["強み・特徴", "主力サービス", "実績・数値"];
    const bw = (CW - 0.3) / 3;
    items.forEach((item, i) => {
      const cx2 = CX + i * (bw + 0.15);
      slide.addShape("rect", { x: cx2, y: CY, w: bw, h: 0.06, fill: { color: RED }, line: { color: RED } });
      box(slide, cx2, CY + 0.1, bw, 0.55, "", LIGHT_BG, "E0E0E0");
      txt(slide, item, cx2, CY + 0.1, bw, 0.55, { fontSize: 11, bold: true, align: "center", valign: "middle", color: DARK });
      box(slide, cx2, CY + 0.75, bw, CH - 0.75, "詳細内容", WHITE, "E0E0E0");
    });
  }
}

// ─── Dispatch ─────────────────────────────────────────────────────────────────
function addContentSlide(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pptx: any,
  selectedSlide: SelectedSlide,
  pageNum: number,
  totalPages: number,
  date: string
) {
  const slide = pptx.addSlide();
  const layoutName = getLayoutName(selectedSlide.sectionKey, selectedSlide.type);
  const sectionLabel =
    selectedSlide.slideIndexInSection > 1
      ? `${selectedSlide.sectionName}（${selectedSlide.slideIndexInSection}）`
      : selectedSlide.sectionName;

  addShell(pptx, slide, sectionLabel, layoutName, pageNum, totalPages, date);

  switch (selectedSlide.sectionKey) {
    case "greeting": renderGreeting(slide, selectedSlide.type); break;
    case "brand": renderBrand(slide, selectedSlide.type); break;
    case "business_scheme": renderScheme(slide, selectedSlide.type); break;
    case "cases": renderCases(slide, selectedSlide.type); break;
    case "items": renderItems(slide, selectedSlide.type); break;
    case "schedule": renderSchedule(slide, selectedSlide.type); break;
    case "pricing": renderPricing(slide, selectedSlide.type); break;
    case "company": renderCompany(slide, selectedSlide.type); break;
    default: box(slide, CX, CY, CW, CH, selectedSlide.sectionName);
  }
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generatePptx(options: GeneratePptxOptions): Promise<void> {
  const { default: PptxGenJS } = await import("pptxgenjs");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pptx = new (PptxGenJS as any)();

  pptx.layout = "LAYOUT_WIDE"; // 16:9

  // Total slides = cover + content slides
  const totalPages = options.selectedSlides.length + 1;

  // Cover
  addCoverSlide(pptx, options.title, options.clientName, options.date, options.companyLogo);

  // Content slides in selectedSlides order
  options.selectedSlides.forEach((s, i) => {
    addContentSlide(pptx, s, i + 2, totalPages, options.date);
  });

  // File name: proposal-YYYYMMDD.pptx
  const now = new Date();
  const ymd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
  const fileName = `proposal-${ymd}.pptx`;

  // Cross-browser download (includes iOS Safari)
  const base64Data = (await pptx.write("base64")) as string;
  const mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const a = document.createElement("a");
  a.href = `data:${mime};base64,${base64Data}`;
  a.setAttribute("download", fileName);
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
