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
const RED       = "C0392B";
const DARK      = "2C2C2C";
const MID       = "666666";
const LIGHT_BG  = "F5F5F5";
const PH        = "DEDEDE";
const PH_BORDER = "C8C8C8";
const WHITE     = "FFFFFF";
const RED_SOFT  = "FEF2F2";
const FONT      = "Arial";

// ─── Geometry (LAYOUT_WIDE = 13.333 × 7.5 inches) ────────────────────────────
const SW       = 13.333; // slide width
const SH       = 7.5;    // slide height
const M        = 0.5;    // horizontal margin
const HEADER_H = 0.8;    // header band height
const FOOTER_Y = 6.9;    // footer top y
const FOOTER_H = SH - FOOTER_Y; // 0.6
const CX       = M;                        // content left
const CY       = HEADER_H + 0.22;          // content top  (1.02)
const CW       = SW - M * 2;               // content width (12.333)
const CH       = FOOTER_Y - CY - 0.16;    // content height (5.72)

// ─── Helpers ──────────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Slide = any;

/** Shape-only box with optional centered label at small font size (placeholder) */
function box(
  slide: Slide,
  x: number, y: number, w: number, h: number,
  label = "",
  fill  = PH,
  border = PH_BORDER,
  labelColor = "AAAAAA"
) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: fill },
    line: { color: border, width: 0.5 },
  });
  if (label) {
    slide.addText(label, {
      x, y, w, h,
      fontSize: 9,
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
  x: number, y: number, w: number, h: number,
  opts: Record<string, unknown> = {}
) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT,
    color: DARK,
    fontSize: 10,
    valign: "top",
    ...opts,
  });
}

function arrow(slide: Slide, x: number, y: number, w: number, h: number) {
  slide.addShape("rightArrow", {
    x, y, w, h,
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
  date: string,
  companyLogo: string
) {
  // Header band (full width)
  slide.addShape("rect", {
    x: 0, y: 0, w: SW, h: HEADER_H,
    fill: { color: LIGHT_BG }, line: { color: LIGHT_BG },
  });
  // Red left accent bar
  slide.addShape("rect", {
    x: 0, y: 0, w: 0.09, h: HEADER_H,
    fill: { color: RED }, line: { color: RED },
  });
  // Section name
  txt(slide, sectionName, 0.25, 0, SW - M - 3.0, HEADER_H, {
    fontSize: 16, bold: true, color: DARK, valign: "middle",
  });
  // Layout badge background
  slide.addShape("rect", {
    x: SW - M - 2.8, y: 0.2, w: 2.8, h: 0.38,
    fill: { color: RED_SOFT }, line: { color: "FECACA", width: 0.5 },
  });
  // Layout badge text
  txt(slide, layoutName, SW - M - 2.8, 0.2, 2.8, 0.38, {
    fontSize: 9, color: RED, align: "center", valign: "middle",
  });

  // Footer separator line (full content width)
  slide.addShape("rect", {
    x: CX, y: FOOTER_Y, w: CW, h: 0.01,
    fill: { color: "E0E0E0" }, line: { color: "E0E0E0" },
  });

  // Footer center: logo image or KANSAI SUPER STUDIO + date
  const footerLogoW = CW - 2.2; // leave room for page number
  const footerLogoX = CX;
  if (companyLogo) {
    try {
      const logoH = FOOTER_H - 0.14;
      const logoW = logoH * 3.5; // assume roughly 3.5:1 aspect
      const logoX = SW / 2 - logoW / 2;
      slide.addImage({
        data: companyLogo,
        x: logoX,
        y: FOOTER_Y + 0.08,
        w: logoW,
        h: logoH,
      });
      if (date) {
        txt(slide, date, logoX + logoW + 0.15, FOOTER_Y + 0.06, 1.8, FOOTER_H - 0.08, {
          fontSize: 7, color: "BBBBBB", valign: "middle",
        });
      }
    } catch {
      // Fallback to text
      const footerText = date ? `KANSAI SUPER STUDIO　${date}` : "KANSAI SUPER STUDIO";
      txt(slide, footerText, footerLogoX, FOOTER_Y + 0.06, footerLogoW, FOOTER_H - 0.08, {
        fontSize: 7.5, color: "BBBBBB", valign: "middle", align: "center",
      });
    }
  } else {
    const footerText = date ? `KANSAI SUPER STUDIO　${date}` : "KANSAI SUPER STUDIO";
    txt(slide, footerText, footerLogoX, FOOTER_Y + 0.06, footerLogoW, FOOTER_H - 0.08, {
      fontSize: 7.5, color: "BBBBBB", valign: "middle", align: "center",
    });
  }

  // Page number (right-aligned, fits within SW)
  txt(slide, `${pageNum} / ${totalPages}`, SW - M - 1.8, FOOTER_Y + 0.06, 1.8, FOOTER_H - 0.08, {
    fontSize: 8, color: "BBBBBB", align: "right", valign: "middle",
  });

  void pptx;
}

// ─── Cover slide ──────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function addCoverSlide(
  pptx: any,
  title: string,
  clientName: string,
  date: string,
  companyLogo: string,
  coverType: PageType
) {
  const slide = pptx.addSlide();

  // Custom cover: essentially blank — user edits in PowerPoint
  if (coverType === "B" || coverType === "C") {
    slide.addShape("rect", {
      x: 0, y: 0, w: SW, h: 0.15,
      fill: { color: RED }, line: { color: RED },
    });
    slide.addShape("rect", {
      x: 0, y: SH - 0.15, w: SW, h: 0.15,
      fill: { color: RED }, line: { color: RED },
    });
    void pptx;
    return;
  }

  // ── Standard cover (A) ────────────────────────────────────────────────────

  // Top red accent strip (full width)
  slide.addShape("rect", {
    x: 0, y: 0, w: SW, h: 0.15,
    fill: { color: RED }, line: { color: RED },
  });
  // Bottom red accent strip (full width)
  slide.addShape("rect", {
    x: 0, y: SH - 0.15, w: SW, h: 0.15,
    fill: { color: RED }, line: { color: RED },
  });

  // Title (largest) — centered vertically in upper half
  txt(slide, title || "企画書タイトル", M, 1.2, CW, 2.2, {
    fontSize: 36,
    bold: true,
    color: title ? DARK : "CCCCCC",
    align: "center",
    valign: "middle",
    wrap: true,
  });

  // Thin separator
  slide.addShape("rect", {
    x: SW / 2 - 2.5, y: 3.6, w: 5.0, h: 0.02,
    fill: { color: "E0E0E0" }, line: { color: "E0E0E0" },
  });

  // Client name (second largest)
  txt(slide, clientName || "提案先会社名", M, 3.7, CW, 1.4, {
    fontSize: 26,
    bold: true,
    color: clientName ? DARK : "CCCCCC",
    align: "center",
    valign: "middle",
  });

  // KANSAI SUPER STUDIO / logo + date (small, lower area)
  if (companyLogo) {
    const logoH = 0.55;
    const logoW = logoH * 3.5;
    const logoX = SW / 2 - logoW / 2 - 0.5;
    try {
      slide.addImage({
        data: companyLogo,
        x: logoX,
        y: 5.5,
        w: logoW,
        h: logoH,
      });
    } catch { /* ignore image errors */ }
    if (date) {
      txt(slide, date, logoX + logoW + 0.2, 5.5, 2.5, logoH, {
        fontSize: 11, color: MID, valign: "middle",
      });
    }
  } else {
    const logoLine = date ? `KANSAI SUPER STUDIO　${date}` : "KANSAI SUPER STUDIO";
    txt(slide, logoLine, M, 5.5, CW, 0.55, {
      fontSize: 11,
      color: MID,
      align: "center",
      valign: "middle",
    });
  }

  void pptx;
}

// ─── Layout renderers ─────────────────────────────────────────────────────────

function renderBrand(slide: Slide, type: PageType) {
  if (type === "A") {
    const leftW = 2.8;
    const gap   = 0.3;
    const rightW = CW - leftW - gap;
    box(slide, CX, CY, leftW, 1.3, "ロゴ", PH, PH_BORDER);
    box(slide, CX, CY + 1.5, leftW, CH - 1.5, "ブランド説明文", WHITE, "E0E0E0");
    box(slide, CX + leftW + gap, CY, rightW, CH, "キービジュアル", PH, PH_BORDER);
  } else if (type === "B") {
    box(slide, CX, CY, CW, CH * 0.72, "世界観ビジュアル", PH, PH_BORDER);
    box(slide, CX, CY + CH * 0.75, CW, CH * 0.25, "キャッチコピー・説明文", WHITE, "E0E0E0");
  } else {
    const colW = (CW - 0.4) / 3;
    const gap  = 0.2;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (colW + gap);
      box(slide, cx2, CY, colW, 1.6, ["特徴", "ターゲット", "強み"][i], PH, PH_BORDER);
      box(slide, cx2, CY + 1.8, colW, CH - 1.8, "説明テキスト", WHITE, "E0E0E0");
    }
  }
}

function renderScheme(slide: Slide, type: PageType) {
  if (type === "A") {
    // 2 boxes + arrow — centered horizontally
    const bw = 4.4;
    const bh = 2.6;
    const arrowW = 1.0;
    const arrowH = 0.55;
    const totalSpan = bw * 2 + arrowW + 0.4; // boxes + arrow + gaps
    const startX = CX + (CW - totalSpan) / 2;
    const by = CY + (CH - bh) / 2;

    // Box 1: 自社
    slide.addShape("rect", {
      x: startX, y: by, w: bw, h: bh,
      fill: { color: WHITE }, line: { color: RED, width: 1.5 },
    });
    txt(slide, "自社", startX, by, bw, bh, {
      fontSize: 18, bold: true, align: "center", valign: "middle", color: DARK,
    });

    // Arrow
    arrow(slide, startX + bw + 0.2, by + bh / 2 - arrowH / 2, arrowW, arrowH);

    // Box 2: パートナー企業
    const box2X = startX + bw + arrowW + 0.4;
    slide.addShape("rect", {
      x: box2X, y: by, w: bw, h: bh,
      fill: { color: WHITE }, line: { color: RED, width: 1.5 },
    });
    txt(slide, "パートナー企業", box2X, by, bw, bh, {
      fontSize: 18, bold: true, align: "center", valign: "middle", color: DARK,
    });

    txt(slide, "取引の流れ・役割分担", CX, CY + CH - 0.55, CW, 0.45, {
      fontSize: 10, color: MID, align: "center",
    });

  } else if (type === "B") {
    // 3 boxes horizontal — centered
    const bw = 3.4;
    const bh = 2.2;
    const arrowW = 0.6;
    const arrowH = 0.5;
    const totalSpan = bw * 3 + arrowW * 2 + 0.4 * 4; // boxes + gaps
    const startX = CX + (CW - (bw * 3 + 2 * (arrowW + 0.3))) / 2;
    const by = CY + (CH - bh) / 2;
    const labels = ["自社", "中間事業者", "エンドユーザー"];

    for (let i = 0; i < 3; i++) {
      const cx2 = startX + i * (bw + arrowW + 0.3);
      slide.addShape("rect", {
        x: cx2, y: by, w: bw, h: bh,
        fill: { color: WHITE }, line: { color: RED, width: 1.5 },
      });
      txt(slide, labels[i], cx2, by, bw, bh, {
        fontSize: 16, bold: true, align: "center", valign: "middle", color: DARK,
      });
      if (i < 2) {
        arrow(slide, cx2 + bw + 0.08, by + bh / 2 - arrowH / 2, arrowW, arrowH);
      }
    }
    txt(slide, "3者間スキームの流れ", CX, CY + CH - 0.55, CW, 0.45, {
      fontSize: 10, color: MID, align: "center",
    });
    void totalSpan;

  } else {
    // Flow: 4 steps
    const steps = ["商品仕入れ", "プロモーション", "販売", "入金・精算"];
    const arrowW = 0.35;
    const gap    = arrowW + 0.05;
    const bw     = (CW - gap * 3) / 4;
    const bh     = 2.4;
    const by     = CY + (CH - bh) / 2;

    for (let i = 0; i < 4; i++) {
      const cx2 = CX + i * (bw + gap);
      slide.addShape("rect", {
        x: cx2, y: by, w: bw, h: bh,
        fill: { color: WHITE }, line: { color: RED, width: 1.2 },
      });
      // Step number circle
      const circleR = 0.35;
      slide.addShape("ellipse", {
        x: cx2 + bw / 2 - circleR, y: by + 0.2,
        w: circleR * 2, h: circleR * 2,
        fill: { color: RED }, line: { color: RED },
      });
      txt(slide, String(i + 1), cx2 + bw / 2 - circleR, by + 0.2, circleR * 2, circleR * 2, {
        fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle",
      });
      txt(slide, steps[i], cx2, by + 1.05, bw, 1.1, {
        fontSize: 12, align: "center", valign: "middle", color: DARK, wrap: true,
      });
      if (i < 3) {
        arrow(slide, cx2 + bw + 0.03, by + bh / 2 - 0.22, arrowW - 0.05, 0.44);
      }
    }
  }
}

function renderCases(slide: Slide, type: PageType) {
  if (type === "A") {
    const imgW = CW * 0.42;
    const textX = CX + imgW + 0.25;
    const textW = CW - imgW - 0.25;
    box(slide, CX, CY, imgW, CH, "導入事例イメージ", PH, PH_BORDER);
    box(slide, textX, CY, textW, 0.85, "企業名・業種", WHITE, "E0E0E0");
    box(slide, textX, CY + 1.0, textW, CH - 1.0 - 0.85, "導入効果・コメント", WHITE, "E0E0E0");
    box(slide, textX, CY + CH - 0.8, textW, 0.75, "「お客様の声」", WHITE, "E0E0E0");
  } else if (type === "B") {
    const colW = (CW - 0.4) / 3;
    const gap  = 0.2;
    for (let i = 0; i < 3; i++) {
      const cx2 = CX + i * (colW + gap);
      box(slide, cx2, CY, colW, CH * 0.45, `事例 ${i + 1} 画像`, PH, PH_BORDER);
      box(slide, cx2, CY + CH * 0.48, colW, CH * 0.52, `事例 ${i + 1} テキスト`, WHITE, "E0E0E0");
    }
  } else {
    const half = (CW - 0.35) / 2;
    // Before
    slide.addShape("rect", {
      x: CX, y: CY, w: half, h: 0.55,
      fill: { color: RED }, line: { color: RED },
    });
    txt(slide, "Before", CX, CY, half, 0.55, {
      fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle",
    });
    box(slide, CX, CY + 0.65, half, CH - 0.65, "課題・現状", WHITE, "E0E0E0");
    // After
    const ax = CX + half + 0.35;
    slide.addShape("rect", {
      x: ax, y: CY, w: half, h: 0.55,
      fill: { color: DARK }, line: { color: DARK },
    });
    txt(slide, "After", ax, CY, half, 0.55, {
      fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle",
    });
    box(slide, ax, CY + 0.65, half, CH - 0.65, "改善後の状態・効果", WHITE, "E0E0E0");
  }
}

function renderItems(slide: Slide, type: PageType) {
  if (type === "A") {
    const imgW = CW * 0.52;
    const textX = CX + imgW + 0.25;
    const textW = CW - imgW - 0.25;
    box(slide, CX, CY, imgW, CH, "商品画像", PH, PH_BORDER);
    box(slide, textX, CY, textW, 0.8, "商品名", WHITE, "E0E0E0");
    box(slide, textX, CY + 0.95, textW, 0.65, "価格", WHITE, "E0E0E0");
    box(slide, textX, CY + 1.75, textW, CH - 1.75, "商品説明", WHITE, "E0E0E0");
  } else if (type === "B") {
    const gw = (CW - 0.25) / 2;
    const gh = (CH - 0.25) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) {
        box(slide, CX + c * (gw + 0.25), CY + r * (gh + 0.25), gw, gh, `商品 ${r * 2 + c + 1}`, PH, PH_BORDER);
      }
    }
  } else {
    box(slide, CX, CY, CW, CH * 0.58, "使用シーン画像", PH, PH_BORDER);
    box(slide, CX, CY + CH * 0.62, CW * 0.55, CH * 0.38, "商品説明", WHITE, "E0E0E0");
    box(slide, CX + CW * 0.58, CY + CH * 0.62, CW * 0.42, CH * 0.38, "商品画像", PH, PH_BORDER);
  }
}

function renderSchedule(slide: Slide, type: PageType) {
  if (type === "A") {
    const milestones = ["Phase 1", "Phase 2", "Phase 3", "Phase 4"];
    const mw = CW / milestones.length;
    // Timeline bar
    slide.addShape("rect", {
      x: CX, y: CY + CH / 2 - 0.05, w: CW, h: 0.1,
      fill: { color: "DDDDDD" }, line: { color: "DDDDDD" },
    });
    milestones.forEach((m, i) => {
      const mx = CX + mw * i + mw / 2;
      slide.addShape("ellipse", {
        x: mx - 0.25, y: CY + CH / 2 - 0.25,
        w: 0.5, h: 0.5,
        fill: { color: RED }, line: { color: RED },
      });
      const isAbove = i % 2 === 0;
      txt(slide, m, mx - mw / 2 + 0.1, isAbove ? CY + 0.2 : CY + CH / 2 + 0.45, mw - 0.2, 1.0, {
        fontSize: 11, bold: true, align: "center", valign: "middle", color: DARK,
      });
      box(slide, mx - mw / 2 + 0.2, isAbove ? CY + CH / 2 + 0.45 : CY + 0.2, mw - 0.4, CH / 2 - 0.75, "タスク内容", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    const months  = ["4月", "5月", "6月", "7月", "8月", "9月"];
    const tasks   = ["企画・設計", "開発", "テスト", "リリース"];
    const labelW  = 2.4;
    const cellW   = (CW - labelW) / months.length;
    const rowH    = CH / (tasks.length + 1);
    // Header row
    box(slide, CX, CY, labelW, rowH, "タスク", LIGHT_BG, "E0E0E0", MID);
    months.forEach((m, i) => {
      box(slide, CX + labelW + i * cellW, CY, cellW, rowH, m, LIGHT_BG, "E0E0E0", MID);
    });
    // Task rows
    tasks.forEach((task, row) => {
      const ry = CY + (row + 1) * rowH;
      box(slide, CX, ry, labelW, rowH, task, WHITE, "E0E0E0", DARK);
      const start = row;
      box(slide, CX + labelW + start * cellW, ry + 0.1, cellW * 2.5, rowH - 0.2, "", RED, RED);
      months.forEach((_, i) => {
        if (i < start || i >= start + 2.5) {
          slide.addShape("rect", {
            x: CX + labelW + i * cellW, y: ry, w: cellW, h: rowH,
            fill: { color: WHITE }, line: { color: "E8E8E8", width: 0.5 },
          });
        }
      });
    });
  } else {
    const phases = [
      { name: "準備期間", sub: "1〜2ヶ月" },
      { name: "試験運用", sub: "2〜3ヶ月" },
      { name: "本格展開", sub: "3〜6ヶ月" },
    ];
    const bw = (CW - 0.5) / 3;
    const gap = 0.25;
    phases.forEach((p, i) => {
      const cx2 = CX + i * (bw + gap);
      const isHighlight = i === 1;
      // Step label bar
      slide.addShape("rect", {
        x: cx2, y: CY, w: bw, h: 0.65,
        fill: { color: isHighlight ? RED : LIGHT_BG },
        line: { color: isHighlight ? RED : "CCCCCC", width: 0.5 },
      });
      txt(slide, `STEP ${i + 1}`, cx2, CY, bw, 0.65, {
        fontSize: 13, bold: true, align: "center", valign: "middle",
        color: isHighlight ? WHITE : MID,
      });
      // Content box — no duplicate label
      slide.addShape("rect", {
        x: cx2, y: CY + 0.8, w: bw, h: CH - 0.8,
        fill: { color: WHITE }, line: { color: "E0E0E0", width: 0.5 },
      });
      txt(slide, p.name, cx2, CY + 1.0, bw, 1.2, {
        fontSize: 14, bold: true, align: "center", valign: "top", color: DARK,
      });
      txt(slide, p.sub, cx2, CY + 2.4, bw, 0.6, {
        fontSize: 11, align: "center", valign: "middle", color: MID,
      });
      box(slide, cx2 + 0.25, CY + 3.2, bw - 0.5, CH - 3.4, "主なタスク", WHITE, "E0E0E0");
    });
  }
}

function renderPricing(slide: Slide, type: PageType) {
  if (type === "A") {
    const rows = ["掛け率", "最低発注数", "支払い条件", "契約期間", "専属条件", "備考"];
    const labelW = 2.6;
    const rowH   = CH / rows.length;
    rows.forEach((r, i) => {
      box(slide, CX, CY + i * rowH, labelW, rowH, r, i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(slide, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    const plans = ["スタンダード", "スタンダード＋", "プレミアム"];
    const bw    = (CW - 0.4) / 3;
    const gap   = 0.2;
    plans.forEach((p, i) => {
      const cx2 = CX + i * (bw + gap);
      const isHighlight = i === 1;
      slide.addShape("rect", {
        x: cx2, y: CY, w: bw, h: 0.8,
        fill: { color: isHighlight ? RED : LIGHT_BG },
        line: { color: isHighlight ? RED : "E0E0E0", width: 0.5 },
      });
      txt(slide, p, cx2, CY, bw, 0.8, {
        fontSize: 13, bold: true, align: "center", valign: "middle",
        color: isHighlight ? WHITE : DARK,
      });
      box(slide, cx2, CY + 0.95, bw, CH - 0.95, "条件詳細", WHITE, isHighlight ? RED : "E0E0E0");
    });
  } else {
    slide.addShape("rect", {
      x: CX, y: CY, w: CW, h: 0.75,
      fill: { color: LIGHT_BG }, line: { color: "E0E0E0", width: 0.5 },
    });
    txt(slide, "ロイヤリティ条件", CX + 0.25, CY, CW - 0.5, 0.75, {
      fontSize: 15, bold: true, valign: "middle", color: DARK,
    });
    const sections2 = ["基本ロイヤリティ率", "インセンティブ条件", "最低保証", "支払いスケジュール"];
    const rowH2 = (CH - 0.85) / sections2.length;
    sections2.forEach((s, i) => {
      const sy = CY + 0.85 + i * rowH2;
      slide.addShape("rect", {
        x: CX, y: sy + 0.1, w: 0.05, h: rowH2 - 0.2,
        fill: { color: RED }, line: { color: RED },
      });
      txt(slide, s, CX + 0.18, sy, CW * 0.28, rowH2, {
        fontSize: 11, bold: true, valign: "middle", color: DARK,
      });
      box(slide, CX + CW * 0.3 + 0.12, sy + 0.1, CW * 0.7 - 0.12, rowH2 - 0.2, "内容", WHITE, "E0E0E0");
    });
  }
}

function renderCompany(slide: Slide, type: PageType) {
  if (type === "A") {
    const fields = ["会社名", "代表者", "設立", "所在地", "事業内容", "資本金"];
    const labelW = 2.4;
    const rowH   = CH / fields.length;
    fields.forEach((f, i) => {
      box(slide, CX, CY + i * rowH, labelW, rowH, f, i % 2 === 0 ? LIGHT_BG : WHITE, "E0E0E0", MID);
      box(slide, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
  } else if (type === "B") {
    const labelW = 2.1;
    const infoH  = CH * 0.42;
    const fields = ["会社名", "代表者", "所在地"];
    const rowH   = infoH / fields.length;
    fields.forEach((f, i) => {
      box(slide, CX, CY + i * rowH, labelW, rowH, f, LIGHT_BG, "E0E0E0", MID);
      box(slide, CX + labelW + 0.1, CY + i * rowH, CW - labelW - 0.1, rowH, "—", WHITE, "E0E0E0");
    });
    const tly = CY + infoH + 0.25;
    txt(slide, "沿革・主な実績", CX, tly, CW, 0.5, { fontSize: 13, bold: true, color: DARK });
    slide.addShape("rect", {
      x: CX + 0.4, y: tly + 0.6, w: 0.05, h: CH - infoH - 0.9,
      fill: { color: "DDDDDD" }, line: { color: "DDDDDD" },
    });
    for (let i = 0; i < 3; i++) {
      const ey = tly + 0.65 + i * 1.0;
      slide.addShape("ellipse", {
        x: CX + 0.24, y: ey, w: 0.32, h: 0.32,
        fill: { color: RED }, line: { color: RED },
      });
      box(slide, CX + 0.8, ey, CW - 0.8, 0.75, `実績 ${i + 1}`, WHITE, "E0E0E0");
    }
  } else {
    const items = ["強み・特徴", "主力サービス", "実績・数値"];
    const bw    = (CW - 0.4) / 3;
    const gap   = 0.2;
    items.forEach((item, i) => {
      const cx2 = CX + i * (bw + gap);
      slide.addShape("rect", {
        x: cx2, y: CY, w: bw, h: 0.07,
        fill: { color: RED }, line: { color: RED },
      });
      slide.addShape("rect", {
        x: cx2, y: CY + 0.12, w: bw, h: 0.7,
        fill: { color: LIGHT_BG }, line: { color: "E0E0E0", width: 0.5 },
      });
      txt(slide, item, cx2, CY + 0.12, bw, 0.7, {
        fontSize: 13, bold: true, align: "center", valign: "middle", color: DARK,
      });
      box(slide, cx2, CY + 0.95, bw, CH - 0.95, "詳細内容", WHITE, "E0E0E0");
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
  date: string,
  companyLogo: string
) {
  const slide = pptx.addSlide();
  const layoutName = getLayoutName(selectedSlide.sectionKey, selectedSlide.type);
  const sectionLabel =
    selectedSlide.slideIndexInSection > 1
      ? `${selectedSlide.sectionName}（${selectedSlide.slideIndexInSection}）`
      : selectedSlide.sectionName;

  addShell(pptx, slide, sectionLabel, layoutName, pageNum, totalPages, date, companyLogo);

  switch (selectedSlide.sectionKey) {
    case "brand":           renderBrand(slide, selectedSlide.type);    break;
    case "business_scheme": renderScheme(slide, selectedSlide.type);   break;
    case "cases":           renderCases(slide, selectedSlide.type);    break;
    case "items":           renderItems(slide, selectedSlide.type);    break;
    case "schedule":        renderSchedule(slide, selectedSlide.type); break;
    case "pricing":         renderPricing(slide, selectedSlide.type);  break;
    case "company":         renderCompany(slide, selectedSlide.type);  break;
    default: box(slide, CX, CY, CW, CH, selectedSlide.sectionName);
  }
}

// ─── Filename helper ──────────────────────────────────────────────────────────
function buildFileName(title: string, date: string): string {
  // Sanitize title: remove chars forbidden in filenames
  const safeTitle = title
    .replace(/[/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 40) || "企画書";

  // Parse date string YYYY/MM/DD or YYYYMMDD
  const match = date.match(/(\d{4})[^\d]?(\d{1,2})[^\d]?(\d{1,2})/);
  let ymd: string;
  if (match) {
    ymd = `${match[1]}${match[2].padStart(2, "0")}${match[3].padStart(2, "0")}`;
  } else {
    const n = new Date();
    ymd = `${n.getFullYear()}${String(n.getMonth() + 1).padStart(2, "0")}${String(n.getDate()).padStart(2, "0")}`;
  }

  return `${safeTitle}_${ymd}.pptx`;
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generatePptx(options: GeneratePptxOptions): Promise<void> {
  const { default: PptxGenJS } = await import("pptxgenjs");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pptx = new (PptxGenJS as any)();

  pptx.layout = "LAYOUT_WIDE"; // 13.333 × 7.5 inches (16:9)

  // Determine cover type
  const greetingSlide = options.selectedSlides.find((s) => s.sectionKey === "greeting");
  const coverType: PageType = (greetingSlide?.type === "B" || greetingSlide?.type === "C") ? "B" : "A";

  // Content slides = all slides except greeting
  const contentSlides = options.selectedSlides.filter((s) => s.sectionKey !== "greeting");
  const totalPages    = contentSlides.length + 1;

  // Cover
  addCoverSlide(pptx, options.title, options.clientName, options.date, options.companyLogo, coverType);

  // Content slides
  contentSlides.forEach((s, i) => {
    addContentSlide(pptx, s, i + 2, totalPages, options.date, options.companyLogo);
  });

  // Filename: 企画書タイトル_YYYYMMDD.pptx
  const fileName = buildFileName(options.title, options.date);

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
