"use client";

import type { PageType } from "@/lib/layout-labels";

export type SlideContent = Record<string, string>;

interface Props {
  sectionKey: string;
  type: PageType;
  content: SlideContent;
  // Cover slide extra props
  isCover?: boolean;
  title?: string;
  clientName?: string;
  date?: string;
}

// ── Shared micro-components ──────────────────────────────────────────────────

function ImgBox({ className = "" }: { className?: string }) {
  return (
    <div className={`bg-gray-200 rounded flex items-center justify-center shrink-0 ${className}`}>
      <div className="w-3 h-3 rounded-full bg-gray-300" />
    </div>
  );
}

function Val({
  v,
  fallback,
  cls = "",
}: {
  v?: string;
  fallback: string;
  cls?: string;
}) {
  return v ? (
    <span className={`truncate ${cls}`}>{v}</span>
  ) : (
    <span className={`text-gray-300 italic ${cls}`}>{fallback}</span>
  );
}

function PlaceholderLines({ n = 2 }: { n?: number }) {
  return (
    <div className="space-y-0.5 w-full">
      {Array.from({ length: n }).map((_, i) => (
        <div
          key={i}
          className="h-1 bg-gray-200 rounded"
          style={{ width: `${90 - i * 15}%` }}
        />
      ))}
    </div>
  );
}

function Box({ label, className = "" }: { label: string; className?: string }) {
  return (
    <div
      className={`border border-gray-300 rounded bg-white flex items-center justify-center text-center px-0.5 ${className}`}
    >
      <span className="text-[7px] leading-tight text-gray-600 line-clamp-2 break-all">{label || "──"}</span>
    </div>
  );
}

function Arrow({ dir = "right" }: { dir?: "right" | "both" | "down" }) {
  return (
    <div className="text-gray-400 text-[9px] shrink-0 leading-none">
      {dir === "both" ? "↔" : dir === "down" ? "↓" : "→"}
    </div>
  );
}

// ── Section layouts ───────────────────────────────────────────────────────────

export function SlideMiniPreview({ sectionKey, type, content, isCover, title, clientName, date }: Props) {
  const h = content.heading ?? "";
  const body = content.body ?? "";

  // ── Cover slide (表紙) ───────────────────────────────────────────────────
  if (isCover || sectionKey === "greeting") {
    return (
      <div className="w-full h-full flex flex-col items-center justify-center bg-white px-2 py-1 gap-0.5">
        {/* Company logo / text */}
        <div className="text-[8px] font-bold text-red-600 tracking-widest text-center leading-tight">
          KANSAI SUPER STUDIO
        </div>
        <div className="w-8 h-px bg-red-400 my-0.5" />
        {/* Title */}
        {(title || h) ? (
          <div className="text-[9px] font-semibold text-gray-800 text-center leading-snug line-clamp-2">
            {title || h}
          </div>
        ) : (
          <div className="text-[9px] text-gray-300 italic text-center">企画書タイトル</div>
        )}
        {/* Client */}
        {clientName ? (
          <div className="text-[7px] text-gray-600 text-center line-clamp-1 mt-0.5">{clientName}</div>
        ) : (
          <div className="text-[7px] text-gray-300 italic text-center mt-0.5">提案先名</div>
        )}
        {/* Date */}
        {date ? (
          <div className="text-[7px] text-gray-400 text-center mt-0.5">{date}</div>
        ) : (
          <div className="text-[7px] text-gray-300 italic text-center mt-0.5">日付</div>
        )}
      </div>
    );
  }

  switch (sectionKey) {
    // ── ブランド紹介 ─────────────────────────────────────────────────────────
    case "brand": {
      if (type === "A") {
        return (
          <div className="w-full h-full flex p-1.5 gap-1.5">
            <div className="flex-1 flex flex-col gap-0.5 min-w-0 justify-between">
              <Val v={content.brandName} fallback="ブランド名" cls="text-[9px] font-bold text-red-600 line-clamp-1" />
              {content.brandDesc ? (
                <p className="text-[7px] text-gray-600 line-clamp-2 leading-snug flex-1 mt-0.5">{content.brandDesc}</p>
              ) : (
                <div className="flex-1 mt-0.5">
                  <PlaceholderLines n={2} />
                </div>
              )}
              <Val v={content.target} fallback="ターゲット層" cls="text-[7px] text-gray-400 line-clamp-1" />
            </div>
            <ImgBox className="w-5/12 rounded" />
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1 gap-1">
            <ImgBox className="flex-1" />
            <div className="flex items-center justify-between gap-1 shrink-0">
              <Val v={content.brandName} fallback="ブランド名" cls="text-[8px] font-bold text-gray-800 flex-1 line-clamp-1" />
              <Val v={content.target} fallback="" cls="text-[7px] text-gray-400 shrink-0 line-clamp-1" />
            </div>
          </div>
        );
      }
      // C: feature boxes
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val v={content.brandName} fallback="ブランド名" cls="text-[8px] font-bold text-gray-800 line-clamp-1" />
          <div className="flex-1 flex gap-1">
            {["特徴1", "特徴2", "特徴3"].map((label) => (
              <div
                key={label}
                className="flex-1 bg-gray-100 border border-gray-200 rounded flex items-center justify-center"
              >
                <span className="text-[7px] text-gray-400">{label}</span>
              </div>
            ))}
          </div>
          <Val v={content.target} fallback="ターゲット層" cls="text-[7px] text-gray-400 line-clamp-1" />
        </div>
      );
    }

    // ── ビジネススキーム ──────────────────────────────────────────────────────
    case "business_scheme": {
      if (type === "A") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-1">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1" />
            <div className="flex-1 flex items-center justify-center gap-1">
              <Box label={content.licensor || "ライセンサー"} className="flex-1 py-2" />
              <Arrow dir="both" />
              <Box label={content.licensee || "ライセンシー"} className="flex-1 py-2" />
            </div>
            <Val v={content.license} fallback="" cls="text-[7px] text-gray-400 line-clamp-1" />
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1" />
            <div className="flex-1 flex flex-col gap-1 justify-center">
              <div className="flex items-center gap-1">
                <Box label={content.licensor || "ライセンサー"} className="flex-1 py-1.5" />
                <Arrow dir="right" />
                <Box label={content.licensee || "ライセンシー"} className="flex-1 py-1.5" />
              </div>
              <div className="flex justify-center">
                <Arrow dir="down" />
              </div>
              <div className="flex justify-center">
                <Box label={content.salesTo || "販売先"} className="w-1/2 py-1.5" />
              </div>
            </div>
          </div>
        );
      }
      // C: flow steps
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1" />
          <div className="flex-1 flex items-center gap-0.5">
            <Box label={content.licensor || "ライセンサー"} className="flex-1 py-2" />
            <Arrow dir="right" />
            <Box label={content.licensee || "ライセンシー"} className="flex-1 py-2" />
            <Arrow dir="right" />
            <Box label={content.salesTo || "販売先"} className="flex-1 py-2" />
          </div>
          <Val v={content.channel} fallback="販売チャネル" cls="text-[7px] text-gray-400 line-clamp-1" />
        </div>
      );
    }

    // ── 事例紹介 ─────────────────────────────────────────────────────────────
    case "cases": {
      if (type === "A") {
        return (
          <div className="w-full h-full flex p-1.5 gap-1.5">
            <ImgBox className="w-5/12 shrink-0" />
            <div className="flex-1 flex flex-col gap-0.5 min-w-0 justify-between">
              <Val
                v={content.caseTitle || h}
                fallback="事例タイトル"
                cls="text-[9px] font-semibold text-gray-800 line-clamp-2 leading-snug"
              />
              <Val v={content.company} fallback="企業名" cls="text-[7px] text-gray-500 line-clamp-1" />
              {content.result ? (
                <p className="text-[7px] text-gray-600 line-clamp-2 leading-snug">{content.result}</p>
              ) : (
                <PlaceholderLines n={2} />
              )}
            </div>
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-1">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1" />
            <div className="flex-1 flex gap-1">
              {[content.caseTitle || "事例1", "事例2", "事例3"].map((t, i) => (
                <div
                  key={i}
                  className="flex-1 border border-gray-200 rounded overflow-hidden flex flex-col"
                >
                  <ImgBox className="flex-1 rounded-none" />
                  <div className="text-[7px] text-gray-600 px-1 py-0.5 truncate">{t}</div>
                </div>
              ))}
            </div>
          </div>
        );
      }
      // C: before/after
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val v={content.caseTitle || h} fallback="事例タイトル" cls="text-[8px] font-semibold text-gray-700 line-clamp-1" />
          <div className="flex-1 flex gap-1">
            <div className="flex-1 bg-gray-50 border border-gray-200 rounded overflow-hidden flex flex-col">
              <div className="bg-gray-500 text-white text-[7px] font-medium px-1.5 py-0.5">Before</div>
              <div className="flex-1 p-1">
                <PlaceholderLines n={2} />
              </div>
            </div>
            <div className="flex-1 bg-red-50 border border-red-200 rounded overflow-hidden flex flex-col">
              <div className="bg-red-500 text-white text-[7px] font-medium px-1.5 py-0.5">After</div>
              <div className="flex-1 p-1">
                {content.result ? (
                  <p className="text-[7px] text-gray-700 line-clamp-3 leading-snug">{content.result}</p>
                ) : (
                  <PlaceholderLines n={2} />
                )}
              </div>
            </div>
          </div>
        </div>
      );
    }

    // ── アイテムイメージ ──────────────────────────────────────────────────────
    case "items": {
      if (type === "A") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-1">
            <ImgBox className="flex-1" />
            <div className="flex items-center justify-between gap-1 shrink-0">
              <Val v={content.productName} fallback="商品名" cls="text-[8px] font-semibold text-gray-800 flex-1 line-clamp-1" />
              <Val v={content.price} fallback="" cls="text-[7px] text-red-500 shrink-0" />
            </div>
            <Val v={content.productCategory} fallback="" cls="text-[7px] text-gray-400 line-clamp-1 shrink-0" />
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-1">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 shrink-0" />
            <div className="flex-1 grid grid-cols-2 grid-rows-2 gap-1">
              <ImgBox />
              <ImgBox />
              <ImgBox />
              <ImgBox />
            </div>
          </div>
        );
      }
      // C: scene + product
      return (
        <div className="w-full h-full flex p-1.5 gap-1.5">
          <ImgBox className="flex-1" />
          <div className="w-5/12 shrink-0 flex flex-col gap-1">
            <ImgBox className="flex-1" />
            <Val v={content.productName} fallback="商品名" cls="text-[7px] text-gray-700 line-clamp-1 shrink-0" />
          </div>
        </div>
      );
    }

    // ── スケジュール ─────────────────────────────────────────────────────────
    case "schedule": {
      if (type === "A") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-1 justify-between">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 shrink-0" />
            <div className="flex items-center gap-1 px-0.5">
              <Val v={content.startDate} fallback="開始" cls="text-[7px] text-gray-400 shrink-0" />
              <div className="flex-1 flex items-center">
                <div className="flex-1 h-px bg-gray-300" />
                <div className="w-2 h-2 rounded-full bg-red-400 shrink-0" />
                <div className="flex-1 h-px bg-gray-300" />
                <div className="w-2 h-2 rounded-full bg-red-400 shrink-0" />
                <div className="flex-1 h-px bg-gray-300" />
              </div>
              <Val v={content.endDate} fallback="終了" cls="text-[7px] text-gray-400 shrink-0" />
            </div>
            {content.milestones ? (
              <p className="text-[7px] text-gray-500 line-clamp-2 leading-snug">{content.milestones}</p>
            ) : (
              <div className="h-1 w-2/3 bg-gray-200 rounded" />
            )}
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 mb-0.5 shrink-0" />
            <div className="flex-1 flex flex-col justify-around">
              {["フェーズ1", "フェーズ2", "フェーズ3"].map((phase, i) => (
                <div key={phase} className="flex items-center gap-1">
                  <span className="text-[7px] text-gray-400 w-10 shrink-0 truncate">{phase}</span>
                  <div className="flex-1 h-2.5 bg-gray-100 rounded relative overflow-hidden">
                    <div
                      className="absolute inset-y-0 left-0 bg-red-400 rounded"
                      style={{ width: `${55 - i * 10}%`, left: `${i * 10}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      }
      // C: step cards
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 shrink-0" />
          <div className="flex-1 flex items-center gap-0.5">
            {["企画", "準備", "実施"].map((step, i) => (
              <div key={step} className="flex items-center gap-0.5 flex-1">
                <div className="flex-1 bg-red-50 border border-red-200 rounded flex flex-col items-center justify-center py-1">
                  <span className="text-[8px] font-bold text-red-600">{i + 1}</span>
                  <span className="text-[6px] text-gray-500 leading-tight">{step}</span>
                </div>
                {i < 2 && <div className="text-gray-300 text-[8px] shrink-0">›</div>}
              </div>
            ))}
          </div>
          {content.milestones && (
            <p className="text-[7px] text-gray-400 line-clamp-1 shrink-0">{content.milestones}</p>
          )}
        </div>
      );
    }

    // ── 料金・条件 ───────────────────────────────────────────────────────────
    case "pricing": {
      if (type === "A") {
        const rows = [
          { label: "契約期間",  value: content.contractPeriod },
          { label: "ロイヤリティ", value: content.royalty },
          { label: "最低保証金", value: content.minGuarantee },
        ];
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 mb-0.5 shrink-0" />
            <div className="flex-1 flex flex-col justify-around">
              {rows.map(({ label, value }) => (
                <div key={label} className="flex items-center gap-1 border-b border-gray-100 pb-0.5">
                  <span className="text-[7px] text-gray-400 w-16 shrink-0">{label}</span>
                  {value ? (
                    <span className="text-[7px] text-gray-700 truncate flex-1">{value}</span>
                  ) : (
                    <div className="flex-1 h-1 bg-gray-200 rounded" />
                  )}
                </div>
              ))}
            </div>
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 mb-0.5 shrink-0" />
            <div className="flex-1 flex gap-0.5">
              {["プランA", "プランB", "プランC"].map((plan, i) => (
                <div
                  key={plan}
                  className={`flex-1 rounded border flex flex-col overflow-hidden ${
                    i === 1 ? "border-red-300 bg-red-50" : "border-gray-200"
                  }`}
                >
                  <div
                    className={`text-[7px] font-medium px-1 py-0.5 text-center ${
                      i === 1 ? "text-red-600 bg-red-100" : "text-gray-500 bg-gray-100"
                    }`}
                  >
                    {plan}
                  </div>
                  <div className="flex-1 p-0.5 space-y-0.5">
                    <div className="h-1 bg-gray-200 rounded w-full" />
                    <div className="h-1 bg-gray-200 rounded w-3/4" />
                    <div className="h-1 bg-gray-200 rounded w-1/2" />
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      }
      // C: royalty focus
      return (
        <div className="w-full h-full flex flex-col items-center justify-center p-1.5 gap-1.5">
          <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 text-center w-full line-clamp-1" />
          <div className="flex flex-col items-center gap-0.5">
            <span className="text-[7px] text-gray-400">ロイヤリティ条件</span>
            {content.royalty ? (
              <span className="text-[13px] font-bold text-red-600 leading-none">{content.royalty}</span>
            ) : (
              <div className="h-3 w-16 bg-gray-200 rounded" />
            )}
          </div>
          <div className="w-full space-y-0.5">
            {[
              { label: "契約期間", value: content.contractPeriod },
              { label: "最低保証金", value: content.minGuarantee },
            ].map(({ label, value }) => (
              <div key={label} className="flex justify-between items-center">
                <span className="text-[7px] text-gray-400">{label}</span>
                {value ? (
                  <span className="text-[7px] text-gray-700">{value}</span>
                ) : (
                  <div className="h-1 w-10 bg-gray-200 rounded" />
                )}
              </div>
            ))}
          </div>
        </div>
      );
    }

    // ── 会社概要 ─────────────────────────────────────────────────────────────
    case "company": {
      if (type === "A") {
        const rows = [
          { label: "会社名",   value: content.companyName },
          { label: "所在地",   value: content.address },
          { label: "事業内容", value: content.business },
        ];
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val v={h} fallback="" cls="text-[8px] font-semibold text-gray-700 line-clamp-1 mb-0.5 shrink-0" />
            <div className="flex-1 flex flex-col justify-around">
              {rows.map(({ label, value }) => (
                <div key={label} className="flex items-start gap-1 border-b border-gray-100 pb-0.5">
                  <span className="text-[7px] text-gray-400 w-12 shrink-0 pt-px">{label}</span>
                  {value ? (
                    <span className="text-[7px] text-gray-700 truncate flex-1">{value}</span>
                  ) : (
                    <div className="flex-1 h-1 bg-gray-200 rounded mt-1" />
                  )}
                </div>
              ))}
            </div>
          </div>
        );
      }
      if (type === "B") {
        return (
          <div className="w-full h-full flex flex-col p-1.5 gap-0.5">
            <Val
              v={content.companyName || h}
              fallback="会社名"
              cls="text-[8px] font-semibold text-gray-700 line-clamp-1 mb-0.5 shrink-0"
            />
            <div className="flex-1 flex flex-col justify-around">
              {["創業", "成長期", "現在"].map((era) => (
                <div key={era} className="flex items-center gap-1.5">
                  <div className="w-1.5 h-1.5 rounded-full bg-red-400 shrink-0" />
                  <div className="flex-1 h-px bg-gray-200" />
                  <span className="text-[7px] text-gray-400 shrink-0">{era}</span>
                </div>
              ))}
            </div>
            {content.achievements && (
              <p className="text-[7px] text-gray-400 line-clamp-1 shrink-0 mt-0.5">{content.achievements}</p>
            )}
          </div>
        );
      }
      // C: strength cards
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val
            v={content.companyName || h}
            fallback="会社名"
            cls="text-[8px] font-semibold text-gray-700 line-clamp-1 shrink-0"
          />
          <div className="flex-1 flex gap-1">
            {["強み1", "強み2", "強み3"].map((s) => (
              <div
                key={s}
                className="flex-1 bg-red-50 border border-red-100 rounded flex items-center justify-center"
              >
                <span className="text-[7px] text-red-400">{s}</span>
              </div>
            ))}
          </div>
          {content.address && (
            <p className="text-[7px] text-gray-400 line-clamp-1 shrink-0">{content.address}</p>
          )}
        </div>
      );
    }

    // ── fallback ─────────────────────────────────────────────────────────────
    default: {
      return (
        <div className="w-full h-full flex flex-col p-1.5 gap-1">
          <Val v={h} fallback="見出し" cls="text-[9px] font-semibold text-gray-800 line-clamp-1 shrink-0" />
          {body ? (
            <p className="text-[7px] text-gray-600 leading-snug line-clamp-4">{body}</p>
          ) : (
            <PlaceholderLines n={3} />
          )}
        </div>
      );
    }
  }
}
