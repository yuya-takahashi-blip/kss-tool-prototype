"use client";

import { useState, useCallback, useEffect, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Textarea } from "@/components/ui/textarea";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/components/ui/collapsible";
import {
  CalendarIcon,
  FileText,
  Download,
  Loader2,
  GripVertical,
  Save,
  FilePlus,
  ChevronDown,
  ChevronUp,
} from "lucide-react";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from "@/components/ui/alert-dialog";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
  DragStartEvent,
  DragOverlay,
  TouchSensor,
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  rectSortingStrategy,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { layoutLabels, getLayoutName, type PageType } from "@/lib/layout-labels";
import { generatePptx } from "@/lib/generate-pptx";

// ─── Types ────────────────────────────────────────────────────────────────────
interface PageConfig {
  sectionKey: string;
  sectionName: string;
  checked: boolean;
  pageCount: number;
  type: PageType;
  previewFields: string[];
}

interface SelectedSlide {
  id: string;
  sectionKey: string;
  sectionName: string;
  type: PageType;
  slideIndexInSection: number;
}

type PageContent = Record<string, string>;

interface DraftData {
  title: string;
  date: string;
  clientName: string;
  pages: Array<{ sectionKey: string; checked: boolean; pageCount: number; type: PageType }>;
  slideOrder: string[];
  pageContents: Record<string, PageContent>;
  savedAt: string;
}

// ─── Field definitions ────────────────────────────────────────────────────────
interface FieldDef {
  key: string;
  label: string;
  multiline?: boolean;
}

const COMMON_FIELDS: FieldDef[] = [
  { key: "heading", label: "見出し" },
  { key: "body",    label: "本文",    multiline: true },
  { key: "memo",    label: "補足メモ", multiline: true },
];

const SECTION_EXTRA_FIELDS: Record<string, FieldDef[]> = {
  greeting: [
    { key: "addressee",   label: "宛名" },
    { key: "responsible", label: "担当者名" },
  ],
  brand: [
    { key: "brandName", label: "ブランド名" },
    { key: "brandDesc", label: "ブランド説明", multiline: true },
    { key: "target",    label: "ターゲット層" },
  ],
  business_scheme: [
    { key: "licensor",  label: "ライセンサー名" },
    { key: "licensee",  label: "ライセンシー名" },
    { key: "salesTo",   label: "販売先名" },
    { key: "license",   label: "許諾内容",     multiline: true },
    { key: "channel",   label: "販売チャネル" },
  ],
  cases: [
    { key: "caseTitle", label: "事例タイトル" },
    { key: "company",   label: "実施企業名" },
    { key: "period",    label: "実施時期" },
    { key: "result",    label: "成果・実績", multiline: true },
  ],
  items: [
    { key: "productName",     label: "商品名" },
    { key: "productCategory", label: "商品カテゴリ" },
    { key: "price",           label: "想定価格" },
    { key: "productDesc",     label: "商品説明", multiline: true },
  ],
  schedule: [
    { key: "startDate",   label: "開始時期" },
    { key: "endDate",     label: "終了時期" },
    { key: "milestones",  label: "主なマイルストーン", multiline: true },
  ],
  pricing: [
    { key: "contractPeriod",  label: "契約期間" },
    { key: "royalty",         label: "ロイヤリティ条件", multiline: true },
    { key: "minGuarantee",    label: "最低保証金" },
    { key: "otherConditions", label: "その他条件", multiline: true },
  ],
  company: [
    { key: "companyName",  label: "会社名" },
    { key: "address",      label: "所在地" },
    { key: "business",     label: "事業内容", multiline: true },
    { key: "achievements", label: "実績",    multiline: true },
  ],
};

function getFieldsForSection(sectionKey: string): FieldDef[] {
  return [...COMMON_FIELDS, ...(SECTION_EXTRA_FIELDS[sectionKey] ?? [])];
}

// Keys shown in the detail preview card (excluding heading/body/memo)
const PREVIEW_HIGHLIGHT_FIELDS: Record<string, string[]> = {
  greeting:        ["addressee", "responsible"],
  brand:           ["brandName", "target"],
  business_scheme: ["licensor", "licensee", "salesTo"],
  cases:           ["caseTitle", "company", "result"],
  items:           ["productName", "productCategory", "price"],
  schedule:        ["startDate", "endDate", "milestones"],
  pricing:         ["contractPeriod", "royalty", "minGuarantee"],
  company:         ["companyName", "address", "business"],
};

const EMPTY_CONTENT: PageContent = {};

// ─── Constants ────────────────────────────────────────────────────────────────
const STORAGE_KEY = "kss-tool-proposal-draft";

const initialPages: PageConfig[] = [
  { sectionKey: "greeting",        sectionName: "ご挨拶",           checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "brand",           sectionName: "ブランド紹介",      checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "business_scheme", sectionName: "ビジネススキーム",  checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "cases",           sectionName: "事例紹介",          checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "items",           sectionName: "アイテムイメージ",   checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "schedule",        sectionName: "スケジュール",       checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "pricing",         sectionName: "料金・条件",         checked: true, pageCount: 1, type: "A", previewFields: [] },
  { sectionKey: "company",         sectionName: "会社概要",           checked: true, pageCount: 1, type: "A", previewFields: [] },
];

// ─── localStorage helpers ─────────────────────────────────────────────────────
function loadDraft(): DraftData | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as DraftData;
    if (!parsed.title || !Array.isArray(parsed.pages) || !Array.isArray(parsed.slideOrder)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function saveDraft(data: DraftData): boolean {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    return true;
  } catch {
    return false;
  }
}

function formatDateTime(iso: string): string {
  try {
    const d = new Date(iso);
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, "0");
    const dy = String(d.getDate()).padStart(2, "0");
    const h  = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${y}/${mo}/${dy} ${h}:${mi}`;
  } catch {
    return "";
  }
}

function todayString(): string {
  const d = new Date();
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
}

// ─── Slide helpers ────────────────────────────────────────────────────────────
function generateSelectedSlides(pages: PageConfig[]): SelectedSlide[] {
  return pages
    .filter((p) => p.checked)
    .flatMap((p) =>
      Array.from({ length: p.pageCount }, (_, i) => ({
        id: `${p.sectionKey}-${i}`,
        sectionKey: p.sectionKey,
        sectionName: p.sectionName,
        type: p.type,
        slideIndexInSection: i + 1,
      }))
    );
}

function applySlideOrder(slides: SelectedSlide[], order: string[]): SelectedSlide[] {
  const map = new Map(slides.map((s) => [s.id, s]));
  const ordered: SelectedSlide[] = [];
  for (const id of order) {
    const s = map.get(id);
    if (s) { ordered.push(s); map.delete(id); }
  }
  map.forEach((s) => ordered.push(s));
  return ordered;
}

function mergePages(base: PageConfig[], saved: DraftData["pages"]): PageConfig[] {
  const savedMap = new Map(saved.map((s) => [s.sectionKey, s]));
  return base.map((p) => {
    const s = savedMap.get(p.sectionKey);
    if (!s) return p;
    return {
      ...p,
      checked:   typeof s.checked   === "boolean" ? s.checked   : p.checked,
      pageCount: typeof s.pageCount  === "number"  && s.pageCount >= 1 ? s.pageCount : p.pageCount,
      type:      (["A", "B", "C"] as PageType[]).includes(s.type as PageType) ? s.type : p.type,
    };
  });
}

// ─── Thumbnail (sortable) ─────────────────────────────────────────────────────
interface ThumbnailProps {
  slide: SelectedSlide;
  index: number;
  heading: string;
}

function SortableThumbnail({ slide, index, heading }: ThumbnailProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: slide.id });

  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;

  return (
    <div
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition, opacity: isDragging ? 0.4 : 1 }}
      className={`bg-white rounded-lg border border-gray-200 overflow-hidden aspect-[4/3] flex flex-col hover:border-red-200 hover:shadow-sm transition-all ${isDragging ? "shadow-lg" : ""}`}
    >
      {/* slide "header" bar */}
      <div className="bg-red-600 px-2 py-1 flex items-center justify-between shrink-0">
        <span className="text-white text-[10px] font-semibold truncate leading-tight">P.{index + 1}</span>
        <div {...attributes} {...listeners} className="cursor-grab active:cursor-grabbing p-0.5 touch-none ml-1">
          <GripVertical className="w-3 h-3 text-red-200" />
        </div>
      </div>
      {/* slide body */}
      <div className="flex-1 flex flex-col justify-between p-2 min-h-0">
        <div>
          <div className="text-[11px] font-semibold text-gray-800 leading-snug line-clamp-1">{displayName}</div>
          {heading ? (
            <div className="text-[10px] text-gray-600 mt-0.5 line-clamp-2 leading-snug">{heading}</div>
          ) : (
            <div className="text-[10px] text-gray-300 mt-0.5 leading-snug">見出し未入力</div>
          )}
        </div>
        <div className="text-[9px] text-red-500 bg-red-50 px-1.5 py-0.5 rounded leading-tight self-start max-w-full truncate mt-1">
          {getLayoutName(slide.sectionKey, slide.type)}
        </div>
      </div>
    </div>
  );
}

function DragOverlayThumbnail({ slide, index, heading }: ThumbnailProps) {
  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;
  return (
    <div className="bg-white rounded-lg border border-red-300 overflow-hidden aspect-[4/3] flex flex-col shadow-xl rotate-2 scale-105">
      <div className="bg-red-600 px-2 py-1 shrink-0">
        <span className="text-white text-[10px] font-semibold">P.{index + 1}</span>
      </div>
      <div className="flex-1 flex flex-col justify-between p-2">
        <div>
          <div className="text-[11px] font-semibold text-gray-800 line-clamp-1">{displayName}</div>
          {heading ? (
            <div className="text-[10px] text-gray-600 mt-0.5 line-clamp-2">{heading}</div>
          ) : (
            <div className="text-[10px] text-gray-300 mt-0.5">見出し未入力</div>
          )}
        </div>
        <div className="text-[9px] text-red-500 bg-red-50 px-1.5 py-0.5 rounded self-start truncate mt-1">
          {getLayoutName(slide.sectionKey, slide.type)}
        </div>
      </div>
    </div>
  );
}

// ─── PageContentForm ──────────────────────────────────────────────────────────
interface PageContentFormProps {
  sectionKey: string;
  sectionName: string;
  layoutName: string;
  slideCount: number;
  content: PageContent;
  onContentChange: (sectionKey: string, fieldKey: string, value: string) => void;
  defaultOpen?: boolean;
}

function PageContentForm({
  sectionKey,
  sectionName,
  layoutName,
  slideCount,
  content,
  onContentChange,
  defaultOpen = false,
}: PageContentFormProps) {
  const [open, setOpen] = useState(defaultOpen);
  const fields = getFieldsForSection(sectionKey);

  return (
    <Collapsible open={open} onOpenChange={setOpen}>
      <Card className="bg-white border-gray-200 overflow-hidden">
        <CollapsibleTrigger asChild>
          <button
            type="button"
            className="w-full text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-red-400 focus-visible:ring-offset-1 rounded-t-lg"
          >
            <div className="flex items-center justify-between px-4 py-3 bg-gray-50 border-b border-gray-200 hover:bg-gray-100 transition-colors">
              <div className="flex items-center gap-3 min-w-0">
                <span className="font-semibold text-sm text-gray-800 shrink-0">{sectionName}</span>
                {slideCount > 1 && (
                  <span className="text-xs text-gray-500 shrink-0">({slideCount}ページ)</span>
                )}
                <span className="text-xs font-medium px-2 py-0.5 bg-red-50 text-red-600 rounded whitespace-nowrap shrink-0">
                  {layoutName}
                </span>
              </div>
              {open ? (
                <ChevronUp className="w-4 h-4 text-gray-400 shrink-0 ml-2" />
              ) : (
                <ChevronDown className="w-4 h-4 text-gray-400 shrink-0 ml-2" />
              )}
            </div>
          </button>
        </CollapsibleTrigger>
        <CollapsibleContent>
          <CardContent className="p-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {fields.map((field) =>
                field.multiline ? (
                  <div key={field.key} className="space-y-1.5 md:col-span-2">
                    <Label className="text-xs font-medium text-gray-600">{field.label}</Label>
                    <Textarea
                      value={content[field.key] ?? ""}
                      onChange={(e) => onContentChange(sectionKey, field.key, e.target.value)}
                      placeholder={`${field.label}を入力`}
                      className="min-h-[88px] text-base resize-y leading-relaxed"
                    />
                  </div>
                ) : (
                  <div key={field.key} className="space-y-1.5">
                    <Label className="text-xs font-medium text-gray-600">{field.label}</Label>
                    <Input
                      value={content[field.key] ?? ""}
                      onChange={(e) => onContentChange(sectionKey, field.key, e.target.value)}
                      placeholder={`${field.label}を入力`}
                      className="h-11 text-base"
                    />
                  </div>
                )
              )}
            </div>
          </CardContent>
        </CollapsibleContent>
      </Card>
    </Collapsible>
  );
}

// ─── SlidePreviewCard ─────────────────────────────────────────────────────────
interface SlidePreviewCardProps {
  sectionKey: string;
  sectionName: string;
  layoutName: string;
  slideCount: number;
  content: PageContent;
}

function SlidePreviewCard({
  sectionKey,
  sectionName,
  layoutName,
  slideCount,
  content,
}: SlidePreviewCardProps) {
  const heading = content.heading || "";
  const body    = content.body    || "";

  // Highlight field entries for this section
  const highlightKeys = PREVIEW_HIGHLIGHT_FIELDS[sectionKey] ?? [];
  const allFields     = getFieldsForSection(sectionKey);
  const highlights    = highlightKeys
    .map((key) => {
      const def = allFields.find((f) => f.key === key);
      return def ? { label: def.label, value: content[key] ?? "" } : null;
    })
    .filter(Boolean) as { label: string; value: string }[];

  const hasAnyContent = heading || body || highlights.some((h) => h.value);

  return (
    <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm flex flex-col">
      {/* Card header – red bar mimicking a slide title area */}
      <div className="bg-red-600 px-4 py-2.5 flex items-center justify-between gap-2">
        <div className="min-w-0">
          <div className="text-white text-sm font-semibold leading-tight truncate">
            {sectionName}
            {slideCount > 1 && (
              <span className="text-red-200 text-xs ml-1.5">({slideCount}ページ)</span>
            )}
          </div>
        </div>
        <span className="text-xs font-medium px-2 py-0.5 bg-red-500 text-red-100 rounded whitespace-nowrap shrink-0">
          {layoutName}
        </span>
      </div>

      {/* Slide body */}
      <div className="p-4 flex-1 flex flex-col gap-3">
        {/* Heading */}
        <div className="border-b border-gray-100 pb-3">
          {heading ? (
            <p className="text-base font-semibold text-gray-800 leading-snug">{heading}</p>
          ) : (
            <p className="text-sm text-gray-300 italic">{hasAnyContent ? "（見出し未入力）" : "見出し"}</p>
          )}
          {/* Body preview */}
          {body ? (
            <p className="text-sm text-gray-600 mt-1.5 leading-relaxed line-clamp-3">{body}</p>
          ) : (
            !hasAnyContent && <p className="text-xs text-gray-300 mt-1 italic">本文</p>
          )}
        </div>

        {/* Highlight fields */}
        {highlights.length > 0 && (
          <div className="grid grid-cols-1 gap-1.5">
            {highlights.map(({ label, value }) => (
              <div key={label} className="flex items-baseline gap-2">
                <span className="text-xs text-gray-400 shrink-0 w-28">{label}</span>
                {value ? (
                  <span className="text-sm text-gray-700 leading-snug line-clamp-2 flex-1">{value}</span>
                ) : (
                  <span className="text-xs text-gray-300 italic flex-1">未入力</span>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────
export default function Home() {
  const [title, setTitle]             = useState("企画書タイトル");
  const [date, setDate]               = useState("2026/06/30");
  const [clientName, setClientName]   = useState("");
  const [pages, setPages]             = useState<PageConfig[]>(initialPages);
  const [selectedSlides, setSelectedSlides] = useState<SelectedSlide[]>(() =>
    generateSelectedSlides(initialPages)
  );
  const [pageContents, setPageContents] = useState<Record<string, PageContent>>({});
  const [activeId, setActiveId]         = useState<string | null>(null);
  const [exporting, setExporting]       = useState(false);
  const [restoreMessage, setRestoreMessage] = useState("");
  const [saveMessage, setSaveMessage]   = useState("");
  const [saveStatus, setSaveStatus]     = useState<"success" | "error" | "">("");
  const [savedAt, setSavedAt]           = useState<string>("");
  const didRestoreRef  = useRef(false);
  const saveTimerRef   = useRef<ReturnType<typeof setTimeout> | null>(null);

  // ── Restore ──────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (didRestoreRef.current) return;
    didRestoreRef.current = true;

    const draft = loadDraft();
    if (!draft) return;

    const restoredPages  = mergePages(initialPages, draft.pages);
    const rawSlides      = generateSelectedSlides(restoredPages);
    const orderedSlides  = applySlideOrder(rawSlides, draft.slideOrder);

    setTitle(draft.title ?? "企画書タイトル");
    setDate(draft.date   ?? "2026/06/30");
    setClientName(draft.clientName ?? "");
    setPages(restoredPages);
    setSelectedSlides(orderedSlides);
    setPageContents(
      draft.pageContents && typeof draft.pageContents === "object" ? draft.pageContents : {}
    );
    setSavedAt(draft.savedAt ?? "");

    setRestoreMessage("前回の内容を復元しました");
    setTimeout(() => setRestoreMessage(""), 3000);
  }, []);

  // ── Content change ───────────────────────────────────────────────────────────
  const handleContentChange = (sectionKey: string, fieldKey: string, value: string) => {
    setPageContents((prev) => ({
      ...prev,
      [sectionKey]: { ...(prev[sectionKey] ?? EMPTY_CONTENT), [fieldKey]: value },
    }));
  };

  // ── Save ──────────────────────────────────────────────────────────────────────
  const handleSave = () => {
    const now = new Date().toISOString();
    const draft: DraftData = {
      title,
      date,
      clientName,
      pages: pages.map(({ sectionKey, checked, pageCount, type }) => ({
        sectionKey, checked, pageCount, type,
      })),
      slideOrder: selectedSlides.map((s) => s.id),
      pageContents,
      savedAt: now,
    };
    const ok = saveDraft(draft);
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    if (ok) {
      setSavedAt(now);
      setSaveMessage("保存しました");
      setSaveStatus("success");
    } else {
      setSaveMessage("保存できませんでした");
      setSaveStatus("error");
    }
    saveTimerRef.current = setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 3000);
  };

  // ── New document ──────────────────────────────────────────────────────────────
  const handleNew = () => {
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    setTitle("企画書タイトル");
    setDate(todayString());
    setClientName("");
    setPages(initialPages);
    setSelectedSlides(generateSelectedSlides(initialPages));
    setPageContents({});
    setSavedAt("");
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setSaveMessage("新しい企画書を作成しました");
    setSaveStatus("success");
    saveTimerRef.current = setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 3000);
  };

  // ── Page config ───────────────────────────────────────────────────────────────
  const updateSelectedSlidesForSection = useCallback(
    (sectionKey: string, checked: boolean, pageCount: number, type: PageType) => {
      setSelectedSlides((prevSlides) => {
        const page = pages.find((p) => p.sectionKey === sectionKey);
        if (!page) return prevSlides;
        const others = prevSlides.filter((s) => s.sectionKey !== sectionKey);
        if (!checked || pageCount === 0) return others;
        return [
          ...others,
          ...Array.from({ length: pageCount }, (_, i) => ({
            id: `${sectionKey}-${i}`,
            sectionKey,
            sectionName: page.sectionName,
            type,
            slideIndexInSection: i + 1,
          })),
        ];
      });
    },
    [pages]
  );

  const handlePageChange = (
    sectionKey: string,
    field: "checked" | "pageCount" | "type",
    value: boolean | number | PageType
  ) => {
    setPages((prev) =>
      prev.map((p) => (p.sectionKey !== sectionKey ? p : { ...p, [field]: value }))
    );
    const page = pages.find((p) => p.sectionKey === sectionKey);
    if (!page) return;
    if (field === "type" && page.checked) {
      setSelectedSlides((prev) =>
        prev.map((s) =>
          s.sectionKey === sectionKey ? { ...s, type: value as PageType } : s
        )
      );
    }
    if (field === "checked" || field === "pageCount") {
      updateSelectedSlidesForSection(
        sectionKey,
        field === "checked" ? (value as boolean) : page.checked,
        field === "pageCount" ? (value as number) : page.pageCount,
        page.type
      );
    }
  };

  // ── DnD ──────────────────────────────────────────────────────────────────────
  const sensors = useSensors(
    useSensor(PointerSensor,  { activationConstraint: { distance: 8 } }),
    useSensor(TouchSensor,    { activationConstraint: { delay: 200, tolerance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  const handleDragStart = (e: DragStartEvent) => setActiveId(e.active.id as string);
  const handleDragEnd   = (e: DragEndEvent) => {
    const { active, over } = e;
    setActiveId(null);
    if (over && active.id !== over.id) {
      setSelectedSlides((items) =>
        arrayMove(
          items,
          items.findIndex((i) => i.id === active.id),
          items.findIndex((i) => i.id === over.id)
        )
      );
    }
  };

  const activeSlide = activeId ? selectedSlides.find((s) => s.id === activeId) : null;
  const activeIndex = activeSlide ? selectedSlides.findIndex((s) => s.id === activeId) : 0;

  // ── Export ────────────────────────────────────────────────────────────────────
  const handleExport = async () => {
    if (selectedSlides.length === 0) {
      setSaveMessage("スライドが選択されていません");
      setSaveStatus("error");
      setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 2500);
      return;
    }
    setExporting(true);
    try {
      await generatePptx({ title, clientName, date, selectedSlides });
      setSaveMessage("ダウンロードしました");
      setSaveStatus("success");
    } catch (err) {
      console.error(err);
      setSaveMessage("書き出しに失敗しました");
      setSaveStatus("error");
    } finally {
      setExporting(false);
      setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 3000);
    }
  };

  // ── Derived ───────────────────────────────────────────────────────────────────
  const slidesBySection = selectedSlides.reduce(
    (acc, slide) => {
      if (!acc[slide.sectionKey]) acc[slide.sectionKey] = [];
      acc[slide.sectionKey].push(slide);
      return acc;
    },
    {} as Record<string, SelectedSlide[]>
  );
  const orderedSectionKeys = Array.from(new Set(selectedSlides.map((s) => s.sectionKey)));

  // ── Render ────────────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-gray-100">
      <header className="bg-white border-b border-gray-200 px-4 lg:px-6 py-3 shadow-sm">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <h1 className="text-xl lg:text-2xl font-bold text-red-600 tracking-wide">KSS TOOL</h1>
            <span className="hidden sm:inline-block text-gray-400 text-sm">営業企画書生成ツール</span>
          </div>
          {restoreMessage && (
            <div className="text-sm px-4 py-2 bg-gray-800 text-white rounded-lg">
              {restoreMessage}
            </div>
          )}
        </div>
      </header>

      <main className="p-3 lg:p-4 max-w-[1400px] mx-auto">
        {/* ── Basic info ── */}
        <Card className="mb-4 bg-white border-gray-200">
          <CardContent className="p-4">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="space-y-2">
                <Label htmlFor="title" className="text-sm font-medium text-gray-700">企画書タイトル</Label>
                <Input
                  id="title" value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  className="h-11 text-base" placeholder="タイトルを入力"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="date" className="text-sm font-medium text-gray-700">日付</Label>
                <div className="relative">
                  <Input
                    id="date" value={date}
                    onChange={(e) => setDate(e.target.value)}
                    className="h-11 text-base pl-10" placeholder="YYYY/MM/DD"
                  />
                  <CalendarIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400" />
                </div>
              </div>
              <div className="space-y-2">
                <Label htmlFor="client" className="text-sm font-medium text-gray-700">提案先名</Label>
                <Input
                  id="client" value={clientName}
                  onChange={(e) => setClientName(e.target.value)}
                  className="h-11 text-base" placeholder="株式会社〇〇"
                />
              </div>
            </div>
          </CardContent>
        </Card>

        {/* ── Page structure + Overall preview ── */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
          {/* Page structure */}
          <Card className="lg:col-span-1 bg-white border-gray-200">
            <CardHeader className="pb-2 px-4 pt-4">
              <CardTitle className="text-lg font-semibold text-gray-800 flex items-center gap-2">
                <FileText className="w-5 h-5 text-red-500" />
                ページ構成選択
              </CardTitle>
            </CardHeader>
            <CardContent className="px-4 pb-4">
              <div className="space-y-3">
                {pages.map((page) => {
                  const labels = layoutLabels[page.sectionKey];
                  return (
                    <div
                      key={page.sectionKey}
                      className="flex flex-col gap-2 p-3 bg-gray-50 rounded-lg border border-gray-200 hover:bg-gray-100 transition-colors"
                    >
                      <div className="flex items-center gap-3">
                        <Checkbox
                          id={page.sectionKey}
                          checked={page.checked}
                          onCheckedChange={(v) => handlePageChange(page.sectionKey, "checked", v as boolean)}
                          className="w-5 h-5 shrink-0"
                        />
                        <Label htmlFor={page.sectionKey} className="text-sm font-medium text-gray-800 cursor-pointer">
                          {page.sectionName}
                        </Label>
                      </div>
                      <div className="flex items-center gap-2 ml-8">
                        <div className="flex items-center gap-1 shrink-0">
                          <Label className="text-xs text-gray-500 whitespace-nowrap">ページ数</Label>
                          <Input
                            type="number" min="1" max="10"
                            value={page.pageCount}
                            onChange={(e) => handlePageChange(page.sectionKey, "pageCount", parseInt(e.target.value) || 1)}
                            className="w-14 h-9 text-center text-sm"
                            disabled={!page.checked}
                          />
                        </div>
                        <Select
                          value={page.type}
                          onValueChange={(v: PageType) => handlePageChange(page.sectionKey, "type", v)}
                          disabled={!page.checked}
                        >
                          <SelectTrigger className="flex-1 min-w-0 h-9 text-xs">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="A">{labels.A}</SelectItem>
                            <SelectItem value="B">{labels.B}</SelectItem>
                            <SelectItem value="C">{labels.C}</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                    </div>
                  );
                })}
              </div>
            </CardContent>
          </Card>

          {/* Overall preview */}
          <Card className="lg:col-span-2 bg-white border-gray-200">
            <CardHeader className="pb-2 px-4 pt-4">
              <CardTitle className="text-lg font-semibold text-gray-800">全体プレビュー</CardTitle>
            </CardHeader>
            <CardContent className="px-4 pb-4">
              <div className="max-h-[420px] overflow-y-auto pr-1">
                {selectedSlides.length === 0 ? (
                  <div className="text-center text-gray-400 py-12">ページが選択されていません</div>
                ) : (
                  <DndContext
                    sensors={sensors}
                    collisionDetection={closestCenter}
                    onDragStart={handleDragStart}
                    onDragEnd={handleDragEnd}
                  >
                    <SortableContext items={selectedSlides.map((s) => s.id)} strategy={rectSortingStrategy}>
                      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
                        {selectedSlides.map((slide, index) => (
                          <SortableThumbnail
                            key={slide.id}
                            slide={slide}
                            index={index}
                            heading={pageContents[slide.sectionKey]?.heading ?? ""}
                          />
                        ))}
                      </div>
                    </SortableContext>
                    <DragOverlay>
                      {activeSlide && (
                        <DragOverlayThumbnail
                          slide={activeSlide}
                          index={activeIndex}
                          heading={pageContents[activeSlide.sectionKey]?.heading ?? ""}
                        />
                      )}
                    </DragOverlay>
                  </DndContext>
                )}
              </div>
            </CardContent>
          </Card>
        </div>

        {/* ── Actions ── */}
        <div className="mb-4 space-y-2">
          <div className="flex flex-wrap items-center gap-3">
            <Button
              onClick={handleSave}
              className="h-11 px-6 bg-gray-700 hover:bg-gray-800 text-white flex items-center gap-2"
            >
              <Save className="w-4 h-4" />保存
            </Button>
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button variant="outline" className="h-11 px-6 border-gray-300 text-gray-700 flex items-center gap-2">
                  <FilePlus className="w-4 h-4" />新規作成
                </Button>
              </AlertDialogTrigger>
              <AlertDialogContent>
                <AlertDialogHeader>
                  <AlertDialogTitle>新規作成</AlertDialogTitle>
                  <AlertDialogDescription>
                    現在の入力内容をリセットして、新しい企画書を作成します。よろしいですか？
                  </AlertDialogDescription>
                </AlertDialogHeader>
                <AlertDialogFooter>
                  <AlertDialogCancel>キャンセル</AlertDialogCancel>
                  <AlertDialogAction onClick={handleNew}>OK</AlertDialogAction>
                </AlertDialogFooter>
              </AlertDialogContent>
            </AlertDialog>
            <Button
              onClick={handleExport}
              disabled={exporting}
              className="h-11 px-6 bg-red-600 hover:bg-red-700 text-white disabled:opacity-70 flex items-center gap-2"
            >
              {exporting ? (
                <><Loader2 className="w-4 h-4 animate-spin" />生成中...</>
              ) : (
                <><Download className="w-4 h-4" />PowerPointを書き出す</>
              )}
            </Button>
            <span className="text-xs text-gray-400 ml-1">
              {savedAt ? `最終保存：${formatDateTime(savedAt)}` : "最終保存：未保存"}
            </span>
          </div>
          {saveMessage && (
            <div
              className={[
                "inline-flex items-center px-4 py-2.5 rounded-lg text-sm font-medium",
                saveStatus === "success"
                  ? "bg-green-50 text-green-800 border border-green-200"
                  : "bg-red-50 text-red-800 border border-red-200",
              ].join(" ")}
            >
              {saveMessage}
            </div>
          )}
        </div>

        {/* ── Page content input ── */}
        <Card className="mb-4 bg-white border-gray-200">
          <CardHeader className="pb-2 px-4 pt-4">
            <CardTitle className="text-lg font-semibold text-gray-800">ページ内容入力</CardTitle>
          </CardHeader>
          <CardContent className="px-4 pb-4">
            {selectedSlides.length === 0 ? (
              <div className="text-center text-gray-400 py-8">ページが選択されていません</div>
            ) : (
              <div className="space-y-3">
                {orderedSectionKeys.map((sectionKey, idx) => {
                  const sectionSlides = slidesBySection[sectionKey];
                  const firstSlide    = sectionSlides[0];
                  return (
                    <PageContentForm
                      key={sectionKey}
                      sectionKey={sectionKey}
                      sectionName={firstSlide.sectionName}
                      layoutName={getLayoutName(sectionKey, firstSlide.type)}
                      slideCount={sectionSlides.length}
                      content={pageContents[sectionKey] ?? EMPTY_CONTENT}
                      onContentChange={handleContentChange}
                      defaultOpen={idx === 0}
                    />
                  );
                })}
              </div>
            )}
          </CardContent>
        </Card>

        {/* ── Page detail preview ── */}
        <Card className="bg-white border-gray-200">
          <CardHeader className="pb-2 px-4 pt-4">
            <CardTitle className="text-lg font-semibold text-gray-800">ページ別プレビュー</CardTitle>
          </CardHeader>
          <CardContent className="px-4 pb-4">
            {selectedSlides.length === 0 ? (
              <div className="text-center text-gray-400 py-8">ページが選択されていません</div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                {orderedSectionKeys.map((sectionKey) => {
                  const sectionSlides = slidesBySection[sectionKey];
                  const firstSlide    = sectionSlides[0];
                  return (
                    <SlidePreviewCard
                      key={sectionKey}
                      sectionKey={sectionKey}
                      sectionName={firstSlide.sectionName}
                      layoutName={getLayoutName(sectionKey, firstSlide.type)}
                      slideCount={sectionSlides.length}
                      content={pageContents[sectionKey] ?? EMPTY_CONTENT}
                    />
                  );
                })}
              </div>
            )}
          </CardContent>
        </Card>
      </main>
    </div>
  );
}
