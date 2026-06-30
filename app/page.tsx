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
import { SlideMiniPreview, type SlideContent } from "@/components/slide-mini-preview";

// ─── Types ────────────────────────────────────────────────────────────────────
interface PageConfig {
  sectionKey: string;
  sectionName: string;
  checked: boolean;
  pageCount: number;
  type: PageType;
}

interface SelectedSlide {
  id: string;
  sectionKey: string;
  sectionName: string;
  type: PageType;
  slideIndexInSection: number;
}

interface DraftData {
  title: string;
  date: string;
  clientName: string;
  companyLogo: string; // base64 data URL or "" for text fallback
  pages: Array<{ sectionKey: string; checked: boolean; pageCount: number; type: PageType }>;
  slideOrder: string[];
  pageContents: Record<string, SlideContent>;
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
    { key: "license",   label: "許諾内容",      multiline: true },
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
    { key: "startDate",  label: "開始時期" },
    { key: "endDate",    label: "終了時期" },
    { key: "milestones", label: "主なマイルストーン", multiline: true },
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

const EMPTY_CONTENT: SlideContent = {};

// ─── Constants ────────────────────────────────────────────────────────────────
const STORAGE_KEY = "kss-tool-proposal-draft";

const initialPages: PageConfig[] = [
  { sectionKey: "greeting",        sectionName: "表紙",             checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "brand",           sectionName: "ブランド紹介",      checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "business_scheme", sectionName: "ビジネススキーム",  checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "cases",           sectionName: "事例紹介",          checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "items",           sectionName: "アイテムイメージ",   checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "schedule",        sectionName: "スケジュール",       checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "pricing",         sectionName: "料金・条件",         checked: true,  pageCount: 1, type: "A" },
  { sectionKey: "company",         sectionName: "会社概要",           checked: true,  pageCount: 1, type: "A" },
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
    return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  } catch { return ""; }
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
      checked:   typeof s.checked === "boolean" ? s.checked : p.checked,
      pageCount: typeof s.pageCount === "number" && s.pageCount >= 1 ? s.pageCount : p.pageCount,
      type:      (["A", "B", "C"] as PageType[]).includes(s.type as PageType) ? s.type : p.type,
    };
  });
}

// ─── Sortable thumbnail ───────────────────────────────────────────────────────
interface ThumbnailProps {
  slide: SelectedSlide;
  index: number;
  content: SlideContent;
  date: string;
  title: string;
  clientName: string;
  onSelect: () => void;
}

function SortableThumbnail({ slide, index, content, date, title, clientName, onSelect }: ThumbnailProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: slide.id });

  const isCover = slide.sectionKey === "greeting";
  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;

  return (
    <div
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition, opacity: isDragging ? 0.35 : 1 }}
      className={`bg-white rounded-lg border border-gray-200 overflow-hidden flex flex-col aspect-[16/9] hover:border-red-300 hover:shadow-md transition-all ${isDragging ? "shadow-xl" : ""}`}
    >
      {/* Header bar */}
      <div className="bg-red-600 px-2 py-[3px] flex items-center justify-between shrink-0">
        <span className="text-white text-[9px] font-semibold truncate leading-tight flex-1 mr-1">
          P.{index + 1} · {displayName}
        </span>
        <div
          {...attributes}
          {...listeners}
          className="cursor-grab active:cursor-grabbing p-0.5 -mr-0.5 touch-none shrink-0"
        >
          <GripVertical className="w-3 h-3 text-red-200" />
        </div>
      </div>

      {/* Slide content — clickable */}
      <div
        className="flex-1 overflow-hidden cursor-pointer min-h-0"
        onClick={onSelect}
        role="button"
        aria-label={`${displayName}の入力フォームへ`}
      >
        <SlideMiniPreview
          sectionKey={slide.sectionKey}
          type={slide.type}
          content={content}
          isCover={isCover}
          title={title}
          clientName={clientName}
          date={date}
        />
      </div>

      {/* Footer: layout name (non-cover) or nothing (cover) */}
      {!isCover && (
        <div className="bg-gray-50 border-t border-gray-100 px-2 py-[2px] shrink-0 flex items-center justify-between gap-1">
          <span className="text-[7px] text-gray-400 truncate flex-1">{getLayoutName(slide.sectionKey, slide.type)}</span>
          <span className="text-[7px] text-gray-400 shrink-0">KANSAI SUPER STUDIO{date ? `　${date}` : ""}</span>
        </div>
      )}
    </div>
  );
}

function DragOverlayThumbnail({ slide, index, content, date, title, clientName }: Omit<ThumbnailProps, "onSelect">) {
  const isCover = slide.sectionKey === "greeting";
  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;
  return (
    <div className="bg-white rounded-lg border border-red-300 overflow-hidden flex flex-col aspect-[16/9] shadow-2xl rotate-1 scale-105">
      <div className="bg-red-600 px-2 py-[3px] flex items-center justify-between shrink-0">
        <span className="text-white text-[9px] font-semibold truncate leading-tight">
          P.{index + 1} · {displayName}
        </span>
        <GripVertical className="w-3 h-3 text-red-200 shrink-0" />
      </div>
      <div className="flex-1 overflow-hidden min-h-0">
        <SlideMiniPreview
          sectionKey={slide.sectionKey}
          type={slide.type}
          content={content}
          isCover={isCover}
          title={title}
          clientName={clientName}
          date={date}
        />
      </div>
      {!isCover && (
        <div className="bg-gray-50 border-t border-gray-100 px-2 py-[2px] shrink-0 flex items-center justify-between gap-1">
          <span className="text-[7px] text-gray-400 truncate flex-1">{getLayoutName(slide.sectionKey, slide.type)}</span>
          <span className="text-[7px] text-gray-400 shrink-0">KANSAI SUPER STUDIO{date ? `　${date}` : ""}</span>
        </div>
      )}
    </div>
  );
}

// ─── PageContentForm ──────────────────────────────────────────────────────────
interface PageContentFormProps {
  sectionKey: string;
  sectionName: string;
  layoutName: string;
  slideCount: number;
  content: SlideContent;
  onContentChange: (sectionKey: string, fieldKey: string, value: string) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

function PageContentForm({
  sectionKey,
  sectionName,
  layoutName,
  slideCount,
  content,
  onContentChange,
  open,
  onOpenChange,
}: PageContentFormProps) {
  const fields = getFieldsForSection(sectionKey);

  return (
    <Collapsible open={open} onOpenChange={onOpenChange}>
      <Card
        id={`form-section-${sectionKey}`}
        className="bg-white border-gray-200 overflow-hidden scroll-mt-4"
      >
        <CollapsibleTrigger asChild>
          <button
            type="button"
            className="w-full text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-red-400 focus-visible:ring-offset-1 rounded-t-lg"
          >
            <div className="flex items-center justify-between px-4 py-3 bg-gray-50 border-b border-gray-200 hover:bg-gray-100 transition-colors">
              <div className="flex items-center gap-2 min-w-0">
                <span className="font-semibold text-sm text-gray-800 shrink-0">{sectionName}</span>
                {slideCount > 1 && (
                  <span className="text-xs text-gray-500 shrink-0">({slideCount}ページ)</span>
                )}
                <span className="text-xs font-medium px-2 py-0.5 bg-red-50 text-red-600 rounded whitespace-nowrap shrink-0 hidden sm:inline">
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

// ─── Main component ───────────────────────────────────────────────────────────
export default function Home() {
  const [title, setTitle]           = useState("企画書タイトル");
  const [date, setDate]             = useState("2026/06/30");
  const [clientName, setClientName] = useState("");
  const [companyLogo, setCompanyLogo] = useState(""); // reserved for future logo upload
  const [pages, setPages]           = useState<PageConfig[]>(initialPages);
  const [selectedSlides, setSelectedSlides] = useState<SelectedSlide[]>(() =>
    generateSelectedSlides(initialPages)
  );
  const [pageContents, setPageContents] = useState<Record<string, SlideContent>>({});

  // Which form sections are expanded
  const [openSections, setOpenSections] = useState<Set<string>>(() => {
    const first = initialPages.find((p) => p.checked)?.sectionKey;
    return first ? new Set([first]) : new Set();
  });

  const [activeId, setActiveId]       = useState<string | null>(null);
  const [exporting, setExporting]     = useState(false);
  const [restoreMessage, setRestoreMessage] = useState("");
  const [saveMessage, setSaveMessage] = useState("");
  const [saveStatus, setSaveStatus]   = useState<"success" | "error" | "">("");
  const [savedAt, setSavedAt]         = useState<string>("");
  const didRestoreRef = useRef(false);
  const saveTimerRef  = useRef<ReturnType<typeof setTimeout> | null>(null);

  // ── Restore ──────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (didRestoreRef.current) return;
    didRestoreRef.current = true;

    const draft = loadDraft();
    if (!draft) return;

    const restoredPages = mergePages(initialPages, draft.pages);
    const orderedSlides = applySlideOrder(generateSelectedSlides(restoredPages), draft.slideOrder);

    setTitle(draft.title ?? "企画書タイトル");
    setDate(draft.date   ?? "2026/06/30");
    setClientName(draft.clientName ?? "");
    setCompanyLogo(typeof draft.companyLogo === "string" ? draft.companyLogo : "");
    setPages(restoredPages);
    setSelectedSlides(orderedSlides);
    setPageContents(
      draft.pageContents && typeof draft.pageContents === "object" ? draft.pageContents : {}
    );
    setSavedAt(draft.savedAt ?? "");

    const firstKey = orderedSlides[0]?.sectionKey;
    if (firstKey) setOpenSections(new Set([firstKey]));

    setRestoreMessage("前回の内容を復元しました");
    setTimeout(() => setRestoreMessage(""), 3000);
  }, []);

  // ── Tap-to-scroll ─────────────────────────────────────────────────────────
  const handleSelectSlide = useCallback((sectionKey: string) => {
    setOpenSections((prev) => new Set(Array.from(prev).concat(sectionKey)));
    requestAnimationFrame(() => {
      document.getElementById(`form-section-${sectionKey}`)?.scrollIntoView({
        behavior: "smooth",
        block: "start",
      });
    });
  }, []);

  const handleSectionOpenChange = useCallback((sectionKey: string, open: boolean) => {
    setOpenSections((prev) => {
      const next = new Set(Array.from(prev));
      if (open) next.add(sectionKey);
      else next.delete(sectionKey);
      return next;
    });
  }, []);

  // ── Content change ────────────────────────────────────────────────────────
  const handleContentChange = (sectionKey: string, fieldKey: string, value: string) => {
    setPageContents((prev) => ({
      ...prev,
      [sectionKey]: { ...(prev[sectionKey] ?? EMPTY_CONTENT), [fieldKey]: value },
    }));
  };

  // ── Save ──────────────────────────────────────────────────────────────────
  const handleSave = () => {
    const now = new Date().toISOString();
    const ok = saveDraft({
      title, date, clientName, companyLogo,
      pages: pages.map(({ sectionKey, checked, pageCount, type }) => ({ sectionKey, checked, pageCount, type })),
      slideOrder: selectedSlides.map((s) => s.id),
      pageContents,
      savedAt: now,
    });
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    if (ok) { setSavedAt(now); setSaveMessage("保存しました"); setSaveStatus("success"); }
    else    { setSaveMessage("保存できませんでした"); setSaveStatus("error"); }
    saveTimerRef.current = setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 3000);
  };

  // ── New document ──────────────────────────────────────────────────────────
  const handleNew = () => {
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    setTitle("企画書タイトル");
    setDate(todayString());
    setClientName("");
    setCompanyLogo("");
    setPages(initialPages);
    setSelectedSlides(generateSelectedSlides(initialPages));
    setPageContents({});
    setSavedAt("");
    const first = initialPages.find((p) => p.checked)?.sectionKey;
    setOpenSections(first ? new Set([first]) : new Set());
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setSaveMessage("新しい企画書を作成しました");
    setSaveStatus("success");
    saveTimerRef.current = setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 3000);
  };

  // ── Page config ───────────────────────────────────────────────────────────
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
        prev.map((s) => s.sectionKey === sectionKey ? { ...s, type: value as PageType } : s)
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

  // ── DnD ──────────────────────────────────────────────────────────────────
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

  // ── Export ────────────────────────────────────────────────────────────────
  const handleExport = async () => {
    if (selectedSlides.length === 0) {
      setSaveMessage("スライドが選択されていません");
      setSaveStatus("error");
      setTimeout(() => { setSaveMessage(""); setSaveStatus(""); }, 2500);
      return;
    }
    setExporting(true);
    try {
      await generatePptx({ title, clientName, date, companyLogo, selectedSlides });
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

  // ── Derived ───────────────────────────────────────────────────────────────
  const slidesBySection   = selectedSlides.reduce(
    (acc, s) => { (acc[s.sectionKey] ??= []).push(s); return acc; },
    {} as Record<string, SelectedSlide[]>
  );
  const orderedSectionKeys = Array.from(new Set(selectedSlides.map((s) => s.sectionKey)));

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-gray-100">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-4 lg:px-6 py-3 shadow-sm sticky top-0 z-10">
        <div className="flex items-center justify-between max-w-[1400px] mx-auto">
          <div className="flex items-center gap-4">
            <h1 className="text-xl lg:text-2xl font-bold text-red-600 tracking-wide">KSS TOOL</h1>
            <span className="hidden sm:inline-block text-gray-400 text-sm">営業企画書生成ツール</span>
          </div>
          {restoreMessage && (
            <div className="text-sm px-4 py-2 bg-gray-800 text-white rounded-lg">{restoreMessage}</div>
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
                <Input id="title" value={title} onChange={(e) => setTitle(e.target.value)} className="h-11 text-base" placeholder="タイトルを入力" />
              </div>
              <div className="space-y-2">
                <Label htmlFor="date" className="text-sm font-medium text-gray-700">日付</Label>
                <div className="relative">
                  <Input id="date" value={date} onChange={(e) => setDate(e.target.value)} className="h-11 text-base pl-10" placeholder="YYYY/MM/DD" />
                  <CalendarIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400" />
                </div>
              </div>
              <div className="space-y-2">
                <Label htmlFor="client" className="text-sm font-medium text-gray-700">提案先名</Label>
                <Input id="client" value={clientName} onChange={(e) => setClientName(e.target.value)} className="h-11 text-base" placeholder="株式会社〇〇" />
              </div>
            </div>
          </CardContent>
        </Card>

        {/* ── Page structure + Overall preview ── */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
          {/* Left: page structure selector */}
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
                    <div key={page.sectionKey} className="flex flex-col gap-2 p-3 bg-gray-50 rounded-lg border border-gray-200 hover:bg-gray-100 transition-colors">
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

          {/* Right: overall preview (slide deck) */}
          <Card className="lg:col-span-2 bg-white border-gray-200">
            <CardHeader className="pb-2 px-4 pt-4">
              <div className="flex items-center justify-between">
                <CardTitle className="text-lg font-semibold text-gray-800">全体プレビュー</CardTitle>
                <span className="text-xs text-gray-400">サムネイルをタップ → 入力フォームへ</span>
              </div>
            </CardHeader>
            <CardContent className="px-4 pb-4">
              <div className="max-h-[560px] overflow-y-auto pr-1">
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
                      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
                        {selectedSlides.map((slide, index) => (
                          <SortableThumbnail
                            key={slide.id}
                            slide={slide}
                            index={index}
                            content={pageContents[slide.sectionKey] ?? EMPTY_CONTENT}
                            date={date}
                            title={title}
                            clientName={clientName}
                            onSelect={() => handleSelectSlide(slide.sectionKey)}
                          />
                        ))}
                      </div>
                    </SortableContext>
                    <DragOverlay>
                      {activeSlide && (
                        <DragOverlayThumbnail
                          slide={activeSlide}
                          index={activeIndex}
                          content={pageContents[activeSlide.sectionKey] ?? EMPTY_CONTENT}
                          date={date}
                          title={title}
                          clientName={clientName}
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
            <Button onClick={handleSave} className="h-11 px-6 bg-gray-700 hover:bg-gray-800 text-white flex items-center gap-2">
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
              onClick={handleExport} disabled={exporting}
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

        {/* ── Page content input forms ── */}
        <Card className="bg-white border-gray-200">
          <CardHeader className="pb-2 px-4 pt-4">
            <CardTitle className="text-lg font-semibold text-gray-800">ページ内容入力</CardTitle>
          </CardHeader>
          <CardContent className="px-4 pb-4">
            {selectedSlides.length === 0 ? (
              <div className="text-center text-gray-400 py-8">ページが選択されていません</div>
            ) : (
              <div className="space-y-3">
                {orderedSectionKeys.map((sectionKey) => {
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
                      open={openSections.has(sectionKey)}
                      onOpenChange={(v) => handleSectionOpenChange(sectionKey, v)}
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
