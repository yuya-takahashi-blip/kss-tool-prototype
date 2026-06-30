"use client";

import { useState, useCallback, useEffect, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { CalendarIcon, FileText, Download, Loader2, GripVertical, Save, FilePlus } from "lucide-react";
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

interface DraftData {
  title: string;
  date: string;
  clientName: string;
  pages: Array<{ sectionKey: string; checked: boolean; pageCount: number; type: PageType }>;
  slideOrder: string[]; // ordered list of slide ids
  savedAt: string;
}

// ─── Constants ────────────────────────────────────────────────────────────────
const STORAGE_KEY = "kss-tool-proposal-draft";

const initialPages: PageConfig[] = [
  { sectionKey: "greeting",        sectionName: "ご挨拶",           checked: true, pageCount: 1, type: "A", previewFields: ["挨拶文", "署名"] },
  { sectionKey: "brand",           sectionName: "ブランド紹介",      checked: true, pageCount: 1, type: "A", previewFields: ["ブランド概要", "特徴", "実績"] },
  { sectionKey: "business_scheme", sectionName: "ビジネススキーム",  checked: true, pageCount: 1, type: "A", previewFields: ["関係図", "フロー図解"] },
  { sectionKey: "cases",           sectionName: "事例紹介",          checked: true, pageCount: 1, type: "A", previewFields: ["導入事例", "効果", "お客様の声"] },
  { sectionKey: "items",           sectionName: "アイテムイメージ",   checked: true, pageCount: 1, type: "A", previewFields: ["商品画像", "商品名", "価格"] },
  { sectionKey: "schedule",        sectionName: "スケジュール",       checked: true, pageCount: 1, type: "A", previewFields: ["タイムライン", "マイルストーン"] },
  { sectionKey: "pricing",         sectionName: "料金・条件",         checked: true, pageCount: 1, type: "A", previewFields: ["料金表", "条件", "特典"] },
  { sectionKey: "company",         sectionName: "会社概要",           checked: true, pageCount: 1, type: "A", previewFields: ["会社名", "代表者", "所在地"] },
];

// ─── localStorage helpers ─────────────────────────────────────────────────────
function loadDraft(): DraftData | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as DraftData;
    // Minimal validation
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
    const h = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${y}/${mo}/${dy} ${h}:${mi}`;
  } catch {
    return "";
  }
}

function todayString(): string {
  const d = new Date();
  const y = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const dy = String(d.getDate()).padStart(2, "0");
  return `${y}/${mo}/${dy}`;
}

// ─── Rebuild slides from pages, then apply a saved order ─────────────────────
function generateSelectedSlides(pages: PageConfig[]): SelectedSlide[] {
  return pages
    .filter((page) => page.checked)
    .flatMap((page) =>
      Array.from({ length: page.pageCount }, (_, i) => ({
        id: `${page.sectionKey}-${i}`,
        sectionKey: page.sectionKey,
        sectionName: page.sectionName,
        type: page.type,
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
  // Append any new slides not in saved order (e.g. newly added sections)
  map.forEach((s) => ordered.push(s));
  return ordered;
}

// ─── Merge saved page settings onto initialPages ─────────────────────────────
function mergePages(
  base: PageConfig[],
  saved: DraftData["pages"]
): PageConfig[] {
  const savedMap = new Map(saved.map((s) => [s.sectionKey, s]));
  return base.map((p) => {
    const s = savedMap.get(p.sectionKey);
    if (!s) return p;
    return {
      ...p,
      checked: typeof s.checked === "boolean" ? s.checked : p.checked,
      pageCount: typeof s.pageCount === "number" && s.pageCount >= 1 ? s.pageCount : p.pageCount,
      type: (["A", "B", "C"] as PageType[]).includes(s.type as PageType) ? s.type : p.type,
    };
  });
}

// ─── Thumbnail components ─────────────────────────────────────────────────────
function SortableThumbnail({ slide, index }: { slide: SelectedSlide; index: number }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: slide.id });

  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;

  return (
    <div
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition, opacity: isDragging ? 0.5 : 1 }}
      className={`bg-gray-50 rounded-lg border border-gray-200 p-3 aspect-[4/3] flex flex-col justify-between hover:border-gray-300 transition-colors ${isDragging ? "shadow-lg" : ""}`}
    >
      <div>
        <div className="flex items-center justify-between mb-1">
          <div className="text-xs text-gray-500">P.{index + 1}</div>
          <div {...attributes} {...listeners} className="cursor-grab active:cursor-grabbing p-1 -m-1 touch-none">
            <GripVertical className="w-4 h-4 text-gray-400" />
          </div>
        </div>
        <div className="font-medium text-sm text-gray-800 truncate">{displayName}</div>
      </div>
      <div className="text-xs text-red-600 bg-red-50 px-2 py-1 rounded leading-tight self-start max-w-full truncate">
        {getLayoutName(slide.sectionKey, slide.type)}
      </div>
    </div>
  );
}

function DragOverlayThumbnail({ slide, index }: { slide: SelectedSlide; index: number }) {
  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;
  return (
    <div className="bg-white rounded-lg border border-red-300 p-3 aspect-[4/3] flex flex-col justify-between shadow-xl rotate-2 scale-105">
      <div>
        <div className="flex items-center justify-between mb-1">
          <div className="text-xs text-gray-500">P.{index + 1}</div>
          <GripVertical className="w-4 h-4 text-gray-400" />
        </div>
        <div className="font-medium text-sm text-gray-800 truncate">{displayName}</div>
      </div>
      <div className="text-xs text-red-600 bg-red-50 px-2 py-1 rounded leading-tight self-start max-w-full truncate">
        {getLayoutName(slide.sectionKey, slide.type)}
      </div>
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────
export default function Home() {
  // Initialize state lazily so localStorage is read only on first render (client only)
  const [title, setTitle] = useState("企画書タイトル");
  const [date, setDate] = useState("2026/06/30");
  const [clientName, setClientName] = useState("");
  const [pages, setPages] = useState<PageConfig[]>(initialPages);
  const [selectedSlides, setSelectedSlides] = useState<SelectedSlide[]>(() =>
    generateSelectedSlides(initialPages)
  );
  const [activeId, setActiveId] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [restoreMessage, setRestoreMessage] = useState(""); // 復元通知（ヘッダー）
  const [saveMessage, setSaveMessage] = useState("");       // 保存通知（ボタン付近）
  const [saveStatus, setSaveStatus] = useState<"success" | "error" | "">("");
  const [savedAt, setSavedAt] = useState<string>("");
  const didRestoreRef = useRef(false);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // ── Restore from localStorage on mount (runs client-side only) ──────────────
  useEffect(() => {
    if (didRestoreRef.current) return;
    didRestoreRef.current = true;

    const draft = loadDraft();
    if (!draft) return;

    const restoredPages = mergePages(initialPages, draft.pages);
    const rawSlides = generateSelectedSlides(restoredPages);
    const orderedSlides = applySlideOrder(rawSlides, draft.slideOrder);

    setTitle(draft.title ?? "企画書タイトル");
    setDate(draft.date ?? "2026/06/30");
    setClientName(draft.clientName ?? "");
    setPages(restoredPages);
    setSelectedSlides(orderedSlides);
    setSavedAt(draft.savedAt ?? "");

    setRestoreMessage("前回の内容を復元しました");
    setTimeout(() => setRestoreMessage(""), 3000);
  }, []);

  // ── Save handler ─────────────────────────────────────────────────────────────
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
    saveTimerRef.current = setTimeout(() => {
      setSaveMessage("");
      setSaveStatus("");
    }, 3000);
  };

  // ── New document handler ──────────────────────────────────────────────────────
  const handleNew = () => {
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
    setTitle("企画書タイトル");
    setDate(todayString());
    setClientName("");
    setPages(initialPages);
    setSelectedSlides(generateSelectedSlides(initialPages));
    setSavedAt("");
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setSaveMessage("新しい企画書を作成しました");
    setSaveStatus("success");
    saveTimerRef.current = setTimeout(() => {
      setSaveMessage("");
      setSaveStatus("");
    }, 3000);
  };

  // ── Page config change ────────────────────────────────────────────────────────
  const updateSelectedSlidesForSection = useCallback(
    (sectionKey: string, checked: boolean, pageCount: number, type: PageType) => {
      setSelectedSlides((prevSlides) => {
        const page = pages.find((p) => p.sectionKey === sectionKey);
        if (!page) return prevSlides;
        const otherSlides = prevSlides.filter((s) => s.sectionKey !== sectionKey);
        if (!checked || pageCount === 0) return otherSlides;
        const newSlides: SelectedSlide[] = Array.from({ length: pageCount }, (_, i) => ({
          id: `${sectionKey}-${i}`,
          sectionKey,
          sectionName: page.sectionName,
          type,
          slideIndexInSection: i + 1,
        }));
        return [...otherSlides, ...newSlides];
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
      prev.map((page) => (page.sectionKey !== sectionKey ? page : { ...page, [field]: value }))
    );

    const page = pages.find((p) => p.sectionKey === sectionKey);
    if (!page) return;

    if (field === "type" && page.checked) {
      setSelectedSlides((prevSlides) =>
        prevSlides.map((slide) =>
          slide.sectionKey === sectionKey ? { ...slide, type: value as PageType } : slide
        )
      );
    }

    if (field === "checked" || field === "pageCount") {
      const newChecked = field === "checked" ? (value as boolean) : page.checked;
      const newPageCount = field === "pageCount" ? (value as number) : page.pageCount;
      updateSelectedSlidesForSection(sectionKey, newChecked, newPageCount, page.type);
    }
  };

  // ── DnD ──────────────────────────────────────────────────────────────────────
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 8 } }),
    useSensor(TouchSensor, { activationConstraint: { delay: 200, tolerance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  const handleDragStart = (event: DragStartEvent) => setActiveId(event.active.id as string);

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
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
  const getPreviewFields = (sectionKey: string): string[] =>
    pages.find((p) => p.sectionKey === sectionKey)?.previewFields || [];

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
        {/* Input Fields */}
        <Card className="mb-4 bg-white border-gray-200">
          <CardContent className="p-4">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="space-y-2">
                <Label htmlFor="title" className="text-sm font-medium text-gray-700">
                  企画書タイトル
                </Label>
                <Input
                  id="title"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  className="h-11 text-base"
                  placeholder="タイトルを入力"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="date" className="text-sm font-medium text-gray-700">
                  日付
                </Label>
                <div className="relative">
                  <Input
                    id="date"
                    value={date}
                    onChange={(e) => setDate(e.target.value)}
                    className="h-11 text-base pl-10"
                    placeholder="YYYY/MM/DD"
                  />
                  <CalendarIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400" />
                </div>
              </div>
              <div className="space-y-2">
                <Label htmlFor="client" className="text-sm font-medium text-gray-700">
                  提案先名
                </Label>
                <Input
                  id="client"
                  value={clientName}
                  onChange={(e) => setClientName(e.target.value)}
                  className="h-11 text-base"
                  placeholder="株式会社〇〇"
                />
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Main Content */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
          {/* Page Structure */}
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
                      {/* 1行目：チェックボックス＋ページ名全文 */}
                      <div className="flex items-center gap-3">
                        <Checkbox
                          id={page.sectionKey}
                          checked={page.checked}
                          onCheckedChange={(checked) =>
                            handlePageChange(page.sectionKey, "checked", checked as boolean)
                          }
                          className="w-5 h-5 shrink-0"
                        />
                        <Label
                          htmlFor={page.sectionKey}
                          className="text-sm font-medium text-gray-800 cursor-pointer"
                        >
                          {page.sectionName}
                        </Label>
                      </div>
                      {/* 2行目：ページ数＋構成選択 */}
                      <div className="flex items-center gap-2 ml-8">
                        <div className="flex items-center gap-1 shrink-0">
                          <Label className="text-xs text-gray-500 whitespace-nowrap">
                            ページ数
                          </Label>
                          <Input
                            type="number"
                            min="1"
                            max="10"
                            value={page.pageCount}
                            onChange={(e) =>
                              handlePageChange(
                                page.sectionKey,
                                "pageCount",
                                parseInt(e.target.value) || 1
                              )
                            }
                            className="w-14 h-9 text-center text-sm"
                            disabled={!page.checked}
                          />
                        </div>
                        <Select
                          value={page.type}
                          onValueChange={(value: PageType) =>
                            handlePageChange(page.sectionKey, "type", value)
                          }
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

          {/* Overall Preview */}
          <Card className="lg:col-span-2 bg-white border-gray-200">
            <CardHeader className="pb-2 px-4 pt-4">
              <CardTitle className="text-lg font-semibold text-gray-800">全体プレビュー</CardTitle>
            </CardHeader>
            <CardContent className="px-4 pb-4">
              <div className="max-h-[400px] overflow-y-auto pr-2">
                {selectedSlides.length === 0 ? (
                  <div className="text-center text-gray-400 py-12">
                    ページが選択されていません
                  </div>
                ) : (
                  <DndContext
                    sensors={sensors}
                    collisionDetection={closestCenter}
                    onDragStart={handleDragStart}
                    onDragEnd={handleDragEnd}
                  >
                    <SortableContext
                      items={selectedSlides.map((s) => s.id)}
                      strategy={rectSortingStrategy}
                    >
                      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
                        {selectedSlides.map((slide, index) => (
                          <SortableThumbnail key={slide.id} slide={slide} index={index} />
                        ))}
                      </div>
                    </SortableContext>
                    <DragOverlay>
                      {activeSlide && (
                        <DragOverlayThumbnail slide={activeSlide} index={activeIndex} />
                      )}
                    </DragOverlay>
                  </DndContext>
                )}
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Actions */}
        <div className="mb-4 space-y-2">
          <div className="flex flex-wrap items-center gap-3">
            <Button
              onClick={handleSave}
              className="h-11 px-6 bg-gray-700 hover:bg-gray-800 text-white flex items-center gap-2"
            >
              <Save className="w-4 h-4" />
              保存
            </Button>
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button
                  variant="outline"
                  className="h-11 px-6 border-gray-300 text-gray-700 flex items-center gap-2"
                >
                  <FilePlus className="w-4 h-4" />
                  新規作成
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
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  生成中...
                </>
              ) : (
                <>
                  <Download className="w-4 h-4" />
                  PowerPointを書き出す
                </>
              )}
            </Button>
            <span className="text-xs text-gray-400 ml-1">
              {savedAt ? `最終保存：${formatDateTime(savedAt)}` : "最終保存：未保存"}
            </span>
          </div>
          {/* Save / action message banner */}
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

        {/* Page-Specific Preview Cards */}
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
                  const firstSlide = sectionSlides[0];

                  return (
                    <Card key={sectionKey} className="bg-gray-50 border border-gray-200 overflow-hidden">
                      <CardHeader className="bg-gray-100 px-4 py-3 border-b border-gray-200">
                        <div className="flex items-center justify-between gap-2">
                          <CardTitle className="text-sm font-semibold text-gray-800 truncate">
                            {firstSlide.sectionName}
                            {sectionSlides.length > 1 && (
                              <span className="text-xs text-gray-500 ml-2">
                                ({sectionSlides.length}ページ)
                              </span>
                            )}
                          </CardTitle>
                          <span className="text-xs font-medium px-2 py-1 bg-red-50 text-red-600 rounded whitespace-nowrap shrink-0">
                            {getLayoutName(sectionKey, firstSlide.type)}
                          </span>
                        </div>
                      </CardHeader>
                      <CardContent className="p-4">
                        <div className="space-y-2">
                          <div className="text-xs font-medium text-gray-700 mb-2">入力項目:</div>
                          {getPreviewFields(sectionKey).map((field, idx) => (
                            <div
                              key={idx}
                              className="text-sm text-gray-600 bg-white px-3 py-2 rounded border border-gray-200"
                            >
                              {field}
                            </div>
                          ))}
                        </div>
                        <div className="mt-3 pt-3 border-t border-gray-200">
                          <div className="aspect-[16/9] bg-gray-200 rounded flex items-center justify-center">
                            <div className="text-xs text-gray-400">簡易プレビュー</div>
                          </div>
                        </div>
                      </CardContent>
                    </Card>
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
