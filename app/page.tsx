"use client";

import { useState, useCallback } from "react";
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
import { CalendarIcon, FileText, Download, Loader2, GripVertical } from "lucide-react";
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

const initialPages: PageConfig[] = [
  {
    sectionKey: "greeting",
    sectionName: "ご挨拶",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["挨拶文", "署名"],
  },
  {
    sectionKey: "brand",
    sectionName: "ブランド紹介",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["ブランド概要", "特徴", "実績"],
  },
  {
    sectionKey: "business_scheme",
    sectionName: "ビジネススキーム",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["関係図", "フロー図解"],
  },
  {
    sectionKey: "cases",
    sectionName: "事例紹介",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["導入事例", "効果", "お客様の声"],
  },
  {
    sectionKey: "items",
    sectionName: "アイテムイメージ",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["商品画像", "商品名", "価格"],
  },
  {
    sectionKey: "schedule",
    sectionName: "スケジュール",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["タイムライン", "マイルストーン"],
  },
  {
    sectionKey: "pricing",
    sectionName: "料金・条件",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["料金表", "条件", "特典"],
  },
  {
    sectionKey: "company",
    sectionName: "会社概要",
    checked: true,
    pageCount: 1,
    type: "A",
    previewFields: ["会社名", "代表者", "所在地"],
  },
];

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

interface SortableThumbnailProps {
  slide: SelectedSlide;
  index: number;
}

function SortableThumbnail({ slide, index }: SortableThumbnailProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: slide.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
  };

  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;
  const layoutName = getLayoutName(slide.sectionKey, slide.type);

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={`bg-gray-50 rounded-lg border border-gray-200 p-3 aspect-[4/3] flex flex-col justify-between hover:border-gray-300 transition-colors ${isDragging ? "shadow-lg" : ""}`}
    >
      <div>
        <div className="flex items-center justify-between mb-1">
          <div className="text-xs text-gray-500">P.{index + 1}</div>
          <div
            {...attributes}
            {...listeners}
            className="cursor-grab active:cursor-grabbing p-1 -m-1 touch-none"
          >
            <GripVertical className="w-4 h-4 text-gray-400" />
          </div>
        </div>
        <div className="font-medium text-sm text-gray-800 truncate">{displayName}</div>
      </div>
      <div className="text-xs text-red-600 bg-red-50 px-2 py-1 rounded leading-tight self-start max-w-full truncate">
        {layoutName}
      </div>
    </div>
  );
}

interface DragOverlayThumbnailProps {
  slide: SelectedSlide;
  index: number;
}

function DragOverlayThumbnail({ slide, index }: DragOverlayThumbnailProps) {
  const displayName =
    slide.slideIndexInSection > 1
      ? `${slide.sectionName} (${slide.slideIndexInSection})`
      : slide.sectionName;
  const layoutName = getLayoutName(slide.sectionKey, slide.type);

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
        {layoutName}
      </div>
    </div>
  );
}

export default function Home() {
  const [title, setTitle] = useState("企画書タイトル");
  const [date, setDate] = useState("2026/06/30");
  const [clientName, setClientName] = useState("");
  const [pages, setPages] = useState<PageConfig[]>(initialPages);
  const [message, setMessage] = useState("");
  const [selectedSlides, setSelectedSlides] = useState<SelectedSlide[]>(() =>
    generateSelectedSlides(initialPages)
  );
  const [activeId, setActiveId] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);

  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 8 } }),
    useSensor(TouchSensor, { activationConstraint: { delay: 200, tolerance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

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
      prev.map((page) =>
        page.sectionKey !== sectionKey ? page : { ...page, [field]: value }
      )
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

  const handleDragStart = (event: DragStartEvent) => setActiveId(event.active.id as string);

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    setActiveId(null);
    if (over && active.id !== over.id) {
      setSelectedSlides((items) => {
        const oldIndex = items.findIndex((i) => i.id === active.id);
        const newIndex = items.findIndex((i) => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const activeSlide = activeId ? selectedSlides.find((s) => s.id === activeId) : null;
  const activeIndex = activeSlide ? selectedSlides.findIndex((s) => s.id === activeId) : 0;

  const handleExport = async () => {
    if (selectedSlides.length === 0) {
      setMessage("スライドが選択されていません");
      setTimeout(() => setMessage(""), 2500);
      return;
    }
    setExporting(true);
    try {
      await generatePptx({ title, clientName, date, selectedSlides });
      setMessage("ダウンロードしました");
    } catch (err) {
      console.error(err);
      setMessage("書き出しに失敗しました");
    } finally {
      setExporting(false);
      setTimeout(() => setMessage(""), 3000);
    }
  };

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

  return (
    <div className="min-h-screen bg-gray-100">
      <header className="bg-white border-b border-gray-200 px-4 lg:px-6 py-3 shadow-sm">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <h1 className="text-xl lg:text-2xl font-bold text-red-600 tracking-wide">KSS TOOL</h1>
            <span className="hidden sm:inline-block text-gray-400 text-sm">営業企画書生成ツール</span>
          </div>
          {message && (
            <div className="text-sm px-4 py-2 bg-gray-800 text-white rounded-lg">{message}</div>
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
        <div className="flex flex-wrap gap-3 mb-4">
          <Button
            onClick={() => {
              setMessage("保存しました");
              setTimeout(() => setMessage(""), 2000);
            }}
            className="h-11 px-6 bg-gray-700 hover:bg-gray-800 text-white"
          >
            保存
          </Button>
          <Button
            onClick={() => {
              setMessage("プレビューを更新しました");
              setTimeout(() => setMessage(""), 2000);
            }}
            variant="outline"
            className="h-11 px-6 border-gray-300 text-gray-700"
          >
            プレビュー更新
          </Button>
          <Button
            onClick={handleExport}
            disabled={exporting}
            className="h-11 px-6 bg-red-600 hover:bg-red-700 text-white disabled:opacity-70"
          >
            {exporting ? (
              <>
                <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                生成中...
              </>
            ) : (
              <>
                <Download className="w-4 h-4 mr-2" />
                PowerPointを書き出す
              </>
            )}
          </Button>
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
                    <Card
                      key={sectionKey}
                      className="bg-gray-50 border border-gray-200 overflow-hidden"
                    >
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
