import type { PageType } from "@/lib/layout-labels";

// Re-export PageType so consumers can import all types from one module
export type { PageType };

export type SlideContent = Record<string, string>;

export interface PageConfig {
  sectionKey: string;
  sectionName: string;
  checked: boolean;
  pageCount: number;
  type: PageType;
}

export interface SelectedSlide {
  id: string;
  sectionKey: string;
  sectionName: string;
  type: PageType;
  slideIndexInSection: number;
}

export interface DraftData {
  title: string;
  date: string;
  clientName: string;
  companyLogo: string;
  pages: Array<{ sectionKey: string; checked: boolean; pageCount: number; type: PageType }>;
  slideOrder: string[];
  pageContents: Record<string, SlideContent>;
  savedAt: string;
}
