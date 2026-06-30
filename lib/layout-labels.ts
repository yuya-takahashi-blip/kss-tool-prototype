export type PageType = "A" | "B" | "C";

export const layoutLabels: Record<string, Record<PageType, string>> = {
  greeting: {
    A: "標準挨拶文",
    B: "担当者コメント付き",
    C: "代表メッセージ風",
  },
  brand: {
    A: "ロゴ＋説明文＋キービジュアル",
    B: "世界観画像メイン",
    C: "特徴・ターゲット整理",
  },
  business_scheme: {
    A: "2社間スキーム",
    B: "3者間スキーム",
    C: "販売までのフロー",
  },
  cases: {
    A: "1事例を大きく紹介",
    B: "3事例を一覧表示",
    C: "Before / After形式",
  },
  items: {
    A: "商品画像1点を大きく表示",
    B: "商品画像4点グリッド",
    C: "使用シーン画像＋商品画像",
  },
  schedule: {
    A: "横軸タイムライン",
    B: "月別ガントチャート",
    C: "フェーズ別ステップ",
  },
  pricing: {
    A: "条件一覧表",
    B: "プラン比較表",
    C: "ロイヤリティ条件中心",
  },
  company: {
    A: "基本情報",
    B: "沿革・実績付き",
    C: "強み・サービス紹介",
  },
};

export function getLayoutName(sectionKey: string, type: PageType): string {
  return layoutLabels[sectionKey]?.[type] ?? type;
}
