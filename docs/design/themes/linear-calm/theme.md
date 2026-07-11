# Theme: Linear Calm

> 對應 Philosophy：[victor-academy](../../philosophy/victor-academy.md)
> 風格定位：Linear + Apple + 文青教育品牌

## 一句話

冷靜、理智、乾淨、有未來感，但不炫技。雲白底、細分隔線、柔陰影、16px 圓角；紫藍光只用在重點。

## 設計方向

### 版面

- 整頁背景使用雲白 `bg/canvas`，不使用滿版漸層。
- 內容區以接近純白的浮層（`bg/surface`）浮在雲白上。
- 區塊分層靠「1px 細分隔線 + 柔和微陰影」，不靠濃重卡片或厚邊框。
- 大量留白，資訊層級靠字級與顏色深淺，不靠陰影堆疊。

### 重點才發光

- `accent/primary`（靛藍紫）只出現在主要 CTA、目前選取狀態、關鍵數字。
- `accent/soft`（柔紫藍）作為 hover、次級強調、背景微光。
- `warn/amber`（晨光金）只做少量狀態提示（待處理、未報到呼吸燈），不作主色。

## 視覺規格

| 項目 | 值 |
|---|---|
| 卡片 / 面板圓角 | 16px |
| 按鈕 / 輸入框圓角 | 12px |
| 標籤圓角 | 8px |
| 卡片陰影 | `0 1px 2px rgba(15,23,42,.04), 0 4px 12px rgba(15,23,42,.05)` |
| 浮層 / modal 陰影 | `0 8px 28px rgba(15,23,42,.12)` |
| 焦點環 | `0 0 0 3px rgba(99,102,241,.18)` |
| 分隔線 | 1px `border/hairline` |
| 標題字體 | Inter / `-apple-system` |
| 內文字體 | Inter / `-apple-system` / Roboto fallback |
| 重要數字 | 字重 700、`accent/primary`、可用 tabular-nums |

## 漸層（產品級）

允許「克制、有質感」的漸層，定位同 Linear / Stripe / Apple：同色系、低飽和、只用在重點。

| Token | 值 | 用途 |
|---|---|---|
| `--lc-grad-accent` | `135deg, #6366F1→#818CF8` | 主按鈕、選取分頁 |
| `--lc-grad-accent-hover` | `135deg, #4F46E5→#6D67EC` | 主按鈕 hover（略深） |
| `--lc-grad-surface` | `180deg, #FFFFFF→#FBFBFE` | 卡片表面極淡層次 |
| `--lc-grad-text` | `135deg, #6366F1→#818CF8` | 重要數字漸層文字 |
| 背景微光 | 頂部 radial `rgba(99,102,241,.07)` + 右上 `rgba(129,140,248,.05)` | 全頁空間感 |

## 邊界

- 不用電競、霓虹、賽博龐克元素。
- 漸層只走低飽和同色系、低透明度；不用高飽和撞色漸層（例如舊版 `#667eea→#764ba2` 滿版）。
- 不用厚重卡片與大陰影。
- DB 只存 token，不存 animation keyframes 與圖片。

## 實作對應

- Token / override 樣式：`public/css/theme-linear-calm.css`
- 狀態色 fallback：`public/js/theme-status-tokens.js`、`public/css/main.css :root`
- 套用流程詳見 [UI Theme 開發指南](../../../UI_THEME_DEVELOPMENT.md)
