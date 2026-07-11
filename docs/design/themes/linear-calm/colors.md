# Linear Calm — 色彩系統

## 基礎調色盤

| 角色 | 色值 | 用途 |
|---|---|---|
| `bg/canvas` | `#F8FAFC` | 全頁背景（雲白） |
| `bg/surface` | `#FFFFFF` | 卡面、面板、modal |
| `bg/subtle` | `#F1F5F9` | 次級區塊、輸入框底、hover 底 |
| `text/primary` | `#1E293B` | 主文字（深藍灰） |
| `text/secondary` | `#64748B` | 次要文字、說明、placeholder |
| `border/hairline` | `#E2E8F0` | 細分隔線、卡片邊、輸入框邊 |
| `accent/primary` | `#6366F1` | 主按鈕、選取、重點數字（靛藍紫） |
| `accent/soft` | `#818CF8` | hover、次級強調、背景微光（柔紫藍） |
| `accent/glow` | `rgba(99,102,241,.12)` | 焦點環底、微光底 |
| `warn/amber` | `#F59E0B` | 提醒、待處理（晨光金，少量） |

## 狀態色（沿用既有 token schema）

| Token | 色值 | 說明 |
|---|---|---|
| `statusNotCheckedIn` | `#818CF8` | 未簽到（淡紫，呼吸燈主色） |
| `statusCheckedIn` | `#10B981` | 已報到（綠色，保留辨識度，調為較冷的祖母綠） |
| `statusCheckedOut` | `#94A3B8` | 已簽退（淺灰） |
| `breathColor` | `#A5B4FC` | 呼吸動畫過渡色（更淡的紫） |
| `breathShadow` | `rgba(129,140,248,.40)` | 呼吸光暈（紫） |

非 token 狀態（定義於 `main.css :root`）：

| 變數 | 色值 | 說明 |
|---|---|---|
| `--status-can-color` | `#475569` | 已取消（深灰文字） |
| `--status-can-bg` | `#E2E8F0` | 已取消底（淺 slate） |
| `--status-can-border` | `#94A3B8` | 已取消邊框 |

> 決策：已報到維持綠色系（非紫），因為點名場景下顏色辨識最直覺。未簽到用淡紫呼吸態、已簽退用淺灰、取消用深灰，整體保持冷靜調性且狀態彼此可辨。

## CSS 變數命名

主題 token 對應到 CSS 變數（定義於 `theme-linear-calm.css`）：

```css
--lc-bg-canvas:    #F8FAFC;
--lc-bg-surface:   #FFFFFF;
--lc-bg-subtle:    #F1F5F9;
--lc-text-primary: #1E293B;
--lc-text-secondary:#64748B;
--lc-border:       #E2E8F0;
--lc-accent:       #6366F1;
--lc-accent-soft:  #818CF8;
--lc-accent-glow:  rgba(99,102,241,.12);
--lc-amber:        #F59E0B;
```

狀態色仍透過既有 `--sr-status-*` / `--status-*` 變數系統供應，由 `theme-status-tokens.js` 注入，確保表格模式與卡片模式同步。
