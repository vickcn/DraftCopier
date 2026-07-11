# Linear Calm — 元件範例

## 主按鈕

實心靛藍紫，hover 轉柔紫藍，focus 用 glow 焦點環。

```css
.btn-primary {
  background: var(--lc-accent);
  color: #fff;
  border-radius: 12px;
}
.btn-primary:hover { background: var(--lc-accent-soft); }
.btn-primary:focus-visible {
  box-shadow: 0 0 0 3px rgba(99,102,241,.18);
}
```

## 次按鈕

白底細邊，hover 帶極淡微光底。

```css
.btn-secondary {
  background: var(--lc-bg-surface);
  color: var(--lc-text-primary);
  border: 1px solid var(--lc-border);
  border-radius: 12px;
}
.btn-secondary:hover { background: var(--lc-bg-subtle); }
```

## 卡片 / 面板

16px 圓角，細邊，柔陰影，不用漸層底。

```css
.card {
  background: var(--lc-bg-surface);
  border: 1px solid var(--lc-border);
  border-radius: 16px;
  box-shadow: 0 1px 2px rgba(15,23,42,.04), 0 4px 12px rgba(15,23,42,.05);
}
```

## 重要數字

字重 700、靛藍紫、tabular-nums，底可加極淡微光。

```css
.stat-number {
  font-weight: 700;
  color: var(--lc-accent);
  font-variant-numeric: tabular-nums;
}
```

## 狀態提示

- 未報到：晨光金，呼吸燈。
- 已報到：綠色 chip / 列底。
- 已簽退：柔紫藍。

```css
.tag-checked-in  { color: var(--sr-status-checked-in); }
.tag-checked-out { color: var(--sr-status-checked-out); }
.tag-pending     { color: var(--sr-status-not-checked-in); }
```

## 分隔線

```css
.divider { border-top: 1px solid var(--lc-border); }
```

## 產品級漸層（允許）

同色系、低飽和、只用在重點。

```css
/* 主按鈕 */
.btn-primary { background: var(--lc-grad-accent); }
.btn-primary:hover { background: var(--lc-grad-accent-hover); }

/* 重要數字漸層文字 */
.stat-number {
  background: var(--lc-grad-text);
  -webkit-background-clip: text; background-clip: text;
  -webkit-text-fill-color: transparent;
}

/* 卡片極淡表面層次 */
.card { background: var(--lc-grad-surface); }

/* 背景微光（全頁）*/
body {
  background:
    radial-gradient(1100px 520px at 50% -8%, rgba(99,102,241,.07), transparent 60%),
    #F8FAFC;
}
```

## 不要這樣做

- ❌ `background: linear-gradient(135deg,#667eea,#764ba2)` 高飽和撞色滿版漸層
- ❌ `box-shadow: 0 20px 40px rgba(0,0,0,.1)` 厚陰影
- ❌ 高飽和霓虹色、外發光
- ❌ 厚重邊框（2px 以上彩色邊框）作為常態
- ❌ 漸層用在大面積背景或非重點區塊
