# Design Documentation

本資料夾管理 Smart Registrator 的設計知識庫。

## 設計原則

本系統支援多主題（Multi Theme）架構，且分成三層管理：

1. `philosophy/`：定義品牌精神、情緒方向、決策原則、禁止事項
2. `themes/`：定義各主題的顏色、圓角、陰影、圖片與視覺規格
3. `design-system/`：定義跨主題共用的 token、元件、間距、排版與互動規則

設計時請先確認：

1. 目前專案對應的 Philosophy
2. 目前使用中的 Theme
3. Design System Token 規範
4. 元件規範與狀態規則

不要直接依照個人審美修改樣式。
不要把 Philosophy、Theme、Design System 混寫在同一份文件。

---

## 文件結構

### Philosophy

描述每個產品脈絡背後的品牌精神與設計哲學。

這一層只回答：

* 我們想讓使用者感受到什麼
* 我們在設計選擇時偏向什麼
* 什麼風格不屬於這個方向
* Theme 與元件實作時應遵守哪些邊界

這一層不回答：

* 具體使用哪個顏色
* border radius 幾 px
* 陰影、圖片、插圖、背景貼圖怎麼做
* animation keyframes 要怎麼寫

位置：

```text
docs/design/philosophy/
```

例如：

```text
victor-academy.md
church.md
space.md
```

目前建議的對應關係：

* `victor-academy.md`：培育、觀察、陪伴
* `space.md`：探索、未知、遠見
* `church.md`：關懷、信任、連結

---

### Themes

每個 Theme 的具體規範。

位置：

```text
docs/design/themes/
```

例如：

```text
linear-calm/
space-observer/
community-builder/
church-care/
```

每個主題應包含：

* theme.md
* colors.md
* examples.md

---

### Design System

跨主題共用規則，負責把 Philosophy 轉成可實作的共用接口。

位置：

```text
docs/design/design-system/
```

包含：

* tokens.md
* typography.md
* spacing.md
* components.md

---

## 設計流程

當需要建立新頁面：

1. 閱讀對應的 Philosophy
2. 閱讀目前 Theme 的 `theme.md`
3. 閱讀 Design System
4. 先提出 Wireframe 或畫面骨架
5. 再修改 UI

---

## 建立新 Theme

建立：

```text
docs/design/themes/<theme-name>/
```

至少包含：

```text
theme.md
colors.md
examples.md
```

建立前先確認它屬於哪個 Philosophy。

---

## 重要原則

* Philosophy 決定方向，不決定數值。
* Theme 是可替換的。
* Design System 是共用的。
* 不要把某個 Theme 的顏色、動畫、視覺語言寫死在共用元件中。
* 若未來 Theme 來自 DB，DB 只提供 token，不直接控制 keyframes 或元件行為。
