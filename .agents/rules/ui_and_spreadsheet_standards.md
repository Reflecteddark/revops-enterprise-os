# UI, Spreadsheet, and Typography Sizing Standards

Whenever creating or modifying spreadsheets (Excel/Google Sheets), web interfaces, landing pages, documents, or reports, ALWAYS ensure strict compliance with the following sizing, typography, and color standards:

## 1. Column Widths & Content Sizing (Zero Text Truncation / No `###`)
- **No Text Clipping**: Never leave text in a column narrower than its content unless `wrap_text=True` is explicitly enabled AND row height is proportionally expanded.
- **Numbers & Currencies**: All columns containing formatted numbers (e.g., `#,##0 "₽"`) must have at least 14–18 character width to guarantee Excel never displays `###` overflow errors.
- **First Column (№ / Code / Nav)**: Minimum width of 12–16 characters so navigation bars, KPI cards, and ordinal numbers never get cramped or clipped.
- **Data Table Text Columns**: Standard text columns (descriptions, symptoms, recommendations) must have widths between 32–60 characters, with `wrap_text=True`, `vertical="center"`, and row height calculated at ≥22–26pt per line.

## 2. KPI Cards & Dashboard Headers
- **Card Merging**: Multi-column cards must span cleanly across defined table column boundaries (e.g. `A:B`, `C:D`, `E:F` or `A:C`, `D:F`), never cramming a 25-character KPI title or 7-digit currency into an unmerged narrow column (like width 6).
- **Proportional Row Heights**:
  - Title Rows (Row 1): 34–38pt
  - Navigation Rows (Row 2): 22–24pt
  - KPI Label Rows: 20–22pt
  - KPI Value Rows: 30–34pt
  - Table Header Rows: 26–28pt
  - Standard Data Rows: 22–26pt (single-line) or 34–52pt (multi-line wrapped quotes/tips)
  - Section Banners / Spacers: 12–16pt (spacers) / 24–28pt (section headers)

## 3. Cohesive Executive Color Palette (60-30-10 Rule)
- **60% Neutral Dominant**:
  - Dark Navy Header / Text: `#0F172A` (Slate 900)
  - Card & Row Zebra Backgrounds: `#F8FAFC` (Slate 50)
  - Clean Border Lines: `#CBD5E1` (Slate 300) / `#E2E8F0` (Slate 200)
- **30% Structural Secondary**:
  - Table Header Fills: `#334155` (Slate 700) with `#FFFFFF` text
  - Sub-headers / Accents: `#1E293B` (Slate 800)
  - Soft Indigo Highlight: `#EEF2FF` (Indigo 50) with `#2563EB` text
- **10% Semantic Functional Accents (Strict Accessibility Contrast)**:
  - Critical / Danger / Loss: Soft Red `#FEF2F2` fill, Deep Crimson `#DC2626` bold text, `#FECACA` border
  - Warning / Risk / Attention: Soft Amber `#FFFBEB` fill, Warm Amber `#B45309` bold text, `#FDE68A` border
  - Success / Growth / Won: Soft Emerald `#ECFDF5` fill, Forest Green `#16A34A` bold text, `#A7F3D0` border

## 4. Typography Scale & Hierarchy
- **Font Family**: Unified `Segoe UI` across all sheets, reports, and UI components.
- **Hierarchy Scale**:
  - Main Title: 14pt Bold
  - KPI Large Value: 14–16pt Bold
  - Section Headers: 11pt Bold
  - Table Column Headers: 9–10pt Bold (White on `#334155`)
  - Body Text / Numbers: 9–10pt Regular / Bold
  - Subtitles, Notes, Disclaimers: 8pt Italic (`#64748B`)

## 5. Web UI & Landing Page Harmony
- **Desktop & Mobile Responsiveness**: All KPI grids and tables must wrap gracefully (`grid-cols-1 sm:grid-cols-2 lg:grid-cols-4`).
- **Button Sizing & Touch Targets**: Minimum 44px height, padding `px-6 py-3.5`, font-semibold.
- **Zero Horizontal Scroll Spills**: Ensure `overflow-x-auto` on wide data tables with rounded container cards and custom scrollbars.
