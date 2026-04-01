# Form Studio

Form Studio includes a Canva-style template builder and a linked filler page.

## What is included

- Startup flow asks for template name and page format first
- Standard paper formats: A5, A4, A3, Letter, Legal, Tabloid, Executive
- Orientation switching: portrait and landscape
- Full-browser dynamic layout with expanded side panels
- **Collapsible/Expandable side panels**: Click − / ◄ ► buttons to toggle visibility
- **Drag-to-resize panels**: Hover and drag the colored edge to resize palette or properties
- Drag-and-drop field placement on a sheet canvas
- Smart snap and alignment guide lines while dragging
- Property panel to edit labels, required flag, size, and position
- Auto-save draft and auto-restore on reload
- Duplicate selected object and clear sheet actions
- Undo/Redo history for layout edits
- Layers panel with bring-front/send-back and lock/unlock
- Keyboard shortcuts for fast layout edits
- Export JSON including page metadata and placed object coordinates
- Download template JSON file
- Open linked filler page with current template auto-loaded

## Filler flow

- Open `filler.html` directly, or click `Go To Filler` in builder
- Filler loads latest saved template from local storage
- Optional: upload a template JSON file manually
- **Collapsible Info sidebar**: Click − / ► button to toggle template info display
- **Drag-to-resize sidebar**: Hover and drag the colored edge to resize the info panel
- Fill data on the same sheet alignment as builder
- Download filled form in multiple formats: JSON, PDF, DOCX, PNG, JPG, TXT
- Export quality selector: Standard, High, or Print (300 DPI)
- Export outputs include constants and only the fields the user filled

## Supported form objects

**Layout & Structure:**

- Section Header (titled section)
- Divider Line (visual separator)
- Constant/Label (static text block)

**Text Input:**

- Short Text (single-line text)
- Email (email field)
- Phone (phone number field)
- URL (web address)
- Long Text / Paragraph (multi-line text)

**Numbers & Currency:**

- Number (numeric input)
- Currency (money amount)
- Percentage (percentage value)

**Selection Fields:**

- Dropdown / Select (single selection from options)
- Radio Buttons (exclusive choice)
- Checkboxes (single or multi-select)
- Multi-Select (multiple selection from list)

**Date & Time:**

- Date Picker (date only)
- Time Picker (time only)
- Date & Time (combined date and time)

**File & Media:**

- File Upload (any file type)
- Photo Upload (images only)

**Special Fields:**

- Signature (signature drawing field)
- Rating (star-based rating, 1-10)

## Keyboard shortcuts (Builder)

- `Delete` or `Backspace`: remove selected object
- `Ctrl + D` / `Cmd + D`: duplicate selected object
- `Ctrl + Z` / `Cmd + Z`: undo
- `Ctrl + Y` or `Cmd + Shift + Z`: redo
- `Arrow keys`: nudge selected object by 1px
- `Shift + Arrow keys`: nudge selected object by 10px

## Builder Features

- **3-column split-pane layout** with draggable dividers
  - Left: Field Palette with all form element types
  - Center: Template Sheet canvas with Canva-style drag/drop
  - Right: Properties panel with Layers
- **Smart alignment** with guide lines during object movement
- **Resize handles** on 8 points (corners + edges) for each selected object
- **Layers panel** with layer ordering (Bring Front, Send Back) and lock/unlock toggles
- **Full undo/redo history** for all layout changes
- **Corner drag-to-resize** handles for precise object sizing
- **Auto-save draft** and restore on page reload
- **Copy/duplicate objects** with Ctrl+D or Cmd+D
- **Clear sheet** to start template over
- **Export JSON** with complete template structure
- **Download template** as JSON file for sharing

## Filler Features

- **2-column split-pane layout** with draggable info sidebar
  - Left: Template info (collapsible)
  - Right: Fill form sheet
- **Field-specific inputs**
  - Text inputs for email, phone, URL with native browser validation
  - Number fields with currency and percentage support
  - Dropdowns, radio buttons, checkboxes, and multi-select lists
  - Date and time pickers
  - File and photo upload fields
  - Signature canvas (ready for drawing)
  - Star rating controls
- **Smart export** (only filled values + constants)
- **Multiple export formats**: JSON, PDF, DOCX, PNG, JPG, TXT
- **Export quality selector**: Standard, High, Print (300 DPI)
- **Auto-load template** from builder via local storage
- **Manual upload** of template JSON files

## Form creation essentials to add next

- Header/footer areas and reusable sections
- Validation rules (pattern, min/max, date range)
- Conditional logic between fields
- File limits (size, count, type)
- Signature block
- Save draft and publish versioning
- Filler runtime and submission API

## Run

Open [index.html](index.html) in your browser.

For fill mode, open [filler.html](filler.html).

## Production Readiness Check

- No syntax errors in workspace files (validated in editor)
- Builder and filler flows linked and working
- Template and filled-form exports tested (JSON, PDF, DOCX, PNG, JPG, TXT)
- Responsive layout tuned for desktop/laptop usage
- Split-pane state persistence enabled
- No temporary debug logs in source files
- Installable web app (PWA) enabled with manifest + service worker

## Web App Features (PWA)

- Install prompt support in compatible browsers
- Standalone app display mode
- App icon and theme color configured
- Offline-ready shell caching for key app files
- Runtime caching for same-origin files and CDN libraries used by filler exports

## GitHub Upload Preparation

Before pushing to GitHub, do this quick checklist:

1. Open builder and filler once and verify your latest UI changes.
2. Hard refresh browser (`Ctrl+F5`) to clear cached CSS/JS.
3. Confirm exports from filler (JSON + one binary format like PDF).
4. Commit with a clean message such as `feat: production-ready Form Studio builder and filler`.

## Optional GitHub Pages Deploy

This project is static (HTML/CSS/JS), so GitHub Pages works directly.

1. Push repository to GitHub.
2. Go to repository `Settings` -> `Pages`.
3. Set Source to `Deploy from a branch`.
4. Choose branch `main` and folder `/ (root)`.
5. Save and wait for the Pages URL.

Use the deployed `index.html` as entry point.

## Notes for Existing GitHub Pages URL

If you already published at:

- [https://vignesh-s-github.github.io/Form-Studio/](https://vignesh-s-github.github.io/Form-Studio/)

then PWA install and service worker work under that same path.

After deploying updates, open the app once and hard refresh (`Ctrl+F5`) to activate the new service worker/cache version.
