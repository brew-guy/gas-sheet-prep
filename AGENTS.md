# AI Agent Instructions for SheetPrep - Google Sheets Data Importer

**Project context:** This is a Google Apps Script (GAS) application to import structured data files directly into Google Sheets. The code uses `clasp` for sync, which converts standard server-side `.js` to `.gs` extensions during the push.

## Core Directives

1. **Google Apps Script Constraints:**
   - Google Sheets cells have a strict limit of 50,000 characters. Implement automatic truncation (slice) strings running over the limit when committing to the active sheet.
   - When importing large arrays into a sheet via `range.setValues(data)`, the data must be a perfectly rectangular 2D array. Otherwise, it will fail. Ensure consistent array row lengths before committing.

2. **Component Architecture & Syntax:**
   - **Frontend UI (`src/ui/ImporterModal.html`):** Do not write `<script>` tags externally in `src/ui/js/` to be compiled natively like standard webapps. GAS requires HTML partials compiled inside scriptlet syntax (`<?!= include(...) ?>`).
   - Thus, UI scripts like `ModalLogic.html` must remain defined inside a root `<html>` tag wrapping a `<script>` tag. The same rule applies to UI styles in `src/ui/styles/` wrapped inside a `<style>` tag.
   - **McPherson Proxy Pattern (Library Support):** Do NOT call backend functions directly via `google.script.run.[methodName]()`. Always use the universal router instead: `google.script.run.runRouter('methodName', args)`. This allows the codebase to be deployed as a pluggable GAS Library. Always add `.withSuccessHandler()` and `.withFailureHandler()` since it's asynchronous.

3. **Feature Specifics - Required Behaviors:**
   - **Header Mismatches:** Warn users strictly with a Confirmation overlay dialog whenever the imported file headers differ from those of the existing selected destination sheet. Wait for explicit user confirmation before overwriting or appending.
   - **Data Previews:** Only parse slices (e.g., first 50 rows) for the frontend DOM table during the Preview stage, while keeping the full CSV/XLS data in memory (`window.appState.fullData`). Rendering tens of thousands of rows at once crashes the side-panel/dialog.
   - **Saved Flows Storage:** The importer remembers settings per "Output Sheet." The configuration parameters (skip top/bottom/left/right, replacement rules json, updated timestamps) are logged in a hidden `_SavedFlows` sheet. Update and retrieve these details seamlessly so it never breaks the UI.
   - **Reserved Sheets:** Sheets prefixed with `_` like `_CleaningRules` and `_SavedFlows` are reserved configurations. Prevent users from selecting or pushing data into these arrays.

## Clasp & Deployments

- Never manually invoke `npm run build` or `npm run dev` as there's no node execution layer required to deploy GAS code. Only execute `clasp push` to sync code arrays directly to the linked script ID.
- Do not edit the `.clasp.json` structure or modify `appsscript.json` indiscriminately, as doing so might detach the script extension associations.

## Formatting Best Practices

- **Variables & DOM IDs:** Use JS descriptive camelCase naming (`fileInput`, `targetSheetSelect`, etc.). Stick to clear semantic IDs (`#helpBtn`, `#previewTbody`).
- **Icons & Visuals:** When updating the extension frontend visually, prefer SVG code definitions or lightweight Vanilla CSS. Avoid complex dependencies like Bootstrap. The design emphasizes a clean, glassmorphic or modern minimalist appearance using custom root CSS variable tokens.
- Avoid using external JS cdns besides trusted minimal parsers, in this case, PapaParse (CSV) and SheetJS (XLS(X)).
