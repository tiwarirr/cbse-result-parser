# AGENTS.md

## Project Overview
- This repository contains a browser-only CBSE result dashboard.
- There is no backend, build tool, or package manager.
- The app is served as static files and runs entirely in the browser.

## Architecture
- The project follows a static three-layer structure:
  - `cbse_dashboard.html`: page shell, core layout containers, and script/style includes
  - `styles.css`: all design tokens, layout rules, responsive behavior, and component styling
  - `app.js`: application logic
- `app.js` is organized conceptually into these subsystems:
  - subject/constants data
  - parser and metadata extraction
  - in-memory state
  - dashboard rendering
  - local persistence and backup restore
  - Excel export
  - UI event wiring
- The app is stateful in-browser only. There is no server persistence.

## Runtime State Model
- Raw uploaded text is stored in:
  - `raw.X`
  - `raw.XII`
- Parsed student data is stored in:
  - `DB.X`
  - `DB.XII`
- Parse diagnostics are stored per class and are used to render warning banners.
- Chart.js instances are tracked in a shared chart registry so old charts can be destroyed before redraw.
- Active view state is controlled by:
  - current class
  - current section/tab
  - current filters/sort settings for student and merit views

## File Layout
- `cbse_dashboard.html`: HTML shell and script/style includes.
- `styles.css`: all visual styling.
- `app.js`: parsing, state, rendering, persistence, and export logic.

## User Workflow
1. User opens the static dashboard page in a browser.
2. User uploads one or both gazette TXT files for Class X and/or Class XII.
3. The app detects class from the file header when possible.
4. On analyse:
   - raw text is parsed into student records
   - diagnostics are collected
   - warnings are shown for recoverable parse issues
   - dashboard views are rendered
   - current session is stored in `localStorage`
5. User explores:
   - summary
   - subject analytics
   - merit list
   - gender analysis
   - all students table
6. User can:
   - export workbook data to Excel
   - export filtered current student view
   - back up saved sessions to JSON
   - restore a backup JSON
   - restore a previously saved browser session

## Run Instructions
- Preferred local run command:
  - `python -m http.server 8000 --bind 127.0.0.1`
- Open:
  - `http://127.0.0.1:8000/cbse_dashboard.html`

## Implementation Rules
- Keep the app fully client-side unless the user explicitly asks for a backend.
- Preserve the current three-file structure.
- Do not introduce a bundler/framework unless explicitly requested.
- Prefer minimal, behavior-preserving changes over rewrites.
- Keep dependencies CDN-based unless the user asks to vendor them locally.

## Core Algorithms

## Parsing Algorithm
- Class detection reads the file header and looks for exam-specific phrases.
- Student parsing is line-oriented:
  - the main student row contains roll number, gender, name, subject codes, and result
  - the following line contains subject mark/grade tokens
- Subject parsing pairs subject codes from the first line with mark-grade tokens from the second line.
- If the student is fully absent and no marks are present:
  - all listed subjects are emitted as `grade: 'AB'` and `marks: 0`
- If a subject token is `AB E`:
  - it is treated as a real subject entry, not missing data
  - marks are stored as `0`
  - grade is stored as `AB`
- If token count and code count do not match:
  - the parser keeps the safely matched subset
  - a warning is recorded

## Metadata Extraction
- Exam year is extracted from the exam title line.
- Region is extracted from the `REGION` label.
- School code and school name are extracted from the `SCHOOL` line.
- These values are reused in header display, local storage keys, backups, and export filenames.

## Merit Algorithm
- Merit is computed only from `PASS` and `COMP` students in the current class view.
- Supported score modes:
  - all subjects
  - best N subjects
  - best N including English
- For `best N including English`:
  - English code is auto-detected from known class-specific English subject codes
  - English is forced into the set if found
  - remaining top subjects are added by marks
- Sorting order:
  - score percentage descending
  - English marks descending as tie-breaker
  - highest single-subject mark as final tie-breaker
- Rank assignment uses dense ranking:
  - equal scores share rank
  - next distinct score gets the next consecutive rank

## Summary / Analytics Logic
- Summary cards use parsed class records to compute:
  - total students
  - pass, compartment, absent, fail counts
  - pass percentage
- Subject analysis aggregates by subject code:
  - appeared count
  - pass count
  - fail count
  - average/highest/lowest marks
  - grade distribution
- Gender analysis splits the same aggregates by `M` and `F`.
- Student table logic supports:
  - result filter
  - gender filter
  - free-text search
  - subject-wise and total sorting

## Persistence Workflow
- After successful analysis, raw class text is saved to `localStorage` using keys:
  - `{schoolCode}-{year}-X`
  - `{schoolCode}-{year}-XII`
- On later page loads:
  - saved sessions are collected from `localStorage`
  - if only one session exists, it is restored automatically
  - if multiple sessions exist, the user must choose which one to restore
- Backup export writes all saved sessions to one JSON file.
- Backup import reads the JSON file, lets the user choose a session, writes that session back to `localStorage`, and restores it into the app.

## Parsing Rules
- Student records must use this shared shape:
  - `rollNo`, `gender`, `name`, `result`, `cls`, `compSub`, `subjects`
- Subject entries must use:
  - `code`, `name`, `marks`, `grade`
- `AB` subject entries are valid and must not be treated as missing data.
- Parse warnings should be used for recoverable issues, not silent data loss.
- If a row is partially parseable, prefer preserving valid parsed data and surfacing a warning.

## Merit Rules
- Merit ranking must use dense ranking:
  - `1, 2, 3, 4, 4, 5`
- Do not switch to competition ranking unless explicitly requested.

## Persistence Rules
- Keep backup JSON backward-compatible with the current format:
  - `version`
  - `exportedAt`
  - `generator`
  - `sessions`
- If multiple saved sessions exist in `localStorage`, restore must be user-selected, not implicit.

## Editing Guidance
- Make focused changes in `app.js` and avoid broad rewrites of analytics/export logic unless necessary.
- Keep UI text clear and plain ASCII where practical.
- When fixing parser issues, verify downstream calculations still behave correctly:
  - summary
  - subject analysis
  - merit
  - gender analysis
  - all students table
  - Excel export
- Prefer changing parsing helpers before changing renderer logic if the root cause is incorrect data shape.
- Do not silently drop subject entries if they can be represented as `AB`.

## Common Change Workflow
1. Identify whether the issue is in:
   - raw parsing
   - derived calculations
   - rendering
   - persistence/restore
   - export formatting
2. If it is parsing-related:
   - inspect raw gazette row format first
   - confirm expected student/subject shape
   - update parser token logic
   - preserve warnings for ambiguous cases
3. If it is analytics-related:
   - verify parsed student objects first
   - then adjust aggregate logic
4. If it is UI-only:
   - prefer changes in `styles.css` or small view-specific render sections
5. Run syntax validation and then browser verification.

## Verification
- Minimum validation after JavaScript edits:
  - `node --check app.js`
- When changing runtime behavior, also manually verify in browser if possible.
- High-value manual checks:
  - one-class upload
  - two-class upload
  - absent-subject case like `AB E`
  - merit ties
  - saved-session chooser
  - backup export/import
