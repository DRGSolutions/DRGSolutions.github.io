# Katapult JSON → Workbook Generator (Static Website)

## What this is
A static, client-only website that:
- lets a user upload a Katapult Job JSON export
- generates an Excel workbook with:
  - a cover sheet: "Make Ready App Info"
  - a "GPS Points" sheet
  - one sheet per pole in a make-ready style layout
- downloads the workbook in-browser (no server upload)

## Optional pole selection
After you load the JSON, the app shows a map of all poles. You can:
- Draw polygons/rectangles to select a subset of poles
- Or click "Select all poles" to export the full map

Only the selected poles will be included in the workbook (including their midspan information to unselected poles).

## Hosting (no localhost needed for end users)
Because this is a static site, you can deploy it to any HTTPS static host:
- Netlify (drag-and-drop the folder)
- GitHub Pages
- Cloudflare Pages
- AWS S3 static hosting
- Any standard web hosting provider

## Files
- index.html
- styles.css
- app.js
- worker.js

## Notes on fidelity vs your Python pipeline
Your Python pipeline is driven by an HTML Make-Ready report and performs several transforms (plus map-image insertion).
This web version is JSON-driven and reproduces the *functional intent*:
- Existing vs Proposed heights are computed from measured/manual heights + Katapult move deltas
- Midspan clearance values are pulled from midpoint-section photos and pivoted by direction
- Make-Ready Notes are generated using the same “moved up/down … (from … to …)” rule

If you want 1:1 output parity with your exact workbook formatting rules,
we can extend the workbook builder further (additional merged columns, special replacements, and optional map image).


## Tab coloring
- The tool tries to color each pole sheet tab based on the Katapult map style (job.map_styles.default.nodes).
- If tabs come out black or incorrect, set **Tab color attribute** in the UI to the attribute your map style uses to drive node colors (e.g., `MR_level`).
  - Matching is case-insensitive.


## Multiple midspan measurements
If a span has multiple midspan measurement nodes (multiple measurements between the same two poles/reference), the workbook will add additional **Existing Midspan** / **Proposed Midspan** column pairs for that direction.
