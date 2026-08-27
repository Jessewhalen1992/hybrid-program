# Hybrid Manager Version History

## 1.1.1 — 2026-08-27

- Preserve the original ID of an ownerless number bubble anchored at a Hybrid vertex or uniquely identified by its persistent ID.
- Reuse one original block and remove owned or legacy duplicate replacements created during migration.
- Stop using another polyline's matching sequence number as ownership evidence.
- Count only the first vertex when multiple vertices share the same stored X/Y coordinate.

## 1.1.0 — 2026-08-27

- Require UTM state for Hybrid selection, geometry, and data-changing workflows.
- Store coordinates to 0.001 while displaying them to 0.01.
- Prevent RC vertices from receiving marker blocks.
- Consolidate polyline and numbering workflows and clarify the UI labels.
- Preserve safe X/Y/Z number-bubble scale editing.
- Add atomic point-and-marker deletion and single-cell copy/fill.
- Prevent duplicate legacy number bubbles from being reinserted.
- Standardize AutoCAD table widths, row height, and text height.
- Add semantic version and version-date metadata for Compass update tracking.
