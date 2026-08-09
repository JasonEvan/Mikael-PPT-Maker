# Addendum: How PPTX placeholder mapping must actually work

This clarifies a point that must not be misread from PROJECT.md: the
PPTX template being "fixed" (one file, reused weekly, edited in place)
does **not** mean the app should hardcode slide numbers or blindly
assume structure without verification. The mapping table in
PROJECT.md's "MVP Mapping Scope" section is a **reference of what to
expect**, not literal array indices to trust unconditionally.

## What the app must actually do, every run

1. **Always scan the uploaded PPTX file itself** — unzip it, walk
   `ppt/slides/slideN.xml` in slide order, find every occurrence of the
   `( teks )` token (after merging runs). Never assume slide 35 is
   Bacaan I without having just read slide 35's content and confirmed
   it contains the expected surrounding text (e.g. "BACAAN PERTAMA").
2. **Match placeholders to liturgy sections by sequence + context, not
   by raw slide index.** Use the fixed heading text already present on
   each slide (e.g. "BACAAN PERTAMA", "DOA UMAT") as the anchor for
   matching to the right parsed Word section — the slide's own visible
   label is the source of truth, not its position in the file.
3. **Validate before writing anything.** Before injecting content,
   compare what was found in the PPTX (count and identity of
   placeholders) against what's expected from the mapping table. If a
   slide's expected label/placeholder isn't found where expected —
   because the user reordered slides, deleted one, or the template
   changed in some other way — **surface this to the user as a warning
   in the review step**, don't silently write to the wrong slide or
   skip it without telling them.
4. **This constraint exists because the template is user-edited between
   runs** (the person edits it directly in PowerPoint from week to
   week). Even though its overall structure is expected to stay the
   same, the app must not treat "we've seen this template before" as
   "we don't need to re-verify this run."

## Why this matters despite the template being fixed

The user confirmed: it's one template file, reused weekly, with only
the placeholder content reset/edited between runs — the structure
itself isn't expected to change. But "expected to stay the same" is a
weekly human editing habit, not a guarantee the code can assume. The
app's job is to verify that expectation holds *this run*, and clearly
flag it in the preview step if it doesn't — not to hardcode trust in
slide numbers.

## Concretely, this means

- The slide-number table in PROJECT.md should be implemented as a
  **lookup/reference for what section label to expect at each position
  in the sequence**, used to validate and to know which slides to skip
  during duplication logic — not as literal `slides[35] = ...`
  assignments made without reading slide 35 first.
- The parser output for the PPTX side should look like: "found 15
  `( teks )` placeholders across N slides, matched 15/15 to known
  sections" (or "matched 13/15, 2 unmatched — see below") — this result
  is what drives the interactive review step, so mismatches are visible
  to the user before export, not hidden inside code that assumed
  success.
