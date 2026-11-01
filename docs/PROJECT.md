# Project Context: Liturgy Slide Automation App

## What this app does

This is a Next.js web app that automates the creation of weekly Catholic
Mass liturgy PowerPoint slides for a church. Every week, the church
office prepares a Word document (.docx) containing the full liturgy text
(prayers, readings, responses) for that Sunday's Mass. That text then
needs to be manually copy-pasted into a fixed 117-slide PowerPoint
template that follows the Mass structure. This app automates that
copy-paste step: parse the Word doc, extract the relevant text per
section, and inject it into the matching placeholders in a duplicate of
the PPTX template — producing a ready-to-use slide deck in seconds
instead of the ~30-60 minutes it currently takes to do by hand.

This is a single-purpose internal tool for one specific church's
workflow, not a generic document-to-slide converter. The Word document
structure and the PPTX template structure are both fixed and known in
advance (see below). Do not over-engineer this into a configurable /
general-purpose mapping system — the mapping rules are intentionally
hardcoded to this one template.

## Source document: the Word file

The weekly Word doc (.docx) contains the full liturgy in a fixed,
predictable structure:

- **Major sections** are marked with **bold + underline** text (e.g.
  "RITUS PEMBUKA", "LITURGI SABDA", "LITURGI EKARISTI", "RITUS PENUTUP").
- **Sub-sections** (e.g. "Doa Kolekta", "Bacaan I", "Bacaan II",
  "Mazmur Tanggapan", "Injil", "Doa Umat", "Doa Atas Persembahan", "Doa
  Sesudah Komuni") are marked with **bold only**, usually followed by an
  italic annotation like _(umat berdiri)_ ("congregation stands") or
  _(umat duduk)_ ("congregation sits").
- Some sub-sections have **no body text** below them (e.g. "Tanda
  Salib", "Kudus", "Homili") — these are sung or led live and should be
  skipped entirely; they never produce a slide.
- Readings (Bacaan I, Bacaan II, Injil) follow this exact sub-pattern:
  1. An **italicized quoted line** in quotes — a short thematic epigraph
     for the reading (e.g. `*"Berdirilah di atas gunung itu di hadapan
Tuhan."*`)
  2. A **bold source line** naming the book/chapter/verse, prefixed with
     the speaker label `L` (Lektor), e.g. `L **Bacaan dari Kitab Pertama
Raja-raja 19:9a.11-13a**`
  3. The **full body text** of the reading (this is long — goes to
     speaker notes, not the visible slide, see mapping below)
  4. A closing formula ("Demikianlah sabda Tuhan" / response)
- **Mazmur Tanggapan** (Responsorial Psalm) has a fixed sung melody
  (already baked into the PPTX template as text using a custom notation
  font) plus 2-4 **numbered verses** (bait) of lyrics that change weekly.
  Verse count varies week to week.
- **Doa Umat** (Prayers of the Faithful) has a fixed intro line, then a
  **variable number of intentions** (typically 3-5), each formatted as
  `L **[bold intention title].** [prayer body text] Kami mohon...`
  followed by a fixed congregational response line, then a fixed closing
  line. The number of intentions varies week to week — this is the key
  case where the number of output slides must scale with the number of
  items found, not be fixed.
- Occasional **one-off prayers** appear in some weeks only (e.g. a
  building fund prayer, a special novena prayer) and are absent in
  others — these map to a placeholder that should simply be left blank
  when not found in a given week's document, not deleted or skipped.
- Song lyrics (Nyanyian Pembuka, Persembahan, Komuni, Penutup) are
  **NOT included** in this Word document — they live in a separate file
  and are explicitly **out of scope for the MVP**. Do not attempt to
  parse or map them.
- Announcements (Pengumuman Gereja) are also **out of scope for the
  MVP** — always skipped.

## Target document: the PPTX template

`Template_Mingguan.pptx` currently contains **112 slides**, with the full
Mass running order, fonts, layout, and styling. It is not a small set of
reusable master slides: the generator edits selected placeholders and adds
continuations to the existing deck. Matching never depends on these ordinals.

- Editable placeholders are literally the text `( teks )` typed inside a
  text box on the slide, in the position/style where the real content
  should go. Some slides have one `( teks )` placeholder, some have two
  (e.g. reading slides have one for the epigraph quote and one for the
  source citation).
- Some content is meant to go into **PowerPoint speaker notes**, not the
  visible slide body — specifically the full body text of readings
  (Bacaan I, II, Injil). The visible slide only shows the short epigraph
  quote and the source citation; the notes pane holds the full text for
  the lector/reader to reference. Existing notes slides already contain
  a `(jangan diblank)` ("don't leave blank") reminder — this convention
  predates this app and must be preserved.
- Musical notation (Mazmur melody, Bait Pengantar Injil, Anamnesis) is
  **already present as styled text** in the template using a custom
  notation encoding (characters like `j`, `k`, `.`, `!`, `#` render as
  pitch/octave marks via a special font). This notation is fixed and
  never changes week to week — only the lyric placeholders below/beside
  it change.
- When a source section is genuinely not present in that week's Word
  doc (e.g. "Seruan Tobat" is sometimes replaced by a different penance
  rite), the matching placeholder is simply **left blank** — the slide
  stays in the deck, it's just empty. It is not deleted.
- When a source section has a **variable number of repeating items**
  (Doa Umat intentions, Mazmur verses), the app must **duplicate** the
  corresponding template slide once per item found, preserving the
  slide's layout/formatting, and insert each item into its own copy.

## Known placeholder-to-slide mapping (MVP scope)

| Content                                             | Slide #            | Mode                                       |
| --------------------------------------------------- | ------------------ | ------------------------------------------ |
| Seruan Tobat (3 lines)                              | 22, 23, 24         | Measured invocation and fixed responses    |
| Doa Kolekta                                         | 33                 | Measured prayer, final response            |
| Bacaan Pertama — epigraph quote                     | 35 (placeholder 1) | Shared measured metadata budget            |
| Bacaan Pertama — source citation                    | 35 (placeholder 2) | Shared measured metadata budget            |
| Bacaan Pertama — full body                          | notes on slide 35  | Simple fill (notes)                        |
| Mazmur Tanggapan verses                             | 39                 | Separate measured page sequence per verse  |
| Bacaan Kedua — epigraph / citation / body           | 41 (+ notes)       | Same pattern as Bacaan Pertama             |
| Bacaan Injil — Gospel-name introduction             | 45                 | Measured dialogue/source introduction      |
| Bacaan Injil — epigraph / citation / body           | 46 (+ notes)       | Metadata visible; body only in notes       |
| Doa Umat — intro line                               | 63                 | Measured pagination                        |
| Doa Umat — each intention (title + body + response) | 64                 | Separate measured sequence per intention   |
| Doa Umat — closing line                             | 66                 | Fixed response only after complete closing |
| Doa Atas Persembahan                                | 72                 | Measured pagination, final response        |
| Doa Sesudah Komuni                                  | 92                 | Measured pagination                        |

Out of scope: song lyrics, fixed musical notation, and church announcements.
These parts are not globally filled or restyled. See `docs/PPTX_MAPPING.md`
for anchor-based weekly and OBS rules.

## Tech stack

- **Next.js** (App Router), TypeScript
- Word parsing: unzip the .docx (it's a ZIP of XML) and read
  `word/document.xml` directly — parse run-level `bold`/`underline`/
  `italic` formatting to detect structure, rather than relying on an
  LLM to interpret the document. The structure is 100% deterministic
  from the XML, so no AI/LLM call is needed or wanted for this step.
- PPTX generation: unzip the template .pptx, locate `( teks )` tokens
  inside `ppt/slides/slideN.xml`, replace them with parsed content,
  duplicate slide XML + relationships when an item needs to repeat,
  write speaker notes into the corresponding `ppt/notesSlides/` file.
  Must handle PowerPoint's tendency to split a single visual string
  across multiple `<a:r>` runs — merge runs within a paragraph before
  searching for the `( teks )` token, don't assume it's a single intact
  run.
- No OCR, no vision LLM calls anywhere in this pipeline — the source is
  already structured text (Word), and the notation graphics are
  explicitly out of scope.

## UX principle

The user (a single church volunteer who prepares slides weekly) should
be able to: upload this week's Word doc → see a preview of what was
detected and where it will go → fix anything mis-detected → export the
final .pptx. Parsing should never silently guess wrong; ambiguous or
missing sections should be visibly flagged in the preview, not hidden.

## English Mass: homily document

- Select **English Mass** on the first screen and upload its liturgy document.
- After the liturgy is parsed, the results/preview screen shows the separate
  **Teks homili (opsional)** uploader. It is not shown on the first screen
  or for Indonesian Mass. It accepts `.doc` and `.docx`; all readable body
  paragraphs are included, without filtering headings or speaker labels.
- The English Mass preview displays the homily's file name and full text.
  Uploading another homily replaces it; **Hapus homili** removes it.
  **Replace file** resets both the liturgy preview and the homily.
- PowerPoint generation follows the template's presentation order, finds
  `HOMILY`, and fills the following `(teks)` placeholder before the creed.
  It never relies on a slide number. Long homilies use continuation copies
  of that placeholder slide, preserving its layout and relationships.
- Without a homily upload, its text placeholder is cleared. If a homily was
  supplied but the template has no matching placeholder after `HOMILY`,
  generation reports an error rather than silently omitting the text.

## English Mass: measured text layout

- All English dynamic text destinations use the shared `pptxTextLayout.ts`:
  song titles, reading introductions/bodies, psalm verses/refrains, acclamation,
  prayers, intercessions, Kyrie, and homily. There are no character-budget
  pagination limits in this English export path. Indonesian weekly now also
  uses measured layout; OBS Hijau/Kuning deliberately retain legacy pagination.
- Layout is measured with Skia Canvas and the exact template fonts. The engine
  resolves text-box extents/insets and inherited paragraph/run properties,
  budgets all paragraphs sharing a shape together, and fills pages with measured
  lines. Source paragraphs and hard breaks become DrawingML paragraphs/breaks.
  Ordinary horizontal whitespace is normalized; nonbreaking spaces are retained.
- Body font size is preserved. The exporter neither clips overflowing content
  nor relies on viewer AutoFit. First/Second Reading introductions, Gospel
  Acclamation (Alleluia), and Psalm slides first retain their template's empty
  spacer paragraphs. Only if the complete shape would overflow are spacers
  removed one at a time, reducing double/triple gaps before isolated gaps and
  stopping as soon as the content fits. Source paragraph and hard line breaks
  are never removed by this fallback. Other edited roles keep their existing
  compact layout. Reading metadata can still continue before the scripture body;
  prayer/Kyrie responses stay after the continued text. Song lyric placeholders
  remain intentionally unfilled.
- Unsupported rotated, grouped, vertical, or multicolumn dynamic text shapes,
  missing fonts, and indivisible content that cannot fit produce actionable
  errors instead of hiding or dropping text.

### Required licensed fonts

The server bundles the exact template font files in `app/fonts/pptx`, shared by
English Mass and Indonesian weekly generation. These are server-only assets,
not files in `public/`, and are not embedded in generated presentations.
Presentation computers still need the same fonts. Bundling does not grant new
font rights: deployment and redistribution must comply with the font licenses.

- The bundle contains Arial and Arial Narrow regular/bold/italic/bold italic,
  Book Antiqua regular/bold/italic/bold italic, Arial Rounded MT Bold,
  Lucida Sans demibold italic, and Wingdings.
- By default, only `app/fonts/pptx` is searched. Generation does not depend on a
  developer's Office installation, macOS system fonts, or the user's home folder.
- Optional `ENGLISH_MASS_FONT_DIR` and `INDONESIAN_MASS_FONT_DIR` overrides remain
  exclusive. Remove old deployment values pointing at local Mac paths to use
  bundled assets. No font-directory environment variable is needed by default.
- Internal family, weight, and italic metadata are verified; renaming a
  substitute font is not supported.
- Next externalizes `skia-canvas` and traces the PPTX templates plus
  `app/fonts/pptx/*.ttf`. Vercel builds must install native dependencies for their
  target OS/architecture, including optional dependencies. Do not deploy a
  `node_modules` directory copied from macOS to Linux.

### Streaming downloads on Vercel

`/api/generate-pptx` explicitly uses the Node.js runtime and returns PPTX files
through a backpressure-driven `ReadableStream` for all templates, including OBS
and English. The response emits at most 64KiB per pull using buffer views,
without a second full-file copy or a `Content-Length` header. Filenames, MIME
type, error responses, and the existing browser `response.blob()` flow remain
compatible. Streaming avoids the non-streaming response payload limit; it does
not change ZIP generation, which still finishes in memory before download starts.

Deploy the font assets with the application and verify the download on a Vercel
preview. A local production build proves packaging and HTTP streaming locally,
not execution inside Vercel.

### Verification

Run `npm run test:english` using the bundled fonts or an explicit font-directory
override. The layout suite covers proportional font widths, shared-shape
capacity, inherited styles, paragraphs/breaks, metadata continuations,
Unicode/nonbreaking groups, changing continuation widths, and final-response
placement. Generator tests exercise every English dynamic role, long Gospel
citations, text conservation/order, and absent homily behavior. Missing required
fonts fail the suite rather than silently skipping font-dependent cases.

Model bounds are not a claim of identical rendering across Office applications.
Validate representative generated decks in Microsoft PowerPoint after template
or typography changes, including First Reading introductions/bodies and homily
first/middle/final pages.

## Indonesian weekly: measured layout

- `generatePPTXBuffer` selects measured Indonesian layout in the existing
  generator. The public template entry point used by OBS retains the legacy
  mode; the API, parser, UI, and OBS wrappers are unchanged.
- Every existing weekly dynamic role is measured, including previously
  unbounded Seruan Tobat, reading metadata, Gospel-name dialogue, and psalm
  verses. Legacy character budgets and combined-response thresholds do not
  participate in weekly layout decisions.
- Complete Bacaan I/II/Injil bodies remain only in the original metadata
  slide's speaker notes, resolved through its notes relationship. Metadata
  continuations do not share that notes part.
- Template spacers are preserved until the edited shape overflows. Heading,
  epigraph, citation, static prompts, and responses share one actual shape
  budget; source paragraphs/hard breaks are never discarded as spacers.
- Structured intentions preserve a bold title, regular body, and the original
  colored/italic response with its `U` tab label. Each verse/intention has its
  own ordered page sequence. A fitting response goes last; an oversized response
  continues after its body without truncation or font shrinking.
- Exact weekly fonts include Arial, Arial Narrow, Lucida Sans demibold italic
  (physical weight600), and Wingdings for the existing book icon. Symbol-only
  runs are measured using `a:sym`, not a Latin theme font.
- Weekly generation uses the same `app/fonts/pptx` bundle as English. An optional
  `INDONESIAN_MASS_FONT_DIR` override remains exclusive. Fonts are not embedded
  in the generated presentation; presentation computers need the same faces.
- Run `npm run test:pptx` for all generator/shared-engine regressions, or
  `npm run test:indonesian` for weekly/OBS plus shared layout. Full acceptance
  requires both English and weekly fonts, logical ZIP-part equivalence for
  unchanged OBS/English scenarios, and Microsoft PowerPoint render inspection.
