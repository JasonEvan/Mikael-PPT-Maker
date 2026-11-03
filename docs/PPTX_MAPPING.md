# PPTX Generation & Structural Mapping Specification

The Indonesian generators consume `MappedLiturgyData` from
`app/utils/docxParser.ts`. The template remains the typography and layout
source; matching uses content anchors and presentation relationships, not
hardcoded slide indices.

## 1. Entry points and scope

| Entry point                   | Template                       | Text layout                         |
| ----------------------------- | ------------------------------ | ----------------------------------- |
| `generatePPTXBuffer`          | `Template_Mingguan.pptx`       | Measured weekly layout              |
| `generatePPTXOBSBuffer`       | `Template_PPT_OBS_Hijau.pptx`  | Measured OBS layout                 |
| `generatePPTXOBSKuningBuffer` | `Template_PPT_OBS_Kuning.pptx` | Measured OBS layout                 |
| `generatePPTXOBSUnguBuffer`   | `Template_PPT_OBS_Ungu.pptx`   | Measured OBS layout                 |

`generatePPTXBufferForTemplate` is the sole shared Indonesian implementation.
Weekly and all OBS wrappers use measured layout with the Indonesian font
profile. No character-count pagination or legacy fallback remains.
English shares the measurement engine while retaining its own generator,
content mapping, font profile, and pagination policies.

The API selects the weekly wrapper only for `template: "mingguan"`. Its existing
OBS/default routing and English early dispatch remain unchanged. DOCX parsing and
the `MappedLiturgyData` contract are unchanged.

The Indonesian template selector includes `obs-ungu` (Template OBS Ungu).
Its generator uses the uploaded Ungu template and the same OBS role profile.
The API download is `Slide_Liturgi_Misa_OBS_Ungu.pptx`; the browser retains the
source-document-based filename with the `OBS_Ungu` suffix. Omitting the template
still selects OBS Hijau. Existing Hijau/Kuning, weekly, and English IDs are unchanged.

## 2. Shared Indonesian measured layout

`pptxTextLayout.ts` owns DrawingML measurement and pagination. The generator owns
liturgy mapping, static first/final content policy, notes, and ZIP assembly.

- Resolve the edited shape's extents, insets, paragraph margins/spacing,
  indentation, and effective run properties through layout/master/theme.
- Register the exact physical fonts and measure shaped candidate lines with
  Skia Canvas. All paragraphs sharing a shape use one height budget, including
  fixed headings, labels, tab stops, responses, and symbols.
- For vertically centered boxes, distinguish nominal line-box height from
  visible glyph ink. Account for first/last-line font leading and the centered
  baseline placement when checking ink against the physical shape edges.
  Keep the font size, line pitch, box geometry and insets unchanged. This lets
  genuinely fitting three-line lower-thirds stay together without imposing a
  three-line special case; longer content still paginates.
- Preserve intended point sizes. Do not use character budgets, clipping, or
  viewer AutoFit to make content fit.
- Preserve template spacer paragraphs first. If the complete edited shape is
  too tall, remove template blanks one at a time, reducing repeated gaps first,
  until it fits or no removable spacers remain. Source paragraph/hard breaks
  are not disposable template spacers.
- Emit actual DrawingML paragraphs and line breaks. Reflow the remaining source
  text for continuation geometry instead of treating previous soft wraps as
  source hard breaks.
- Keep each psalm verse and each intention as its own ordered page sequence.
  A new item never shares a preceding item's page merely because space remains.
- Match split-run `(teks)` / `( teks )` tokens, retaining unaffected styled
  prefixes/suffixes. Do not run legacy placeholder replacement over final
  measured slide XML.

String replacements retain the established normalization of ordinary horizontal
whitespace and CRLF. Structured paragraphs additionally retain semantic tabs,
including the congregational `U\t` label. Nonbreaking spaces and grapheme
boundaries are retained by the shared engine's wrapping rules.

Structured paragraphs accept either `text` or inline `runs` with bold/italic
overrides inherited from the original token. Rich-run whitespace is normalized
across style seams without joining words; adjacent matching styles are coalesced.
An optional paragraph `policy: "last"` overrides the replacement-level policy,
allowing a same-token response to use the existing final-page reservation and
oversized-response flow. Existing string and structured-text defaults retain
their established behavior.

## 3. Weekly section mapping

| Content               | Selection/input                                                                                   | Output contract                                                                                     |
| --------------------- | ------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| Seruan Tobat          | `KASIHANILAH KAMI`; `seruanTobat`                                                                 | Invocation text followed by the template's fixed invocation/response; all share the measured budget |
| Doa Kolekta           | `MARILAH KITA BERDOA` before Bacaan Pertama; `doaKolekta`                                         | Opening prompt only at the start; fixed response after the final body page                          |
| Bacaan Pertama        | `BACAAN PERTAMA` / `BACAAN I`; `bacaanPertama`                                                    | Formatted epigraph and cleaned citation visible; complete `notesBody` in original slide notes       |
| Mazmur Tanggapan      | Psalm anchor or first eligible placeholder between Bacaan Pertama/Kedua; `mazmurTanggapan.verses` | One measured page sequence per verse, in source order                                               |
| Bacaan Kedua          | `BACAAN KEDUA` / `BACAAN II`; `bacaanKedua`                                                       | Same metadata/notes separation as Bacaan Pertama                                                    |
| Injil introduction    | `INILAH INJIL SUCI MENURUT`; `bacaanInjil.gospelName`                                             | Preserve the template's dialogue, label, and response order                                         |
| Injil metadata        | `INJIL`, excluding `PENGANTAR`; `bacaanInjil`                                                     | Epigraph/citation visible; complete `notesBody` in original metadata slide notes                    |
| Doa Umat introduction | First one-placeholder slide after the Doa Umat title; `doaUmat.intro`                             | Measured introductory text                                                                          |
| Doa Umat intentions   | Following two-placeholder slide; `doaUmat.intentions`                                             | Bold title once, regular body, original-style response after its own body                           |
| Doa Umat closing      | Following placeholder/closing slide; `doaUmat.closing`                                            | Complete closing text, then fixed `U Amin` only at the end                                          |
| Doa Atas Persembahan  | `DOA ATAS PERSEMBAHAN`; `doaAtasPersembahan`                                                      | Heading only initially; fixed response only after the final body                                    |
| Doa Sesudah Komuni    | `SESUDAH KOMUNI` state; `doaSesudahKomuni`                                                        | Opening prompts initially; preserve any actual fixed closing after the final body                   |

`formatEpigraph` retains the existing curly-quote normalization. Citation cleaners
retain the existing speaker/prefix rules. Weekly does not adopt OBS's optional
Gospel-source-prefix omission.

### Doa Umat typography and responses

Structured replacements inherit the appropriate template token's face, size, and
fill. Only the title/body bold override is explicit: title bold, body unbold.
The response retains its separate token's italic/color properties and `U\t` label.
A fitting response is scheduled at the end, carrying a tail body line with it
when possible. A response larger than an entire page is ordered content after
the complete body and may span final pages; it is never dropped or shrunk.
There is no combined-character threshold or length-based sentence-fit heuristic.

### Missing content and excluded slides

The existing empty-data behavior is retained. Intended empty destinations remain
in the deck; zero intentions clear their tokens. Zero psalm verses leave the
original verse placeholder unchanged. Fewer Seruan items retain the existing
selection/fallback behavior rather than changing parser semantics.

Slides containing `PENGUMUMAN` or `LIRIK LAGU` are skipped. Unmatched placeholders,
song lyrics, static notation, announcements, artwork, and unrelated shapes are
not globally filled or restyled. The broad psalm fallback still depends on the
fixed template sequence; new placeholder slides in that interval need review.

## 4. Speaker notes and slide-package integrity

The complete scripture bodies remain **speaker notes only**. Do not borrow
English's visible reading-body pagination for the Indonesian templates.

`updateSlideNotes` follows each original slide's own `notesSlide` relationship
and normalizes its target path. It does not assume `notesSlide2/3/6.xml` or matching
slide/notes ordinals. Nonempty notes replace the notes body-placeholder paragraphs;
empty notes leave existing template notes unchanged. Existing missing-notes-part
behavior is unchanged; this cutover does not create new notes parts.

Every continuation uses the existing `duplicateSlide` machinery:

- Allocate unique numeric slide IDs, relationship IDs, and slide part names.
- Register the part in `[Content_Types].xml` and presentation relationships.
- Insert after the previous page's relationship to preserve item/page order.
- Copy layout/media relationships, but strip the notes relationship from clones.
- Keep complete scripture notes attached to the original metadata slide, not
  shared between multiple slides.

## 5. Exact fonts and deployment

The current weekly template requires Arial regular/bold/italic, Arial Narrow
regular/bold/italic variants, Lucida Sans bold italic, and Wingdings. Headings
include 60pt text and the symbol-only Wingdings book icon; body/epigraph/source
runs are typically 48pt, while `(hening sejenak)` uses 44pt. Do not measure all
runs as Arial48 or inherit a Latin theme face for an `a:sym` run.

Lucida Sans's physical demibold italic face has weight600 even though the
DrawingML run requests bold. The engine verifies that exact face instead of
synthesizing weight700 or replacing the family. Symbol runs retain their source
`a:sym` and glyph; measurement selects the symbol face.

All measured generators load the exact bundled `.ttf` files in `app/fonts/pptx`
by default. There is no default search of macOS Office/system/user directories.
The shared bundle includes Arial, Arial Narrow, Book Antiqua, Arial Rounded MT
Bold, Lucida Sans demibold italic, Wingdings, and Calibri Bold/Bold Italic. Optional
`INDONESIAN_MASS_FONT_DIR` and `ENGLISH_MASS_FONT_DIR` overrides are exclusive:
remove obsolete Mac-specific values from Vercel to use the bundle.

These are server assets, not public web fonts or fonts embedded in the exported
PPTX. Presentation computers still need matching fonts. Deployment and
redistribution must comply with the proprietary font licenses.
Next externalizes `skia-canvas` and traces `app/fonts/pptx/*.ttf` and templates.
Install native and optional dependencies for the deployment OS/architecture.
OBS now requires the exact Indonesian-profile faces, including Calibri inherited
by lower-third text from the template's theme. Missing fonts fail explicitly;
OBS does not substitute Arial for Calibri or bypass measurement.

The download route uses the Node.js runtime and returns a `ReadableStream`
for every template. It emits buffer views in chunks of at most 64KiB per pull,
honoring backpressure without a second full-file copy. Responses omit
`Content-Length` and preserve attachment filenames/MIME type and JSON errors.
The browser's existing `response.blob()` download flow remains valid.
This is transport streaming: ZIP generation still finishes in memory first.
Verify response streaming on a Vercel preview as well as locally.

## 6. OBS-specific mapping and appearance

Hijau, Kuning, and Ungu use the same role rules, but their slide counts, positions,
font sizes, spacing and notes relationships differ. Read each original template,
not a color-swapped copy or a shared set of numerical slide indices.

- After `SERUAN TOBAT`, consume the first three eligible placeholder-only slides
  in order. Missing invocations clear their own slot without repeating the first.
  Each invocation's continuations precede its existing separate response slide.
- The Gospel introduction takes `gospelName`. Its following anonymous two-token
  metadata slide takes epigraph and source, omitting the already-spoken citation
  prefix. Complete scripture notes belong to that metadata slide only.
- Doa Umat uses three successive one-placeholder roles: introduction, intentions,
  closing. Empty content still advances the role state. Each intention uses one
  rich replacement: initial sentence/title-line emphasis, regular remainder,
  then a regular `U\t` response with paragraph-level last policy.
- Retain the OBS token's face, size, fill and italic settings. Do not impose the
  weekly separate response token's color/italic style. Inline sentence emphasis
  preserves its separating spaces. When a separate title is supplied, title and
  body are distinct justified paragraphs: the wrapped title's final line remains
  a normal left-aligned paragraph tail rather than a force-justified soft break.
- Measure lower-third boxes and full-content boxes independently from actual
  geometry, insets and paragraph spacing. Keep template gaps first, then remove
  only decorative blanks if necessary before paginating. Fixed prompts and
  same-shape final responses share this budget; separate fixed slides/shapes
  are not moved into a dynamic text box.
- Preserve backgrounds, artwork, untouched shapes, hidden slides, excluded
  placeholders and original notes ownership. No legacy replacement, character
  budgets, sentence-tail fit heuristic or measured/legacy mode switch remains.

## 7. Verification

- `npm run test:pptx`: shared layout plus English and Indonesian generator tests.
- `npm run test:english`: English generator and shared layout regressions.
- `npm run test:indonesian`: weekly/OBS generator and shared layout regressions.
- Type/lint/build checks do not prove Office rendering.

Use exact licensed fonts for full acceptance. Check source-text conservation and
order, mixed fonts/sizes/symbols, rich intention typography, final responses,
spacer-first fitting, long responses, notes-only scripture, and valid package
relationships. Compare short/long weekly and English logical ZIP-part contents
before/after shared changes, not timestamp-sensitive ZIP bytes. Exercise both
OBS colors with long text, inline emphasis, oversized responses and absent data.
Render short/stress exports in Microsoft PowerPoint and inspect edited
first/middle/final pages against their text-box frames, not only slide edges.
Skia model bounds alone are not a claim of identical rendering in every viewer.
