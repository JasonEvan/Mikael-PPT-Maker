# PPTX Generation & Structural Mapping Specification

The Indonesian generators consume `MappedLiturgyData` from
`app/utils/docxParser.ts`. The template remains the typography and layout
source; matching uses content anchors and presentation relationships, not
hardcoded slide indices.

## 1. Entry points and scope

| Entry point                   | Template                       | Text layout                         |
| ----------------------------- | ------------------------------ | ----------------------------------- |
| `generatePPTXBuffer`          | `Template_Mingguan.pptx`       | Measured weekly layout              |
| `generatePPTXOBSBuffer`       | `Template_PPT_OBS_Hijau.pptx`  | Existing OBS character-based layout |
| `generatePPTXOBSKuningBuffer` | `Template_PPT_OBS_Kuning.pptx` | Existing OBS character-based layout |

`generatePPTXBufferForTemplate` remains the legacy public template entry point
used by OBS. Only the weekly wrapper selects the private measured mode in
`pptxGenerator.ts`. OBS does not initialize the measured layout/font profile.
English uses the same measurement implementation, but retains its own generator,
content mapping, font profile, and pagination policies.

The API selects the weekly wrapper only for `template: "mingguan"`. Its existing
OBS/default routing and English early dispatch remain unchanged. DOCX parsing,
UI, and the `MappedLiturgyData` contract are unchanged.

## 2. Weekly measured layout

`pptxTextLayout.ts` owns DrawingML measurement and pagination. The generator owns
liturgy mapping, static first/final content policy, notes, and ZIP assembly.

- Resolve the edited shape's extents, insets, paragraph margins/spacing,
  indentation, and effective run properties through layout/master/theme.
- Register the exact physical fonts and measure shaped candidate lines with
  Skia Canvas. All paragraphs sharing a shape use one height budget, including
  fixed headings, labels, tab stops, responses, and symbols.
- Preserve intended point sizes. Do not use character budgets, clipping, or
  viewer AutoFit to make weekly content fit.
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
There is no weekly combined-character threshold or 35%-length sentence heuristic.

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
English's visible reading-body pagination for the Indonesian template.

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

## 5. Exact weekly fonts and deployment

The current weekly template requires Arial regular/bold/italic, Arial Narrow
regular/bold/italic variants, Lucida Sans bold italic, and Wingdings. Headings
include 60pt text and the symbol-only Wingdings book icon; body/epigraph/source
runs are typically 48pt, while `(hening sejenak)` uses 44pt. Do not measure all
runs as Arial48 or inherit a Latin theme face for an `a:sym` run.

Lucida Sans's physical demibold italic face has weight600 even though the
DrawingML run requests bold. The engine verifies that exact face instead of
synthesizing weight700 or replacing the family. Symbol runs retain their source
`a:sym` and glyph; measurement selects the symbol face.

Set `INDONESIAN_MASS_FONT_DIR` to a directory containing licensed exact faces.
That override is exclusive. Otherwise the engine searches
`app/fonts/indonesian-mass` and standard macOS Office/system/user font directories.
Typical filenames include `Arial.ttf`, `Arial Bold.ttf`, `Arial Italic.ttf`,
`Arial Narrow Bold.ttf`, `Arial Narrow Bold Italic.ttf`,
`Lucida Sans Demibold Italic.ttf`, and `Wingdings.ttf`.

Fonts must also exist on presentation computers. This repository does not embed
or redistribute proprietary fonts. Next externalizes `skia-canvas` and traces
both profile font folders and templates. Install native dependencies for the
deployment OS/architecture, and provision/mount licensed fonts at runtime.
English retains its exclusive `ENGLISH_MASS_FONT_DIR` and existing font behavior.
OBS does not acquire these font-provisioning requirements.

## 6. OBS remains legacy

Both OBS wrappers still call `generatePPTXBufferForTemplate` with their existing
structural flags and limits: Kolekta200, one-placeholder intentions210, and
postcommunion210 characters. Existing shared legacy branches retain their other
260/240/250 heuristics where applicable, placeholder-only Seruan/Gospel mapping,
one-placeholder intention bolding, response placement, and notes handling.
Those helpers remain live OBS code; they are not weekly fallback paths.

## 7. Verification

- `npm run test:pptx`: shared layout plus English and Indonesian generator tests.
- `npm run test:english`: English generator and shared layout regressions.
- `npm run test:indonesian`: weekly/OBS generator and shared layout regressions.
- Type/lint/build checks do not prove Office rendering.

Use exact licensed fonts for full acceptance. Check source-text conservation and
order, mixed fonts/sizes/symbols, rich intention typography, final responses,
spacer-first fitting, long responses, notes-only scripture, and valid package
relationships. Compare representative short/long OBS and English logical ZIP
part contents before/after shared changes, not timestamp-sensitive ZIP bytes.
Render weekly short/stress exports in Microsoft PowerPoint and inspect edited
first/middle/final pages, icons, responses, spacing, and glyph bounds. Skia model
bounds alone are not a claim of identical rendering in every Office viewer.
