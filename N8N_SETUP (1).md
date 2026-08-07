# Document Translation — n8n Setup (pure-n8n, Anthropic)

No host script, no `/opt/translate`, no python, no `merge_runs.py`. Everything
runs inside n8n Code nodes using **Node built-ins only** (`zlib`) — so it works
regardless of whether your instance allows external npm modules like jszip.

**Flow:** `Webhook → Extract Text Nodes → Build Batches → Translate Batch → Rebuild DOCX → Respond to Webhook`

Provider: **Anthropic** (`api.anthropic.com/v1/messages`, model
`claude-sonnet-4-6`), via an httpRequest node. Gated on the Anthropic BAA before
real client documents flow through it; non-PHI template docs (HR forms,
newsletters, notices) can be tested/demoed before the BAA lands.

---

## 1. Set the API key in the n8n environment

The Translate Batch node reads `{{ $env.ANTHROPIC_API_KEY }}`. Put the key in
the n8n process env (docker-compose / systemd unit), not hardcoded in the node:

```
ANTHROPIC_API_KEY=sk-ant-...
```

If your instance doesn't expose `$env` to expressions, swap the node to an
**httpRequest credential** (Header Auth: `x-api-key` = the key) instead — same
result, no env dependency.

---

## 2. Import the workflow

n8n → Workflows → **Import from File** → `translate_workflow.json`. Five nodes,
wired left to right. Node param names drift across n8n versions — if a field is
empty/red after import, re-set it. Verify these:

1. **Webhook** — POST, path `translate-doc`, **Respond = "Using Respond to
   Webhook Node."** Confirm multipart/binary is accepted so the upload arrives as
   binary property `file`. (Options → Binary Property = `file`.)
2. **Translate Batch** (httpRequest) — confirm the header `x-api-key` resolves to
   the env var, `anthropic-version: 2023-06-01` is present, and the JSON body
   expression survived import (should be a single `JSON.stringify({...})`
   expression, not fixed text). **Toggle the body field to expression mode if
   import reset it to fixed** — this is the #1 thing that breaks on import.
3. **Respond to Webhook** — respond with **Binary**, property `data`, headers set
   to the docx mime type + `Content-Disposition: attachment`.

The two Code nodes (Extract, Rebuild) carry all the zip/xml logic inline; nothing
to configure in them.

---

## 3. Point the front end at it

Activate the workflow, copy the **Production** webhook URL, set it in `index.html`:

```js
const WEBHOOK_URL = "https://your-n8n-host/webhook/translate-doc";
```

Repo path: **`materials/translate/index.html`** (GitHub Pages, same as Compass).
The committed copy is **unbranded** (neutral palette, "Wellness Tools" kicker).
Per-client branding is injected on the admin side, same pattern as Check Queue
white-label — do not bake Concentra colors into the repo copy.

---

## 4. Test

Raw, without the page:

```bash
curl -X POST https://your-n8n-host/webhook/translate-doc \
  -F "file=@some_doc.docx" \
  -F "language=Spanish" \
  -F "register=conversational" \
  -o out.docx
open out.docx
```

Then through the page: upload → pick style → Translate → file downloads. Confirm
on a doc with a **table, bullets, and a header/footer** that layout is intact and
only the words changed.

`register` accepts `basic | business | conversational` (Plain / Business /
Conversational buttons). `language` is the dropdown value.

---

## How it works (one paragraph)

A `.docx` is a zip of XML. **Extract Text Nodes** unzips it in-memory (built-in
`zlib`), pulls every `<w:t>` text span from `word/document.xml` + any
`header*.xml` / `footer*.xml`, recording the exact byte offsets so we can splice
back precisely. **Build Batches** chunks the strings (40/call) and attaches the
register-specific system prompt. **Translate Batch** sends each chunk to Anthropic
as a numbered JSON array and gets a same-length array back. **Rebuild DOCX**
splices each translation into its original byte span (back-to-front so offsets
stay valid), rezips starting from the *original* archive so images/styles/rels
pass through untouched, and returns the binary. Layout survives because only the
text inside each node changes — no structure is ever touched.

---

## Known caveats

- **No run-merge.** We deliberately skip the docx-skill `merge_runs.py` step (it
  rewrites run structure, the one operation that can disturb formatting). The cost:
  a sentence Word split across runs is translated as separate fragments. The batch
  prompt tells the model to translate fragments independently; in HR forms /
  newsletters this is rarely visible. If a specific client doc shows awkward
  fragment seams, that's the tradeoff — revisit run-merge only if it actually bites.
- **Batch pairing.** Rebuild pairs each API response with its Build Batches item
  **by index** (`$('Build Batches').all()[i]`). n8n preserves item order through
  the httpRequest node, so index alignment holds. If you ever see a translation
  landing in the wrong slot, that ordering assumption is the first thing to check.
- **max_tokens 8000** per batch is comfortable for 40 short strings. Very long
  paragraphs in a single doc could approach it; lower the batch size in Build
  Batches (the `BATCH` const) if you hit truncation.
- **v1 = .docx only.** PDF is fixed-box layout (ES runs ~20–30% longer than EN and
  overflows) — separate approach, later phase. Front end restricts to `.docx`.
