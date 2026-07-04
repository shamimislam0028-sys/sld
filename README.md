# MCQ PowerPoint Generator

Generate `.pptx` MCQ decks from a web form, **reusing your existing PowerPoint
template verbatim**. The app never recreates the layout. It opens the real
`.pptx` (which is a zip of XML), clones the question/answer slide pair once per
MCQ, and replaces **only**:

- the question text
- the four option texts
- the "MCQ N" number
- the **course icon** — placed inside the top-left "Icon" box on the layout
- the **course mnemonic image** — placed inside the top-right box on the layout
- the **correct option's box color** on the answer slide (pink → blue)

Everything else — fonts, colors, rounded corners, shadows, spacing, the Shikho
background, the blue question box, the ক/খ/গ/ঘ letter circles, the GPA5/Shikho
logos — comes straight from your template and is never touched.

> **No `pptxgenjs`.** `pptxgenjs` builds slides from scratch and cannot preserve
> a Google-Slides-exported template's exact shapes. To honour the "never redraw
> shapes" requirement, this project edits the template's XML directly.

---

## How each MCQ becomes two slides

| | Question slide | Answer slide |
|---|---|---|
| Question box | your question | your question |
| Options A–D | all **pink** | correct one turns **blue** |
| MCQ number | `MCQ N` | `MCQ N` |

The correct answer is chosen with an A/B/C/D dropdown. In the template the option
letters are the Bengali glyphs **ক / খ / গ / ঘ**, so A→ক, B→খ, C→গ, D→ঘ
automatically. Only the option *value* box is recolored; the letter circle is
left as-is.

---

## Project structure

```
mcq-pptx-generator/
├── backend/
│   ├── server.js              # Express API (/api/generate, /api/health)
│   ├── test.js                # standalone engine test (no server needed)
│   ├── src/
│   │   └── pptEngine.js       # the verbatim-template XML engine (core)
│   └── templates/
│       └── master.pptx        # YOUR template — swap this to restyle
├── frontend/
│   ├── index.html
│   ├── vite.config.js
│   └── src/
│       ├── main.jsx
│       ├── App.jsx            # step flow, autosave, drag-reorder, generate
│       ├── CourseInfo.jsx     # step 1: mnemonic + icon upload
│       ├── QuestionCard.jsx   # one MCQ editor card
│       ├── utils.js           # validation, JSON import/export, localStorage
│       └── styles.css
├── sample-questions.json      # example import file
└── README.md
```

---

## Requirements

- **Node.js 18+** (for `crypto.randomUUID` in the browser and `--watch`)
- `sharp` (icon transcoding) compiles a native binary — Node 18/20 recommended.

---

## Install & run locally

Two terminals.

### 1. Backend

```bash
cd backend
npm install
npm start          # http://localhost:4000
```

Quick engine sanity check without the server or frontend:

```bash
npm test           # writes ./test_out.pptx using built-in sample questions
```

### 2. Frontend

```bash
cd frontend
npm install
npm run dev        # http://localhost:5173  (proxies /api to :4000)
```

Open **http://localhost:5173**, fill in the course mnemonic, upload an icon,
add questions, and click **Generate PowerPoint**.

### Production (single server)

```bash
cd frontend && npm run build      # outputs frontend/dist
cd ../backend && npm start        # serves the built UI + API on :4000
```

---

## Using the app

**Step 1 — Course info**
- Course mnemonic image: PNG / JPG / SVG (required) — fills the top-right box
- Course icon: PNG / JPG / SVG (required). Placed inside the top-left "Icon"
  box on every slide. Both images are auto-converted to PNG server-side.

**Step 2 — Questions**
- Add / Duplicate / Delete questions (auto-numbered)
- Drag the `⋮⋮` handle to reorder
- Correct answer dropdown (A/B/C/D)
- **Export JSON** / **Import JSON** (see `sample-questions.json`)
- **Auto-save**: mnemonic + questions persist in `localStorage` (the icon is a
  file and is re-selected each session)
- Live counter + progress bar
- Inline + summary validation; generation is blocked until everything is valid

---

## JSON import format

An array of objects (see `sample-questions.json`):

```json
[
  {
    "question": "Your question?",
    "optionA": "First",
    "optionB": "Second",
    "optionC": "Third",
    "optionD": "Fourth",
    "correct": "B"
  }
]
```

`correct` must be `"A"`, `"B"`, `"C"`, or `"D"`. A wrapper object
`{ "mcqs": [ ... ] }` is also accepted.

---

## Swapping the template later

The whole point of the design: **to restyle, replace one file.**

1. Drop your new deck in `backend/templates/master.pptx`.
2. The new template must keep the same two-slide pattern: slide 1 = a question
   slide, slide 2 = its answer slide.
3. If your new template uses **different shape IDs** for the question / options /
   number, update the `TEMPLATE` map at the top of
   `backend/src/pptEngine.js`:

   ```js
   const TEMPLATE = {
     questionSlide: 'ppt/slides/slide1.xml',
     answerSlide:   'ppt/slides/slide2.xml',
     shapeIds: {
       mcqNumber: '202',
       question:  '203',
       options: { A: '206', B: '209', C: '212', D: '215' },
     },
     layout: {                                  // icon box + mnemonic box
       path:     'ppt/slideLayouts/slideLayout7.xml',
       relsPath: 'ppt/slideLayouts/_rels/slideLayout7.xml.rels',
       iconBoxId:     '2',   // top-left box  — image fills this
       mnemonicBoxId: '6',   // top-right box — mnemonic text
     },
     colors: { pink: 'E2098D', blue: '002060' }, // option fill / highlight
   };
   ```

   To find shape IDs in a new template, unzip it and look at
   `ppt/slides/slide1.xml` for `<p:cNvPr id="…">` next to each text.

No other code changes are needed.

---

## How it works (engine internals)

`pptEngine.js`:

1. Loads `master.pptx` with `JSZip`.
2. Reads the canonical question slide (`slide1.xml`) and answer slide
   (`slide2.xml`) plus their `.rels`.
3. For each MCQ it clones both slides, then:
   - `setShapeText` finds a shape by its `<p:cNvPr id>` via balanced
     `<p:sp>…</p:sp>` matching (shapes are nested inside Google-export groups,
     so naive regex would corrupt them), keeps the first run's `<a:rPr>`
     formatting exactly, swaps its `<a:t>` text, and drops the extra runs.
   - On the answer slide, `highlightOption` resets every option box to pink,
     then recolors only the correct option's first `solidFill` to blue.
4. Optionally replaces `ppt/media/image1.png` with the uploaded icon.
5. Rewrites `presentation.xml` (`sldIdLst`), `presentation.xml.rels`, and
   `[Content_Types].xml`, then repacks the zip.

---

## License

MIT.
