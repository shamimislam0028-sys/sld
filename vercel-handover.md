# Technical Handover: Vercel Multi-Service Deployment & PPTX Engine Fixes

This document serves as a complete technical handover detailing the architecture, challenges faced, and modifications made to deploy the MCQ PowerPoint Generator on Vercel and fix the engine's core bugs.

---

## 1. Project Goal & Architecture
The project is a React (Vite) frontend and Express (Node.js) backend monorepo that processes user inputs and a `.pptx` template directly (via balanced XML tag manipulations inside the presentation zip) to output generated PowerPoint decks.

We converted the project into a Vercel-ready monorepo using Vercel's **Multi-Service Architecture** (in Beta) which allows deploying independent services in a single repository sharing the same domain.

---

## 2. Key Challenges & Solutions

### Challenge A: Native OS Binaries Mismatch (Backend)
- **Problem**: The backend used the `sharp` library for image conversion. `sharp` relies on platform-specific C++ binary additions. In a Windows development environment, `npm install` locks the Windows binary in `package-lock.json`. When deployed to Vercel's Linux x64 containers, it crashed with `Error: Could not load the "sharp" module using the linux-x64 runtime`.
- **Solution**: We replaced `sharp` with **`jimp`** in `backend/package.json` and `backend/server.js`. Jimp is a 100% pure JavaScript image library. It runs identically on Windows and Vercel without compiling or fetching native OS binaries.

### Challenge B: Native OS Binaries Mismatch (Frontend)
- **Problem**: The frontend used Vite v8 (`"vite": "^8.1.0"`), which uses **Rolldown** as its bundler. Rolldown relies on native Rust bindings (`@rolldown/binding-linux-x64-gnu`), causing the same platform-specific binary lock crash during Vercel's build phase.
- **Solution**: We downgraded Vite to **`^5.4.11`** (Vite 5) in `frontend/package.json`. Vite 5 uses **Rollup**, which is written in standard JS, does not require native OS binaries, and builds perfectly on Vercel.

### Challenge C: PowerPoint File Corruption (Engine Bug)
- **Problem**: When generating decks with more than 1 question, the resulting `.pptx` file was marked as corrupt by PowerPoint. The template has speaker notes files (`notesSlide1.xml`, etc.). The engine duplicated slides and copied their relationship files (`slide1.xml.rels` pointing to `notesSlide1.xml`). This caused multiple slides to link to the *same* notes slide, violating the OpenXML standard's strict 1-to-1 notes constraint.
- **Solution**: Since generated MCQ slides don't require speaker notes, we updated `backend/src/pptEngine.js` to strip the `notesSlide` relationship declarations from the slide `.rels` files using a regex before appending slides.

### Challenge D: Vercel Reverse Proxy Routing
- **Problem**: The frontend called `/api/generate`. Vercel's multi-service router mapped `/api` to the backend service but stripped the `/api` prefix from the path before it reached the Express container, resulting in a 404 response from Express.
- **Solution**: We updated `backend/server.js` route handlers to accept arrays of paths: `app.post(['/api/generate', '/generate'], ...)` and `app.get(['/api/health', '/health'], ...)`.

---

## 3. File-by-File Changes Summary

### 1. Root Configurations
- **[vercel.json](file:///vercel.json)**:
  Defines the multi-service build instructions and reverse proxy rewrites:
  ```json
  {
    "experimentalServices": {
      "frontend": {
        "root": "frontend",
        "framework": "vite",
        "routePrefix": "/"
      },
      "backend": {
        "root": "backend",
        "entrypoint": "server.js",
        "routePrefix": "/api"
      }
    }
  }
  ```
- **[package.json](file:///package.json)**:
  Configured npm workspaces to tie the monorepo together:
  ```json
  "workspaces": [
    "backend",
    "frontend"
  ]
  ```

### 2. Backend Files
- **[backend/package.json](file:///backend/package.json)**:
  Removed `sharp` and `@img/sharp-linux-x64`; added `"jimp": "^0.22.12"`.
- **[backend/server.js](file:///backend/server.js)**:
  - Replaced `sharp` imports with `Jimp`.
  - Updated `imageToPng` to use `Jimp.read(file.buffer)` and `getBufferAsync(Jimp.MIME_PNG)`.
  - Allowed both `/api/...` and `/...` route paths.
  - Wrapped `app.listen()` inside `if (require.main === module)` and exported `app` using `module.exports = app` so Vercel can run the app serverlessly.
- **[backend/src/pptEngine.js](file:///backend/src/pptEngine.js)**:
  Added a clean helper to strip out notes slide relationships:
  ```javascript
  const cleanRels = (xml) => xml.replace(/<Relationship\s+[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*\/>/g, '');
  ```

### 3. Frontend Files
- **[frontend/package.json](file:///frontend/package.json)**:
  Downgraded devDependencies to use `"vite": "^5.4.11"`.
- **[frontend/vite.config.js](file:///frontend/vite.config.js)**:
  Reverted custom root build output paths so Vite outputs to standard local `dist/` relative to the frontend directory, matching Vercel's isolated service context.

---

## 4. How to Verify & Run Locally
1. Run workspace installation: `npm install`
2. Run backend test suite: `npm test --workspace=backend` (should output generated bytes without error).
3. Run local dev:
   - Terminal 1 (Backend): `npm run start:backend`
   - Terminal 2 (Frontend): `npm run dev:frontend`
