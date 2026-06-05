# WordForge

A smart Word document formatter + `.ipynb` to `.docx` converter for college students, with Gemini AI-powered formatting hints.

## Run & Operate

- `cd /home/runner/workspace/artifacts/api-server && uvicorn main:app --host 0.0.0.0 --port $PORT --reload` — run the FastAPI backend (port 8080)
- Frontend runs via the wordforge workflow at port 26132
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- Required env: `GEMINI_API_KEY` — Gemini AI API key for formatting suggestions

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- Frontend: React + Vite + Tailwind CSS + shadcn/ui
- Backend: **FastAPI (Python)** + uvicorn
- Doc processing: `python-docx`, `nbformat`, `mistletoe`, `Pillow`
- AI: Google Gemini (`gemini-2.5-flash`) via `google-genai` Python SDK
- API codegen: Orval (from OpenAPI spec)

## Where things live

- `artifacts/wordforge/` — React Vite frontend
- `artifacts/api-server/main.py` — FastAPI app entry point
- `artifacts/api-server/formatter.py` — Word document formatting logic
- `artifacts/api-server/nb_converter.py` — Jupyter notebook → docx conversion
- `artifacts/api-server/ai_hints.py` — Gemini AI suggestions
- `lib/api-spec/openapi.yaml` — API contract (source of truth)
- `lib/api-client-react/src/generated/` — generated React Query hooks

## Architecture decisions

- FastAPI replaces the Express api-server: better fit for binary file I/O (python-docx, nbformat) and Python ecosystem
- Gemini `gemini-2.5-flash` for AI hints: fast, cheap, and reliable for short contextual suggestions
- File upload endpoints use direct fetch + FormData on the client (not generated hooks) since they return binary blobs
- Only `POST /api/ai-suggest` uses the generated React Query mutation hook (JSON request/response)
- All formatting defaults match Indian university submission standards (Times New Roman 12pt, 1.5 spacing, 2.5cm margins)

## Product

- **Word Formatter**: Upload an existing `.docx` or start fresh. Full typography controls (font, size, color per heading level), header/footer editor with page numbers, margin controls, live HTML preview, download reformatted `.docx`
- **Notebook Converter**: Upload `.ipynb` → download formatted `.docx` preserving markdown headings, code cell shading, output text, embedded images, and DataFrame tables
- **AI Suggestions**: Gemini-powered non-blocking hints triggered by user actions (large H1 → suggest complementary H2, dark color → ask about header background, double spacing → note Indian university standard)

## User preferences

- Use Gemini API for AI features (GEMINI_API_KEY in secrets)

## Gotchas

- Python `uvicorn` server runs from `artifacts/api-server/` directory — the artifact.toml uses absolute path `cd /home/runner/workspace/artifacts/api-server && uvicorn ...`
- File upload endpoints return binary blobs — use direct fetch on the frontend, not generated React Query hooks
- The Node.js api-server source files (src/, build.mjs) are still present but unused — the Python files (main.py, formatter.py, etc.) are what actually run

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
