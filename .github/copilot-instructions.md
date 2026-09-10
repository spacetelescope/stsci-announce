# stsci-announce

A JupyterLab 4.x extension (forked from NERSC's nersc-refresh-announcements) that polls an
external API for announcements and shows a status-bar button / modal dialog with the content.

## Architecture

This is a hybrid TypeScript/Python JupyterLab "prebuilt extension":

- `src/index.ts` — the entire frontend logic in one file. Key pieces:
  - `Message` / `MessageBlock` / `AnnouncementsData` classes validate and sanitize raw JSON from
    the announcements API before rendering. Each constructor throws (`log_and_throw`) on malformed
    input — treat the API response as untrusted.
  - `sanitizeHtml()` (DOMPurify, allowlist of tags/attrs) sanitizes message/title HTML content;
    `escapeHtml()` is used for plain-text metadata fields (username, timestamp, level). Don't mix
    these up when handling new fields — HTML content must go through `sanitizeHtml`, plain text
    through `escapeHtml`.
  - `RefreshAnnouncements` owns the polling loop (`updateAnnouncements`), status-bar button state
    machine (`serviceState`: `normal` / `degraded` / `failed`), and exponential backoff retry logic
    (`retryDelay`, doubling up to `maxRetryDelay`, reset on success).
  - The plugin (`PLUGIN_ID = 'stsci-announce:plugin'`) reads its `url` and `refresh-interval`
    settings from `schema/plugin.json` via `ISettingRegistry` at activation.
- `schema/plugin.json` — user-configurable settings schema (announcement API URL, refresh
  interval in ms). Mirrors `example_overrides.json`, which shows how admins set these system-wide.
- Compiled TS (`lib/`) and the bundled labextension assets (`stsci_announce/labextension/`) are
  build artifacts — do not hand-edit them; regenerate via the build commands below.
- `stsci_announce/` is the thin Python package (hatch + `hatch-jupyter-builder`) that wraps the
  built JS assets so `pip install` deploys the prebuilt extension into JupyterLab.
- `tests/test-server/server.js` is a minimal Express server used only to serve fake announcement
  JSON payloads (normal or `--empty-announcement`) for the Selenium-based integration tests.
- `tests/test_nonempty_announcement.py` / `test_empty_announcement.py` are Selenium/pytest
  end-to-end tests that drive a real running JupyterLab + Firefox/geckodriver and assert on the
  rendered status-bar button and dialog content — they require the test server and a running
  `jupyter lab` instance (see `.github/workflows/main.yml` for the exact sequence).

## Build / lint commands

```bash
jlpm                      # install JS deps (or `yarn`/`npm`)
jlpm run build            # build:lib (tsc) + build:labextension:dev
jlpm run build:prod       # clean + build:lib + build:labextension (production)
jlpm run watch            # watch:src (tsc -w) + watch:labextension, for live development
jlpm run eslint:check     # ESLint check only (no fixes)
jlpm run eslint           # ESLint with --fix
```

Python side:

```bash
pip install -e .                              # editable install, builds JS via hatch-jupyter-builder
jupyter labextension develop . --overwrite    # symlink dev build into JupyterLab
jupyter labextension list                     # verify "stsci-announce ... OK"
```

There is no unit test runner for the TS code and no single-test invocation for the Python tests
beyond standard pytest file selection, e.g. `pytest tests/test_empty_announcement.py`. Both
Python tests require a live JupyterLab instance on `localhost:8888` and the fake announcement
server (`node tests/test-server/server.js` / `--empty-announcement`) already running — they are
not standalone unit tests.

## Conventions

- Single quotes, no semicolon-omission issues (`prettier`: singleQuote, no trailing comma,
  printWidth 140) — run through `eslint`/`prettier`, don't hand-format differently.
- Interfaces must be named `I<PascalCase>` (enforced by `@typescript-eslint/naming-convention`).
- `dlog()` is a no-op wrapper around `console.log` used for verbose/debug tracing in `index.ts`;
  prefer it (or leave it commented) over adding new raw `console.log` debug statements.
- Any new field coming from the external announcements JSON must be validated in its
  corresponding class constructor (regex/length/enum checks) before use, following the existing
  `Message`/`MessageBlock`/`AnnouncementsData` pattern — never trust the API payload directly.
