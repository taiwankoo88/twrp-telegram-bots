# Claude Code handoff: Telegram Fastener Order Bot

This package transfers the working context and a safe copy of the Telegram fastener-order project to Claude Code.

Start with these files:

1. `CLAUDE.md` — operating rules for Claude Code.
2. `CONVERSATION_HANDOFF.md` — business rules, history, current state and known limitations.
3. `telegram-fastener-bot/` — source code, launchers, tests and the HDG/SS316 Excel templates.

## Safe transfer

The package deliberately excludes:

- `token.dpapi` — encrypted Telegram bot token.
- `config.json` — administrator pairing code.
- `state.json` — Telegram administrator and approved user IDs.
- `status.json`, `bot.log`, `error.log`.
- Generated Telegram order folders, requester details and delivery records.
- `node_modules` junction.

The live installation remains at:

`C:\Users\nicholas.koo\Documents\Codex\2026-08-14\i\work\telegram-fastener-bot`

If Claude Code runs on this same computer and needs to edit the live bot, point it to that live folder after it has read this handoff. Do not paste the Telegram token into an AI chat.

## Current status

- Telegram bot username: `@TWRPC3D1C3D2Bot`.
- Local Windows long-polling bot is operational.
- The user has completed administrator pairing.
- English Telegram guidance and errors are enabled.
- HDG and SS316 Excel generation has been tested.
- Every generated bolt description now begins with `FULL THREADED HEXAGON BOLT`.
- Gasket (EPDM) orders are supported since 15 September 2026: a message starting with `C3D1 Gasket` or `C3D2 Gasket` with `DN400 PN10 = 10` lines produces `C3D1 EPDM Gasket Order.xlsx` from `templates/Gasket.xlsx`.
- A restart is required after source changes: run `Restart.cmd` in the live folder.

## Use with Claude Code

Extract the ZIP, open a terminal in the extracted folder, and start Claude Code there. Paste the content of `PROMPT_TO_CLAUDE.txt` as the first message. If Claude Code is on this same computer, tell it that the live installation path above is the deployment target.


## Second bot: PDF Table to Excel (added 15 September 2026)

`telegram-pdf-table-bot/` is a separate Telegram bot with its own BotFather token, chat and allowlist. A user sends a PDF, or a photo or screenshot of a table, and receives an Excel file of its tables. Photos are read offline on the host PC with OpenCV and RapidOCR, which were added on 15 September 2026. It runs on Python (PyMuPDF, openpyxl, OpenCV, RapidOCR), not on the Codex Node runtime. Live folder: `C:\Users\nicholas.koo\Documents\Codex\2026-08-14\i\work\telegram-pdf-table-bot`. See `使用说明.md` in that folder. Its `token.dpapi`, `config.json`, `state.json`, `status.json`, logs and `outputs/` are private and must not be included in a handoff.
