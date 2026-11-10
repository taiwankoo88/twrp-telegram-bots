# Gemini Enterprise handoff

Start with [README.md](README.md) for WN setup and [CLAUDE.md](CLAUDE.md) for existing bot behavior and operating rules. The PDF table bot is a separate service. Its OCR scripts also support photo input in the Fastener bot.

- telegram-fastener-bot/bot.mjs routes Telegram messages. steel.mjs detects WN submissions, checks required fields and writes revisions for repeated references.
- telegram-fastener-bot/wn/outputs/update_steel_forms.py updates the two approved templates. steel_verify.py checks the output fields and preserved template content. RenderDocx.ps1 uses desktop Word to create PDF previews and check page layout.
- Preserve C3D1 and C3D2 behavior, Singapore dates, multiline descriptions, both Request Form text-box copies, signatures, and revision history. Verify changes with the existing test.mjs using approved local templates and a temporary output folder.
- This repository is public. Keep credentials, business documents, deployment configuration, chat IDs and logs out of commits. Ask the maintainer for the approved templates and authorized runtime before running end-to-end tests. Live deployment remains on the configured Windows host.
