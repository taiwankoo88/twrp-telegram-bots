# TWRP Telegram bots

This repository contains the [Fastener and WN bot](telegram-fastener-bot/) and the [PDF table bot](telegram-pdf-table-bot/). Start with [GEMINI.md](GEMINI.md) for the maintainer handoff, or [README_START_HERE.md](README_START_HERE.md) for the existing setup and history.

## Written Notification setup

The Fastener bot accepts C3D1 and C3D2 steel submissions and returns a Written Notification and Request Form. See the [usage guide](telegram-fastener-bot/使用说明.md) for the message format.

1. Set up the existing Windows bot runtime. The Fastener bot uses the private Codex runtime dependency @oai/artifact-tool. A standard npm install alone does not supply it. If the Node path on your host differs, adjust the launcher locally or start bot.mjs using the configured runtime.
2. Install the WN Python dependency with python -m pip install -r telegram-fastener-bot/wn/requirements.txt.
3. Copy telegram-fastener-bot/steel.example.json to steel.json in that same bot folder. Set root to the absolute path to telegram-fastener-bot/wn and python to the interpreter from step 2. An existing deployment can keep using its configured handoff folder.
4. Supply the four approved Word templates locally, using these paths relative to the repository root:

   - telegram-fastener-bot/wn/outputs/C3D1 Forms/SS_Steel Work Written Notification C3D1_Batch_2.docx
   - telegram-fastener-bot/wn/outputs/C3D1 Forms/SS_Request Form for Issuing Wirtten Notification_Batch_2 Amended.docx
   - telegram-fastener-bot/wn/outputs/C3D2 Forms/SS_Steel Work Written Notification C3D2.docx
   - telegram-fastener-bot/wn/outputs/C3D2 Forms/SS_Request Form.docx

5. Run node test.mjs from telegram-fastener-bot. The WN generation checks use a temporary copy of the templates and verify both projects and revision handling when steel.json is configured. Set STEEL_SKIP_RENDER=1 to test without desktop Word. For PDF previews and visual layout checks, run with desktop Word available on the Windows host.

Keep Word templates, production credentials, chat state, generated documents and submissions local. Use a separate test bot for Telegram delivery checks.
