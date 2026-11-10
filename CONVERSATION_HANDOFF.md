# Conversation and business handoff

Prepared 14 September 2026 (Asia/Singapore). This is a structured record of the relevant work in the Codex conversation. It is not a verbatim chat export and contains no secrets.

## Original objective

The user maintains bolt, nut, washer and stud-rod purchase requests for the TWRP C3D1/C3D2 project. They first supplied Excel templates and screenshots, then asked the assistant to normalize new request lists, apply the saved material standards, clean previous records, retain each work title, consolidate eligible quantities and generate Excel files.

The workflow was later connected to a Telegram bot so colleagues can submit a text order and receive the generated Excel file directly.

## Templates and material standards

Two reusable templates were established:

- HDG Bolt and Nut Template.
- SS316 Bolt and Nut Template. The user often writes `SS16`; treat it as SS316.

Saved standards:

| Material | Item | Standard | Strength/grade |
| --- | --- | --- | --- |
| SS316 | Bolt | BS EN ISO 3506 | A4-70 |
| SS316 | Nut | BS EN ISO 3506 | A4-80 |
| SS316 | Flat washer | BS EN ISO 3506 | A4-70 |
| SS316 | Stud rod | BS EN ISO 3506 | A4-70 |
| HDG | Bolt | BS EN ISO 4190 | G 8.8 |
| HDG | Nut | BS EN ISO 4190 | G 8.8 |
| HDG | Flat washer | BS 4320 | G 8.8 |

Washer thickness convention used by the user:

- M10, M12, M16 and M20: 3 mm.
- M24, M27 and M30: 4 mm.

Every bolt description must now read `FULL THREADED HEXAGON BOLT: Mxx, L=...mm`.

## Gasket orders (added 15 September 2026)

The user supplied `TWRP C3D1 C3D2 Gasket Order.xlsx` as the EPDM gasket template, stored as `templates/Gasket.xlsx`. Colleagues write only the project, title, DN, PN and quantity; every other cell follows the template.

- First line `C3D1 Gasket` or `C3D2 Gasket` (`EPDM`, `EPDM Gasket` and `Gasket EPDM` are also accepted). The parser stores material `EPDM`; the file is named `C3D1 EPDM Gasket Order.xlsx`.
- Item lines: `DN400 PN10 = 10`. Thickness is 3 mm unless written, e.g. `DN400 PN10 5mm = 10`. The words `EPDM`, `gasket`, `Full-Face` and `with bolt holes` are ignored on item lines.
- Each title is its own section, written as `Title: …` in column B without a section number, as in the template. Gaskets combine only when DN, PN and thickness match within a title; combined quantities keep the `Combined: 30 + 20` remark and a `=SUM(...)` formula. `Sets:` works as for fasteners.
- Fixed per template: Size `SIZE: DN400 3MM THK Full-Face with Bolt Holes` (the template's `wth` was changed to `with` at the user's request on 15 September 2026), Standard `BS 1154 Class Y3`, Reinforcement `Single layers of fabric in accordance with BS EN 751` with bold `BS EN 751`, PN as a number, Type `EPDM`, subject `RFQ for Gasket EPDM`, `Document Required: Certificate for all items to be attached together with DO.`, blank unit prices, and amount/total formulas in columns I and J.
- Fasteners in a gasket order, or gaskets in a fastener order, are rejected with a message asking for a separate order. Other gasket materials and ANSI class ratings are not supported.
- Weekly summaries list gasket items as `SIZE: …, PN10 × qty`; the Excel list's `Strength / PN` column holds `PN10`.

## Consolidation and layout rules

- Remove the template's previous bolt/nut records before filling a new order.
- Put every item/location title into the workbook as its own section.
- Keep different bolt lengths as separate items.
- Combine nuts of the same diameter within the same title/location.
- Combine flat washers only when diameter and thickness match within the same title/location.
- Do not consolidate across different titles, projects or materials.
- A `Sets:` order multiplies each stated per-set quantity by the set count.
- Do not add an extra 5% unless explicitly instructed for that order.
- Add `Requester Name` to every workbook. When Telegram input omits it, use the sender's Telegram display name.
- Preserve numeric quantity cells and the template's unit-price/amount formulas.

## Historical request records supplied in the conversation

The following records show prior inputs and decisions. They are useful examples; do not automatically append them to a new order.

### WDL12 Level 3 Pump Gallery 1, SS316

- M24 x 90 bolt: 1,200; M24 nut: 1,200; M24 washer: 2,400.
- M20 x 80 bolt: 2,000; M20 nut: 2,000; M20 washer: 4,000.
- This was used to establish the clean SS316 template.

### C3D1 WDL12 FRP Volume Control Damper and Isolation Valve, SS316

- M16 x 65 bolt 720; nut 720; washer 1,440.
- M16 x 70 bolt 160; nut 160; washer 320.
- M16 x 80 bolt 120; nut 120; washer 240.
- M20 x 65 bolt 600; nut 600; washer 1,200.
- M20 x 70 bolt 950; nut 950; washer 1,900.
- M20 x 80 bolt 700; nut 700; washer 1,400.
- A correction established that quantities must be read exactly, e.g. M16 x 70 is 160 pieces, not 3,800.

### HDG handrail set definition

One set:

- M12 x 135 mm HDG hex bolt: 1.
- M12 HDG nut: 3.
- M12 x 3 mm HDG flat washer: 4.

Locations previously supplied:

- C3D1 WDL15 RO Shelter high-platform handrail: 75 sets.
- C3D1 WDL15 RO Shelter rooftop: 360 sets.
- C3D1 WDL06 Access Platform handrail + WDL11 Grit Classifier Platform handrail: 490 sets.
- C3D1 WDL13 CIP Platform handrail: 20 sets.
- C3D2 WDL21 Grit Classifier Platform handrail: 50 sets.
- C3D2 WDL21 IFC Platform handrail: 16 sets.
- C3D2 WDL22 PST Tank Platform handrail: 350 sets.

C3D2 WDL23 bioreactor piperack installation used 50 sets of:

- M16 x 100 mm HDG bolt: 1.
- M16 HDG nut: 3.
- M16 x 3 mm HDG flat washer: 4.

### Large SS316 replacement list supplied earlier

- WDL12 chemical hanging piperack replacement: M16 x 2 m stud rod 20; M16 x 3 washer 200; M16 nut 400.
- WDL12 cantilever piperack replacement: M20 nut 120; M20 x 3 washer 60; M20 x 2 m stud rod 8.
- WDL12 MBR fine screen piperack replacement: M16 nut 140; M16 x 3 washer 70; M16 x 2 m stud rod 7.
- WDL13 bioreactor PR: M16 nut 60; M20 nut 1,100; M30 nut 120; M16 x 3 washer 50; M20 x 3 washer 1,000; M30 x 4 washer 110; M30 x 2 m stud rod 22.
- WDL13 CIP Platform: M16 x 2 m stud rod 23; M16 x 3 washer 185; M16 nut 370.
- WDL15 RO Feed PR: M20 nut 600; M20 x 3 washer 300; M20 x 2 m stud rod 40.
- WDL13 MBR filtrate pump monorail: M10 nut 100; M12 nut 400; M10 x 3 washer 65; M12 x 3 washer 285; M10 x 2 m stud rod 10; M12 x 2 m stud rod 10.
- WDL13 DEOX RAS crane: M12 nut 70; M16 nut 200; M20 nut 180; M24 nut 450; M12 x 3 washer 45; M16 x 3 washer 135; M20 x 3 washer 120; M24 x 4 washer 300.
- WDL13 ROHP EOTC: M16 nut 120; M16 x 3 washer 80.
- WDL13 ST237: M20 nut 70; M20 x 3 washer 35.
- WDL15 UV staircase: M20 nut 260; M20 x 3 washer 130; M20 x 2 m stud rod 21.
- WDL15 UV third batch: M20 x 2 m stud rod 85; M20 nut 920; M20 x 3 washer 460.
- WDL23 Central Gallery Piperack replacement: M24 x 2 m stud rod 5.
- WDL23 MBR Tank SS structure for 20 tanks: M12 x 2 m stud rod 600; M16 x 2 m stud rod 160.

### WDL22 Lamella Truss, SS316

- M20 nut 560; M24 nut 2,080; M30 nut 330.
- M16 x 3 washer 60; M24 x 4 washer 930.
- M16 x 65 bolt 180; M20 x 65 bolt 860; M20 x 85 bolt 1,300; M24 x 125 bolt 540; M30 x 125 bolt 330.

### WDL12 FCV for ALP service, SS316

- Reference: WDL12-033-FCV-0101 to 0601.
- M20 x 85 bolt 150; M20 nut 150; M20 washer 300.

### Later C3D1/C3D2 additions

- C3D1 WDL12 Chemical Hanging Piperack add-on, replacement to ST311-22048, no extra 5%: M16 x 60 HDG bolt 400; M16 HDG washer 500; M16 HDG nut 555.
- C3D2 WDL23 MBR filtrate pipe rack, back charge SMZ: M20 x 3 SS washer 105, M20 SS nut 210; M24 x 4 SS washer 110, M24 SS nut 220; M30 x 4 SS washer 255, M30 SS nut 510; M20 SS stud rod 5, M24 SS stud rod 5, M30 SS stud rod 10.
- C3D2 WDL23 MBR filtrate tank EOTC ST345, back charge SMZ: M12 x 3 SS washer 155; M12 SS nut 310.

### C3D1 WDL12 individual pipe-support chemical anchor bolts, SS316

- Nut 320; washer 320.
- M16 x 2 m stud rod 73.
- M12 x 2 m stud rod 29.
- An M10 total row was present but had no supplied quantity.

### Pipe connections, SS316

- WDL13 ALP pipe into MBR tank connection: M16 x 170 bolt 200; M16 nut 400; M16 washer 450; DN200 gas-cut gasket 50. The user clarified that DN200 gas-cut means gasket and accepted the proposed washer thickness.
- WDL11 G2D pipe connection: M27 x 330 bolt 10; M27 nut 20; M27 washer 10. M20 x 230 bolt 24; M20 nut 24; M20 washer 48. M20 x 180 bolt 100; M20 nut 100; M20 washer 200. M20 nuts and washers were combined by size in the finished order.

### C3D1 WDL13 membrane tank pipe header valve connection, SS316

- M16 x 110 bolt 2,880.
- M16 nut 2,880.
- M16 x 130 bolt 5,760.
- M16 nut 5,760.
- M16 flat washer 17,280.
- Final consolidation: M16 nut 8,640; M16 x 3 mm flat washer 17,280; bolts remained separated by length.
- Saved template grades were applied: nuts A4-80, bolts and washers A4-70, even though the source screenshot displayed A4-70 on nut rows.

## Telegram bot implementation history

- User requested that colleagues write orders in Telegram and receive files directly.
- Chosen deployment: local Windows PC using Telegram Bot API long polling. The PC must remain logged in, awake and online.
- Bot created as `@TWRPC3D1C3D2Bot` with display name `TWRP C3D1/C3D2 Fasteners`.
- Token is entered locally through `Setup.cmd`, validated through Telegram and stored with Windows DPAPI.
- First administrator pairs using `/pair CODE`.
- Colleagues send `/id`; administrator sends `/allow NUMERIC_ID`. `/remove NUMERIC_ID` revokes access.
- Private chats only; group messages are ignored.
- All guidance and errors were changed to English because some users do not read Chinese.
- Telegram startup now retries transient network failures such as `ECONNRESET`.
- `Restart.cmd` stops and starts the bot, waits for status and displays ONLINE when connected.
- A test order successfully returned an Excel file through Telegram.
- The latest change made every bolt designation `FULL THREADED HEXAGON BOLT`.

## Supported Telegram text format

```text
C3D1 SS316
Requester Name: Nicholas
Title: WDL13 Membrane Tank Pipe Header Valve Connection
M16x110mm bolt = 2880
M16 nut = 2880
M16x130mm bolt = 5760
M16 nut = 5760
M16x3mm flat washer = 17280
```

Set-based example:

```text
C3D2 HDG
Title: WDL21 Platform handrail
Sets: 50
M12x135mm bolt = 1
M12 nut = 3
M12x3mm washer = 4
```

Gasket example:

```text
C3D1 Gasket
Requester Name: Nicholas
Title: WDL12 Pump Gallery 2 and 3 PSD
DN400 PN10 = 10
DN200 PN10 = 30
```

## Current limitations and likely next work

- Text input only; no image OCR, Excel/PDF attachment parsing or general conversational AI.
- Only EPDM gaskets in the gasket template's form are handled (DN and PN); other gasket materials or ANSI class ratings are rejected.
- HDG stud rods are rejected because a saved standard has not been confirmed.
- One Telegram message can contain only one project/material combination.
- Local host must remain online; automatic Windows login/startup was not configured.
- The live implementation depends on Codex's bundled Node runtime and `@oai/artifact-tool` through a junction. A portable deployment requires a replacement dependency or a compatible runtime.
- The BotFather token was visible in a screenshot during setup. The user was told to revoke it and use a new token. Do not retain or reproduce the exposed token.

