# Offline PO Management

The **offline / general-trade** counterpart to `online_po_management/`, built
to the **same shape** so the two can be merged under one **OMT (Order
Management Tool)** umbrella later.

Where `online_po_management` handles marketplace POs through the
`online_po_processor` package, this handles the offline channels that were
previously loose scripts in `standalone_files/`.

## Structure (mirrors the Online side)

```
offline_po_management/
├── main.py                         # → from offline_po_processor.app import main
├── README.md
├── offline_po_processor/           # the launcher package (cf. online_po_processor)
│   ├── __init__.py
│   ├── app.py                      #   main(): bootstrap + open the launcher
│   ├── config/
│   │   └── channels.py             #   the scalable Channel registry
│   └── gui/
│       └── launcher_window.py      #   channel chooser (Tkinter)
└── channels/                       # the channel tools (relocated as-is)
    ├── eka/
    │   ├── standalone_EKA_constructor.py
    │   └── Calculation_Data_EKA/   #   EKA_DATA.xlsx + Items_March.xlsx (fallback)
    ├── gt_mass/
    │   └── standalone_gt_mass_automation.py   # self-contained (CWD-relative output/)
    └── mt_select/
        ├── standalone_mt_select_automation.py
        ├── Calculation_Data_MT/    #   MT_Masters.xlsx
        ├── mt_select_config.json
        └── mt_select_seq.json
```

## How it runs

```
python main.py            →  offline_po_processor.app.main()
                          →  LauncherWindow  (Tkinter chooser)
                          →  click a channel → launched as an INDEPENDENT
                             subprocess (own Tk loop, CWD = its folder)
```

Each channel is its original standalone tool **with core logic unchanged** —
only relocated. Running them as subprocesses keeps them decoupled (one
crashing can't take down the launcher or the others) and is what lets a
future refactor swap a script for a proper sub-package without touching the
launcher.

## Channels

| Channel | Folder | What it does | Status |
|---|---|---|---|
| **EKA** | `channels/eka/` | Transfer / Sales Order constructor (EKA branches) | ✅ |
| **GT Mass** | `channels/gt_mass/` | GT Mass dump generator (PO → D365 import) | ✅ |
| **MT Select** | `channels/mt_select/` | Modern-trade multi-channel PO processor (e.g. H&G) | ✅ |

### Shared notes
- **EKA** and **MT Select** read the **Item Master** from the shared
  `Online_B2B_Dump_Compilation.xlsx` (the same master the Online tool uses).
  EKA reads it file-open-safe (private in-memory snapshot, Windows shared
  mode), so it works while the compilation is open in Excel.

## Adding a channel

1. Drop its tool under `channels/<key>/` (with any bundled data beside it).
2. Add one `Channel(...)` entry to `offline_po_processor/config/channels.py`.
That's it — the launcher picks it up.

## Roadmap

1. **Now** — offline mirrors online; **EKA + GT Mass + MT Select** migrated in.
2. **Later** — merge `online_po_management` + `offline_po_management` under
   one **OMT** shell (shared launcher / master / history DB).

> Keep this file current as channels are added or restructured.
