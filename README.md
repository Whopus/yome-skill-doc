# @yome/doc

Microsoft Word document editing commands for Yome agents — full Word for Mac
AppleScript surface (file ops + paragraph read/write/format + tables, headers /
footers / sections, find / replace, bookmarks, hyperlinks, comments, images /
shapes, styles, lists, table-of-contents, page setup, view, track changes,
print / export). One of the official skills.

## Layout (spec v0.1, section 4)

```
yome-skill-doc/
├── yome-skill.json                       manifest (slug / domain / delivery / capabilities)
├── README.md
├── signature/
│   └── doc.signature.json                LLM-facing command signature (truth)
├── backends/
│   └── macos/                            Declarative manifest + .applescript templates,
│                                         consumed by cli/src/skills/runner/dispatcher.ts.
│                                         Mirrors Yome/macOS/Bridge/WordBridge.swift.
└── (ios / node / sandbox come later)
```

## Status during v0.1 monorepo phase

The signature in `signature/doc.signature.json` is byte-aligned with the
runtime descriptor `Server/agent/commands/docCommands.ts`. The macOS
implementation also lives inside the Yome app target as
`Yome/macOS/Bridge/WordBridge.swift`; `backends/macos/` ships an OTA-style
declarative manifest + AppleScript templates so the **CLI hub-skill
dispatcher** (`cli/src/skills/runner/dispatcher.ts`) can run it without
needing the bundled app — that's what makes Word installable as a hub
skill from the CLI.

When the skill is split out of the monorepo (spec 8.5), the Swift sources
will be `git mv`d into `backends/macos/Sources/DocBackend/`.
