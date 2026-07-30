# mailOpt

Outlook VBA framework that processes mail when Outlook starts and when new messages arrive. Built to filter mail into folders, mark unimportant messages as read, and delete noise — rules start as safe stubs so nothing is changed until you add conditions.

## Architecture

| Module | Role |
|---|---|
| `ThisOutlookSession` | Wires `Application_Startup`, `Application_NewMailEx`, and `Application_Quit` |
| `MailProcessor` | Startup scan, on-demand inbox run, per-item pipeline |
| `MailRules` | Rule evaluation (meeting responses, delete → move → mark read / flag) |
| `FolderHelpers` | Resolve / create folders under the Inbox |
| `Logger` | Writes to the Immediate Window (`Ctrl+G`); optional file log |

Flow:

1. Outlook opens → `Application_Startup` → `MailProcessor.Initialize` → scan unread Inbox items
2. New mail arrives → `Application_NewMailEx` → resolve each EntryID → `ProcessOutlookItem`
3. `MailRules.EvaluateItem` returns an action → move / mark read / delete / flag / none
4. On demand → `MailProcessor.RunInboxFilters` processes every Inbox item

## Install in Outlook

1. Open Outlook → press `Alt+F11` to open the VBA editor.
2. Enable macros: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → choose *Notifications for all macros* or *Enable all macros* (prefer signed macros in production).
3. Import the standard modules from `src/`:
   - **File → Import File…** → select `MailProcessor.bas`, `MailRules.bas`, `FolderHelpers.bas`, `Logger.bas`
4. Open **Microsoft Outlook Objects → ThisOutlookSession** and paste the body from `src/ThisOutlookSession.cls` (skip the `VERSION` / `Attribute` header lines — paste from `Option Explicit` downward).
5. Save the VBA project (`Ctrl+S`). Restart Outlook.
6. Confirm in the Immediate Window (`Ctrl+G`) that you see: `Outlook started — mail framework initialized.`

## Adding rules

Edit matchers / `EvaluateItem` in `MailRules.bas`:

- `MatchesDeleteRule` — return `True` to soft-delete (Deleted Items)
- `MatchesMoveRule` — set `folderPath` and return `True`
- `MatchesMarkReadRule` — return `True` to mark unread mail as read
- Meeting responses and other item types can be handled in `EvaluateItem`

Helpers: `SenderStartsWith`, `SenderContains`, `SubjectContains`, `SenderAddress`.

## Run filters on demand

From the Immediate Window (`Ctrl+G`):

```vba
MailProcessor.RunInboxFilters
```

Or call that same public sub from a Quick Access Toolbar / ribbon macro button.

## Next steps

- Add more conditions in `MailRules`
- Optionally expand the startup scan beyond unread mail
- Optionally enable file logging in `Logger` (`ENABLE_FILE_LOG`)
