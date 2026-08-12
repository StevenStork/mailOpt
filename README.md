# mailOpt

Outlook VBA framework that processes mail when Outlook starts and when new messages arrive. Built to filter mail into folders, mark unimportant messages as read, and delete noise — rules start as safe stubs so nothing is changed until you add conditions.

## Architecture

| Module | Role |
|---|---|
| `ThisOutlookSession` | Wires `Application_Startup`, `Application_NewMailEx`, and `Application_Quit` |
| `MailProcessor` | Startup scan, on-demand inbox run, per-item pipeline |
| `MailRules` | Rule evaluation (meeting responses, delete → move → mark read / flag) |
| `SortRules` | File-based sender routing from sort text files |
| `FolderHelpers` | Resolve / create folders under the Inbox |

Flow:

1. Outlook opens → `Application_Startup` → `MailProcessor.Initialize` → scan unread Inbox items
2. New mail arrives → `Application_NewMailEx` → resolve each EntryID → `ProcessOutlookItem`
3. `MailRules.EvaluateItem` returns an action → move / mark read / delete / flag / none
4. On demand → `MailProcessor.RunInboxFilters` processes every Inbox item

## Install in Outlook

1. Open Outlook → press `Alt+F11` to open the VBA editor.
2. Enable macros: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → choose *Notifications for all macros* or *Enable all macros* (prefer signed macros in production).
3. Import the standard modules from `src/`:
   - **File → Import File…** → select `MailProcessor.bas`, `MailRules.bas`, `SortRules.bas`, `FolderHelpers.bas`
4. Set `SORT_RULES_FOLDER` in `SortRules.bas` to the folder containing your sort text files (`sortComms`, `sortProductLines`, `sortTickets`).
5. Copy your sort files into that folder (tab- or comma-delimited: `Email`, `Destination`, `Name`).
6. Open **Microsoft Outlook Objects → ThisOutlookSession** and paste the body from `src/ThisOutlookSession.cls` (skip the `VERSION` / `Attribute` header lines — paste from `Option Explicit` downward).
7. Save the VBA project (`Ctrl+S`). Restart Outlook.

## Adding rules

Edit matchers / `EvaluateItem` in `MailRules.bas`:

- `MatchesDeleteRule` — return `True` to soft-delete (Deleted Items)
- `MatchesMoveRule` — set `folderPath` and return `True`
- `MatchesMarkReadRule` — return `True` to mark unread mail as read
- Meeting responses and other item types can be handled in `EvaluateItem`

Helpers: `SenderUserNameIs`, `SenderStartsWith`, `SenderContains`, `SubjectContains`, `SenderAddress`.

## Run filters on demand

From the Immediate Window (`Ctrl+G`):

```vba
MailProcessor.RunInboxFilters
```

Or call that same public sub from a Quick Access Toolbar / ribbon macro button.

## Next steps

- Add more conditions in `MailRules`
- Optionally expand the startup scan beyond unread mail
