# mailOpt

Outlook VBA framework that processes mail on startup and when new messages arrive. Routes senders into folders from sort text files, handles meeting responses/requests, and includes a form to add sort rules from the open mail.

## Architecture

| Module | Role |
|---|---|
| `ThisOutlookSession` | `Application_Startup`, `Application_NewMailEx`, `Application_Quit` |
| `MailProcessor` | Startup scan, on-demand full-mailbox sort, per-item pipeline |
| `MailRules` | Rule evaluation, sort-file load/match/upsert, Add Sort Rule macro |
| `FolderHelpers` | Resolve / create folders; list child folders for the UI |
| `frmAddSortRule` | UserForm: map current sender → parent folder → destination |

Flow:

1. Outlook opens → `MailProcessor.Initialize` → scan unread items in all mail folders
2. New mail → `ProcessNewMail` → `EvaluateItem`
3. Meetings first, then each sort file (conversation **root** sender via conversation table, then **current** sender; conversation failures never skip current-sender matching)
4. On demand → `SortAllEmails` or `AddSortRuleFromCurrentMail`

## Install in Outlook

1. Open Outlook → `Alt+F11`.
2. Enable macros under **Trust Center → Macro Settings**.
3. Import from `src/`: `MailProcessor.bas`, `MailRules.bas`, `FolderHelpers.bas`.
4. Import `frmAddSortRule.frm` (remove any old copy first). Forms build controls in code — no `.frx` needed.
5. Set `SORT_RULES_FOLDER` in `MailRules.bas` to your sort-files folder.
6. Copy `sortComms`, `sortTickets`, `sortProductLines` (tab- or comma-delimited: `Email`, `Destination`, `Name`) into that folder.
7. Paste `ThisOutlookSession.cls` body (from `Option Explicit` down) into **Microsoft Outlook Objects → ThisOutlookSession**.
8. Save and restart Outlook.

If you previously imported `SortRules.bas` or `SortRuleUI.bas`, remove those modules — they are now part of `MailRules.bas`.

## Macros

**`SortAllEmails`** — reload sort files and process every item in every folder (`Alt+F8`, Immediate Window `MailProcessor.SortAllEmails`, or Quick Access Toolbar).

**`AddSortRuleFromCurrentMail`** — with a message open or selected:

1. Confirm sender email / display name
2. Choose parent folder (`BAE Comms`, `Tickets`, `Program Groups`)
3. Pick or type a destination (prompted to create if missing)
4. Save upserts the matching sort file and reloads rules

```vba
MailRules.AddSortRuleFromCurrentMail
```
