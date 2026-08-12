# mailOpt

Outlook VBA framework that processes mail when Outlook starts and when new messages arrive. Built to filter mail into folders and organize meeting-related items automatically.

## Architecture

| Module | Role |
|---|---|
| `ThisOutlookSession` | Wires `Application_Startup`, `Application_NewMailEx`, and `Application_Quit` |
| `MailProcessor` | Startup scan, on-demand full-mailbox run, per-item pipeline |
| `MailRules` | Rule evaluation (meeting responses/requests, move, flag) |
| `SortRules` | File-based sender routing from sort text files (load + upsert) |
| `FolderHelpers` | Resolve / create folders; list child folders for the UI |
| `SortRuleUI` | Macro entry point for the Add Sort Rule form |
| `frmAddSortRule` | UserForm: map current sender → sort file → destination folder |

Flow:

1. Outlook opens → `Application_Startup` → `MailProcessor.Initialize` → scan unread items in all mail folders
2. New mail arrives → `Application_NewMailEx` → resolve each EntryID → `ProcessOutlookItem`
3. `MailRules.EvaluateItem` returns an action → move / mark read & move / flag / none
4. On demand → `MailProcessor.SortAllEmails` processes every item in every mail folder
5. On demand → `SortRuleUI.AddSortRuleFromCurrentMail` updates a sort text file from the open mail

## Install in Outlook

1. Open Outlook → press `Alt+F11` to open the VBA editor.
2. Enable macros: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → choose *Notifications for all macros* or *Enable all macros* (prefer signed macros in production).
3. Import the standard modules from `src/`:
   - **File → Import File…** → select `MailProcessor.bas`, `MailRules.bas`, `SortRules.bas`, `FolderHelpers.bas`, `SortRuleUI.bas`
4. Import the UserForm:
   - **File → Import File…** → select `frmAddSortRule.frm`
   - If a broken `frmAddSortRule` already exists in the project, remove it first (right-click → Remove), then import again.
   - The form builds its controls in code (`cboSortFile` is the sort-file dropdown). No `.frx` designer file is required.
5. Set `SORT_RULES_FOLDER` in `SortRules.bas` to the folder containing your sort text files (`sortComms`, `sortProductLines`, `sortTickets`).
6. Copy your sort files into that folder (tab- or comma-delimited: `Email`, `Destination`, `Name`).
7. Open **Microsoft Outlook Objects → ThisOutlookSession** and paste the body from `src/ThisOutlookSession.cls` (skip the `VERSION` / `Attribute` header lines — paste from `Option Explicit` downward).
8. Save the VBA project (`Ctrl+S`). Restart Outlook.

## Adding rules

Edit matchers / `EvaluateItem` in `MailRules.bas`:

- `MatchesDeleteRule` — return `True` to soft-delete (Deleted Items)
- `MatchesMoveRule` — set `folderPath` and return `True`
- `MatchesMarkReadRule` — return `True` to mark unread mail as read
- Meeting responses and other item types can be handled in `EvaluateItem`

Helpers: `SenderUserNameIs`, `SenderStartsWith`, `SenderContains`, `SubjectContains`, `SenderAddress`.

## Sort all mail on command

Run the **`SortAllEmails`** macro to apply every rule to all mail (read and unread) in every folder. It also reloads your sort text files first.

**From the macro dialog:** press `Alt+F8`, choose `SortAllEmails`, click **Run**.

**From the Immediate Window** (`Ctrl+G`):

```vba
MailProcessor.SortAllEmails
```

**Assign to a button:** **File → Options → Quick Access Toolbar** → choose *Macros*, add `SortAllEmails`, then click the button whenever you want a full sort.

`RunAllFilters` and `RunInboxFilters` call the same macro.

## Add a sender to a sort file

Use the **`AddSortRuleFromCurrentMail`** macro while a message is open or selected:

1. Open (or select) the email whose sender you want to route.
2. Press `Alt+F8`, choose `AddSortRuleFromCurrentMail`, click **Run**.
3. Confirm/edit the sender email and display name.
4. Choose the parent folder (`BAE Comms`, `Tickets`, or `Program Groups`).
5. Pick a destination from the dropdown of folders under that parent. You can also type a new folder name; if it does not exist you will be asked whether to create it.
6. Click **Save** — the matching sort text file is updated and rules are reloaded.

**Immediate Window:**

```vba
SortRuleUI.AddSortRuleFromCurrentMail
```

**Quick Access Toolbar:** add the `AddSortRuleFromCurrentMail` macro the same way as `SortAllEmails`.
