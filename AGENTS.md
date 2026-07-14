# VBA Coding Guidelines for This Project

## CRITICAL: Do NOT use legacy VBA file/directory functions

This project runs in a **UNC path environment** (`\\server\share\...`).  
Legacy VBA functions listed below are thin wrappers around old Windows APIs and are **unreliable under UNC paths**. They must never be used in generated or suggested code.

Additionally, `Dir` retains state in a **module-level global variable**, making it unsafe for any loop that may re-enter or nest calls — a common source of silent data-loss bugs.

**Always use `Scripting.FileSystemObject` (FSO) instead.**

---

### Banned functions and their FSO replacements

#### File operations

| Banned | Use instead |
|--------|-------------|
| `Kill path` | `fso.DeleteFile path, True` |
| `Name src As dst` | `fso.MoveFile src, dst` (rename) / `fso.MoveFile src, dstFolder` (move) |
| `FileCopy src, dst` | `fso.CopyFile src, dst` |
| `FileAttr(fn, 1)` | `fso.GetFile(path).Attributes` |
| `GetAttr(path)` | `fso.GetFile(path).Attributes` |
| `SetAttr path, attr` | `fso.GetFile(path).Attributes = attr` |
| `FileDateTime(path)` | `fso.GetFile(path).DateLastModified` |
| `FileLen(path)` | `fso.GetFile(path).Size` |

#### Directory (folder) operations

| Banned | Use instead |
|--------|-------------|
| `MkDir path` | `fso.CreateFolder path` |
| `RmDir path` | `fso.DeleteFolder path, True` |
| `ChDir path` | (avoid; work with absolute paths only) |
| `ChDrive drv` | (avoid; work with absolute paths only) |
| `CurDir` / `CurDir(drv)` | (avoid; use explicit absolute paths) |

#### File search and sequential I/O

| Banned | Use instead |
|--------|-------------|
| `Dir(pattern)` / `Dir` loop | `fso.GetFolder(path).Files` / `.SubFolders` collections |
| `FreeFile` / `Open` / `Close` | `fso.OpenTextFile` / `fso.CreateTextFile` |
| `Input`, `Line Input`, `InputB` | `TextStream.ReadLine` / `.ReadAll` |
| `Print #n`, `Write #n` | `TextStream.WriteLine` / `.Write` |

---

### Standard FSO setup pattern

```vba
Dim fso As Object   ' Scripting.FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
```

Prefer late binding (`CreateObject`) so the code does not require a reference added to every workbook.

---

### Why this matters

- **UNC paths** (`\\server\share\folder`): The legacy functions resolve paths through the Windows CRT, which may silently fall back to the current drive/directory or fail entirely on UNC targets. FSO delegates to the Shell infrastructure, which handles UNC correctly.
- **`Dir` global state**: A single module-level iteration counter is shared across all callers. Any re-entrant call (callback, event, nested loop) resets it and corrupts the outer loop — with no error raised.
- **Error handling**: FSO raises proper COM errors that `Err.Number` / `On Error` can catch. The legacy functions often fail silently or return ambiguous empty strings.

---

## Source files

### ASCII-only VBA source

- `.bas` / `.frm` files must be **ASCII only**: write em-dashes as `--`, no Japanese
  comments, introduce no new non-ASCII characters.
- Pre-existing em-dashes in `xlsm_devkit.bas` are out of scope -- leave them as they are,
  but do not add more.

### User-facing strings (i18n)

- Any string shown to the user through `MsgBox` must be resolved via `t()` (lookup) and,
  when it contains placeholders, formatted with `Fmt`. Log / `Debug.Print` messages may
  stay English.

### Language resource files

- `lang/*.ini` are **UTF-8 without BOM**. A new key must be added to **all 27** language
  files. (`ReadUTF8` / ADODB strips any leading BOM on read; keep the files BOM-less for
  consistency with the rest of the repo, which writes BOM-less UTF-8 everywhere.)

### Binary form resources

- `src/*.frx` are binary UserForm resources. If an unintended Excel operation touches them,
  restore with `git restore src/*.frx` rather than committing the churn.

---

## Git commit

- Stage all modified and untracked files that are **not excluded by `.gitignore`**.
- Commit with a concise message to the local repository.
- Do **not** stage files listed in `.gitignore`.

## Pushing to GitHub

- Push to `origin main` **only when explicitly asked by the user**.
- Do not push automatically after every commit.

## Release management

When a released source module has changed since the last release -- the core
`xlsm_devkit.bas` **or** an optional harness module (`src/devkit_*.bas`) -- create a new
GitHub release:

1. **Determine the version** using Semantic Versioning based on the `[Unreleased]` section of `CHANGELOG.md`:
   - New feature added → bump MINOR (e.g. `1.0.x` → `1.1.0`)
   - Bug fix or non-breaking change only → bump PATCH (e.g. `1.1.0` → `1.1.1`)
   - Breaking change → bump MAJOR (e.g. `1.1.0` → `2.0.0`)

2. **Update `CHANGELOG.md`**: replace `## [Unreleased]` with a new `## [Unreleased]` section (empty) followed by `## [x.y.z] - YYYY-MM-DD`, then commit and push.

3. **Tag and release**:
   ```
   git tag vX.Y.Z
   git push origin main
   git push origin vX.Y.Z
   gh release create vX.Y.Z \
     --title "vX.Y.Z" \
     --notes "<content of the new CHANGELOG section>"
   ```

4. **Do not attach individual `.bas` files.** The modules are interdependent (the optional
   `devkit_*` modules are useless without the core, and vice versa), so a partial asset set
   is misleading. Rely on GitHub's auto-generated **Source code (zip / tar.gz)** archives,
   which contain the whole repository at the tag; users extract the module(s) they need.

---

## Session load token

This file is the canonical rule set and must be in context every session (Claude Code
auto-loads it via `CLAUDE.md`'s `@AGENTS.md` import; other tools read it directly). The
line below is a load canary: the SESSION_BOOTSTRAP.md startup handshake asks the assistant
to quote it. If the assistant cannot quote this token from context, AGENTS.md was not
loaded -- inspect with `/memory` (Claude Code) or restart the session. Keep this as the
last line of the file so a correct quote also evidences a full load.

`AGENTS-LOAD-TOKEN: af39c7`
