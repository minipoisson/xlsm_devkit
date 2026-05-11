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
