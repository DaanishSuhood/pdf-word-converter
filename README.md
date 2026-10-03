# PDF – Word Converter

Batch-converts between Word and PDF in both directions, using Microsoft Word over COM. Drop a mixed
pile of files on the window and it works out which way each one needs to go.

Script: [`pdf-word-converter.ahk`](pdf-word-converter.ahk)

---

## Requirements

| Requirement | Notes |
| --- | --- |
| [AutoHotkey v2.0+](https://www.autohotkey.com/) | Declares `#Requires AutoHotkey v2.0` |
| Microsoft Word (desktop) | Does all the conversion. PDF → Word needs Word 2013 or newer |

## Getting started

1. Run `pdf-word-converter.ahk`.
2. Drag files onto the window, or click **Browse Files**.
3. Optionally set the Word output format and an output folder.
4. Click **Convert All**.

## Direction is automatic

There is one convert button, not two. Each file's extension decides its direction:

| Input | Output |
| --- | --- |
| `.doc`, `.docx` | `.pdf` |
| `.pdf` | `.docx` (or `.doc`) |

Anything else is rejected as it is added, with the reason shown in the status line.

## The window

| Control | Purpose |
| --- | --- |
| Drop zone | Drag and drop target. Dropping anywhere on the window works — the bordered box is a visual cue |
| **Browse Files** | Multi-select picker for `*.docx; *.doc; *.pdf` |
| File list | Filename, detected type (`Word` / `PDF`), and status (`Pending` → `Converting…` → `✓ Completed` / `✗ Failed`) |
| **Remove** / **Clear All** | Drop the selected row, or empty the list |
| **Word Output Format** | `DOCX` (default) or `DOC` — applies to PDF → Word only |
| **Output Folder** | Blank means *same folder as source*. The `📁` button opens a folder picker |
| **Convert All** | Runs the batch |
| Progress bar + status line | Per-file position and a `Converting n/total: <name>` label |

Duplicate paths are ignored, so dropping the same file twice adds one row.

A summary message box at the end reports success and failure counts and where the output went.

## How it works

One hidden Word instance (`Visible := false`, `DisplayAlerts := 0`) handles the whole batch; if Word
cannot be created the run stops immediately with an explanatory dialog.

Each file is opened with `Documents.Open` and re-saved with `SaveAs2` using a Word format constant:

| Direction | Constant | Meaning |
| --- | --- | --- |
| Word → PDF | `17` | `wdFormatPDF` |
| PDF → `.docx` | `16` | `wdFormatDocumentDefault` |
| PDF → `.doc` | `0` | `wdFormatDocument` |

Failures are caught per file and marked `✗ Failed` in the list without stopping the batch. Word is
quit at the end.

## Limitations

- **It quits Word when it finishes**, so close any Word documents you have open before converting —
  the `wordApp.Quit()` can take your session with it.
- **PDF → Word quality depends entirely on Word's PDF reflow.** Word rebuilds the document from
  scratch; complex layouts, columns, and tables often come out rearranged, and a scanned PDF (no text
  layer) produces a page of nothing. Word → PDF is reliable; treat the reverse direction as a
  starting point for editing, not a faithful copy.
- Existing output files are overwritten without asking.
- Failures are reported as `✗ Failed` only — the underlying COM error message is not surfaced
  anywhere.
- The GUI is unresponsive during the batch and there is no cancel button.
- The window is a fixed 600×670 layout with absolute coordinates; resizing is enabled but controls do
  not move.
