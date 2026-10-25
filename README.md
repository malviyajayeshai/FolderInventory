# FolderInventory

`FolderInventory.ps1` scans a Windows drive or folder and writes one self-contained HTML report: summary cards, a sortable/paginated file table, a lazy folder tree with sizes, combined filters (type, extension, modified date, minimum size, folder, search), a size-analysis tab, a skipped-and-errors tab, and CSV export. The report opens from disk with no internet, server or extra software.

The script only reads. It never deletes, moves or modifies files, and never changes permissions or ownership.

Requirements: Windows 10 or 11 with Windows PowerShell 5.1 (built in) or PowerShell 7+. The report works in current Chrome, Edge and Firefox.

---

## 1. Install

1. Save `FolderInventory.ps1` to your Downloads folder.
2. Open **Windows PowerShell** (Start menu → type `PowerShell`).
3. Files downloaded from the internet are blocked by default. Unblock this one and allow scripts for the current window only:

```powershell
Unblock-File "$HOME\Downloads\FolderInventory.ps1"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

The `-Scope Process` setting lasts only until you close that PowerShell window. Nothing is changed system-wide.

## 2. Scan D:\ and save the report in Downloads

```powershell
Unblock-File "$HOME\Downloads\FolderInventory.ps1"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

& "$HOME\Downloads\FolderInventory.ps1" `
  -RootPath "D:\" `
  -OutputHtml "$HOME\Downloads\FolderInventory.html" `
  -ThrottleLimit 4

Start-Process "$HOME\Downloads\FolderInventory.html"
```

To also include hidden and system items (for example `System Volume Information`, which is usually access-denied and will then be listed as skipped):

```powershell
& "$HOME\Downloads\FolderInventory.ps1" -RootPath "D:\" -OutputHtml "$HOME\Downloads\FolderInventory.html" -ThrottleLimit 4 -IncludeHidden -IncludeSystem
```

## 3. Scanning another drive or folder

Change `-RootPath`, and usually `-OutputHtml` so you don't overwrite the previous report:

```powershell
# Another drive
& "$HOME\Downloads\FolderInventory.ps1" -RootPath "E:\" -OutputHtml "$HOME\Downloads\Inventory-E.html"

# A folder whose name contains spaces
& "$HOME\Downloads\FolderInventory.ps1" -RootPath "C:\Users\$env:USERNAME\My Projects" -OutputHtml "$HOME\Downloads\Projects inventory.html"

# A network share (use 1-2 workers for network drives)
& "$HOME\Downloads\FolderInventory.ps1" -RootPath "\\server\share\team" -OutputHtml "$HOME\Downloads\Share.html" -ThrottleLimit 2

# The current folder, report written to the current folder as FolderInventory.html
cd "D:\Photos"
& "$HOME\Downloads\FolderInventory.ps1"
```

`D:`, `D:\` and paths with or without a trailing backslash all work. `D:` on its own is treated as the root of drive D.

## 4. Parameters

There is a single parameter set. Three parameters can also be given by position, in this order: `RootPath`, `OutputHtml`, `ThrottleLimit`. Switches are named only.

| Parameter | Position | Default | Meaning |
|---|---|---|---|
| `-RootPath` | 0 | current folder | Drive or folder to scan. Must exist and be a folder. |
| `-OutputHtml` | 1 | `.\FolderInventory.html` | Report path. Missing parent folders are created. If you pass an existing folder, `FolderInventory.html` is written inside it. Writability is checked **before** the scan starts. |
| `-ThrottleLimit` | 2 | `4` | Concurrent scanning workers, 1–64. Use 1–2 for spinning hard drives and network shares, 4–8 for SSDs. |
| `-IncludeHidden` | named | off | Include items with the Hidden attribute. |
| `-IncludeSystem` | named | off | Include items with the System attribute. An item that is both Hidden and System needs both switches. |
| `-FollowReparsePoints` | named | off | Descend into junctions and directory symbolic links (see section 7). |
| `-TimeoutMinutes` | named | `0` (none) | Stop after this many minutes and write a Partial report from what was collected. |

Positional example: `& .\FolderInventory.ps1 "E:\Photos" "$HOME\Desktop\Photos.html" 2`

**Exit codes** (in `$LASTEXITCODE`): `0` complete scan, report written · `2` partial scan, report written · `1` fatal error, no report.

## 5. Using the report

- **Summary cards**: total size, files, folders, largest file, most common extension, skipped folders, scan status (Complete / Partial). A Partial scan also shows a banner, and totals are labelled as lower bounds.
- **Files tab**: click a column header to sort (sizes and dates sort numerically); click again to reverse. Hover a size for the exact byte count. Choose 25–1000 rows per page.
- **Search** matches file names and folder paths, case-insensitively. Typing a path fragment like `photos\2020` also works.
- **Type**: category chips (Documents, Images, Videos, Audio, Archives, Executables, Code, Other, Extensionless files) and an **Extensions** picker. Multiple selections combine as "any of these types".
- **Modified**: last 24 hours / 7 / 30 / 90 days / last year, or a custom from–to range. Custom dates are inclusive (from 00:00:00.000 on the start date to 23:59:59.999 on the end date, local time).
- **Minimum size**: number plus unit (bytes, KB, MB, GB, TB; 1 KB = 1024 bytes). Compares real byte values.
- **Folder tree**: click a folder name to show only that folder and its subfolders; click ▸ to expand. *Show all files* returns to the full inventory. Sort by name or size. Folders are marked *empty*, *no files*, *Access denied*, *partial*. Partial totals are prefixed with ≥.
- All filters combine and none resets the others. Active filters appear as removable chips. **Clear all filters** resets everything except sort order.
- **Export CSV** exports **every** record matching the current filters (not just the visible page), in the current sort order. It's UTF-8 with BOM so Excel shows Unicode names correctly. Text cells starting with `=`, `+`, `-`, `@`, tab or CR are prefixed with `'` so spreadsheet programs don't run them as formulas.
- **Size analysis tab**: largest folders (total and direct files), largest files, distribution by category and extension (by size and by count). Click any entry to jump to it in the Files tab.
- **Skipped & errors tab**: every folder that could not be read, partly listed folders, file-level errors, with reason and full path; CSV export; plus hidden/system items excluded by options and links not followed.
- The footer shows a data check: file records, per-folder counts and root totals are re-verified in the browser when the page loads.

Times are shown in the local time zone of the computer viewing the report.

## 6. Extension → category mapping

Matching is by extension only, case-insensitive. The extension is the text after the last dot (`archive.tar.gz` → `gz`; `.gitignore` → `gitignore`).

| Category | Extensions |
|---|---|
| Documents | pdf doc docx docm dot dotx odt rtf txt md markdown tex wpd pages xls xlsx xlsm xlsb xlt xltx ods csv tsv ppt pptx pptm pps ppsx odp key numbers epub mobi azw azw3 one onepkg vsd vsdx pub xps oxps msg eml log nfo |
| Images | jpg jpeg jpe jfif png gif bmp dib tif tiff webp heic heif avif ico cur svg svgz raw cr2 cr3 nef nrw arw dng orf rw2 raf srw pef psd psb ai eps xcf jxl |
| Videos | mp4 m4v mkv avi mov wmv flv f4v webm mpg mpeg mpe m2v m2ts mts ts vob 3gp 3g2 ogv rm rmvb asf divx |
| Audio | mp3 wav flac aac m4a m4b ogg oga opus wma aiff aif aifc alac ape mid midi amr ac3 dts mka |
| Archives | zip 7z rar tar gz tgz bz2 tbz2 xz txz zst lz lzma cab arj z iso img vhd vhdx vmdk wim esd dmg jar war apk |
| Executables | exe msi msix msixbundle appx appxbundle com scr dll sys ocx drv cpl bat cmd vbs vbe wsf wsh ps1 psm1 psd1 lnk reg |
| Code | c h cpp cc cxx hpp hh cs vb fs java kt kts scala go rs py pyw rb php pl pm swift m mm js mjs cjs ts tsx jsx html htm css scss sass less json xml yaml yml toml ini cfg conf sql sh bash zsh r lua dart groovy gradle ipynb sln csproj vbproj vcxproj |
| Other | any other extension |
| No extension | files without an extension (shown as the *Extensionless files* chip) |

`ts` appears under Videos (MPEG transport stream) and not Code: the first category listing an extension wins. To change the mapping, edit `$script:CategoryExtensions` near the top of the script.

## 7. How it works

- **Scanning.** A bounded runspace pool (`-ThrottleLimit` workers; no background jobs) shares one work queue. The root is queued once. A worker takes a folder, lists only its immediate entries with `DirectoryInfo.EnumerateFileSystemInfos()` (no extra disk calls per file), sends back one result for that folder, and queues each subfolder exactly once. A countdown of outstanding folders tells the workers when everything is done. The main thread collects results, shows progress (`Write-Progress`, plus a console line every 30 s; never one line per file), and rejects any duplicate folder result, so a file can't be counted twice.
- **Totals.** After the scan, folder totals are computed in one bottom-up pass (deepest folders first), straight from the collected file records. Nothing is rescanned. The script checks that the root total equals the sum of all file records before writing the report.
- **Errors.** Each folder listing and each entry is wrapped individually. Access-denied, not-found (deleted during the scan), path-too-long and other I/O errors are classified and recorded with the path. A folder that can't be opened is *Skipped*. One that fails part-way is *Partially listed*, and the entries read before the failure are kept. Any of these, a worker failure or a time-out makes the status **Partial**, and ≥ marks totals on the affected branch. Unexpected script errors print the exception type, line number and stack trace.
- **Reparse points.** Directory junctions and symbolic links are **not followed** by default, so a scan can't loop or leave the selected root (for example `C:\Users\Name\Application Data` points back into the profile). They are listed in the report. `-FollowReparsePoints` follows them, with a nesting limit of 256 levels as a loop guard. Files that are links or cloud placeholders (OneDrive "online-only") are listed with the size Windows reports. Enumerating them does **not** download them.
- **Hidden/System.** Excluded by default and counted in the report. The root folder itself is always scanned, even though drive roots carry the Hidden and System attributes.
- **Long paths.** Folders with paths of 240+ characters are opened with the `\\?\` long-path prefix. If that still fails, the folder is recorded as skipped (`PathTooLong`) instead of being silently omitted.
- **Report safety.** Data is embedded as JSON inside a `<script type="application/json">` block, with `<`, `>`, `&`, quotes, control characters, U+2028/2029 and unpaired surrogates escaped as `\uXXXX`. The page renders every file name with `textContent`, never `innerHTML`. A Content-Security-Policy blocks all network access. The report is written to a temporary file and then renamed, so a failed run never leaves a half-written report.

## 8. Troubleshooting

| Symptom | Fix |
|---|---|
| `... cannot be loaded because running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in the same window, then run the script again. |
| `... is not digitally signed` / script won't run after download | `Unblock-File "$HOME\Downloads\FolderInventory.ps1"` |
| `A positional parameter cannot be found that accepts argument ...` | You passed more than three unnamed values, or a path with spaces without quotes. Quote paths and use the parameter names. |
| `Parameter set cannot be resolved ...` | Shouldn't happen (there's only one set). Check for typos in parameter names; PowerShell may be matching a partial name. |
| Running from `cmd.exe`: `powershell -File FolderInventory.ps1 -RootPath "D:\" -OutputHtml ...` scans the wrong path | In cmd.exe, `\"` escapes the quote. Write `-RootPath D:\` or `-RootPath "D:\\"`. The script also strips a stray trailing quote. Example: `powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\Downloads\FolderInventory.ps1" -RootPath D:\ -OutputHtml "%USERPROFILE%\Downloads\FolderInventory.html"` |
| Status is **Partial** | Open the *Skipped & errors* tab. Access-denied folders are normal for system areas. Run PowerShell **as Administrator** to read more of them; the script never changes permissions itself. |
| Scan is slow on a USB hard drive or NAS | Use `-ThrottleLimit 1` or `2`. More workers make a spinning disk seek more. |
| A network share hangs | Use `-TimeoutMinutes 30` (or similar) to get a Partial report instead of waiting indefinitely. |
| The report is very large / the browser is slow | Large drives produce large reports (roughly 50–60 bytes per file). Scan a sub-folder, or use Chrome/Edge, which handle large pages best. |
| `Cannot write to output folder` | Choose a folder you can write to, such as Downloads or Desktop. This is checked before scanning, so no time is lost. |
| Ctrl+C during the scan | The scan stops and its workers are cleaned up. No report is written for an interrupted run. Use `-TimeoutMinutes` if you need a time-boxed report. |

## 9. Test plan and results

### What was actually run

The script was developed and executed with **PowerShell 7.4.6 on Linux** (the only PowerShell available in the build environment). **It has not been run on Windows or in Windows PowerShell 5.1.** For 5.1 compatibility, the script was checked statically with PSScriptAnalyzer 1.23 compatibility rules (`PSUseCompatibleSyntax` for 3.0/5.1/7.0, and `PSUseCompatibleCommands` / `PSUseCompatibleTypes` against the Windows 10 / PowerShell 5.1 / .NET Framework 4.x profile), with **no findings**. The script is pure ASCII, so 5.1 reads it correctly without a BOM. Run the Windows procedure below before relying on it.

Test tree: 26 files and 21 folders, including different extensions, names with spaces, `[brackets]`, `(parentheses)`, apostrophes, Unicode (`Ünïcødé 日本語 файл.jpg`, `emoji 😀.png`), HTML/script-like names (`"><img src=x onerror=alert(1)>.txt`), CSV-formula names (`=HYPERLINK(1).csv`, `-minus @at +plus.log`), known sizes, two empty folders, 5-level nesting, duplicate names in different folders, extensionless files, a hidden file and folder, a directory symlink, and timestamps from 2015, 2020, 200/40/3 days ago and today.

The file list, sizes and every folder's six totals were compared with an independent Python walk of the same tree.

| # | Test | Result |
|---|---|---|
| 1 | Named parameters, root with spaces and trailing slash, output in a not-yet-existing folder with spaces | PASS: 24 files, 20 folders, 1,962,897 bytes, matches independent walk exactly; folder created; exit 0 |
| 2 | Default invocation `& .\FolderInventory.ps1` | PASS: scanned the current folder, wrote `.\FolderInventory.html` |
| 3 | Positional invocation `root output 1` | PASS: identical results |
| 4 | `pwsh -File` style invocation | PASS: exit 0 |
| 5 | Inaccessible folder (run as non-root user, folder mode 000), `-IncludeHidden` | PASS: folder listed as *Access denied*, parent and root marked partial (≥), status Partial, exit 2, all other 25 files present |
| 6 | Every file counted once, duplicates in different folders kept | PASS: no duplicate paths; `same.txt`/`readme` in two folders both present |
| 7 | Folder totals (direct/total size and files, direct/total subfolders) | PASS: all 20 folders match independent computation |
| 8 | ThrottleLimit 1, 4 and 8 give identical inventories | PASS |
| 9 | `-FollowReparsePoints` | PASS: link followed (22 folders, 28 files); default lists the link as not followed |
| 10 | Output path is an existing folder | PASS: writes `FolderInventory.html` inside it |
| 11 | Root missing / root is a file / ThrottleLimit 0 / extra positional argument | PASS: clear message; exit 1 (or PowerShell's own validation error before the script runs) |
| 12 | Stray quote in RootPath (the cmd.exe `"D:\"` problem) | PASS: path cleaned, scan correct |
| 13 | Time limit path (test copy with limit shortened to 3 s, 150k-file tree, 1 worker) | PASS: Partial report, 8,863 unscanned folders listed as `Timeout`, totals consistent, exit 2 |
| 14 | Scale: 150,000 files in 15,781 folders, 4 workers | PASS: file set and all folder totals identical to independent walk; 21.8 s scan, 28 s total, 8.5 MB report |
| 15 | Browser tests (headless Chromium, offline, `file://`): 57 checks | PASS: see below |

Browser checks (all passed in Chromium): loads offline with zero network requests; zero JS errors; no dialog fired and no `<img>` injected by the HTML-like file name, which shows literally; search by name, folder path, path across separator and case-insensitive Unicode; size/date/name sorting numeric and stable, record count unchanged; Videos → +last 90 days → +min 100 MB (0 results, empty message) → min 0.5 MB → +folder, other filters preserved, 4 active chips; clear all; extensionless chip and extension picker combined; inclusive custom date range; tree click filters to folder and descendants; expand/collapse all; empty-folder badges; tree sort by size; pagination; CSV BOM, row count equals filtered count, formula guard, quote escaping, Unicode preserved; analysis tab content and click-through; partial report shows banner, *Access denied* and partial badges, ≥ totals and the error row. With 150,000 files: load 0.5 s, name sort 0.25 s, search 0.3 s, *Expand all* capped at ~4,000 rows, CSV contains all 15,000 matches across pages.

**Not tested here:** Windows / Windows PowerShell 5.1 execution, NTFS junctions, Hidden/System attributes as Windows sets them, Windows access-denied folders, `\\?\` long paths, UNC shares, OneDrive placeholders, Edge and Firefox (Firefox wasn't installed in the environment).

### Manual test procedure for Windows (about 10 minutes)

Run in **Windows PowerShell 5.1** (`powershell.exe`, not `pwsh`):

```powershell
Unblock-File "$HOME\Downloads\FolderInventory.ps1"
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
$PSVersionTable.PSVersion          # expect 5.1.x

# 1. Build a test tree
$T = Join-Path $env:TEMP 'FI Test'
Remove-Item -LiteralPath $T -Recurse -Force -ErrorAction SilentlyContinue
function New-TestFile($rel, $bytes) { $p = Join-Path $T $rel; New-Item -ItemType Directory -Force -Path (Split-Path $p) | Out-Null; [IO.File]::WriteAllBytes($p, (New-Object byte[] $bytes)) }
New-TestFile 'docs\report final.pdf' 1000
New-TestFile 'docs\2023\Q1 [draft] (v2).docx' 5000
New-TestFile "docs\2023\it's mine.xlsx" 3000
New-TestFile ("media\" + [char]0x00DC + 'nicode ' + [char]0x65E5 + [char]0x672C + '.jpg') 4096
New-TestFile 'media\big movie.mp4' 1048576
New-TestFile 'dupes\a\readme' 10
New-TestFile 'dupes\b\readme' 20
New-TestFile 'weird\=HYPERLINK(1).csv' 6
New-Item -ItemType Directory -Force -Path "$T\empty1", "$T\nested\l1\l2\leaf" | Out-Null
(Get-Item -LiteralPath "$T\docs\report final.pdf").LastWriteTime = '2015-06-01 12:00'
New-TestFile 'hidden.txt' 9;  (Get-Item -LiteralPath "$T\hidden.txt").Attributes = 'Hidden'
New-TestFile 'locked\secret.txt' 100
icacls "$T\locked" /deny "${env:USERNAME}:(RX)" | Out-Null            # inaccessible folder
cmd /c mklink /J "$T\junction-to-docs" "$T\docs" | Out-Null           # junction

# 2. Named invocation, output path with spaces
& "$HOME\Downloads\FolderInventory.ps1" -RootPath "$T\" -OutputHtml "$env:TEMP\FI out\My Report.html" -ThrottleLimit 4
$LASTEXITCODE    # expect 2 (Partial, because of 'locked')
# Expect: 8 files, 1,061,708 bytes, 'locked' listed as AccessDenied, 1 hidden file excluded, 1 link not followed

# 3. Default invocation
Push-Location $T; & "$HOME\Downloads\FolderInventory.ps1"; Pop-Location
# Expect: $T\FolderInventory.html written

# 4. Independent cross-check of counts and bytes (excluding hidden, junction and locked folder)
Get-ChildItem -LiteralPath $T -Recurse -File -Attributes !Hidden -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -notlike "$T\junction-to-docs\*" -and $_.Name -ne 'FolderInventory.html' } |
  Measure-Object Length -Sum      # expect Count 8, Sum 1061708

# 5. Open and check the report
Start-Process "$env:TEMP\FI out\My Report.html"
#   - Status 'Partial scan', banner, 'locked' shows 'Access denied'; root total prefixed with >=
#   - Search 'readme' -> 2; Videos chip -> 1; Modified 'Last 30 days' + Documents -> 3 (report final.pdf is from 2015; .csv counts as a document)
#   - Custom range 2015-06-01..2015-06-01 -> 1; Minimum size 1 MB -> 1
#   - Click 'docs' in the tree -> 3 files; Sort by Size and Last modified both directions
#   - Export CSV -> open in Excel: Unicode name correct, '=HYPERLINK(1).csv' shown as text
#   - Disconnect from the network and reload: everything still works
#   - Repeat in Edge and Firefox

# 6. Clean up (restores access so the folder can be deleted)
icacls "$T\locked" /remove:d "${env:USERNAME}" | Out-Null
cmd /c rmdir "$T\junction-to-docs"
Remove-Item -LiteralPath $T, "$env:TEMP\FI out" -Recurse -Force
```

## 10. Known limitations

- **Not yet run on Windows** (see section 9). The code uses only APIs present in Windows PowerShell 5.1 / .NET Framework 4.x, but behaviour on real NTFS volumes (junctions, ACLs, long paths) still has to be confirmed with the procedure above.
- **Memory.** The scanner keeps every file record in memory until the report is written: roughly 300 bytes per file in PowerShell (about 300 MB for a million files). The browser needs a similar amount to open a million-file report. Above a few million files, scan sub-folders separately.
- **Throughput.** Per-file processing runs in PowerShell, so expect tens of thousands of files per second at best. Disk speed and `-ThrottleLimit` usually matter more.
- **Long paths.** Paths of 260+ characters depend on .NET Framework support for the `\\?\` prefix. If it isn't available, those folders are reported as skipped (`PathTooLong`) and the scan is marked Partial, never silently omitted.
- **Time-outs** are checked between folders. A single folder stuck on an unresponsive network share is forcibly stopped 60 seconds after the limit and listed as not scanned.
- **Hard links** are listed once per path, so a hard-linked file counts toward each folder that contains it, as in Explorer. Sizes are logical file sizes, not size on disk, so compression, sparse files and cluster slack aren't reflected.
- **Interrupted runs** (Ctrl+C, closing the window) produce no report.
- **Snapshot.** Files that change during the scan are reported as they were when their folder was listed. A folder deleted mid-scan is reported as `NotFound`.
