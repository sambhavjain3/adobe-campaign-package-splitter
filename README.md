# Adobe Campaign Package Splitter

Splits a large Adobe Campaign (v7 / v8) package XML into smaller packages that import cleanly, one after another. It keeps the XML structure valid and handles dependencies between objects, so no part duplicates or overwrites another.

It runs on any Windows 10/11 or Windows Server 2016+ PC with nothing to install. It uses the Windows PowerShell 5.1 and .NET Framework that come with Windows.

## Files

| File | Purpose |
|---|---|
| `Split-ACPackage.bat` | Launcher. Double-click it. |
| `ACPackageSplitter.ps1` | The utility. Keep it in the same folder as the `.bat`. |

## Quick start

1. Download the repo as a ZIP (**Code > Download ZIP**) and extract it.
2. Double-click `Split-ACPackage.bat`, or drag a package XML onto it.
3. Pick the package and set **Max size per part (MB)** (default 10). You can also set a maximum number of entities per part.
4. Click **Analyze** for a report only, or **Split** to write the parts.
5. Open the `<package name>_split` folder created next to the package and read `00_IMPORT_ORDER.txt`.
6. On the **target** instance, paste `check_target_clashes.js` into a JavaScript activity in a test workflow. Run it and fix every `CLASH` it logs.
7. Import the parts one at a time, in the listed order (**Tools > Advanced > Import package**). Check each import log before starting the next part.

## What it guarantees

- **Valid packages.** Every part keeps the same `<package>` header (author, build number, version) and the original `<entities schema="...">` wrappers. Parts contain only whole top-level objects, copied node for node from the source.
- **Written exactly once.** Every top-level object lands in one part only.
- **Linked objects stay together.** This covers a delivery activity that points to a delivery owned by another campaign, or the same workflow defined under two campaigns.
- **Dependency order.** If a linked group is bigger than the size limit, it is spread over consecutive parts. What is referenced always comes first, and the manifest lists which parts depend on which.
- **Oversized objects.** An object bigger than the limit cannot be cut. It gets a part of its own and a warning.
- **Self-check.** After writing, every part is re-read and compared with the source. Do not import unless it reports **Verification: PASSED**.

## Output

| File | Contents |
|---|---|
| `00_IMPORT_ORDER.txt` | Import order, which parts each part depends on, contents, warnings, duplicate definitions, links by numeric record ID |
| `<name>_Part-NN_of_MM.xml` | The packages |
| `split_map.csv` | Which object went into which part |
| `external_dependencies.csv` | Objects the package references but does not contain (folders, operators, typologies, mappings, routings…). They must already exist on the target. |
| `check_target_clashes.js` | Pre-import check to run on the target. It lists workflows, deliveries and campaigns that already exist there under the same internal name but belong to something else. That is the cause of `duplicate key value violates unique constraint "xtkworkflow_id"`. |

## Analyze several files

Select several package files and click **Analyze** to check a set you already split by hand. The report covers:

- objects that appear in more than one file
- links between files
- files imported in the wrong order

## Command line

```bat
Split-ACPackage.bat -InputFile "C:\pkg\big.xml" -MaxSizeMB 10 [-MaxEntities 50] [-OutputFolder "C:\out"] [-NoTargetCheck]
Split-ACPackage.bat -Analyze -InputFile "C:\pkg\part1.xml" "C:\pkg\part2.xml"
```

Exit codes:

- `0`: OK
- `2`: verification failed
- `1`: error

## Troubleshooting

- **Nothing happens when you double-click.** Run `Split-ACPackage.bat -Gui` from a Command Prompt to see the error.
- **Script blocked after download.** Files downloaded from the internet or received by email can be blocked. Right-click `ACPackageSplitter.ps1`, open **Properties** and tick **Unblock**.
- **Locked-down corporate PCs.** PCs that enforce PowerShell *Constrained Language Mode* or AppLocker can block scripts entirely. Use another PC or ask IT for an exception.

## Notes

- The source file is never modified. Output files are UTF-8.
- Splitting several files at once splits each one on its own. Links *between* input files are reported by **Analyze** but not fixed.
- Links by numeric record ID (for example `<delivery id="974442">` inside a query filter) point to IDs on the source instance. They are listed so you can re-link them after import.
