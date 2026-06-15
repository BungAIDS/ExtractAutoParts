# ExtractAutoParts

SolidWorks macro that extracts the equation-driven AUTO parts from the DAG folder
(`Z:\DAG\`) into a SolidWorks job folder. Run it, confirm the order number, check the
items you want, click **Extract**.

## What the dialog shows

* **Order #** - if a job is open in SolidWorks, the macro reads its order number from the
  job folder the active document sits in (falling back to the file name) and offers it as
  **Use open order: \<number\>** (selected by default); pick **Other order #** to type a
  different one. With nothing open, it falls back to a plain `Order #:` box.
* **Parts** - one checkbox per `AUTO *` folder under `Z:\DAG\` (each holds a single zip),
  labeled with the folder name minus the leading `AUTO `. Folders without a zip are not
  listed.
* **Models** - one checkbox per zip inside `AUTO MODELS`, labeled with the text in
  parentheses (`EXTRACT ME!! (S&BG ONLY).zip` → `S&BG ONLY`). The all-in-one zip is shown
  as **A8 + Accessories**; checking it unchecks and disables the other model zips, since
  it already contains everything they do.

The checkboxes are built by scanning `Z:\DAG\` when the dialog opens, so adding or
renaming an AUTO folder (or a zip in `AUTO MODELS`) needs no code change.

## What extraction does

For every checked item:

1. The zip is extracted into its own subfolder of the job folder so nothing gets mixed
   up, named after the source folder with the leading `AUTO ` removed - e.g.
   `AUTO INLET IVC LINKAGE` → `<job folder>\INLET IVC LINKAGE\`. All `AUTO MODELS` zips
   extract into the same `<job folder>\MODELS\` (the other model zips are subsets of the
   A8 + Accessories zip, so they share one folder).
2. The drawing inside the extraction (`.SLDDRW` / `.DWG`, at most one per zip) is renamed
   to `<job>-<xx>`, where `<xx>` is the suffix of the original name: for job `512345`,
   `412345-63.SLDDRW` becomes `512345-63.SLDDRW`. Drawings from the A8 + Accessories zip
   are left untouched.

A summary box then lists everything extracted and renamed, plus anything skipped or
failed, and the job folder opens in Windows Explorer so the new files are right there.

## Job folder

Located exactly like [Solidworks-Open](../Solidworks-Open): the macro probes
`GENERAL LINE`, `HD-PFD`, `HDX` and `AXIAL` under `Z:\Solidworks\Current\JOBS\` using the
same intermediate-folder rules (range folders, `NNXXXX` buckets, none for AXIAL) and uses
the first folder that exists for the job number.

## Install

Both files are plain code, ready to paste as-is (no headers to trim).

1. In SolidWorks: **Tools > Macro > New...**, save as `ExtractAutoParts.swp`.
2. Paste the contents of `ExtractAutoParts.bas` into the module the new macro started
   with (or **File > Import File...** it and delete the starter module).
3. **Insert > UserForm**, set its `(Name)` to `ExtractForm` in the Properties window
   (F4), press F7 for its code window and paste in the entire contents of
   `ExtractForm.frm`. Leave the form itself empty - the whole dialog is built at run
   time.
4. Save. Run with **Tools > Macro > Run...** (entry point `main` in the
   `ExtractAutoParts` module).

## Notes

* Paths are constants at the top of `ExtractAutoParts.bas`: `DAG_ROOT` (`Z:\DAG\`) and
  `SW_ROOT` (`Z:\Solidworks\Current\JOBS\`). Change them there if anything moves.
* The all-in-one model zip (shown as **A8 + Accessories**) is recognized by `daddy`
  appearing anywhere in its file name.
* If the destination subfolder already exists (e.g. from an earlier run), the macro asks
  before touching it: **Yes** extracts into a new `NAME (2)` folder (then `(3)`, and so
  on), **No** extracts into the existing folder overwriting its files, **Cancel** skips
  that item. The `AUTO MODELS` zips picked in one run still share a single folder, so
  the prompt appears at most once per name.
* If you choose to reuse an existing folder and the renamed drawing name is already in
  it, the freshly extracted file keeps its original name and the summary says so.
* Unzipping uses Windows PowerShell's `Expand-Archive` (built into Windows 10/11).
