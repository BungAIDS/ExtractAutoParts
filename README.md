# ExtractAutoParts

SolidWorks macro that extracts the equation-driven AUTO parts from the DAG folder
(`Z:\DAG\`) into a SolidWorks job folder. Run it, enter the job number, check the items
you want, click **Extract**.

## What the dialog shows

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
failed.

## Job folder

Located exactly like [Solidworks-Open](../Solidworks-Open): the macro probes
`GENERAL LINE`, `HD-PFD`, `HDX` and `AXIAL` under `Z:\Solidworks\Current\JOBS\` using the
same intermediate-folder rules (range folders, `NNXXXX` buckets, none for AXIAL) and uses
the first folder that exists for the job number.

## Install

1. In SolidWorks: **Tools > Macro > New...**, save as `ExtractAutoParts.swp`.
2. In the VBA editor: **File > Import File...** and import `ExtractAutoParts.bas`, then
   `ExtractForm.frm`. Delete the empty module the new macro started with.
3. Save. Run with **Tools > Macro > Run...** (entry point `main` in the
   `ExtractAutoParts` module).

If importing `ExtractForm.frm` gives an error, build the form manually instead:
**Insert > UserForm**, set its `(Name)` to `ExtractForm` in the Properties window, press
F7 and paste in everything below the `Attribute` lines of `ExtractForm.frm`. The form has
no designed controls - the entire dialog is built at run time - so an empty form is all
it needs.

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
