Option Explicit

' ExtractAutoParts.bas - SolidWorks VBA macro
'
' Copies the equation-driven AUTO parts out of the DAG folder into a
' SolidWorks job folder. ExtractForm lists every "AUTO *" folder under
' DAG_ROOT (one zip each) plus the individual zips inside AUTO MODELS;
' the user enters a job number and checks what to extract. Each checked
' item is unzipped into its own subfolder of the job folder, named after
' the source folder with the leading "AUTO " removed; every AUTO MODELS
' zip lands in MODELS. The drawing inside each extraction is renamed to
' <jobNum>-<xx>, keeping the suffix of the original file name, except
' drawings from the "daddy" zip, which are left as they are.
'
' The job folder is located exactly like Solidworks-Open / Pack-n-Go:
' probing each job type folder under SW_ROOT.

Public Const DAG_ROOT As String = "Z:\DAG\"
Public Const MODELS_FOLDER As String = "AUTO MODELS"  ' one checkbox per zip inside
Public Const MODELS_DEST As String = "MODELS"         ' job subfolder for AUTO MODELS zips
Public Const SW_ROOT As String = "Z:\Solidworks\Current\JOBS\"

' ---------------------------------------------------------------------
' Job folder lookup (same hierarchy as Solidworks-Open)
' ---------------------------------------------------------------------

' Range folder name based on first 3 digits, in groups of 5.
' Special case: 401-405 is rolled into "400-405".
Private Function ComputeRangeFolder(jobNum As String) As String
    Dim prefix As Long: prefix = CLng(Left$(jobNum, 3))
    Dim n As Long:      n = -Int(-prefix / 5)               ' ceil(prefix / 5)
    Dim startN As Long: startN = 5 * (n - 1) + 1
    Dim endN As Long:   endN = 5 * n
    If startN = 401 And endN = 405 Then
        ComputeRangeFolder = "400-405"
    Else
        ComputeRangeFolder = startN & "-" & endN
    End If
End Function

' SolidWorks intermediate: HD-PFD lives in a "NNXXXX" bucket keyed on the
' first two digits of the job number (40XXXX, 41XXXX, ... 49XXXX);
' AXIAL has no intermediate at all (jobs sit directly under AXIAL\);
' everyone else uses a range folder on the first 3 digits.
Private Function ComputeSwIntermediate(swType As String, jobNum As String) As String
    Select Case UCase$(swType)
        Case "HD-PFD": ComputeSwIntermediate = Left$(jobNum, 2) & "XXXX"
        Case "AXIAL":  ComputeSwIntermediate = ""
        Case Else:     ComputeSwIntermediate = ComputeRangeFolder(jobNum)
    End Select
End Function

Private Function FolderExists(p As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' Probes every SolidWorks job-type folder; returns the type that contains
' <jobNum> and writes the matching SW job folder path to swJobFolder.
Public Function FindSwJobFolder(jobNum As String, ByRef swJobFolder As String) As String
    Dim swTypes As Variant
    swTypes = Array("GENERAL LINE", "HD-PFD", "HDX", "AXIAL")
    Dim i As Long, candidate As String, intermediate As String
    For i = LBound(swTypes) To UBound(swTypes)
        intermediate = ComputeSwIntermediate(CStr(swTypes(i)), jobNum)
        If Len(intermediate) > 0 Then intermediate = intermediate & "\"
        candidate = SW_ROOT & swTypes(i) & "\" & intermediate & jobNum & "\"
        If FolderExists(candidate) Then
            FindSwJobFolder = CStr(swTypes(i))
            swJobFolder = candidate
            Exit Function
        End If
    Next i
    FindSwJobFolder = ""
End Function

' ---------------------------------------------------------------------
' DAG folder scanning (used by ExtractForm to build the checkboxes)
' ---------------------------------------------------------------------

' Subfolders of DAG_ROOT named "AUTO ...", sorted by name.
Public Function ListAutoFolders() As Collection
    Dim names As New Collection
    Dim entry As String
    On Error Resume Next
    entry = Dir$(DAG_ROOT, vbDirectory)
    On Error GoTo 0
    Do While Len(entry) > 0
        If entry <> "." And entry <> ".." Then names.Add entry
        entry = Dir$()
    Loop

    Dim result As New Collection
    Dim n As Variant, attrs As Long
    For Each n In names
        If UCase$(n) Like "AUTO *" Then
            attrs = 0
            On Error Resume Next
            attrs = GetAttr(DAG_ROOT & n)
            On Error GoTo 0
            If (attrs And vbDirectory) <> 0 Then result.Add CStr(n)
        End If
    Next n
    Set ListAutoFolders = SortedByName(result)
End Function

' Zip file names sitting directly in folderPath (must end with "\"), sorted.
Public Function ListZipsIn(folderPath As String) As Collection
    Dim result As New Collection
    Dim entry As String
    On Error Resume Next
    entry = Dir$(folderPath & "*.zip")
    On Error GoTo 0
    Do While Len(entry) > 0
        result.Add entry
        entry = Dir$()
    Loop
    Set ListZipsIn = SortedByName(result)
End Function

Private Function SortedByName(items As Collection) As Collection
    Dim result As New Collection
    Dim item As Variant, i As Long, inserted As Boolean
    For Each item In items
        inserted = False
        For i = 1 To result.Count
            If StrComp(CStr(item), CStr(result(i)), vbTextCompare) < 0 Then
                result.Add item, Before:=i
                inserted = True
                Exit For
            End If
        Next i
        If Not inserted Then result.Add item
    Next item
    Set SortedByName = result
End Function

' ---------------------------------------------------------------------
' Extraction
' ---------------------------------------------------------------------

Public Sub main()
    If Not FolderExists(DAG_ROOT) Then
        MsgBox "DAG folder not found: " & DAG_ROOT, vbExclamation, "Extract Auto Parts"
        Exit Sub
    End If

    Dim dlg As New ExtractForm
    If dlg.UnitCount = 0 Then
        MsgBox "No AUTO folders with zip files found under " & DAG_ROOT, _
               vbExclamation, "Extract Auto Parts"
        Unload dlg
        Exit Sub
    End If

    dlg.Show
    If dlg.Cancelled Then Unload dlg: Exit Sub

    ' Destination folders are resolved once per name and remembered, so the
    ' AUTO MODELS zips picked in one run share a single MODELS folder and
    ' the already-exists prompt only appears once.
    Dim resolvedDests As Object
    Set resolvedDests = CreateObject("Scripting.Dictionary")
    resolvedDests.CompareMode = vbTextCompare

    Dim report As String
    report = "Job " & dlg.JobNumber & " - " & dlg.JobFolder & vbCrLf & vbCrLf
    Dim unit As Variant
    For Each unit In dlg.SelectedUnits
        Dim destName As String: destName = CStr(unit(2))
        If Not resolvedDests.Exists(destName) Then
            resolvedDests.Add destName, ResolveDestFolder(dlg.JobFolder, destName)
        End If
        Dim destPath As String: destPath = CStr(resolvedDests(destName))
        If Len(destPath) = 0 Then
            report = report & destName & vbCrLf & "    skipped" & vbCrLf & vbCrLf
        Else
            report = report & ExtractUnit(unit, destPath, dlg.JobNumber)
        End If
    Next unit
    Unload dlg

    MsgBox report, vbInformation, "Extract Auto Parts"
End Sub

' Picks the subfolder of the job folder a unit extracts into. If destName
' already exists there, asks whether to create "destName (2)" (then (3),
' and so on) instead. Returns the full path ending in "\", or "" to skip.
Private Function ResolveDestFolder(jobFolder As String, destName As String) As String
    Dim destPath As String: destPath = jobFolder & destName & "\"
    If Not FolderExists(destPath) Then
        ResolveDestFolder = destPath
        Exit Function
    End If

    Dim n As Long: n = 2
    Dim altName As String
    Do
        altName = destName & " (" & n & ")"
        If Not FolderExists(jobFolder & altName & "\") Then Exit Do
        n = n + 1
    Loop

    Select Case MsgBox("""" & destName & """ already exists in the job folder." & vbCrLf & _
                       jobFolder & vbCrLf & vbCrLf & _
                       "Yes - extract into a new """ & altName & """ folder" & vbCrLf & _
                       "No - extract into the existing folder (overwrites its files)" & vbCrLf & _
                       "Cancel - skip this item", _
                       vbYesNoCancel + vbQuestion, "Extract Auto Parts")
        Case vbYes
            ResolveDestFolder = jobFolder & altName & "\"
        Case vbNo
            ResolveDestFolder = destPath
        Case Else
            ResolveDestFolder = ""
    End Select
End Function

' unit = Array(srcSpec, srcIsZip, destName, renameDrawings, label):
'   srcSpec  - a zip file path (srcIsZip True), or an AUTO folder path
'              whose zips should all be extracted (srcIsZip False)
'   label    - friendly name shown in the summary (the checkbox caption)
' destPath is the already-resolved job subfolder (may carry a "(2)" suffix).
Private Function ExtractUnit(unit As Variant, destPath As String, jobNum As String) As String
    Dim srcSpec As String:     srcSpec = CStr(unit(0))
    Dim srcIsZip As Boolean:   srcIsZip = CBool(unit(1))
    Dim renameDwgs As Boolean: renameDwgs = CBool(unit(3))
    Dim label As String:       label = CStr(unit(4))

    Dim unitLog As String: unitLog = label & "  ->  " & LeafName(destPath) & "\" & vbCrLf

    Dim zips As New Collection
    If srcIsZip Then
        zips.Add srcSpec
    Else
        Dim zipName As Variant
        For Each zipName In ListZipsIn(srcSpec)
            zips.Add srcSpec & zipName
        Next zipName
    End If

    If zips.Count = 0 Then
        ExtractUnit = unitLog & "    ! no zip file found in " & srcSpec & vbCrLf & vbCrLf
        Exit Function
    End If

    Dim zipPath As Variant, anyFailed As Boolean
    For Each zipPath In zips
        If Not ExpandArchive(CStr(zipPath), destPath) Then anyFailed = True
    Next zipPath
    If anyFailed Then
        unitLog = unitLog & "    ! extraction FAILED" & vbCrLf
    Else
        unitLog = unitLog & "    extracted" & vbCrLf
    End If

    If renameDwgs Then
        Dim renameLog As String
        renameLog = RenameDrawingsIn(destPath, jobNum)
        If Len(renameLog) = 0 Then renameLog = "    (no drawing found to rename)" & vbCrLf
        unitLog = unitLog & renameLog
    End If

    ExtractUnit = unitLog & vbCrLf
End Function

' Unzips with PowerShell's Expand-Archive: synchronous, overwrites, and
' creates the destination folder when missing.
Private Function ExpandArchive(zipPath As String, destPath As String) As Boolean
    On Error GoTo fail
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
          """Expand-Archive -LiteralPath '" & PsQuote(zipPath) & _
          "' -DestinationPath '" & PsQuote(destPath) & "' -Force"""
    ExpandArchive = (wsh.Run(cmd, 0, True) = 0)
    Exit Function
fail:
    ExpandArchive = False
End Function

Private Function PsQuote(s As String) As String
    PsQuote = Replace(s, "'", "''")
End Function

' Last segment of a folder path: "...\JOBS\512345\MODELS (2)\" -> "MODELS (2)"
Private Function LeafName(folderPath As String) As String
    Dim p As String: p = folderPath
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    LeafName = Mid$(p, InStrRev(p, "\") + 1)
End Function

' Renames every drawing under folderPath (recursively) from <old>-<xx> to
' <jobNum>-<xx>, keeping everything after the first "-" and the extension.
Private Function RenameDrawingsIn(folderPath As String, jobNum As String) As String
    If Not FolderExists(folderPath) Then Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim walkLog As String
    RenameWalk fso, fso.GetFolder(folderPath), jobNum, walkLog
    RenameDrawingsIn = walkLog
End Function

Private Sub RenameWalk(fso As Object, fld As Object, jobNum As String, ByRef walkLog As String)
    ' Snapshot the file list first: renaming while enumerating fld.Files
    ' can skip entries.
    Dim snapshot As New Collection
    Dim f As Object
    For Each f In fld.Files
        snapshot.Add f
    Next f

    Dim fileItem As Object
    For Each fileItem In snapshot
        Dim ext As String: ext = UCase$(fso.GetExtensionName(fileItem.Name))
        If ext = "SLDDRW" Or ext = "DWG" Then
            Dim base As String: base = fso.GetBaseName(fileItem.Name)
            Dim dashAt As Long: dashAt = InStr(base, "-")
            If dashAt = 0 Then
                walkLog = walkLog & "    drawing " & fileItem.Name & _
                          " has no ""-"", left as is" & vbCrLf
            Else
                Dim newName As String
                newName = jobNum & Mid$(base, dashAt) & "." & fso.GetExtensionName(fileItem.Name)
                If StrComp(newName, fileItem.Name, vbTextCompare) <> 0 Then
                    If fso.FileExists(fso.BuildPath(fld.path, newName)) Then
                        walkLog = walkLog & "    ! " & newName & " already exists, left " & _
                                  fileItem.Name & vbCrLf
                    Else
                        Dim oldName As String: oldName = fileItem.Name
                        fileItem.Name = newName
                        walkLog = walkLog & "    renamed " & oldName & " -> " & newName & vbCrLf
                    End If
                End If
            End If
        End If
    Next fileItem

    Dim subFld As Object
    For Each subFld In fld.SubFolders
        RenameWalk fso, subFld, jobNum, walkLog
    Next subFld
End Sub
