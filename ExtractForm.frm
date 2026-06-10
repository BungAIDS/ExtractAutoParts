VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExtractForm
   Caption         =   "Extract Auto Parts"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExtractForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Dialog for ExtractAutoParts. Every control is created at run time in
' UserForm_Initialize (which is why this .frm needs no .frx blob): one
' checkbox per "AUTO *" folder that contains a zip, plus one checkbox
' per zip inside AUTO MODELS. Checking the "daddy" model zip locks out
' the other model zips - it already contains everything they do.
'
' The module reads Cancelled / JobNumber / JobFolder / SelectedUnits
' after Show returns.

Public Cancelled As Boolean
Public JobNumber As String
Public JobFolder As String
Public SelectedUnits As Collection   ' of Array(srcSpec, srcIsZip, destName, renameDrawings, label)

Private mChecks As Collection        ' every checkbox, index-aligned with mUnits
Private mUnits As Collection
Private mModelChecks As Collection   ' AUTO MODELS checkboxes other than daddy
Private mDetectedJob As String       ' order # of the open SolidWorks doc, "" if none
Private mTxtJob As MSForms.TextBox
Private WithEvents mOptDetected As MSForms.OptionButton  ' "use the open order"
Private WithEvents mOptOther As MSForms.OptionButton     ' "type another order"
Private WithEvents mChkDaddy As MSForms.CheckBox
Private WithEvents mBtnExtract As MSForms.CommandButton
Private WithEvents mBtnCancel As MSForms.CommandButton

Private Const INNER_WIDTH As Single = 312
Private Const MARGIN As Single = 8
Private Const ROW_HEIGHT As Single = 16

Public Property Get UnitCount() As Long
    UnitCount = mUnits.Count
End Property

Private Sub UserForm_Initialize()
    Cancelled = True
    Set mChecks = New Collection
    Set mUnits = New Collection
    Set mModelChecks = New Collection

    Me.Caption = "Extract Auto Parts"

    Dim y As Single: y = MARGIN
    y = BuildJobChooser(y)
    y = BuildPartsFrame(y)
    y = BuildModelsFrame(y)

    Set mBtnExtract = Me.Controls.Add("Forms.CommandButton.1", "btnExtract")
    mBtnExtract.Caption = "Extract"
    mBtnExtract.Default = True
    mBtnExtract.Width = 70: mBtnExtract.Height = 22
    mBtnExtract.Left = INNER_WIDTH - MARGIN - 70 - 6 - 70
    mBtnExtract.Top = y

    Set mBtnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    mBtnCancel.Caption = "Cancel"
    mBtnCancel.Cancel = True
    mBtnCancel.Width = 70: mBtnCancel.Height = 22
    mBtnCancel.Left = INNER_WIDTH - MARGIN - 70
    mBtnCancel.Top = y
    y = y + 22 + MARGIN

    Me.Width = INNER_WIDTH + (Me.Width - Me.InsideWidth)
    Me.Height = y + (Me.Height - Me.InsideHeight)
End Sub

' Order # entry at the top. When a job number can be read off the open
' SolidWorks document, offer two radio buttons - use that one, or type
' another - with the detected order selected by default. With nothing
' open to detect, fall back to a plain "Order #:" text box.
Private Function BuildJobChooser(ByVal startY As Single) As Single
    mDetectedJob = DetectActiveJob()
    Dim y As Single: y = startY

    Set mTxtJob = Me.Controls.Add("Forms.TextBox.1", "txtJob")
    mTxtJob.Height = 18

    If Len(mDetectedJob) > 0 Then
        Set mOptDetected = Me.Controls.Add("Forms.OptionButton.1", "optDetected")
        mOptDetected.Caption = "Use open order:  " & mDetectedJob
        mOptDetected.GroupName = "order"
        mOptDetected.Left = MARGIN: mOptDetected.Top = y
        mOptDetected.Width = INNER_WIDTH - 2 * MARGIN: mOptDetected.Height = 16
        mOptDetected.Value = True
        y = y + 18

        Set mOptOther = Me.Controls.Add("Forms.OptionButton.1", "optOther")
        mOptOther.Caption = "Other order #:"
        mOptOther.GroupName = "order"
        mOptOther.Left = MARGIN: mOptOther.Top = y + 2
        mOptOther.Width = 96: mOptOther.Height = 16

        mTxtJob.Left = MARGIN + 100: mTxtJob.Top = y
        mTxtJob.Width = 90: mTxtJob.Enabled = False
        y = y + 18 + MARGIN
    Else
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblJob")
        lbl.Caption = "Order #:"
        lbl.Left = MARGIN: lbl.Top = y + 3: lbl.Width = 48: lbl.Height = 12

        mTxtJob.Left = MARGIN + 50: mTxtJob.Top = y: mTxtJob.Width = 90
        y = y + 18 + MARGIN
    End If

    BuildJobChooser = y
End Function

' The order # the user wants: detected order when its radio is selected,
' otherwise whatever is typed in the text box. (VBA's And does not short-
' circuit, so the Nothing check has to be a separate If.)
Private Function ChosenJob() As String
    If mOptDetected Is Nothing Then
        ChosenJob = Trim$(mTxtJob.Text)
    ElseIf mOptDetected.Value Then
        ChosenJob = mDetectedJob
    Else
        ChosenJob = Trim$(mTxtJob.Text)
    End If
End Function

' Order # of the active SolidWorks document, read from the leading digits
' of its file name (jobs are saved as "<order>-01.SLDDRW" etc.). Returns
' "" when nothing is open or the name has no 3+ digit prefix.
Private Function DetectActiveJob() As String
    On Error Resume Next
    Dim swApp As Object: Set swApp = Application.SldWorks
    If swApp Is Nothing Then Exit Function
    Dim doc As Object: Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Function

    Dim nm As String: nm = doc.GetPathName        ' full path, "" if unsaved
    If Len(nm) = 0 Then nm = doc.GetTitle
    If InStrRev(nm, "\") > 0 Then nm = Mid$(nm, InStrRev(nm, "\") + 1)
    If InStrRev(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)

    Dim digits As String: digits = LeadingDigits(nm)
    If Len(digits) >= 3 Then DetectActiveJob = digits
End Function

Private Function LeadingDigits(s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        Select Case Mid$(s, i, 1)
            Case "0" To "9": LeadingDigits = LeadingDigits & Mid$(s, i, 1)
            Case Else: Exit Function
        End Select
    Next i
End Function

' Selecting "Other" frees the text box for typing; selecting the detected
' order locks it again. Only one handler is needed - the pair toggles
' together.
Private Sub mOptOther_Change()
    mTxtJob.Enabled = mOptOther.Value
    If mOptOther.Value Then
        On Error Resume Next
        mTxtJob.SetFocus
    End If
End Sub

' One checkbox per AUTO folder (except AUTO MODELS) that has a zip in it.
Private Function BuildPartsFrame(ByVal startY As Single) As Single
    Dim fra As MSForms.Frame
    Set fra = Me.Controls.Add("Forms.Frame.1", "fraParts")
    fra.Caption = "Parts"
    fra.Left = MARGIN: fra.Top = startY: fra.Width = INNER_WIDTH - 2 * MARGIN

    Dim y As Single: y = 10
    Dim folderName As Variant
    For Each folderName In ListAutoFolders()
        If StrComp(CStr(folderName), MODELS_FOLDER, vbTextCompare) <> 0 Then
            Dim folderPath As String: folderPath = DAG_ROOT & folderName & "\"
            If ListZipsIn(folderPath).Count > 0 Then
                Dim display As String: display = StripAutoPrefix(CStr(folderName))
                AddUnitCheckBox fra, y, display, Array(folderPath, False, display, True, display)
                y = y + ROW_HEIGHT
            End If
        End If
    Next folderName

    fra.Height = y + 14
    BuildPartsFrame = startY + fra.Height + MARGIN
End Function

' One checkbox per zip inside AUTO MODELS; they all extract into MODELS_DEST.
Private Function BuildModelsFrame(ByVal startY As Single) As Single
    Dim modelsPath As String: modelsPath = DAG_ROOT & MODELS_FOLDER & "\"
    Dim zips As Collection
    Set zips = ListZipsIn(modelsPath)
    If zips.Count = 0 Then
        BuildModelsFrame = startY
        Exit Function
    End If

    Dim fra As MSForms.Frame
    Set fra = Me.Controls.Add("Forms.Frame.1", "fraModels")
    fra.Caption = "Models (all extract into """ & MODELS_DEST & """)"
    fra.Left = MARGIN: fra.Top = startY: fra.Width = INNER_WIDTH - 2 * MARGIN

    Dim y As Single: y = 10
    Dim zipName As Variant
    For Each zipName In zips
        Dim isDaddy As Boolean
        isDaddy = (InStr(1, CStr(zipName), "daddy", vbTextCompare) > 0)
        Dim label As String
        If isDaddy Then
            label = "A8 + Accessories"          ' the all-in-one zip
        Else
            label = ZipLabel(CStr(zipName))
        End If
        Dim chk As MSForms.CheckBox
        Set chk = AddUnitCheckBox(fra, y, label, _
                                  Array(modelsPath & zipName, True, MODELS_DEST, Not isDaddy, label))
        If isDaddy Then
            Set mChkDaddy = chk
            chk.Font.Bold = True
        Else
            mModelChecks.Add chk
        End If
        y = y + ROW_HEIGHT
    Next zipName

    fra.Height = y + 14
    BuildModelsFrame = startY + fra.Height + MARGIN
End Function

Private Function AddUnitCheckBox(fra As MSForms.Frame, ByVal y As Single, _
                                 displayText As String, unit As Variant) As MSForms.CheckBox
    Dim chk As MSForms.CheckBox
    Set chk = fra.Controls.Add("Forms.CheckBox.1", "chkUnit" & (mChecks.Count + 1))
    chk.Left = 8: chk.Top = y
    chk.Width = fra.Width - 20: chk.Height = ROW_HEIGHT
    chk.Caption = displayText   ' MSForms shows captions literally; no & escaping
    mChecks.Add chk
    mUnits.Add unit
    Set AddUnitCheckBox = chk
End Function

' "AUTO INLET IVC LINKAGE" -> "INLET IVC LINKAGE"
Private Function StripAutoPrefix(folderName As String) As String
    If UCase$(folderName) Like "AUTO *" Then
        StripAutoPrefix = Trim$(Mid$(folderName, 6))
    Else
        StripAutoPrefix = folderName
    End If
End Function

' "EXTRACT ME!! (A8 ANGLE BASE ONLY).zip" -> "A8 ANGLE BASE ONLY"
Private Function ZipLabel(zipName As String) As String
    Dim base As String: base = zipName
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    Dim openAt As Long: openAt = InStr(base, "(")
    Dim closeAt As Long: closeAt = InStrRev(base, ")")
    If openAt > 0 And closeAt > openAt Then
        ZipLabel = Mid$(base, openAt + 1, closeAt - openAt - 1)
    Else
        ZipLabel = base
    End If
End Function

' daddy contains everything the other model zips do, so checking it locks
' them out.
Private Sub mChkDaddy_Change()
    Dim chk As Variant
    For Each chk In mModelChecks
        If mChkDaddy.Value Then chk.Value = False
        chk.Enabled = Not mChkDaddy.Value
    Next chk
End Sub

Private Sub mBtnExtract_Click()
    Dim jobNum As String: jobNum = ChosenJob()
    If Len(jobNum) < 3 Or Not jobNum Like String$(Len(jobNum), "#") Then
        MsgBox "Order number must be numeric and at least 3 digits.", vbExclamation
        Exit Sub
    End If

    Dim picked As New Collection
    Dim i As Long
    For i = 1 To mChecks.Count
        If mChecks(i).Enabled And mChecks(i).Value Then picked.Add mUnits(i)
    Next i
    If picked.Count = 0 Then
        MsgBox "Check at least one item to extract.", vbExclamation
        Exit Sub
    End If

    Dim jobFolderPath As String
    If Len(FindSwJobFolder(jobNum, jobFolderPath)) = 0 Then
        MsgBox "No SolidWorks job folder found for job " & jobNum & "." & vbCrLf & _
               "Searched all type folders under " & SW_ROOT, vbExclamation
        Exit Sub
    End If

    JobNumber = jobNum
    JobFolder = jobFolderPath
    Set SelectedUnits = picked
    Cancelled = False
    Me.Hide
End Sub

Private Sub mBtnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

' Treat the title-bar X like Cancel; hiding (not unloading) keeps the
' public fields readable after Show returns.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Cancelled = True
        Me.Hide
    End If
End Sub
