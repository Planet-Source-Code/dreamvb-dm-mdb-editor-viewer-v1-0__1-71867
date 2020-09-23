VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   3930
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1185
   End
   Begin VB.PictureBox pBar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   6435
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6435
      Begin VB.Image ImgLogo 
         Height          =   780
         Left            =   75
         Picture         =   "frmView.frx":0000
         Top             =   -15
         Width           =   720
      End
      Begin VB.Line lnSpacer 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   405
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the following form below to Add/Edit the record"
         Height          =   195
         Left            =   1020
         TabIndex        =   7
         Top             =   300
         Width           =   3645
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "#0"
      Height          =   350
      Left            =   2670
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1185
   End
   Begin VB.PictureBox pHolder 
      Height          =   4665
      Left            =   0
      ScaleHeight     =   307
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   424
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   6420
      Begin VB.PictureBox pMover 
         BorderStyle     =   0  'None
         Height          =   2865
         Left            =   0
         ScaleHeight     =   2865
         ScaleWidth      =   6045
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   6045
         Begin MDBEditor.dField dField1 
            Height          =   555
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   60
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   979
         End
      End
      Begin VB.VScrollBar vBar1 
         Height          =   750
         Left            =   6075
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   350
      Left            =   5190
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1185
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HighlightField(ByVal Index As Integer)
On Error Resume Next
    'Highlight the field
    dField1(Index).SelStart = 0
    dField1(Index).SelLength = Len(dField1(Index).Text)
    dField1(Index).SetFocus
End Sub

Private Sub FixAutoField()
Dim rc As Recordset
Dim Count As Integer

    'Open the recordset
    Set rc = db.OpenRecordset(CurrentTable)
    For Count = 0 To rc.Fields.Count
        'Check if we have a AutoNumber field
        If (rc.Fields(Count).Attributes And dbAutoIncrField) = dbAutoIncrField Then
            dField1(Count).Text = "AutoInc"
            dField1(Count).Tag = "AutoInc"
            dField1(Count).Enabled = False
        End If
    Next Count
    
    Set rc = Nothing
End Sub

Private Sub SetupControls()
Dim fd As Field
Dim fCount As Integer

    'This loads the text fields controls
    For Each fd In db.TableDefs(CurrentTable).Fields
        fCount = (fCount + 1)
        Load dField1(fCount)
        dField1(fCount).Top = dField1(fCount - 1).Top + dField1(0).Height
        dField1(fCount - 1).Visible = True
        dField1(fCount - 1).Caption = fd.Name
        'Set TextField Size
        If (fd.Type = dbText) Then
            dField1(fCount - 1).MaxLength = fd.Size
        End If
    Next fd
    
    'Unload the last field
    Call Unload(dField1(dField1.Count - 1))
    
    'Position the scrollbar
    With vBar1
        .Left = (pHolder.ScaleWidth - vBar1.Width)
        .Height = pHolder.ScaleHeight
        pMover.Height = (dField1(dField1.Count - 1).Top \ 15) + dField1(0).Height \ 15
        .Max = (pMover.Height - pHolder.Height) + 15
        .Enabled = (.Max > 0)
    End With

End Sub

Private Sub ShowRecord()
Dim rc As Recordset
Dim Count As Integer

    Set rc = db.OpenRecordset(CurrentTable)
    'Move to the selected record
    rc.Move CurrentRecord
    For Count = 0 To rc.Fields.Count
        dField1(Count).Text = rc.Fields(Count).Value
    Next Count
    
    'Clear Up
    Set rc = Nothing
    Count = 0
End Sub

Private Function AddRecord(AddRec As RecOp) As Integer
Dim Count As Integer
Dim rc As Recordset
Dim fAttr As Integer

On Error GoTo AddErr:
    
    AddRecord = -1
    'Open recordset
    Set rc = db.OpenRecordset(CurrentTable)
    
    'Add new record
    If (AddRec = rAddNew) Then
        rc.AddNew
    End If
    
    'Edit exsiting record
    If (AddRec = rEdit) Then
        rc.Move CurrentRecord
        rc.Edit
    End If
    
    For Count = 0 To rc.Fields.Count - 1
        'Check if we are dealing with an autonumber field
        fAttr = (rc.Fields(Count).Attributes And dbAutoIncrField)
        If (fAttr <> dbAutoIncrField) Then
            'Not autonumber so add the field
            rc.Fields(Count).Value = dField1(Count).Text
        End If
    Next Count
    'Update record
    rc.Update
    fAttr = 0
    
    Exit Function
AddErr:
    AddRecord = Count
    MsgBox Err.Description, vbInformation, frmView.Caption
End Function

Private Sub cmdClear_Click()
Dim Count As Integer
    'Clear all the text fields
    For Count = 0 To (dField1.Count - 1)
        If dField1(Count).Tag <> "AutoInc" Then
            dField1(Count).Text = ""
        Else
            dField1(Count).Text = "AutoInc"
        End If
    Next Count
    
    vBar1.Value = 0
    dField1(0).SetFocus
    
End Sub

Private Sub cmdClose_Click()
    'Unload this form
    ButtonPress = vbCancel
    Unload frmView
End Sub

Private Sub cmdOK_Click()
Dim Idx As Integer

    If (EditOp = rAddNew) Then
        'Add record
        Idx = AddRecord(rAddNew)
        
        If (Idx <> -1) Then
            If (Idx = 1) Then Idx = 0
            'Highlight the field
            Call HighlightField(Idx)
            Exit Sub
        End If
    End If
    
    If (EditOp = rEdit) Then
        'Edit record
        Idx = AddRecord(rEdit)
        If (Idx <> -1) Then
            If (Idx = 1) Then Idx = 0
            'Highlight the field
            Call HighlightField(Idx)
            Exit Sub
        End If
    End If
    
    'Unload the form
    ButtonPress = vbOK
    Unload frmView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmView = Nothing
End Sub

Private Sub pBar_Resize()
    lnSpacer.X2 = pBar.ScaleWidth
End Sub

Private Sub vBar1_Change()
    pMover.Top = (-vBar1.Value)
End Sub

Private Sub vBar1_Scroll()
    Call vBar1_Change
End Sub

Private Sub Form_Load()
Dim fCount As Integer
On Error Resume Next
    
    Set frmView.Icon = Nothing
    
    'Setup the Controls
    Call SetupControls
    
    If (EditOp = rEdit) Then
        cmdOK.Caption = "Update"
        frmView.Caption = "Update Record"
        Call ShowRecord
    End If
    
    If (EditOp = rAddNew) Then
        cmdOK.Caption = "Add"
        frmView.Caption = "Add Record"
    End If
    
    'Fix AutoNumber Field
    Call FixAutoField
End Sub

