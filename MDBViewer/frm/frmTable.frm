VERSION 5.00
Begin VB.Form frmTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Table"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraProp 
      Caption         =   "Field Properties"
      Height          =   2520
      Left            =   2910
      TabIndex        =   12
      Top             =   1650
      Width           =   3225
      Begin VB.CheckBox chkZero 
         Caption         =   "Allow Zero Length"
         Enabled         =   0   'False
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   1785
         Width           =   1665
      End
      Begin VB.CheckBox ChkReq 
         Caption         =   "Required"
         Enabled         =   0   'False
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   2055
         Width           =   2445
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "AutoIncement"
         Enabled         =   0   'False
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   1515
         Width           =   2835
      End
      Begin VB.TextBox txtSize 
         Enabled         =   0   'False
         Height          =   345
         Left            =   735
         TabIndex        =   18
         Top             =   1065
         Width           =   810
      End
      Begin VB.TextBox txtType 
         Enabled         =   0   'False
         Height          =   345
         Left            =   735
         TabIndex        =   16
         Top             =   675
         Width           =   810
      End
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   750
         Width           =   405
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6315
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   6315
      Begin VB.Label lblTitleTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   20
         Top             =   120
         Width           =   1110
      End
      Begin VB.Line lnSpacer 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   1320
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use the form below to start creating your table"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   3240
      End
      Begin VB.Image ImgLogo 
         Height          =   570
         Left            =   60
         Picture         =   "frmTable.frx":0000
         Top             =   30
         Width           =   930
      End
   End
   Begin MDBEditor.Line3D Line3D1 
      Height          =   30
      Left            =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1350
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Field"
      Height          =   360
      Left            =   105
      TabIndex        =   2
      Top             =   3810
      Width           =   1245
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Field"
      Enabled         =   0   'False
      Height          =   360
      Left            =   1410
      TabIndex        =   3
      Top             =   3810
      Width           =   1320
   End
   Begin VB.ListBox LstFields 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   2610
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   350
      Left            =   4950
      TabIndex        =   5
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Table"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3585
      TabIndex        =   4
      Top             =   4590
      Width           =   1245
   End
   Begin VB.TextBox txtTable 
      Height          =   350
      Left            =   1170
      TabIndex        =   0
      Top             =   825
      Width           =   3180
   End
   Begin MDBEditor.Line3D Line3D2 
      Height          =   30
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4425
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   53
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field List:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1470
      Width           =   660
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      Caption         =   "Table Name:"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   900
      Width           =   915
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Delete(ByVal Index)
Dim iCount As Long
Dim cTop As Long
Dim TmpIdx As Long
On Error Resume Next

    'Get array upper bound.
    cTop = UBound(TempField) - 1
    
    'Loop tho the array and shift all the items up
    For iCount = Index To cTop
        TempField(iCount).fAttr = TempField(iCount + 1).fAttr
        TempField(iCount).fName = TempField(iCount + 1).fName
        TempField(iCount).fSize = TempField(iCount + 1).fSize
        TempField(iCount).fType = TempField(iCount + 1).fType
    Next iCount
    
    'if top less then zero clear the array.
    If (cTop < 0) Then
        FieldCounter = 0
        Erase TempField
    Else
        'Resize the array remoevng the last index.
        ReDim Preserve TempField(cTop)
    End If
    
    FieldCounter = (cTop + 1)
    'Clear up
    iCount = 0
    cTop = 0
End Sub

Private Sub cmdAdd_Click()
    'Show the form
    frmFields.Show vbModal, frmTable
End Sub

Private Sub cmdClose_Click()
    ButtonPress = vbCancel
    Unload frmTable
End Sub

Private Sub cmdCreate_Click()
    TmpTable = Trim(txtTable.Text)
    
    If Len(TmpTable) = 0 Then
        MsgBox "Table name is required.", vbInformation, frmTable.Caption
        Exit Sub
    End If
    
    ButtonPress = vbOK
    Unload frmTable
End Sub

Private Sub cmdDelete_Click()
    Call Delete(LstFields.ListIndex)
    'Reset controls
    LstFields.RemoveItem LstFields.ListIndex
    txtType.Text = ""
    txtName.Text = ""
    txtSize.Text = ""
    ChkAuto.Value = 0
    chkReq.Value = 0
    chkZero.Value = 0
    
    cmdDelete.Enabled = (LstFields.ListCount)
    cmdCreate.Enabled = (LstFields.ListCount)
    
    If (LstFields.ListCount > 0) Then
        LstFields.ListIndex = 0
    End If
    
End Sub

Private Sub Form_Load()
    Set frmTable.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTable = Nothing
End Sub

Private Sub LstFields_Click()
Dim Idx As Integer
On Error Resume Next
    cmdDelete.Enabled = True
    'Get Listbox index
    Idx = LstFields.ListIndex
    
    txtName.Text = TempField(Idx).fName
    txtType.Text = GetFieldTypeStr(TempField(Idx).fType)
    txtSize.Text = TempField(Idx).fSize
    ChkAuto.Value = Abs(TempField(Idx).fAttr = dbAutoIncrField)
    chkReq.Value = Abs(TempField(Idx).fRequired)
    chkZero.Value = Abs(TempField(Idx).fAllowZeroLength)
End Sub

Private Sub pTop_Resize()
    lnSpacer.X2 = pTop.ScaleWidth
End Sub
