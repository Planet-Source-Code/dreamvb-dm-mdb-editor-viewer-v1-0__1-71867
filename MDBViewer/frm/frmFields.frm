VERSION 5.00
Begin VB.Form frmFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fields"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkZeroLen 
      Caption         =   "Allow Zero Length"
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1740
      Width           =   4275
   End
   Begin VB.CheckBox chkReq 
      Caption         =   "Required"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1995
      Width           =   4275
   End
   Begin MDBEditor.Line3D Line3D1 
      Height          =   30
      Left            =   75
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2355
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   53
   End
   Begin VB.CheckBox ChkAuto 
      Caption         =   "AutoIncField"
      Enabled         =   0   'False
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1485
      Width           =   4275
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   350
      Left            =   3270
      TabIndex        =   5
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Field"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1890
      TabIndex        =   4
      Top             =   2475
      Width           =   1215
   End
   Begin VB.TextBox txtSize 
      Enabled         =   0   'False
      Height          =   350
      Left            =   1050
      TabIndex        =   2
      Top             =   1020
      Width           =   780
   End
   Begin VB.ComboBox cboFType 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   645
      Width           =   1635
   End
   Begin VB.TextBox txtFieldName 
      Height          =   350
      Left            =   1050
      TabIndex        =   0
      Top             =   165
      Width           =   3360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblFieldName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "frmFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function ItemInList(TListBox As ListBox, ByVal FindItem As String) As Integer
Dim Cnt As Integer
Dim Idx As Integer
    
    Idx = -1
    'Returns one -1 if item is found in a listbox
    For Cnt = 0 To TListBox.ListCount
        If LCase(TListBox.List(Cnt)) = LCase(FindItem) Then
            Idx = Cnt
            Exit For
        End If
    Next Cnt
    
    ItemInList = Idx
End Function

Private Sub cboFType_Click()
    'Show Fieldsize
    txtSize.Text = GetFieldSize(cboFType.Text)
    'Enable/Disable controls
    ChkAuto.Enabled = (cboFType.ListIndex = 3)
    txtSize.Enabled = (cboFType.ListIndex = 8)
    chkZeroLen.Enabled = (cboFType.ListIndex = 8) Or (cboFType.ListIndex = 9)
End Sub

Private Sub cmdAdd_Click()
Dim Fld As Field

    'Add fieldname to the table list.
    If ItemInList(frmTable.LstFields, txtFieldName.Text) = -1 Then
        frmTable.LstFields.AddItem txtFieldName.Text
        'Fill in Table def

        ReDim Preserve TempField(FieldCounter)
        
        With TempField(FieldCounter)
            'Fill in FieldInfo
            .fName = txtFieldName.Text
            .fSize = Val(txtSize.Text)
            .fType = GetFieldType(cboFType.Text)
            .fRequired = chkReq.Value
            .fAllowZeroLength = chkZeroLen.Value
            
            'Check if adding an AutoNumber field
            If (cboFType.ListIndex = 3) And (ChkAuto.Value) Then
                .fAttr = (.fAttr Or dbAutoIncrField)
            End If
        End With
        
        'INC FieldCount
        FieldCounter = (FieldCounter + 1)
    Else
        MsgBox "'" & txtFieldName.Text & "' is already in the list.", vbInformation, frmFields.Caption
    End If
    
    'Reset defaults
    frmTable.cmdCreate.Enabled = True
    frmTable.LstFields.ListIndex = (frmTable.LstFields.ListCount - 1)
    txtFieldName.Text = vbNullString
    ChkAuto.Value = 0
    chkReq.Value = 0
    chkZeroLen.Value = 0
    cboFType.ListIndex = 8
    txtFieldName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload frmFields
End Sub

Private Sub Form_Load()
    Set frmFields.Icon = Nothing
    'Add some field Types
    cboFType.AddItem "Boolean"
    cboFType.AddItem "Byte"
    cboFType.AddItem "Integer"
    cboFType.AddItem "Long"
    cboFType.AddItem "Currency"
    cboFType.AddItem "Single"
    cboFType.AddItem "Double"
    cboFType.AddItem "Date/Time"
    cboFType.AddItem "Text"
    cboFType.AddItem "Memo"
    cboFType.ListIndex = 8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFields = Nothing
End Sub

Private Sub txtFieldName_Change()
    cmdAdd.Enabled = Len(Trim(txtFieldName.Text)) > 0 And Len(Trim(txtSize.Text)) > 0
End Sub

Private Sub txtFieldName_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        If (cmdAdd.Enabled) Then
            'Click button
            Call cmdAdd_Click
        End If
    End If
End Sub

Private Sub txtSize_Change()
    Call txtFieldName_Change
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Then
        Exit Sub
    End If
    
    If CBool((KeyAscii >= 48) And (KeyAscii <= 57)) = False Then
        KeyAscii = 0
    End If
    
End Sub
