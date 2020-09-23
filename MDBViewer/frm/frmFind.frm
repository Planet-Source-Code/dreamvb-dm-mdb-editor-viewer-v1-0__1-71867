VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3330
      TabIndex        =   2
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4635
      TabIndex        =   3
      Top             =   675
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   3360
      TabIndex        =   1
      Top             =   225
      Width           =   2460
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Left            =   2865
      TabIndex        =   5
      Top             =   285
      Width           =   405
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field Name:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   285
      Width           =   840
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmFind
End Sub

Private Sub cmdOK_Click()
Dim rc As Recordset

    ButtonPress = vbOK
    'Store field index
    FindFieldIdx = cboField.ListIndex
    'Store find value
    FindValue = txtValue.Text
    
    Unload frmFind
End Sub

Private Sub Form_Load()
Dim fd As Field

    Set frmFind.Icon = Nothing
    'Load the field names into the combo box
    For Each fd In db.TableDefs(CurrentTable).Fields
        cboField.AddItem fd.Name
    Next fd
    
    'Set the first index
    cboField.ListIndex = 0
    Set fd = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub

Private Sub txtValue_Change()
    cmdOK.Enabled = Len(Trim(txtValue.Text)) > 0
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    'Check if [enter] was pressed
    If (KeyAscii = 13) Then
        'Cancel out keycode
        KeyAscii = 0
        If (cmdOK.Enabled) Then
            'Only click button if enabled
            Call cmdOK_Click
        End If
    End If
End Sub
