VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuery 
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   900
      Left            =   2505
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   900
      Left            =   1530
      TabIndex        =   1
      Top             =   0
      Width           =   915
   End
   Begin VB.TextBox txtSql 
      Height          =   900
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1410
   End
   Begin MSComctlLib.ListView LstQ 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   945
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   3270
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9922
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload frmQuery
End Sub

Private Sub cmdExecute_Click()
    'Execute Query
    If GetFieldData(txtSql.Text, LstQ) <> 1 Then
        MsgBox DBErrMsg, vbInformation, frmQuery.Caption
    End If
End Sub

Private Sub Form_Load()
    Set frmQuery.Icon = Nothing
    frmQuery.Caption = "Query: " & CurrentTable
    txtSql.Text = "Select * from [" & CurrentTable & "]"
    'Get the fieldnames
    Call GetFieldNames(CurrentTable, LstQ)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Position and resize controls
    cmdClose.Left = (frmQuery.ScaleWidth - cmdClose.Width)
    cmdExecute.Left = (cmdClose.Left - cmdExecute.Width) - 15
    txtSql.Width = (cmdExecute.Left - 30)
    LstQ.Width = (frmQuery.ScaleWidth - LstQ.Left)
    LstQ.Height = (frmQuery.ScaleHeight - sBar1.Height - LstQ.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmQuery = Nothing
End Sub

Private Sub txtSql_Change()
    cmdExecute.Enabled = Len(Trim(txtSql.Text))
End Sub
