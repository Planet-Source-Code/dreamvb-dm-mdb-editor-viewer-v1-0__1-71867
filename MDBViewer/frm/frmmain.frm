VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmmain 
   Caption         =   "DM MDB Editor"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8625
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin MDBEditor.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   53
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   1290
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip sTab1 
      Height          =   2790
      Left            =   0
      TabIndex        =   2
      Top             =   390
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   4921
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   720
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1080
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":13D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1708
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":181A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":192C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1A3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MnuD"
                  Text            =   "Database"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "MnuT"
                  Text            =   "New Table"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FIND"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FIRST"
            Object.ToolTipText     =   "First"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PRE"
            Object.ToolTipText     =   "previous"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NEXT"
            Object.ToolTipText     =   "Next"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "LAST"
            Object.ToolTipText     =   "Last"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NEWREC"
            Object.ToolTipText     =   "New Record"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "EDITREC"
            Object.ToolTipText     =   "Edit recored"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "DELREC"
            Object.ToolTipText     =   "Delete Record"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuNew 
         Caption         =   "New Database"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuCloseDb 
         Caption         =   "Close Database"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "Table"
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Table Info"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSer 
      Caption         =   "&Search"
      Begin VB.Menu mnuSerVal 
         Caption         =   "Search Value"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSqu 
         Caption         =   "&Query"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCompact 
         Caption         =   "CompactDB"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnurCount 
         Caption         =   "Record Count"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCon 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMouseButton As MouseButtonConstants
Private CurrentTab As Integer

Private Const dFilter1 = "Database Files(*.mdb)|*.mdb|"
Private Const dFilter2 = "Access 97(*.mdb)|*.mdb|Access 2000(*.mdb)|*.mdb|"
Private dFilterIdx As Integer
Private dInitDir As String

Private Sub LMoveItem(MoveUp As Boolean)
Dim Idx As Integer

    Idx = LstV.SelectedItem.Index
    
    If (MoveUp) Then
        Idx = (Idx - 1)
        If (Idx <= 1) Then
            Idx = 1
        End If
    Else
        Idx = (Idx + 1)
        If (Idx >= LstV.ListItems.Count) Then
            Idx = LstV.ListItems.Count
        End If
    End If
    
    Call SelectLItem(Idx)

End Sub

Private Sub SelectLItem(ByVal Index As Integer)
    If (LstV.ListItems.Count) Then
        'Select an item in the listview control
        LstV.ListItems(Index).Selected = True
        LstV.ListItems(Index).EnsureVisible
        LstV.SetFocus
        Call LstV_Click
    End If
End Sub

Private Sub SelectTab(ByVal Index As Integer)
    'Only select the tabs if we have items
    If (sTab1.Tabs.Count) Then
        sTab1.Tabs(Index).Selected = True
        Call sTab1_Click
    End If
End Sub

Private Sub DoRefresh(Optional Index As Integer = 1)
On Error Resume Next
    'List the tables
    Call GetTables(sTab1)
    'Select the tab
    Call SelectTab(Index)
End Sub

Private Function GetDLGName(Optional dShowOpen As Boolean = True, Optional dlgFilter As String = dFilter1, _
Optional dTitle As String = "Open") As String
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = dTitle
        .Filter = dlgFilter
        .InitDir = dInitDir
        .Filename = ""
        
        If (dShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        'Set InitDir
        dInitDir = .InitDir
        dFilterIdx = .FilterIndex
        'Return filename
        GetDLGName = .Filename
    End With
    
    Exit Function
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub Form_Load()
    dInitDir = App.Path
    ButtonPress = vbCancel
    sTab1.Tabs.Clear
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resize the controls
    Line3D1.Width = frmmain.ScaleWidth
    sTab1.Width = frmmain.ScaleWidth
    sTab1.Height = (frmmain.ScaleHeight - sBar1.Height - sTab1.Top)

    LstV.Width = (frmmain.ScaleWidth - LstV.Left) - 30
    LstV.Height = (frmmain.ScaleHeight - sBar1.Height - LstV.Top) - 30

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub LstV_Click()
    If (LstV.ListItems.Count) Then
        'Store the selected index
        CurrentRecord = (LstV.SelectedItem.Index - 1)
        'Enable/Disable buttons
        tBar1.Buttons(11).Enabled = True
        tBar1.Buttons(12).Enabled = True
    End If
End Sub

Private Sub LstV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static sSort As Integer
    sSort = (Not sSort)
    LstV.SortKey = (ColumnHeader.Index - 1)
    LstV.SortOrder = Abs(sSort)
    LstV.Sorted = True
End Sub

Private Sub LstV_DblClick()

    If (mMouseButton = vbLeftButton) And (LstV.ListItems.Count) Then
        EditOp = 0
        'Display the data form
        frmView.Show vbModal, frmmain
        If (ButtonPress <> vbCancel) Then
            'Reload the Records
            Call SelectTab(CurrentTab)
            'Select the item
            Call SelectLItem(CurrentRecord + 1)
        End If
    End If
    
    ButtonPress = vbCancel

End Sub

Private Sub LstV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseButton = Button
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & " v1.0" & vbCrLf & vbTab & "By DreamVB", vbInformation, "About"
End Sub

Private Sub mnuCloseDb_Click()
    'Close the database and clear the controls
    Call CloseDataBase
    LstV.ColumnHeaders.Clear
    LstV.ListItems.Clear
    sTab1.Tabs.Clear
    
    mnuCompact.Enabled = False
    tBar1.Buttons(1).ButtonMenus(2).Enabled = False
    tBar1.Buttons(4).Enabled = False
    tBar1.Buttons(5).Enabled = False
    tBar1.Buttons(6).Enabled = False
    tBar1.Buttons(7).Enabled = False
    tBar1.Buttons(8).Enabled = False
    tBar1.Buttons(10).Enabled = False
    tBar1.Buttons(11).Enabled = False
    tBar1.Buttons(12).Enabled = False
    
    mnuRename.Enabled = False
    mnuDelete.Enabled = False
    mnuInfo.Enabled = False
    mnuCloseDb.Enabled = False
    mnuSerVal.Enabled = False
    mnuSqu.Enabled = False
    mnurCount.Enabled = False
    
    sBar1.Panels(1).Text = ""
    sBar1.Panels(2).Text = ""
End Sub

Private Sub mnuCompact_Click()
Dim tmpFile As String

    If (Not DBOpen) Then
        Exit Sub
    Else
        tmpFile = FixPath(App.Path) & "tmp.mdb"
        'Close the open Database
        db.Close
        'Compact the database
        Call CompactDatabase(DataBaseFile, tmpFile)
        'Delete the original database
        Call SetAttr(DataBaseFile, vbNormal)
        Call Kill(DataBaseFile)
        'Rename the temp database as the original
        Name tmpFile As DataBaseFile
        'Reopen the database
        DBOpen = OpenDB(DataBaseFile)
        MsgBox "Finished", vbInformation, "CompactDB"
        'Clear up
        tmpFile = vbNullString
    End If
    
End Sub

Private Sub mnuCon_Click()
    RunApp frmmain.hwnd, "open", FixPath(App.Path) & "Readme.rtf"
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("Are you sure you want to delete the table '" & CurrentTable & "'", vbYesNo Or vbExclamation) = vbYes Then
        'Delete the table
        Call DeleteTable
        Call DoRefresh
        '
        If (sTab1.Tabs.Count = 0) Then
            'Clear the listview
            LstV.ListItems.Clear
            LstV.ColumnHeaders.Clear
        End If
        
        'Update statusbar
        sBar1.Panels(2).Text = "Records: " & LstV.ListItems.Count
        mnuRename.Enabled = sTab1.Tabs.Count
        mnuDelete.Enabled = sTab1.Tabs.Count
        mnuInfo.Enabled = sTab1.Tabs.Count
        mnuSerVal.Enabled = sTab1.Tabs.Count
        mnuSqu.Enabled = sTab1.Tabs.Count
    End If
End Sub

Private Sub mnuExit_Click()
    Call CloseDataBase
    Unload frmmain
End Sub

Private Sub mnuInfo_Click()
    frmInfo.Show vbModal, frmmain
End Sub

Private Sub MnuNew_Click()
Dim lFile As String
Dim dbVer As DatabaseTypeEnum
On Error GoTo FErr:

    'Get New Filename
    lFile = GetDLGName(False, dFilter2, "Create DB")
    'Check if the database isalready found
    If FindFile(lFile) Then
        If MsgBox(lFile & " already exists." & vbCrLf & "Do you want to replace it?", vbYesNo Or vbExclamation) = vbNo Then
            Exit Sub
        Else
            Call Kill(lFile)
        End If
    End If
    
    If Len(lFile) Then
        'Set database version
        If (dFilterIdx = 1) Then
            dbVer = dbVersion30
        Else
            dbVer = dbVersion40
        End If
        
        'Create the new blank database.
        If CreateBlankDataBase(lFile, dbVer) <> 1 Then
            MsgBox "There was an error createing the database.", vbInformation, frmmain.Caption
        End If
    End If
    
    Exit Sub
    'Error flag
FErr:
    MsgBox Err.Description, vbInformation, frmmain.Caption
End Sub

Private Sub mnuOpen_Click()
    DataBaseFile = GetDLGName()
    
    If Len(DataBaseFile) Then
        'Close the database if it's already open
        Call mnuCloseDb_Click
        'Load the database
        DBOpen = OpenDB(DataBaseFile)
        If (Not DBOpen) Then
            MsgBox "Cannot open database.", vbInformation, frmmain.Caption
        Else
            Call DoRefresh
            sBar1.Panels(1).Text = DataBaseFile
        End If
    End If
    
    mnuCloseDb.Enabled = DBOpen
    mnuCompact.Enabled = DBOpen
    mnurCount.Enabled = DBOpen
    tBar1.Buttons(1).ButtonMenus(2).Enabled = DBOpen
    tBar1.Buttons(4).Enabled = LstV.ListItems.Count
End Sub

Private Sub mnurCount_Click()
Dim Count As Integer
    'Get record count
    Count = GetRecordCount
    'Check for error
    If (Count = -1) Then
        MsgBox DBErrMsg, vbInformation, frmmain.Caption
    Else
        MsgBox "Recordcount: " & Count
    End If
    
End Sub

Private Sub mnuRename_Click()
Dim tName As String

    'Rename Table
    tName = Trim$(InputBox("Enter a new name for the table '" & CurrentTable & "'", "Rename Table", CurrentTable))
    If Len(tName) > 0 Then
        'Check if the table was renamed.
        If RenameTable(tName) <> 1 Then
            MsgBox DBErrMsg, vbInformation, frmmain.Caption
        Else
            'Refresh eveything
            Call DoRefresh(CurrentTab)
            Call SelectLItem(CurrentRecord + 1)
        End If
    End If
    
End Sub

Private Sub mnuSerVal_Click()
Dim lItem As ListItem

    frmFind.Show vbModal, frmmain
    If (ButtonPress <> vbCancel) Then
        'This is used to find a field value
        For Each lItem In LstV.ListItems
            'Check if we need to serach subitems
            If (FindFieldIdx > 0) Then
                If lItem.SubItems(FindFieldIdx) = FindValue Then
                    Call SelectLItem(lItem.Index)
                    Exit For
                End If
            Else
                'Selecting the text item
                If (lItem.Text = FindValue) Then
                    Call SelectLItem(lItem.Index)
                    Exit For
                End If
            End If
        Next lItem
    End If
    
End Sub

Private Sub mnuSqu_Click()
    frmQuery.Show vbModal, frmmain
End Sub

Private Sub sTab1_Click()

    If (sTab1.Tabs.Count) Then
        'Get Tab Index
        CurrentTab = sTab1.SelectedItem.Index
        'Get Table name
        CurrentTable = sTab1.SelectedItem.Caption
        'Add field names
        Call GetFieldNames(CurrentTable, LstV)
        'Add the data
        If GetFieldData(CurrentTable, LstV) <> 1 Then
            MsgBox DBErrMsg, vbInformation, frmmain.Caption
        End If
        
        tBar1.Buttons(4).Enabled = LstV.ListItems.Count
        tBar1.Buttons(5).Enabled = LstV.ListItems.Count
        tBar1.Buttons(6).Enabled = LstV.ListItems.Count
        tBar1.Buttons(7).Enabled = LstV.ListItems.Count
        tBar1.Buttons(8).Enabled = LstV.ListItems.Count
        
        tBar1.Buttons(10).Enabled = True
        mnuRename.Enabled = True
        mnuDelete.Enabled = True
        mnuInfo.Enabled = True
        mnuSerVal.Enabled = True
        mnuSqu.Enabled = True
    End If
    
    'Update statusbar
    sBar1.Panels(2).Text = "Records: " & GetRecordCount
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lItem As ListItem
Dim lIdx As Integer

    Select Case Button.Key
        Case "OPEN"
            Call mnuOpen_Click
        Case "INFO"
            'Show table fields info
            frmInfo.Show vbModal, frmmain
        Case "FIND"
            'Find Value
            Call mnuSerVal_Click
        Case "FIRST"
            'Move to record
            Call SelectLItem(1)
        Case "PRE"
            'move to previous record
            Call LMoveItem(True)
        Case "NEXT"
            'Move to next record
            Call LMoveItem(False)
        Case "LAST"
            'Move to the last record
            Call SelectLItem(LstV.ListItems.Count)
        Case "NEWREC"
            'Add new record
            EditOp = rAddNew
            frmView.Show vbModal, frmmain
            If (ButtonPress <> vbCancel) Then
                'Reload the Records
                Call SelectTab(CurrentTab)
                'Select the last item added
                Call SelectLItem(LstV.ListItems.Count)
            End If
        Case "EDITREC"
            'Edit record
            EditOp = rEdit
            'Display the data form
            frmView.Show vbModal, frmmain
            If (ButtonPress <> vbCancel) Then
                'Reload the Records
                Call SelectTab(CurrentTab)
                'Select the item
                Call SelectLItem(CurrentRecord + 1)
            End If
        Case "DELREC"
            'Delete Record
            If MsgBox("Are you sure you want to delete this record.", vbYesNo Or vbQuestion) = vbNo Then
                Exit Sub
            Else
                If DeleteRecord(CurrentRecord) <> 1 Then
                    MsgBox DBErrMsg, vbInformation, frmmain.Caption
                Else
                    'Reload the Records
                    Call SelectTab(CurrentTab)
                    'Select the last item added
                    Call SelectLItem(LstV.ListItems.Count)
                End If
                'Enable/Disable buttons
                tBar1.Buttons(11).Enabled = LstV.ListItems.Count
                tBar1.Buttons(12).Enabled = LstV.ListItems.Count
            End If
            
    End Select
    
    ButtonPress = vbCancel
    
End Sub

Private Sub tBar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Ret As Integer

    Select Case ButtonMenu.Key
        Case "MnuD"
            Call MnuNew_Click
        Case "MnuT"
            frmTable.Show vbModal, frmmain
            'Check button press
            If (ButtonPress = vbOK) Then
                'Create table
                Ret = CreateTable

                If (Ret = 0) Then
                    'Missing table name
                    MsgBox DBErrMsg, vbInformation, frmmain.Caption
                ElseIf (Ret = 1) Then
                    MsgBox DBErrMsg, vbInformation, frmmain.Caption
                Else
                    'Refresh
                    Call DoRefresh(sTab1.Tabs.Count + 1)
                    'Enable/Disable Table info
                    tBar1.Buttons(4).Enabled = sTab1.Tabs.Count
                End If
            End If
    End Select
    
    ButtonPress = vbCancel
End Sub
