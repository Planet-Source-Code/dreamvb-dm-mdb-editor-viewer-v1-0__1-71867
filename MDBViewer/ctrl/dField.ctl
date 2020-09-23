VERSION 5.00
Begin VB.UserControl dField 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   45
      TabIndex        =   1
      Top             =   210
      Width           =   4440
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FieldName"
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
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "dField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Event Change()

Private Sub txtValue_Change()
    RaiseEvent Change
End Sub

Private Sub txtValue_Click()
    RaiseEvent Click
End Sub

Private Sub txtValue_GotFocus()
    txtValue.BackColor = &HFEEDE9
End Sub

Private Sub txtValue_LostFocus()
    txtValue.BackColor = vbWhite
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    BackColor = vbButtonFace
    Text = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Text = PropBag.ReadProperty("Text", "")
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    Enabled = PropBag.ReadProperty("Enabled", True)
    MaxLength = PropBag.ReadProperty("MaxLength", 0)
End Sub

Private Sub UserControl_Resize()
    txtValue.Width = (UserControl.ScaleWidth - txtValue.Left)
End Sub

Public Property Get Caption() As String
    Caption = lblName.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lblName.Caption = NewCaption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Text", Text, "")
    Call PropBag.WriteProperty("BackColor", BackColor, vbButtonFace)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("MaxLength", MaxLength, 0)
End Sub

Public Property Get Text() As String
    Text = txtValue.Text
End Property

Public Property Let Text(ByVal NewText As String)
    txtValue.Text = NewText
    PropertyChanged "Text"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
    PropertyChanged "BackColor"
End Property

Public Property Get SelStart() As Integer
    SelStart = txtValue.SelStart
End Property

Public Property Let SelStart(ByVal vNewSelStart As Integer)
    txtValue.SelStart = vNewSelStart
End Property

Public Property Get SelLength() As Integer
    SelLength = txtValue.SelLength
End Property

Public Property Let SelLength(ByVal vNewSelLength As Integer)
    txtValue.SelLength = vNewSelLength
End Property

Public Property Get Enabled() As Boolean
    Enabled = txtValue.Enabled
End Property

Public Property Let Enabled(ByVal vNewEanble As Boolean)
    txtValue.Enabled = vNewEanble
    PropertyChanged "Enabled"
End Property

Public Property Get MaxLength() As Integer
    MaxLength = txtValue.MaxLength
End Property

Public Property Let MaxLength(ByVal vNewMaxLen As Integer)
    txtValue.MaxLength = vNewMaxLen
    PropertyChanged "MaxLength"
End Property
