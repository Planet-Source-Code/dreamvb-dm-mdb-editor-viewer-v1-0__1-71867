VERSION 5.00
Begin VB.UserControl Line3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   ScaleHeight     =   90
   ScaleWidth      =   1905
   ToolboxBitmap   =   "Line3d.ctx":0000
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const DM_LINE_SHADOW = 16
Private Const COLOR_BTNHIGHLIGHT As Long = 20

Private Sub DrawLine()
    UserControl.Line (0, 15)-(ScaleWidth, 15), GetSysColor(COLOR_BTNHIGHLIGHT)
    UserControl.Line (0, 0)-(ScaleWidth, 0), GetSysColor(DM_LINE_SHADOW)
End Sub

Private Sub UserControl_Initialize()
    Call DrawLine
End Sub

Private Sub UserControl_Resize()
 On Error Resume Next
    UserControl.Height = 30
    DrawLine
End Sub

