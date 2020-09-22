Attribute VB_Name = "Skins"
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Public Sub DragForm(Who As Form)
Call ReleaseCapture
Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub


'dragimgl.picture=skinz.dragimgl.picture
'dragimgr.picture=skinz.dragimgr.picture
'borderdl.picture=skinz.borderdl.picture
'borderdr.picture=skinz.borderdr.picture
'.ForeColor = &H8000000E
'.ForeColor = &H80000012
