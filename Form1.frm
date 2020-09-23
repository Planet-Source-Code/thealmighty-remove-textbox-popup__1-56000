VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu EditSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelectAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnufake 
      Caption         =   "Fake"
      Visible         =   0   'False
      Begin VB.Menu mnufakeitem1 
         Caption         =   "Item1"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'windows API
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'function to check if a desired key is pressed
Private Function KeyIsDown(vKeyCode) As Boolean
 KeyIsDown = (GetAsyncKeyState(vKeyCode) < 0)
End Function

'Edit Menu Functions using Clipboard Properties
'''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditCut_Click()
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
Screen.ActiveControl.SelText = ""
End Sub
Private Sub EditCopy_Click()
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub EditDelete_Click()
Screen.ActiveControl.SelText = ""
End Sub

Private Sub EditSelectAll_Click()
Screen.ActiveControl.SelStart = 0
Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Text1_KeyPress(KeyAscii As Integer)
'disable Ctrl+V (paste function) keyascii=22 for Ctrl+V(combined)
If KeyAscii = 22 Then KeyAscii = 0
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' check if right button was pressed ,if so popup edit menu
If Button = vbRightButton Then
Text1.Enabled = False 'first disable the textbox so that handle to default popup menu
Text1.Enabled = True  'is not available, then immediately enable it. set focus on it
Text1.SetFocus
PopupMenu mnuEdit     'popup menu of ur choice
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'check if shift+insert(45) was pressed
If KeyIsDown(vbKeyShift) And KeyCode = 45 Then KeyCode = 0
'check if user pressed applications key/context menu key (immediate right to
'RightWindows keys)
'this is important coz even if u disable user's right click on the textbox
'so that default textbox menu doesn't pop up, this key seems to override those things
If KeyCode = 93 Then PopupMenu mnuEdit
End Sub
