VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox replace 
      Height          =   1425
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox find 
      Height          =   1425
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "1234567890XX"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldtext As String
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Form_Load()
    Dim txt, txt2 As Variant
    Form1.Left = -10000
    Open App.Path + "\timeradjust.txt" For Input As #1
        Input #1, txt
        Timer1.Interval = txt
    Close
    Open App.Path + "\replace.txt" For Input As #1
        Do
        Input #1, txt, txt2
        find.AddItem (txt)
        replace.AddItem (Trim(txt2))
        Loop Until EOF(1)
    Close #1
    
    set_autorun
    
End Sub
Function ShiftKey() As Boolean
    ShiftKey = (GetAsyncKeyState(vbKeyShift) And &H8000)
End Function

Private Sub Timer1_Timer()

Dim newstr As String
For i = 1 To 256
result = 0
result = GetAsyncKeyState(i)

If result = -32767 Then
If ShiftKey() = True Then
     Text1.Text = Right$(Text1.Text, 20) + UCase(Chr(i))
Else
     Text1.Text = Right$(Text1.Text, 20) + LCase(Chr(i))
End If
End If
Next i
If Text1 <> oldtext Then
    'check replacements
    For i = 0 To find.ListCount - 1
        If Right$(Text1, Len(find.List(i))) = find.List(i) Then
            newstr = "{BS " + Str(Len(find.List(i))) + "}" + replace.List(i)
 
            SendKeys newstr, True
            Text1.Text = Space(20)
        End If
    Next i
    If LCase(Right$(Text1, 7)) = "rplquit" Then
        MsgBox "Replacer was closed by user", 64, "Shutting down"
        End
    End If
End If
oldtext = Text1

End Sub
