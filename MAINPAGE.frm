VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "SALARY CALCULATION"
   ClientHeight    =   9315
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   19020
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu ADD 
      Caption         =   "ADD SALARY"
   End
   Begin VB.Menu VIEW 
      Caption         =   "VIEW SALARY"
   End
   Begin VB.Menu PRINT 
      Caption         =   "PRINT"
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ADD_Click()
Form2.Show


End Sub

Private Sub exit_Click()
Dim frm As Form

Set frm = Form1
If frm.Visible = True Then
frm.Hide
Else
frm.Hide
End If
Form1.Hide
View_details.Hide
Me.Hide
End Sub

Private Sub MDIForm_Load()
Me.Height = 10000
Me.Width = 10000

End Sub

Private Sub PRINT_Click()
Form1.Show
End Sub

Private Sub VIEW_Click()
View_details.Show
End Sub
