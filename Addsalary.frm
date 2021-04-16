VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   6945
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Top             =   5160
      Width           =   3000
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Top             =   4680
      Width           =   3000
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   3000
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   6
      Top             =   3720
      Width           =   3000
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   3000
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   3000
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   3
      Top             =   1695
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   10
      Top             =   5640
      Width           =   3000
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   6120
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   0
      Text            =   "----Select Staff----"
      Top             =   720
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   720
      Width           =   420
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   " A/C No"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit to"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "I/R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PROF-TAX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   23
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Pay"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   20
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EPF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   19
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ESI"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   18
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   840
      Width           =   1200
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   2280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label6 
      Caption         =   "ADD SALARY DETAILS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Public counted As Integer

Private Sub Command2_Click()
Text1.Visible = False
Combo1.Visible = True
Command2.Visible = False
Command3.Visible = True
End Sub

Private Sub Form_Load()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"
Me.Height = 8295
Me.Width = 7185
Combo1.Visible = True
Call combofill
Command2.Visible = False
Command3.Visible = True

End Sub
Private Sub combofill()
If rs.State = 1 Then rs.Close
rs.Open "select * from Salary_Table", con, adOpenDynamic

Dim i As Integer

With MSFlexGrid1

If rs.EOF <> True Then
     Do
     Combo1.AddItem (rs(0))
        i = i + 1
        rs.MoveNext
        
    Loop Until rs.EOF = True
    Combo1.Visible = True
    Text1.Visible = False
    Command3.Visible = True
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from Salary_Table", con, adOpenDynamic
If rs.EOF = True Then
'Text1.Visible = True
'Text1.SetFocus
MsgBox "No Staff details found. Please add details"
End If

End With
End Sub

Private Sub Combo1_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"

If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim Str As String

J = 0
k = 0
i = 1
If rs.EOF <> True Then
     Do
     Dim a As String
     a = rs(0)
     If a = Combo1.Text Then
     Dim b As Integer
     If rs.State = 1 Then rs.Close
     rs.Open " select ID from Salary_Table where (Sname='" & Combo1.Text & "')"
     b = rs(0)
     End If
     rs.MoveNext
     Loop Until rs.EOF = True
     If b <> 0 Then
     If rs.State = 1 Then rs.Close
     rs.Open " select * from Salary_Table where (ID=" & Val(b) & ")"
     
     
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Text5.Text = rs(4)
Text6.Text = rs(5)
Text7.Text = rs(6)
Text8.Text = rs(7)
Text9.Text = rs(8)
Text10.Text = rs(9)
Text11.Text = rs(10)
End If
End If
End Sub

Private Sub Command1_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"


Do While Combo1.Text <> "----Select Staff----"
If Combo1.Text = "----Select Staff----" Then
MsgBox "Please select Staff name"
Exit Do
Else
Exit Do
End If
Loop

If Combo1.Visible = True Then
If Combo1.Text <> "----Select Staff----" Then
Call saved
End If
End If

Do While Text1.Visible = True
If Text1.Text = "" Then
MsgBox "Please enter Staff name"
Exit Do
Else
Exit Do
End If
Loop

If Text1.Visible = True Then
If Text1.Text <> "" Then
Call saved
End If
End If
MsgBox "Record Saved/Updated Successfully", vbInformation, "Information"
End Sub

Private Sub saved()
If Text4.Text = "" Then
Text4.Text = 0
End If

If Text5.Text = "" Then
Text5.Text = 0
End If

If Text6.Text = "" Then
Text6.Text = 0
End If

If Text7.Text = "" Then
Text7.Text = 0
'Label24.Caption = 0
End If

If Text8.Text = "" Then
Text8.Text = 0
End If

If Text9.Text = "" Then
Text9.Text = 0
End If

If Text10.Text = "" Then
Text10.Text = 0
End If

If Text11.Text = "" Then
Text11.Text = "Nil"
End If

If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Salary_Table", con, adOpenDynamic
z = 0
If rs.EOF <> True Then
     Do
     If rs(11) > z Then
     z = rs(11)
     End If
     rs.MoveNext
     Loop Until rs.EOF = True
End If
Dim count As Integer
count = 1
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Salary_Table", con, adOpenDynamic
If rs.BOF = True Then
count = 0
z = 1
End If

If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim Str As String

J = 0
k = 0
i = 1
If rs.EOF <> True Then
     Do
     Dim a As String
     a = rs(0)
     If a = Combo1.Text Then
     J = 1
     con.Execute ("update Salary_Table set Salary_Table.BasicPay=" & Text3.Text & ",DA= " & Text5.Text & ",allowances=" & Text4.Text & ",EPF=" & Text6.Text & ",ESI=" & Text7.Text & ",proftax=" & Text8.Text & ",ir=" & Text9.Text & ",total=" & Text10.Text & ",acno='" & Text11.Text & "' where (Sname='" & Combo1.Text & "')")
     
     End If
     rs.MoveNext
     Loop Until rs.EOF = True
    
     z = z + 1
     If J <> 1 Then
     If Text1.Visible = True Then
     
     con.Execute ("Insert Into Salary_Table values('" & Text1.Text & "','" & Text2.Text & "'," & Text3.Text & "," & Text4.Text & "," & Text5.Text & "," & Text6.Text & "," & Text7.Text & "," & Text8.Text & "," & Text9.Text & "," & Text10.Text & ",'" & Text11.Text & "'," & Val(z) & ")")
     End If
     End If
   
    
'Else
     'Set rs = con.Execute("Insert Into Salary_Table (Name,Basic Pay,DA,allowances,EPF,ESI,prof-tax,ir,total,acno) values(" & 1 & "," & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & Text4.Text & "," & Text5.Text & "," & Text6.Text & "," & Text7.Text & "," & Text8.Text & "," & Text9.Text & "," & Text10.Text & ",'" & Text11.Text & "')")
     'Set rs = con.Execute("Insert Into Salary_Table (ID,Name) values (" & Text1.Text & ",'" & Text2.Text & "')")
End If

If count = 0 Then
con.Execute ("Insert Into Salary_Table values('" & Text1.Text & "','" & Text2.Text & "'," & Text3.Text & "," & Text4.Text & "," & Text5.Text & "," & Text6.Text & "," & Text7.Text & "," & Text8.Text & "," & Text9.Text & "," & Text10.Text & ",'" & Text11.Text & "'," & Val(z) & ")")
End If

Command1.Visible = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""

Combo1.Visible = True
'Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
Text11.Visible = True


Combo1.Visible = True
Combo1.Clear
Combo1.AddItem ("----Select Staff----")
Call combofill
Combo1.SetFocus
Combo1.SelText = "----Select Staff----"
Command2.Visible = False
Command3.Visible = True
End Sub

Private Sub Command3_Click()
Text1.Visible = True
Combo1.Visible = False
Command2.Visible = True
Command3.Visible = False
Label19.Visible = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End Sub

Private Sub Text10_Click()
Dim i As Double
Dim J As Double

i = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
J = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)

Text10.Text = i - J
Text11.SetFocus
End Sub


